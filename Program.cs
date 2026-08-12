using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.Json;

Console.OutputEncoding = Encoding.UTF8;

var config = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .Build();

var connectionString = config.GetConnectionString("Default")
    ?? throw new InvalidOperationException("ConnectionStrings:Default bulunamadı. appsettings.json kontrol et.");

var binaryAssembliesFolder = config["BinaryAssembliesFolder"]
    ?? Path.Combine(AppContext.BaseDirectory, "libs");
var binaryMode = (config["BinaryMode"] ?? "auto").Trim().ToLowerInvariant();

BinaryDeserializer.Initialize(binaryAssembliesFolder);

var sqlFile = args.Length > 0 ? args[0] : "query.sql";

if (!File.Exists(sqlFile))
{
    var besideExe = Path.Combine(AppContext.BaseDirectory, sqlFile);
    if (File.Exists(besideExe))
        sqlFile = besideExe;
    else
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"SQL dosyası bulunamadı: {sqlFile}");
        Console.ResetColor();
        Console.WriteLine();
        Console.WriteLine("Kullanım:");
        Console.WriteLine("  dotnet run");
        Console.WriteLine("  dotnet run -- query.sql");
        return 1;
    }
}

var sql = await File.ReadAllTextAsync(sqlFile);

Console.WriteLine("Bağlantı kuruluyor...");
Console.WriteLine($"SQL dosyası: {Path.GetFullPath(sqlFile)}");
Console.WriteLine($"BinaryMode: {binaryMode}");
Console.WriteLine($"Binary DLL klasörü: {binaryAssembliesFolder}");
Console.WriteLine($"DLL bulundu: {(BinaryDeserializer.HasAssemblies ? "evet" : "hayır")}");
Console.WriteLine(new string('-', 60));

try
{
    await using var connection = new SqlConnection(connectionString);
    await connection.OpenAsync();

    await using var command = new SqlCommand(sql, connection)
    {
        CommandTimeout = config.GetValue("CommandTimeoutSeconds", 120)
    };

    await using var reader = await command.ExecuteReaderAsync();
    var table = new DataTable();
    table.Load(reader);

    if (table.Rows.Count == 0)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("Sonuç yok (0 satır).");
        Console.ResetColor();
        return 0;
    }

    var outputDir = config["ExcelOutputFolder"];
    if (string.IsNullOrWhiteSpace(outputDir))
        outputDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "fortuna script");

    Directory.CreateDirectory(outputDir);

    var excelPath = Path.Combine(outputDir, $"result_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
    ExportExcel(table, excelPath, binaryMode);

    Console.ForegroundColor = ConsoleColor.Green;
    Console.WriteLine($"Toplam satır: {table.Rows.Count}");
    Console.WriteLine($"Excel kaydedildi: {excelPath}");
    Console.ResetColor();

    if (config.GetValue("OpenExcel", true))
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = excelPath,
            UseShellExecute = true
        });
    }

    return 0;
}
catch (Exception ex)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("Hata:");
    Console.WriteLine(ex.Message);
    if (ex.InnerException is not null)
        Console.WriteLine(ex.InnerException.Message);
    Console.ResetColor();
    return 1;
}

static void ExportExcel(DataTable table, string path, string binaryMode)
{
    using var workbook = new XLWorkbook();
    var sheet = workbook.Worksheets.Add("Sonuc");

    for (var c = 0; c < table.Columns.Count; c++)
        sheet.Cell(1, c + 1).Value = table.Columns[c].ColumnName;

    var header = sheet.Range(1, 1, 1, table.Columns.Count);
    header.Style.Font.Bold = true;
    header.Style.Fill.BackgroundColor = XLColor.FromHtml("#1F4E79");
    header.Style.Font.FontColor = XLColor.White;
    header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

    for (var r = 0; r < table.Rows.Count; r++)
    {
        for (var c = 0; c < table.Columns.Count; c++)
        {
            var cell = sheet.Cell(r + 2, c + 1);
            var value = table.Rows[r][c];

            if (value is null or DBNull)
            {
                cell.Value = Blank.Value;
                continue;
            }

            switch (value)
            {
                case DateTime dt:
                    cell.Value = dt;
                    cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss";
                    break;
                case bool b:
                    cell.Value = b;
                    break;
                case byte[] bytes:
                    cell.Value = FormatBinary(bytes, binaryMode);
                    cell.Style.Alignment.WrapText = true;
                    break;
                case short or int or long or byte or sbyte or ushort or uint or ulong:
                    cell.Value = Convert.ToDouble(value);
                    break;
                case float or double or decimal:
                    cell.Value = Convert.ToDouble(value);
                    break;
                case Guid g:
                    cell.Value = g.ToString();
                    break;
                default:
                    cell.Value = value.ToString();
                    break;
            }
        }
    }

    var used = sheet.RangeUsed();
    if (used is not null)
    {
        used.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        used.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        sheet.SheetView.FreezeRows(1);
        sheet.Columns().AdjustToContents(1, 80);
    }

    workbook.SaveAs(path);
}

static string FormatBinary(byte[] bytes, string binaryMode)
{
    var jsonOpts = new JsonSerializerOptions
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    if (bytes.Length == 0)
        return "null";

    // Düz JSON/text ise hızlı dön
    if (TryDecodeCleanUtf8(bytes, out var text))
    {
        if ((text.StartsWith('{') || text.StartsWith('[')) && IsValidJson(text))
            return text;
        return text;
    }

    var wantDeserialize = binaryMode is "auto" or "deserialize";
    if (wantDeserialize && BinaryDeserializer.HasAssemblies)
    {
        if (BinaryDeserializer.TryDeserialize(bytes, out var obj, out var error))
        {
            var graph = ObjectToJsonGraph.ToGraph(obj);
            return JsonSerializer.Serialize(new
            {
                format = "binary-formatter",
                type = obj?.GetType().FullName,
                data = graph
            }, jsonOpts);
        }

        return JsonSerializer.Serialize(new
        {
            format = "binary-formatter",
            error,
            length = bytes.Length,
            hint = "libs klasöründeki DLL'ler eksik/uyumsuz olabilir (Entities + InterFrame.Messaging + bağımlılıkları)."
        }, jsonOpts);
    }

    if (binaryMode == "base64")
        return Convert.ToBase64String(bytes);

    // Hızlı varsayılan: decode deneme, parse yok
    return JsonSerializer.Serialize(new
    {
        format = "binary-formatter",
        length = bytes.Length,
        hint = "Düzgün okumak için Callcenter messaging DLL'lerini libs klasörüne koy. BinaryMode=deserialize"
    }, jsonOpts);
}

static bool TryDecodeCleanUtf8(byte[] bytes, out string text)
{
    text = "";
    try
    {
        var decoder = Encoding.UTF8.GetDecoder();
        decoder.Fallback = DecoderFallback.ExceptionFallback;
        var chars = new char[Encoding.UTF8.GetMaxCharCount(bytes.Length)];
        var count = decoder.GetChars(bytes, 0, bytes.Length, chars, 0, flush: true);
        text = new string(chars, 0, count).Trim('\0').Trim();
        if (text.Length == 0)
            return false;

        var control = text.Count(ch => char.IsControl(ch) && ch is not ('\r' or '\n' or '\t'));
        return control <= text.Length * 0.05;
    }
    catch
    {
        return false;
    }
}

static bool IsValidJson(string text)
{
    try
    {
        using var _ = JsonDocument.Parse(text);
        return true;
    }
    catch
    {
        return false;
    }
}

static class BinaryDeserializer
{
    static string _folder = "";
    static bool _initialized;

    public static bool HasAssemblies { get; private set; }

    public static void Initialize(string folder)
    {
        _folder = folder;
        Directory.CreateDirectory(folder);

        HasAssemblies = Directory.EnumerateFiles(folder, "*.dll").Any();
        if (_initialized)
            return;

        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name = new AssemblyName(args.Name).Name;
            if (string.IsNullOrWhiteSpace(name))
                return null;

            var path = Path.Combine(_folder, name + ".dll");
            return File.Exists(path) ? Assembly.LoadFrom(path) : null;
        };

        // libs'teki DLL'leri önceden yükle (tip çözümlemesi için)
        foreach (var dll in Directory.EnumerateFiles(folder, "*.dll"))
        {
            try { Assembly.LoadFrom(dll); }
            catch { /* bağımlılık eksikse sonra resolve dener */ }
        }

        _initialized = true;
    }

#pragma warning disable SYSLIB0011
    public static bool TryDeserialize(byte[] bytes, out object? obj, out string? error)
    {
        obj = null;
        error = null;
        try
        {
            using var ms = new MemoryStream(bytes);
            var formatter = new BinaryFormatter();
            obj = formatter.Deserialize(ms);
            return obj is not null;
        }
        catch (Exception ex)
        {
            error = ex.GetBaseException().Message;
            return false;
        }
    }
#pragma warning restore SYSLIB0011
}

static class ObjectToJsonGraph
{
    public static object? ToGraph(object? value, int depth = 0, HashSet<object>? seen = null)
    {
        if (value is null)
            return null;
        if (depth > 8)
            return value.ToString();

        seen ??= new HashSet<object>(ReferenceEqualityComparer.Instance);
        var type = value.GetType();

        if (value is string or decimal or DateTime or DateTimeOffset or Guid or bool
            || type.IsEnum)
            return value;

        if (value is byte[] raw)
            return System.Convert.ToBase64String(raw);

        if (type.IsPrimitive)
            return value;

        if (!type.IsValueType && !seen.Add(value))
            return "[circular]";

        if (value is IDictionary dict)
        {
            var map = new Dictionary<string, object?>();
            foreach (DictionaryEntry entry in dict)
                map[entry.Key?.ToString() ?? ""] = ToGraph(entry.Value, depth + 1, seen);
            return map;
        }

        if (value is IEnumerable enumerable and not string)
        {
            var list = new List<object?>();
            foreach (var item in enumerable)
                list.Add(ToGraph(item, depth + 1, seen));
            return list;
        }

        var result = new Dictionary<string, object?>();
        foreach (var prop in type.GetProperties(BindingFlags.Instance | BindingFlags.Public))
        {
            if (!prop.CanRead || prop.GetIndexParameters().Length > 0)
                continue;
            try
            {
                result[prop.Name] = ToGraph(prop.GetValue(value), depth + 1, seen);
            }
            catch
            {
                result[prop.Name] = "[error]";
            }
        }

        foreach (var field in type.GetFields(BindingFlags.Instance | BindingFlags.Public))
        {
            try
            {
                result[field.Name] = ToGraph(field.GetValue(value), depth + 1, seen);
            }
            catch
            {
                result[field.Name] = "[error]";
            }
        }

        foreach (var field in type.GetFields(BindingFlags.Instance | BindingFlags.NonPublic))
        {
            var name = field.Name;
            if (name.Contains("k__BackingField", StringComparison.Ordinal))
            {
                var propName = name.Replace("<", "", StringComparison.Ordinal)
                    .Replace(">k__BackingField", "", StringComparison.Ordinal);
                if (result.ContainsKey(propName))
                    continue;
                try
                {
                    result[propName] = ToGraph(field.GetValue(value), depth + 1, seen);
                }
                catch
                {
                    result[propName] = "[error]";
                }
            }
        }

        result["_type"] = type.FullName;
        return result;
    }
}
