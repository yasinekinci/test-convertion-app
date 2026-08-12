Bu klasöre Callcenter BinaryFormatter deserialize için gereken DLL'leri koy:

Zorunlu (tipik):
- Intertech.Callcenter.Host.Entities.dll
- Intertech.InterFrame.Messaging.dll

Genelde yanlarında gerekenler de olur (ör. Intertech.InterFrame.*.dll ve diğer bağımlılıklar).
Callcenter Host / messaging uygulamasının bin klasöründen kopyala.

Sonra appsettings.json:
  "BinaryMode": "deserialize"

Not: Bu DLL'ler çoğunlukla .NET Framework 4.x. net8 üzerinde çalışmazsa
ayrı net48 console gerekir; o durumda söyle.
