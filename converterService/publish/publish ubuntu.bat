dotnet publish ..\SquidService\SquidGUI\ -o ..\Release\itTLS_Final\itTLSUbuntu\ -r linux-x64 --configuration Release -p:TrimMode=CopyUsed -p:PublishTrimmed=True -p:PublishSingleFile=true --self-contained true -p:EnableCompressionInSingleFile=true -p:Version=2.0.0

dotnet publish ..\SquidService\itTLSWebServer\ -o ..\Release\itTLS_Final\itTLSUbuntu\ -r linux-x64 --configuration Release -p:TrimMode=CopyUsed -p:PublishTrimmed=True -p:PublishSingleFile=true --self-contained true -p:EnableCompressionInSingleFile=true -p:Version=2.0.0
