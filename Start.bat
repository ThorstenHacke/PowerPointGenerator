For /F "Tokens=*" %%I in ('dotnet --version') Do set theVersion=%%I
IF %theVersion% GTR 2.1 (
    dotnet build PowerPointGenerator.sln
    dotnet DemoConsole\bin\Debug\netcoreapp2.1\PowerPointGeneratorExecutor.dll
) ELSE (
    echo Bitte dotnetcore >= 2.1 installieren
)
pause