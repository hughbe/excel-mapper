cd tests
rmdir TestResults /s /q
rmdir coverage /s /q
dotnet test --settings coverage.runsettings
if %errorlevel% neq 0 exit /b %errorlevel%
dotnet reportgenerator -reports:".\**\TestResults\**\coverage.cobertura.xml" -targetdir:"coverage" -reporttypes:Html