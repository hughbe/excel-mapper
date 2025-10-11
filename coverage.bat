cd tests
if exist TestResults rmdir TestResults /s /q
if exist coverage rmdir coverage /s /q
dotnet test --collect:"XPlat Code Coverage" --settings coverage.runsettings
if %errorlevel% neq 0 exit /b %errorlevel%
dotnet reportgenerator -reports:".\**\TestResults\**\coverage.cobertura.xml" -targetdir:"coverage" -reporttypes:Html