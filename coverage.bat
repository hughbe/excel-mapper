@echo off

set project_name=ExcelMapper
set user_directory=C:\Users\hugh

set open_cover_version=4.6.519
set report_generator_version=2.5.7
set xunit_console_version=2.3.0-beta3-build3705

set packages_directory=%user_directory%\.nuget\packages
set open_cover=%packages_directory%\opencover\%open_cover_version%\tools\OpenCover.Console.exe
set report_generator=%packages_directory%\reportgenerator\%report_generator_version%\tools\ReportGenerator.exe

set xunit=%packages_directory%\xunit.runner.console\%xunit_console_version%\tools\net452\xunit.console.exe

set test_name=src\%project_name%.Tests\bin\Debug\net46\%project_name%.Tests.exe

set report_name=coverage.xml
set report_path=resources\coverage

dotnet build src

%open_cover% -register:user -output:"%report_name%" -filter:"+[*]* -[%project_name%.Tests]*" -target:"%xunit%" -targetargs:"%test_name% -noshadow"

%report_generator% -reports:"%report_name%" -targetdir:"%report_path%"
