@echo off

dotnet test tests /p:CollectCoverage=true /p:CoverletOutputFormat=opencover

set report_generator_version=3.1.2
set user_directory=C:\Users\hughb
set packages_directory=%user_directory%\.nuget\packages
set report_generator=%packages_directory%\reportgenerator\%report_generator_version%\tools\ReportGenerator.exe

set report_path=resources\coverage
%report_generator% -reports:"tests\coverage.opencover.xml" -targetdir:"%report_path%"
