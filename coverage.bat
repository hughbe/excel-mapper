cd tests
dotnet test -f netcoreapp3.1 /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
dotnet reportgenerator -reports:coverage.opencover.xml -targetdir:../resources/coverage
