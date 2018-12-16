#!/bin/bash

set -e

cd tests
dotnet test -f netcoreapp2.1 /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
dotnet reportgenerator -reports:coverage.opencover.xml -targetdir:../resources/coverage
