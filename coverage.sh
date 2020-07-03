#!/bin/bash

set -e

cd tests
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
dotnet reportgenerator -reports:coverage.netcoreapp3.1.opencover.xml -targetdir:../coverage
