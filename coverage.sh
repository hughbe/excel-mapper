#!/bin/bash

set -e

cd tests
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
dotnet reportgenerator -reports:coverage.net5.opencover.xml -targetdir:../coverage
