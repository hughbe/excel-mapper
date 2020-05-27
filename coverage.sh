#!/bin/bash

set -e

cd tests
dotnet test /p:CollectCoverage=true /p:CoverletOutputFormat=opencover
dotnet reportgenerator -reports:coverage.opencover.xml -targetdir:../resources/coverage
