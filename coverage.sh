#!/bin/bash

cd tests
if [ -d "TestResults" ]; then
    rm -rf TestResults
fi
if [ -d "coverage" ]; then
    rm -rf coverage
fi
dotnet test --settings coverage.runsettings
if [ $? -ne 0 ]; then
    exit $?
fi
dotnet reportgenerator -reports:"./**/TestResults/**/coverage.cobertura.xml" -targetdir:"coverage" -reporttypes:Html
