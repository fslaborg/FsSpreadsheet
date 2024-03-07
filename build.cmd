@echo off

set PYTHONIOENCODING=utf-8
dotnet tool restore
cls 
dotnet run --project ./build/build.fsproj %*