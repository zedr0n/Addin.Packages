@echo off
set Configuration=%1
if "%Configuration%"=="" set Configuration=Debug 

REM for /F %%x in ('dir /b *.sln') do msbuild %%x /t:Build /p:Configuration=%Configuration%
REM SourceLink.exe index -nvg -pr CommonAddin/CommonAddin.csproj -pp Configuration Release -u "https://zperforce.cloudapp.net/swarm/view/depot/excel/addins/TestAddin/%%var2%%"
REM SourceLink.exe index -nvg -pr ExcelInterfaces/ExcelInterfaces.csproj -pp Configuration %Configuration% -u "http://zperforce.cloudapp.net/view/depot/excel/addins/TestAddin/%%var2%%"
popd
REM nuget pack -IncludeReferencedProjects -Prop Configuration=%Configuration% ExcelInterfaces.csproj
dotnet pack -o c:\NuGet -c %Configuration%
