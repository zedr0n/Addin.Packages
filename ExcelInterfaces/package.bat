@echo off
del *.nupkg
pushd ..
msbuild TestAddin.sln /t:Build /p:Configuration=Release
REM SourceLink.exe index -nvg -pr CommonAddin/CommonAddin.csproj -pp Configuration Release -u "https://zperforce.cloudapp.net/swarm/view/depot/excel/addins/TestAddin/%%var2%%"
SourceLink.exe index -nvg -pr ExcelInterfaces/ExcelInterfaces.csproj -pp Configuration Release -u "http://zperforce.cloudapp.net/view/depot/excel/addins/TestAddin/%%var2%%"
popd
nuget pack -IncludeReferencedProjects -Prop Configuration=Release
copy *.nupkg c:\NuGet 