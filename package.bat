@echo off

pushd ExcelInterfaces
call package.bat
popd

pushd CommonAddin
call package.bat
popd  