@echo off
if [%1]==[] ( set outputPath="..\..\Packages" 
) else ( 
set outputPath=%1 )
echo %outputPath%

if not exist %outputPath% mkdir %outputPath%

rd %outputPath%

for /f %%G in ('dir /b /o:n /ad') do  (
if "%%G" NEQ ".git" (
cd %%G 
nuget restore
msbuild /t:Build /p:Configuration=Release p:TargetFrameworkVersion=4.6.1 /v:q /nologo
nuget pack -OutputDirectory %outputPath% -BasePath output\ 
cd .. )
)
cd %outputPath%
for /f %%f in ('dir /b %outputPath%') do echo %%f