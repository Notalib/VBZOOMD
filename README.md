# VBZOOMD
VB.NET Wrapper for YAZ ZOOM Z39.50 Libraries

This is a VB.NET version of an old VB6 ActiveX library.  

The original VB6 code and documentation can be found here: http://vb-zoom.sourceforge.net/

## Create nuget package
``` powershell
cd \VBZOOMD\VBZOOMD
.\nuget.exe pack .\VBZOOMD.vbproj -properties Configuration=Release
```
## Publish nuget package
``` powershell
.\nuget.exe push .\VBZOOMD.x.x.x.nupkg -Source https://nuget.pkg.github.com/Notalib/index.json -ApiKey [GITHUB-PAT-TOKEN]
```
