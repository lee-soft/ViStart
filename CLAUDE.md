# ViStart

VB6 start-menu replacement. Never build or run via the VB6 IDE — it is unreliable for
apps that subclass their own windows.

## Build

```powershell
.\build.ps1                    # -> C:\Users\Lee\AppData\Roaming\ViStart\ViStart.exe
.\build.ps1 -OutDir <path>     # build elsewhere
.\build.ps1 -KeepRunning       # don't kill a running ViStart.exe first
```

The script registers the required typelibs (winU, GDIplus, SHELLLNK from C:\TLBS) if
needed, kills a running ViStart.exe, compiles with VB6.EXE /make, then embeds
ViStart.exe.manifest into the exe (RT_MANIFEST, via the Windows SDK's mt.exe). Each
build auto-increments the revision and rewrites prjVistaStart.vbp — that diff is
expected.
