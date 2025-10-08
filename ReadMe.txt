TA-Tools for AutoCAD (TraceAir)

Version: 1.0.0

Overview:
TraceAir CAD team tools for request processing using AutoCAD. Installs as an AutoCAD bundle and adds a ribbon panel with commands.

Supported Platforms:
- Windows 32-bit and 64-bit
- AutoCAD (and most verticals) R18.2+ (see Autodesk documentation)

Install:
1) Close AutoCAD.
2) Copy the entire folder "TaTools.bundle" to:
   %PROGRAMFILES%\Autodesk\ApplicationPlugins
3) Start AutoCAD. The TA-Tools ribbon panel should load automatically.

Update:
- Replace the existing "TaTools.bundle" in the same folder with the new version, then restart AutoCAD.

Uninstall:
- Delete the folder:
  %PROGRAMFILES%\Autodesk\ApplicationPlugins\TaTools.bundle

Loaded Components (auto-load on startup):
- Custom UI: Contents\Runtime\tatools.cuix
- LISP: Contents\Runtime\TA-cadtools-geometry.lsp

Available Commands:
- TA-SET-ELEV-FOR-PADS
- TA-POINTS-FROM-MLEADERS
- TA-CONVER-BROKEN-LEADER
- TA-SCALE-LIST-RESET
- TA-MULTY-POLYLINE-OFFSET
- TA-ADD-PREFIX-SUFFIX-TO-TEXT
- TA-POINTS-AT-POLY-ANGLES
- TA-3dPOLY-BY-POINTS-BLOCKS
- TA-EXP-SLOPE

If Ribbon Does Not Appear:
1) Verify the bundle is installed exactly at:
   %PROGRAMFILES%\Autodesk\ApplicationPlugins\TaTools.bundle
2) In AutoCAD, run CUILOAD and load:
   <install path>\TaTools.bundle\Contents\Runtime\tatools.cuix
3) Restart AutoCAD.

More Info from Autodesk on application bundles:
https://help.autodesk.com/view/ACDLT/2024/RUS/?guid=GUID-5E50A846-C80B-4FFD-8DD3-C20B22098008

Author: Nikita Prokhor
Company: TraceAir — https://traceair.net
