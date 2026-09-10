' ============================================================================
'  Readme.vb
'
'  Documentation, not code. There is deliberately nothing in this file but
'  comments, so it compiles to nothing and can be read without building
'  anything. Do not delete it as dead code.
'
'  It says what has to be on a machine before this project will build and look
'  right, and how to put each of those things there. Every section is written
'  twice: once for a person with Visual Studio open, and once as instructions an
'  AI coding agent can carry out on its own. Pointing an agent at this file and
'  asking it to install the prerequisites is what it is for.
'
'  It lives inside the project rather than beside it so that it travels with the
'  source and shows up in Solution Explorer, where somebody will actually see
'  it.
'
'  Contents
'
'      1. What you need                      - the short list
'      2. NuGet package: System.IO.Ports     - the serial port
'      3. Font: Segoe Fluent Icons           - every glyph in the UI
'      4. Nothing else                       - and please keep it that way
' ============================================================================


' ============================================================================
'  1. WHAT YOU NEED
' ============================================================================
'
'  - Windows. This is a WPF application and it calls into SetupAPI to find the
'    probe on the USB bus, so it is Windows only by design, not by accident.
'  - The .NET 8 SDK. The project targets net8.0-windows with Option Strict On.
'  - One NuGet package, System.IO.Ports 8.0.0. Section 2.
'  - One font, Segoe Fluent Icons. Section 3. Windows 11 already has it.
'
'  Visual Studio 2022 or later opens BusMaster.sln directly. There is no build
'  script, no code generation step and nothing to run before the first build.


' ============================================================================
'  2. NUGET PACKAGE: System.IO.Ports 8.0.0
'
'  Why it is needed
'
'  The probe is a USB CDC serial device, and Modules\ProbeControl.vb drives it
'  with a real System.IO.Ports.SerialPort. That type used to be part of the
'  framework; since .NET Core it is not, and it is not in the .NET 8 shared
'  framework either, so it has to be referenced as a package. Without it the
'  build fails in ProbeControl.vb on SerialPort, Handshake and Parity.
'
'  It is already written into BusMaster.vbproj:
'
'      <ItemGroup>
'        <PackageReference Include="System.IO.Ports" Version="8.0.0" />
'      </ItemGroup>
'
'  so on a machine that has ever restored this project there is nothing to do.
'  A fresh clone or a fresh extract of a source archive needs the package
'  downloaded once, which any restore or build does by itself.
' ============================================================================
'
'  FOR A PERSON
'
'  Easiest: open BusMaster.sln and build. Visual Studio restores packages
'  before the first build and fetches it without being asked.
'
'  If that fails, or you want to be sure:
'
'      1. Right-click the BusMaster project in Solution Explorer.
'      2. Manage NuGet Packages...
'      3. If it is listed under Installed with a version of 8.0.0, you are
'         done. Otherwise go to Browse, search for System.IO.Ports, pick
'         version 8.0.0 from the version list, and Install.
'
'  Or, from Tools > NuGet Package Manager > Package Manager Console:
'
'      Install-Package System.IO.Ports -Version 8.0.0
'
'  A restore needs to reach nuget.org. On a machine with no internet access,
'  copy %USERPROFILE%\.nuget\packages\system.io.ports from a machine that has
'  restored it, or point a local package source at the .nupkg.
'
'  FOR AN AI AGENT
'
'  Work in the folder holding BusMaster.vbproj. Check before changing anything:
'
'      Select-String -Path BusMaster.vbproj -Pattern 'System.IO.Ports'
'
'  If a PackageReference is already there, the project file is correct and only
'  the package itself may be missing from the machine. Fetch it with:
'
'      dotnet restore BusMaster.vbproj
'
'  If there is no PackageReference at all, add it pinned to the same version -
'  do not let the tool choose the latest:
'
'      dotnet add package System.IO.Ports --version 8.0.0
'
'  Then confirm the whole thing actually builds. Read Modules\AppSettings.vb
'  for where settings live, and note that the user usually has BusMaster
'  running from Visual Studio, which locks bin\Debug - build to a scratch
'  directory instead of terminating their instance:
'
'      dotnet build BusMaster.vbproj -c Debug -o <some scratch folder>
'
'  A clean build is not proof the program starts: a bad StaticResource key
'  throws XamlParseException at window load. Run it, and if it dies at start-up
'  read the Windows Application event log for the .NET Runtime entry, which
'  gives the exception and the XAML line number.


' ============================================================================
'  3. FONT: SEGOE FLUENT ICONS
'
'  Why it is needed
'
'  Every icon in this program is a character from an icon font - there are no
'  image files anywhere in the project. The toolbar glyphs, the padlock and
'  visibility switches, the disclosure arrows and the device panel buttons are
'  all codepoints in the Private Use Area, written in the XAML as entities:
'
'      Content="&#xE74E;"
'
'  The font is named in exactly one place, so it can be changed in exactly one
'  place. In Themes\VisualStudioDark.xaml:
'
'      <FontFamily x:Key="Font_Glyph">Segoe Fluent Icons</FontFamily>
'      <FontWeight x:Key="Weight_Glyph">Bold</FontWeight>
'
'  Every style that draws a glyph uses those two resources. Never hardcode a
'  font name on a control. The font has no bold face of its own, so Bold is
'  synthesised - that thickening is deliberate and is what was wanted.
'
'  Where it comes from
'
'  Segoe Fluent Icons ships with Windows 11 as C:\Windows\Fonts\SegoeIcons.ttf.
'  It does NOT ship with Windows 10. On Windows 10 the glyphs come out as
'  hollow boxes, or as whatever another font happens to have at those
'  codepoints, and the program looks broken while working perfectly.
'
'  Windows 10 also has the older Segoe MDL2 Assets (segmdl2.ttf), and every one
'  of the 24 codepoints this project currently uses exists in BOTH fonts -
'  checked by rendering each one in each font on 2026-09-10. The shapes differ
'  a little; nothing goes missing. So the fallback below is safe today, and
'  worth re-checking if new glyphs are added.
' ============================================================================
'
'  FOR A PERSON
'
'  On Windows 11 there is nothing to do.
'
'  On Windows 10, either install the font or use the fallback.
'
'  To install it:
'
'      1. Download https://aka.ms/SegoeFluentIcons - it redirects to
'         Segoe-Fluent-Icons.zip on download.microsoft.com.
'      2. Unzip it.
'      3. Right-click the .ttf and choose "Install for all users". Installing
'         for one user works too, but then the font is missing for anything
'         running as another account.
'      4. Restart Visual Studio, or at least the program, so the new font is
'         picked up.
'
'  Microsoft publish the full list of glyphs and their names at
'  learn.microsoft.com/windows/apps/design/style/segoe-fluent-icons-font, which
'  is the place to look when picking a new one.
'
'  To use the fallback instead, change the one line in
'  Themes\VisualStudioDark.xaml:
'
'      <FontFamily x:Key="Font_Glyph">Segoe MDL2 Assets</FontFamily>
'
'  FOR AN AI AGENT
'
'  Check whether the font is there before doing anything:
'
'      Add-Type -AssemblyName System.Drawing
'      [System.Drawing.FontFamily]::Families |
'          Where-Object { $_.Name -eq 'Segoe Fluent Icons' }
'
'  or just test for the file:
'
'      Test-Path C:\Windows\Fonts\SegoeIcons.ttf
'
'  If it is present, stop - there is nothing to install.
'
'  If it is missing, do NOT try to install a font silently. Writing to
'  C:\Windows\Fonts or to the font registry keys needs elevation, and a font
'  installed behind the user's back is not yours to install. Instead:
'
'      - Tell the user the download is at https://aka.ms/SegoeFluentIcons and
'        walk them through the four steps above, OR
'      - Offer the fallback, which is a one-line edit: set the Font_Glyph
'        resource in Themes\VisualStudioDark.xaml to "Segoe MDL2 Assets".
'        Leave Weight_Glyph alone; MDL2 has no bold face either, so the weight
'        is synthesised there in the same way.
'
'  Either way, look at the result. A codepoint a font does not define is not
'  blank - GDI and WPF both draw a hollow box, so a wrong or missing glyph
'  still occupies its space and only a screenshot will tell you. To check a
'  set of codepoints in bulk, render one codepoint that no icon font defines
'  (U+F8F0 is a good one), keep its pixels as a signature, and compare every
'  other rendering against it: anything that matches is undefined.


' ============================================================================
'  4. NOTHING ELSE
'
'  System.IO.Ports is the only NuGet package in this project, and that is a
'  decision rather than an accident. Everything else is in the .NET 8 shared
'  framework:
'
'      - JSON for the settings, project, workspace and device files is
'        System.Text.Json. That is why the config format is JSON and not TOML.
'      - The USB bus scan that finds the probe is SetupAPI through P/Invoke in
'        Modules\ProbeControl.vb, not a library.
'      - The icons are a font, as above, not an image pack.
'      - The theme is plain WPF styles in Themes\, not a UI toolkit.
'
'  Before adding a package, ask whether it is needed the way SerialPort was
'  needed - there was no in-framework way to open a COM port. Anything that can
'  be done with what is already here should be.
' ============================================================================
