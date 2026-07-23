# OutlookTools Setup Guide

## Quick Install (No Admin Required)

1. **Build the project** (see the main README) **or** download a pre‑built release from GitHub.
2. Open the `Setup` folder located at the root of the repository.
3. Run **`Setup.bat`** – this will:
   - Copy the compiled `OutlookTools.dll` to `%LOCALAPPDATA%\OutlookTools`.
   - Register the DLL for the current user (HKCU). No elevation needed.
   - Close any running Outlook instance and restart it.
4. Open Outlook – you will see a new **OutlookTools** tab on the ribbon.

> ⚠️ **Important:** The installer works completely per‑user. Do **not** run it as Administrator.

## Inno Setup Installer (Optional)

If you prefer a single `.exe` installer you can build one with Inno Setup:

1. Make sure **Inno Setup** is installed (download from https://jrsoftware.org/isinfo.php).
2. Build the project in Release mode – the output DLL will be in `OutlookTools/bin/Release/`.
3. Open `OutlookToolsSetup.iss` in the Inno Setup Compiler.
4. Press **Ctrl+F9** (or *Build → Compile*) to generate `OutlookTools_Setup_1.2.0.exe` in the `Output` folder.
5. Distribute the generated `.exe` – double‑click to install. It performs the same per‑user registration as `Setup.bat`.

## Uninstall

To completely remove OutlookTools:

- Run **`Setup/Uninstall.bat`** – it will unregister the COM add‑in and delete the files under `%LOCALAPPDATA%\OutlookTools`.
- No admin rights are required.

## Troubleshooting

- **Add‑in does not appear:** Ensure Outlook was restarted after installation.
- **Registry entry missing:** Verify `HKCU\Software\Microsoft\Office\Outlook\Addins\OutlookTools.AddIn` exists.
- **DLL registration failed:** Make sure the .NET Framework 4.8 is installed and that `regasm.exe` is present at `%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe`.

---

Enjoy OutlookTools – a fully open‑source, per‑user Outlook add‑in with no admin or network requirements.
