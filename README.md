# ListInstalledSoftware — GitHub Actions Builder

This repository builds a **standalone** Windows EXE (bundled Python, no Python needed on end-user PCs)
using **GitHub Actions** (Windows runner).

## How to use

1. Create a new private repo on GitHub.
2. Download this bundle and extract its contents.
3. Commit & push everything (including `.github/workflows/build.yml`).
4. In GitHub → **Actions** tab → run **Build Windows EXE** (Workflow Dispatch).
5. When it finishes, open the run and download the artifact **ListInstalledSoftware** → `ListInstalledSoftware.exe`.

No Python install needed on your machine or users' PCs.

## Notes
- Built with PyInstaller `--onefile`; the EXE is self-contained.
- If SmartScreen warns, use *More info → Run anyway* or code-sign the EXE.
- To change the app name or icon, edit the build step in the workflow:
  ```yaml
  pyinstaller --onefile --icon youricon.ico --name ListInstalledSoftware list_software.py
  ```

Generated on 2025-11-19T07:31:31.321714
