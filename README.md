# MySlice Office Add-in — Desktop EXE (rootLuncher)

Windows **installer + protocol launcher** for the MySlice Word / Excel add-in.

This is not the add-in UI. Taskpane, ribbon, and save-to-S3 live in `MY NEW X1 OFFIce ADD INS`.

| Doc | Content |
| --- | --- |
| [markdown/INSTALLER-AND-LAUNCHER-FLOW.md](markdown/INSTALLER-AND-LAUNCHER-FLOW.md) | Install, `mysliceLTS://`, download, open Office |
| `MY NEW X1 OFFIce ADD INS/markdown/` | End-to-end Edit/save, APIs, legal / enterprise |

## Two EXEs

| Folder | Product | Role |
| --- | --- | --- |
| `officeLuncher` | **Installer** (`MySlice Setup`) | NSIS setup (Admin). Deploys launcher, registers `mysliceLTS://`, copies `manifest.xml`, creates Office trusted catalog share. |
| `mysliceLauncher` | **Protocol handler** (`mysliceLTS.exe`) | Every Document Library **Edit**. Downloads file, puts file UUID in the temp name, opens Word/Excel. |

The setup EXE wraps **one** Electron app (the launcher). It is not an Electron app itself.

## Local build

```bash
cd mysliceLauncher
npm install
npm run build

cd ../officeLuncher
npm install
npm run dist
```

Output: `officeLuncher/dist/MySlice Setup 1.0.0.exe`

Installed files: `C:\ProgramData\myslice\mysliceLTS\`

Uninstall: Settings → Apps → MySlice (run as Administrator). Dev: `cd officeLuncher && npm run start:uninstall`
