# 🚀 PC Activity Logger — Final Ultimate Edition

**Every connect · disconnect · change event on your PC, logged automatically to a colour-coded Excel file — accessible from your phone anywhere in the world.**

---

## 🎯 What This Project Does

PC Activity Logger runs silently in the background and writes **every hardware, network, and system event** to a live Excel workbook that syncs to your phone via Google Drive, OneDrive, or Dropbox every 30 seconds.

- **Zero configuration** after one-time setup
- **Runs automatically every 15 minutes** via Task Scheduler / LaunchAgent / systemd
- **30 colour-coded Excel sheets** — one per category
- **Works on Windows, macOS, and Linux** with dedicated OS-specific scripts

---

## 📦 What's Inside This Release

| File | OS | Description |
|---|---|---|
| `pc_logger_windows.py` | 🪟 Windows 10/11 | Full logger using WMI, PowerShell, win32 APIs |
| `pc_logger_mac.py` | 🍎 macOS 12+ | Full logger using system_profiler, airport, Quartz |
| `pc_logger_linux.py` | 🐧 Ubuntu/Debian/Fedora | Full logger using iwgetid, bluetoothctl, pactl, xrandr |
| `setup_windows.bat` | 🪟 Windows | One-click install + Task Scheduler automation |
| `setup_mac.sh` | 🍎 macOS | One-click install + LaunchAgent automation |
| `setup_linux.sh` | 🐧 Linux | One-click install + systemd timer automation |
| `README.md` | All | Full documentation and setup guide |

---

## 📊 All 30 Logged Categories

```
⚡ Power           🖥️  Applications    🔌 USB Devices     📶 Wi-Fi
🔵 Bluetooth       🌐 Network          📡 Internet         🔋 Battery
🖵  Screen          💾 Drives           🔊 Audio            🎙️  Microphone
📷 Webcam          🖨️  Printers         📺 Monitors         🌍 VPN
📱 Hotspot         🗂️  Network Drives   🎮 Controllers      🔐 Remote Sessions
🛡️  Security        👤 Users            ⌨️  Keyboard/Mouse   ⚙️  Power Plan
🔧 Services        📊 System Perf      🔑 Smart Card       🌙 Display Mode
📋 Clipboard       📊 Daily Summary
```

---

## ⚡ Quick Start — 3 Steps

### Windows
```bat
# 1. Double-click or run in terminal:
setup_windows.bat
```

### macOS
```bash
# 1. Open Terminal and run:
chmod +x setup_mac.sh && ./setup_mac.sh
```

### Linux
```bash
# 1. Open Terminal and run:
chmod +x setup_linux.sh && ./setup_linux.sh
```

> The setup script installs dependencies, picks your sync folder, and schedules automatic runs. That's it — fully hands-off after first run.

---

## 🔄 Automation Details

| OS | Method | Schedule |
|---|---|---|
| Windows | Task Scheduler | On login + every 15 minutes |
| macOS | LaunchAgent plist | On login + every 15 minutes |
| Linux | systemd user timer | On boot + every 15 minutes |

The logger saves a fresh Excel snapshot every 30 seconds, so your phone always sees data that is less than a minute old.

---

## 📱 Mobile Access

1. Install Google Drive / OneDrive / Dropbox on your phone
2. Sign in with the same account as your PC
3. Open `PC_ActivityLog.xlsx` — it refreshes every 30 seconds

**Best apps:** Android → Google Sheets or Microsoft Excel · iPhone → Microsoft Excel or Numbers

---

## 🛠️ Manual Install (without setup script)

```bash
# Windows
pip install psutil openpyxl pywin32

# macOS
pip3 install psutil openpyxl

# Linux (Ubuntu/Debian)
pip3 install psutil openpyxl
sudo apt install iwgetid bluetooth bluez-tools pulseaudio-utils xclip xinput
```

Then run the correct file for your OS:
```bash
python  pc_logger_windows.py   # Windows
python3 pc_logger_mac.py       # macOS
python3 pc_logger_linux.py     # Linux
```

---

## ⚙️ Key Configuration Options

All settings are at the top of each OS file:

| Setting | Default | What It Does |
|---|---|---|
| `LOG_FILE` | `~/OneDrive/PC_ActivityLog.xlsx` | Where the log file is saved |
| `LOG_CLIPBOARD` | `False` | Enable clipboard text logging |
| `LOG_APPS` | `True` | Track app open/close |
| `LOG_PERF` | `True` | Alert on CPU/RAM spikes |
| `CPU_SPIKE_THRESH` | `85` | CPU % that triggers a spike log |
| `RAM_SPIKE_THRESH` | `85` | RAM % that triggers a spike log |
| `SAVE_EVERY` | `30` | Auto-save interval in seconds |

---

## 🐛 Troubleshooting

| Problem | Fix |
|---|---|
| `Save error` on startup | Close `PC_ActivityLog.xlsx` — can't save while it's open |
| Bluetooth not detected (Windows) | Run as Administrator |
| Bluetooth not detected (macOS) | Grant Full Disk Access in System Preferences → Privacy |
| No USB events on Windows | `pip install pywin32` then restart |
| Excel not syncing to phone | Check that Google Drive / OneDrive is running in the system tray |
| Too many app entries | Set `LOG_APPS = False` |
| Stop the logger | Press `Ctrl+C` — saves final snapshot before stopping |

---

## 📁 Project Structure

```
pc-activity-logger/
├── pc_logger_windows.py   ← Windows-only logger (WMI / PowerShell)
├── pc_logger_mac.py       ← macOS-only logger   (system_profiler / airport)
├── pc_logger_linux.py     ← Linux-only logger   (iwgetid / bluetoothctl / pactl)
├── setup_windows.bat      ← Windows install + Task Scheduler automation
├── setup_mac.sh           ← macOS install + LaunchAgent automation
├── setup_linux.sh         ← Linux install + systemd timer automation
└── README.md              ← Full documentation
```

---

## 📄 License

MIT — free to use, modify, and distribute.

---

## 💬 Share This

> **"I built a Python tool that logs every USB plug/unplug, Wi-Fi connect, Bluetooth device, app open/close, battery change, screen lock, VPN connection, and 22 more categories — all to a colour-coded Excel file that syncs live to your phone. Works on Windows, Mac, and Linux. Setup takes 60 seconds."**

⭐ Star this repo if it helped you!
🍴 Fork it and make it yours.
🐛 Found a bug? Open an issue.
