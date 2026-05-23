# 🖥️ PC Activity Logger — Final Ultimate Edition

> **Every connect · disconnect · change event on your PC — logged automatically to Excel, accessible from your phone anywhere.**



## 📋 What It Does

PC Activity Logger silently monitors your computer in the background and writes every hardware, network, and system event to a colour-coded Excel workbook — which syncs to your phone via Google Drive, OneDrive, or Dropbox every 30 seconds.

**Runs once on setup. Repeats every 15 minutes automatically. Zero maintenance.**

---

## 📊 All 30 Logged Categories

| Category | Events Logged |
|---|---|
| ⚡ Power | Startup · Shutdown · Sleep · Wake · Hibernate · Resume |
| 🖥️ Applications | Open & close with exact duration |
| 🔌 USB Devices | Insert · Remove — all types, serial ID |
| 📶 Wi-Fi | Connect · Disconnect — SSID / signal / band / IP / MAC |
| 🔵 Bluetooth | Connect · Disconnect — name / MAC / type / RSSI |
| 🌐 Network | Interface up/down · IP changes · Ethernet plug/unplug |
| 📡 Internet | Online · Offline · latency · outage duration |
| 🔋 Battery | Level % · charging · discharging · low · critical |
| 🖵 Screen | Lock · Unlock with exact duration |
| 💾 Drives | Mount · Unmount · filesystem · size · free space |
| 🔊 Audio | Speaker / headphone connect · disconnect |
| 🎙️ Microphone | Mic connect · disconnect |
| 📷 Webcam | Camera connect · disconnect |
| 🖨️ Printers | Install · Remove · port · driver |
| 📺 Monitors | Display connect · disconnect · resolution · port |
| 🌍 VPN | Connect · Disconnect · server · protocol · duration |
| 📱 Mobile Hotspot | USB tethering · mobile hotspot events |
| 🗂️ Network Drives | Map · Unmap · UNC path · server |
| 🎮 Controllers | Game controller connect · disconnect |
| 🔐 Remote Sessions | RDP · SSH · TeamViewer sessions |
| 🛡️ Security | Firewall · Defender · UAC status changes |
| 👤 Users | Login · Logout · session switch |
| ⌨️ Keyboard / Mouse | USB keyboard & mouse connect · disconnect |
| ⚙️ Power Plan | Balanced · Performance · Battery Saver changes |
| 🔧 Services | Key Windows service start · stop |
| 📊 System Perf | CPU / RAM spike events (>85%) |
| 🔑 Smart Card | Smart card reader connect · disconnect |
| 🌙 Display Mode | Night mode · HDR · colour profile changes |
| 📋 Clipboard | Text changes *(optional — disabled by default)* |
| 📊 Daily Summary | All-day totals per category |

---

## 🚀 Quick Start

### Windows
```bat
setup_windows.bat
```
*Double-click or run from Command Prompt. The script installs everything, picks your sync folder, and schedules automatic runs.*

### macOS
```bash
chmod +x setup_mac.sh && ./setup_mac.sh
```

### Linux
```bash
chmod +x setup_linux.sh && ./setup_linux.sh
```

---

## ⚙️ Manual Installation

### 1 · Install Python 3.9+
Download from [python.org/downloads](https://python.org/downloads).  
On Windows, tick **"Add Python to PATH"** during install.

### 2 · Install packages

**Windows:**
```bash
pip install psutil openpyxl pywin32
```

**macOS / Linux:**
```bash
pip3 install psutil openpyxl
# Linux extras:
sudo apt install iwgetid bluetooth bluez-tools pulseaudio-utils xclip
```

### 3 · Configure the log path

Open `pc_logger.py` and set your sync folder on the `LOG_FILE` line (around line 60):

```python
# OneDrive (Windows default)
LOG_FILE = os.path.join(os.path.expanduser("~"), "OneDrive", "PC_ActivityLog.xlsx")

# Google Drive
LOG_FILE = os.path.join(os.path.expanduser("~"), "Google Drive", "PC_ActivityLog.xlsx")

# Dropbox
LOG_FILE = os.path.join(os.path.expanduser("~"), "Dropbox", "PC_ActivityLog.xlsx")

# Desktop (no mobile sync)
LOG_FILE = os.path.join(os.path.expanduser("~"), "Desktop", "PC_ActivityLog.xlsx")
```

### 4 · Run
```bash
python pc_logger.py        # Windows
python3 pc_logger.py       # macOS / Linux
```

---

## 🔄 Automation (every 15 minutes)

The setup scripts configure this automatically. To set it up manually:

### Windows — Task Scheduler
```bat
schtasks /create /tn "PC_ActivityLogger" /tr "python C:\path\to\pc_logger.py" /sc ONLOGON /ru %USERNAME% /rl HIGHEST /f
```

### macOS — LaunchAgent
```bash
# Place com.pclogger.plist in ~/Library/LaunchAgents/
# (see setup_mac.sh for the full plist template)
launchctl load -w ~/Library/LaunchAgents/com.pclogger.plist
```

### Linux — systemd timer
```bash
systemctl --user enable --now pc-logger.timer
# (see setup_linux.sh for the full .service and .timer templates)
```

---

## 📱 Mobile Access

1. Install **Google Drive / OneDrive / Dropbox** on your phone
2. Sign in with the same account as your PC
3. Find `PC_ActivityLog.xlsx` in your sync folder
4. Open it — it updates every 30 seconds automatically

> **Best mobile apps:** Android → Google Sheets or Excel · iPhone → Microsoft Excel or Numbers

---

## 🛠️ Configuration Options

All settings are at the top of `pc_logger.py`:

| Setting | Default | Description |
|---|---|---|
| `LOG_FILE` | `~/OneDrive/PC_ActivityLog.xlsx` | Output file path |
| `LOG_CLIPBOARD` | `False` | Enable clipboard text logging |
| `LOG_APPS` | `True` | Track app open/close |
| `LOG_PERF` | `True` | Log CPU/RAM spikes |
| `CPU_SPIKE_THRESH` | `85` | CPU % to trigger a spike alert |
| `RAM_SPIKE_THRESH` | `85` | RAM % to trigger a spike alert |
| `SAVE_EVERY` | `30` | Auto-save interval (seconds) |
| `BATTERY_DELTA` | `2` | Min battery % change to log |
| `WATCHED_SERVICES` | `[list]` | Windows services to monitor |

---

## ❓ Troubleshooting

| Problem | Solution |
|---|---|
| `Save error` on startup | Close `PC_ActivityLog.xlsx` on your PC — it can't save while open |
| Bluetooth not detected | Run as Administrator (Windows) or grant Full Disk Access (macOS) |
| No USB events on Windows | Run `pip install pywin32` then restart |
| Excel not syncing to phone | Check that Google Drive / OneDrive is running in the system tray |
| Too many app log entries | Set `LOG_APPS = False` in the config section |
| Stop the logger | Press `Ctrl+C` — it saves a final snapshot before stopping |

---

## 📁 File Structure

```
pc-activity-logger/
├── pc_logger.py        ← Main logger (run this)
├── setup_windows.bat   ← Windows: install + automate
├── setup_mac.sh        ← macOS:   install + automate
├── setup_linux.sh      ← Linux:   install + automate
└── README.md
```

---

## 🤝 Contributing

Pull requests are welcome. For major changes, please open an issue first.

---

## 📄 License

MIT — free to use, modify, and distribute.