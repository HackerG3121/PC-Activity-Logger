#!/bin/bash
# ══════════════════════════════════════════════════════════════
#  PC Activity Logger — macOS Setup & Automation
#  Installs deps, configures sync folder, schedules every 15min
# ══════════════════════════════════════════════════════════════

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOGGER="$SCRIPT_DIR/pc_logger.py"
PLIST="$HOME/Library/LaunchAgents/com.pclogger.plist"

echo ""
echo "  ╔══════════════════════════════════════════════════════╗"
echo "  ║   PC Activity Logger — macOS Setup                  ║"
echo "  ╚══════════════════════════════════════════════════════╝"
echo ""

# ── Step 1: Check Python ──────────────────────────────────────
if ! command -v python3 &>/dev/null; then
    echo "  [ERROR] python3 not found."
    echo "  Install from: https://python.org/downloads"
    exit 1
fi
echo "  [OK] Python: $(python3 --version)"

# ── Step 2: Install packages ─────────────────────────────────
echo ""
echo "  Installing required packages..."
python3 -m pip install psutil openpyxl --upgrade --quiet
echo "  [OK] Packages installed."

# ── Step 3: Choose sync folder ───────────────────────────────
echo ""
echo "  Where should the log file be saved?"
echo "  1. iCloud Drive   (recommended)"
echo "  2. Google Drive"
echo "  3. Dropbox"
echo "  4. Desktop (no mobile sync)"
echo ""
read -p "  Enter 1-4: " CHOICE

case "$CHOICE" in
    1) FOLDER="$HOME/Library/Mobile Documents/com~apple~CloudDocs"
       FOLDER_LABEL="iCloud Drive" ;;
    2) FOLDER="$HOME/Google Drive"
       FOLDER_LABEL="Google Drive" ;;
    3) FOLDER="$HOME/Dropbox"
       FOLDER_LABEL="Dropbox" ;;
    4) FOLDER="$HOME/Desktop"
       FOLDER_LABEL="Desktop" ;;
    *) FOLDER="$HOME/Desktop"
       FOLDER_LABEL="Desktop"
       echo "  Invalid choice — defaulting to Desktop." ;;
esac

# Update path in script
ESCAPED="${FOLDER//\//\\/}"
sed -i.bak "s|LOG_FILE = .*PC_ActivityLog.*|LOG_FILE = \"${FOLDER}/PC_ActivityLog.xlsx\"|" "$LOGGER" && rm -f "$LOGGER.bak"
echo "  [OK] Log path → ${FOLDER}/PC_ActivityLog.xlsx"

# ── Step 4: macOS permissions reminder ───────────────────────
echo ""
echo "  ⚠️  REQUIRED macOS permissions:"
echo "     System Preferences → Privacy → Full Disk Access"
echo "     Add Terminal (or the app running this script)."
echo "     Bluetooth logging also needs this permission."
echo ""
read -p "  Press Enter to continue once permissions are set..."

# ── Step 5: Create LaunchAgent (run every 15 min + at login) ─
PYTHON_PATH="$(command -v python3)"
mkdir -p "$HOME/Library/LaunchAgents"
mkdir -p "$HOME/Library/Logs"

cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
    "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.pclogger</string>

    <key>ProgramArguments</key>
    <array>
        <string>${PYTHON_PATH}</string>
        <string>${LOGGER}</string>
    </array>

    <!-- Run at login -->
    <key>RunAtLoad</key>
    <true/>

    <!-- Repeat every 15 minutes (900 seconds) -->
    <key>StartInterval</key>
    <integer>900</integer>

    <!-- Restart if it crashes -->
    <key>KeepAlive</key>
    <false/>

    <key>StandardOutPath</key>
    <string>${HOME}/Library/Logs/pc_logger.log</string>
    <key>StandardErrorPath</key>
    <string>${HOME}/Library/Logs/pc_logger_err.log</string>
</dict>
</plist>
EOF

# Load the agent
launchctl unload "$PLIST" 2>/dev/null || true
launchctl load -w "$PLIST"
echo "  [OK] LaunchAgent installed — runs at login and every 15 minutes."

# ── Step 6: Run now ──────────────────────────────────────────
echo ""
echo "  ╔══════════════════════════════════════════════════════╗"
echo "  ║   Setup complete! Starting logger now...            ║"
echo "  ║   Press Ctrl+C at any time to stop.                 ║"
echo "  ╚══════════════════════════════════════════════════════╝"
echo ""
python3 "$LOGGER"
