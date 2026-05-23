#!/bin/bash
# ══════════════════════════════════════════════════════════════
#  PC Activity Logger — Linux Setup & Automation
#  Installs deps, configures sync folder, schedules every 15min
# ══════════════════════════════════════════════════════════════

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOGGER="$SCRIPT_DIR/pc_logger.py"
SERVICE_FILE="$HOME/.config/systemd/user/pc-logger.service"
TIMER_FILE="$HOME/.config/systemd/user/pc-logger.timer"

echo ""
echo "  ╔══════════════════════════════════════════════════════╗"
echo "  ║   PC Activity Logger — Linux Setup                  ║"
echo "  ╚══════════════════════════════════════════════════════╝"
echo ""

# ── Step 1: Check Python ──────────────────────────────────────
if ! command -v python3 &>/dev/null; then
    echo "  [ERROR] python3 not found. Install with:"
    echo "    sudo apt install python3 python3-pip   # Ubuntu/Debian"
    echo "    sudo dnf install python3 python3-pip   # Fedora"
    exit 1
fi
echo "  [OK] Python: $(python3 --version)"

# ── Step 2: Install system tools & Python packages ───────────
echo ""
echo "  Installing system tools..."
if command -v apt &>/dev/null; then
    sudo apt install -y iwgetid bluetooth bluez-tools pulseaudio-utils xclip xinput 2>/dev/null || true
elif command -v dnf &>/dev/null; then
    sudo dnf install -y iwconfig bluez pulseaudio-utils xclip xinput 2>/dev/null || true
fi

echo "  Installing Python packages..."
python3 -m pip install psutil openpyxl --upgrade --quiet
echo "  [OK] Packages installed."

# ── Step 3: Choose sync folder ───────────────────────────────
echo ""
echo "  Where should the log file be saved?"
echo "  1. Google Drive  (~/.google-drive or rclone mount)"
echo "  2. Dropbox       (~/Dropbox)"
echo "  3. Nextcloud     (~/Nextcloud)"
echo "  4. Desktop       (no mobile sync)"
echo ""
read -p "  Enter 1-4: " CHOICE

case "$CHOICE" in
    1) FOLDER="$HOME/Google Drive"   ;;
    2) FOLDER="$HOME/Dropbox"        ;;
    3) FOLDER="$HOME/Nextcloud"      ;;
    4) FOLDER="$HOME/Desktop"        ;;
    *) FOLDER="$HOME/Desktop"
       echo "  Invalid choice — defaulting to Desktop." ;;
esac

# Update path in script
ESCAPED_FOLDER="${FOLDER//\//\\/}"
sed -i.bak "s|LOG_FILE = .*PC_ActivityLog.*|LOG_FILE = \"${FOLDER}/PC_ActivityLog.xlsx\"|" "$LOGGER" && rm -f "$LOGGER.bak"
echo "  [OK] Log path → ${FOLDER}/PC_ActivityLog.xlsx"

# ── Step 4: Ensure log folder exists ─────────────────────────
mkdir -p "$FOLDER"

# ── Step 5: Install systemd user service + timer ─────────────
PYTHON_PATH="$(command -v python3)"
mkdir -p "$HOME/.config/systemd/user"
mkdir -p "$HOME/.local/share/pc-logger"

# Service file
cat > "$SERVICE_FILE" <<EOF
[Unit]
Description=PC Activity Logger
After=network.target

[Service]
Type=simple
ExecStart=${PYTHON_PATH} ${LOGGER}
WorkingDirectory=${SCRIPT_DIR}
Restart=on-failure
RestartSec=30
StandardOutput=append:${HOME}/.local/share/pc-logger/pc_logger.log
StandardError=append:${HOME}/.local/share/pc-logger/pc_logger_err.log
Environment=DISPLAY=:0
Environment=DBUS_SESSION_BUS_ADDRESS=unix:path=/run/user/%U/bus

[Install]
WantedBy=default.target
EOF

# Timer file (every 15 minutes)
cat > "$TIMER_FILE" <<EOF
[Unit]
Description=PC Activity Logger — Run every 15 minutes
Requires=pc-logger.service

[Timer]
OnBootSec=30sec
OnUnitActiveSec=15min
AccuracySec=1min
Persistent=true

[Install]
WantedBy=timers.target
EOF

# Enable and start
systemctl --user daemon-reload
systemctl --user enable pc-logger.timer
systemctl --user start  pc-logger.timer
systemctl --user enable pc-logger.service

# Enable lingering so it runs without login
loginctl enable-linger "$USER" 2>/dev/null || true

echo "  [OK] systemd service + timer installed — runs every 15 minutes."
echo "  To check status: systemctl --user status pc-logger.timer"
echo "  To view logs   : journalctl --user -u pc-logger -f"

# ── Step 6: Run now ──────────────────────────────────────────
echo ""
echo "  ╔══════════════════════════════════════════════════════╗"
echo "  ║   Setup complete! Starting logger now...            ║"
echo "  ║   Press Ctrl+C at any time to stop.                 ║"
echo "  ╚══════════════════════════════════════════════════════╝"
echo ""
python3 "$LOGGER"
