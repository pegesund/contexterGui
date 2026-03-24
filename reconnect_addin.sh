#!/bin/bash
# Reconnect the Word add-in: ensure app is running, then click "Prøv på nytt" button
# Usage: ./reconnect_addin.sh

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Step 1: Ensure acatts-rust is running on port 3000
if ! lsof -i:3000 -sTCP:LISTEN > /dev/null 2>&1; then
    echo "App not running — starting..."
    cd "$SCRIPT_DIR"
    ./target/debug/acatts-rust &
    # Wait for port to be ready
    for i in 1 2 3 4 5; do
        lsof -i:3000 -sTCP:LISTEN > /dev/null 2>&1 && break
        sleep 1
    done
else
    echo "App already running on port 3000"
fi

# Step 2: Click "Prøv på nytt" button in the Tillegg sidebar
echo "Clicking 'Prøv på nytt' in sidebar..."
osascript -e '
tell application "Microsoft Word" to activate
delay 0.5
tell application "System Events"
    tell process "Microsoft Word"
        -- Find and click the "Prøv på nytt" button in the sidebar/taskpane
        set found to false
        set allButtons to every button of window 1
        repeat with b in allButtons
            try
                if name of b is "Prøv på nytt" then
                    click b
                    set found to true
                    exit repeat
                end if
            end try
        end repeat
        if not found then
            -- Try searching deeper in groups/web areas
            set allElems to entire contents of window 1
            repeat with e in allElems
                try
                    if class of e is button and name of e is "Prøv på nytt" then
                        click e
                        set found to true
                        exit repeat
                    end if
                end try
            end repeat
        end if
        return found
    end tell
end tell
' 2>&1

sleep 3

# Step 3: Verify connection
CHANGES=$(grep -ac "Addin changed\|HTTP /reset" /var/folders/5d/hbyysq95053_twbcvx4g650r0000gn/T/acatts-rust.log 2>/dev/null)
if [ "$CHANGES" -gt 0 ] 2>/dev/null; then
    echo "OK: Add-in connected ($CHANGES events)"
else
    echo "WARNING: No add-in events yet — may need more time"
fi
