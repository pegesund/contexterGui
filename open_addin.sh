#!/bin/bash
# Open NorskTale add-in taskpane in Word
# This only needs to be done once after a fresh Word restart.
# After that, the taskpane stays open across document changes.

echo "Opening NorskTale add-in..."
echo "If the taskpane doesn't appear automatically:"
echo "  1. In Word, go to Insert > My Add-ins (or Sett inn > Mine Tillegg)"
echo "  2. Click NorskTale to open the taskpane"
echo "  3. Or: Insert > Add-ins > My Add-ins"
echo ""

# Try to open via the Insert ribbon
osascript -e '
tell application "Microsoft Word" to activate
delay 0.5
tell application "System Events"
    tell process "Microsoft Word"
        -- Click Insert tab
        click radio button "Sett inn" of tab group 1 of window 1
        delay 1
        -- Look for any add-in related button by clicking through menu buttons
        set menuBtns to every menu button of tab group 1 of window 1
        repeat with mb in menuBtns
            try
                if name of mb contains "tillegg" or name of mb contains "Add" then
                    click mb
                    delay 0.5
                end if
            end try
        end repeat
    end tell
end tell
' 2>/dev/null

# Wait and check
sleep 5
CONNECTED=$(curl -sk https://127.0.0.1:3000/errors 2>/dev/null | head -c 1)
if [ "$CONNECTED" = "[" ]; then
    echo "Add-in connected!"
else
    echo "Add-in NOT connected. Please open it manually in Word."
fi
