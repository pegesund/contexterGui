#!/bin/bash
# Reload the Word add-in taskpane by right-clicking and selecting "Last inn på nytt"
osascript <<'EOF'
tell application "Microsoft Word" to activate
delay 0.2
tell application "System Events"
    tell process "Microsoft Word"
        set w to window 1
        set {x, y} to position of w
        set {width, height} to size of w
        click at {x + width - 100, y + 300}
        delay 0.1
        perform action "AXShowMenu" of (click at {x + width - 100, y + 300})
        delay 0.3
        key code 125
        delay 0.1
        key code 36
    end tell
end tell
EOF
