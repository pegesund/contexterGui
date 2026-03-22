tell application "Microsoft Word"
    set d to active document
    set t to text object of d
    set wc to count of words of t
    set results to ""
    set uCount to 0

    repeat with i from 1 to wc
        try
            set w to word i of t
            set f to font object of w
            set uType to underline of f
            if uType is not underline none then
                set theWord to content of w
                set uCount to uCount + 1
                set results to results & uCount & ": " & theWord & " type=" & uType & linefeed
            end if
        end try
    end repeat

    if results is "" then
        return "0 underlined words"
    else
        return uCount & " underlined words:" & linefeed & results
    end if
end tell
