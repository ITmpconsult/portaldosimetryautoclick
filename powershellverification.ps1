$wshell = New-Object -ComObject wscript.shell;
$signature=@'
[DllImport("user32.dll",CharSet=CharSet.Auto,CallingConvention=CallingConvention.StdCall)]
public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@

$SendMouseClick = Add-Type -MemberDefinition $signature -name "WinreMouseEventNew" -Namespace Win32Functions -PassThru

#click "Planning"
$x=180
$y=69
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y)
sleep -Seconds 03
$SendMouseClick::mouse_event(0x0002,0,0,0,0);
$SendMouseClick::mouse_event(0x0004,0,0,0,0);

#click "Create verification plan"
$x=216
$y=240
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y)
sleep -Seconds 03
$SendMouseClick::mouse_event(0x0002,0,0,0,0);
$SendMouseClick::mouse_event(0x0004,0,0,0,0);

#click "Create new course"
$x=1176
$y=528
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y)
sleep -Seconds 03
$SendMouseClick::mouse_event(0x0002,0,0,0,0);
$SendMouseClick::mouse_event(0x0004,0,0,0,0);

#double click to name course "RAQA"
$x=786
$y=325
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y)
sleep -Seconds 03
$SendMouseClick::mouse_event(0x0002,0,0,0,0);
$SendMouseClick::mouse_event(0x0004,0,0,0,0);
$SendMouseClick::mouse_event(0x0002,0,0,0,0);
$SendMouseClick::mouse_event(0x0004,0,0,0,0);
$wshell.SendKeys("RAQA");

#click "Click to show dropdown list"
$x=1025
$y=370
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y)
sleep -Seconds 03
$SendMouseClick::mouse_event(0x0002,0,0,0,0);
$SendMouseClick::mouse_event(0x0004,0,0,0,0);

#click "Select Physics QA"
$x=1005
$y=541
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y)
sleep -Seconds 03
$SendMouseClick::mouse_event(0x0002,0,0,0,0);
$SendMouseClick::mouse_event(0x0004,0,0,0,0);

#click "Click OK"
$x=770
$y=768
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y)
sleep -Seconds 03
$SendMouseClick::mouse_event(0x0002,0,0,0,0);
$SendMouseClick::mouse_event(0x0004,0,0,0,0);

for ($i=0; $i -lt 6; $i++) {
    #click "Click proceed next step"
    $x=1195
    $y=811
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x,$y)
    sleep -Seconds 03
    $SendMouseClick::mouse_event(0x0002,0,0,0,0);
    $SendMouseClick::mouse_event(0x0004,0,0,0,0);
}

