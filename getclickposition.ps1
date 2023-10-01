Add-Type -AssemblyName System.Windows.Forms


#this will get the mouse positions, so you can set them correctly with differenct screen sizes
while ($true) {
    $X = [System.Windows.Forms.Cursor]::Position.X
    $Y = [System.Windows.Forms.Cursor]::Position.Y
    Write-Output "X:$X | Y:$Y";
    Start-Sleep -Milliseconds 500
}
