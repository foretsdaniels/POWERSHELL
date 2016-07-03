# Define the media file path:
$AudioFile = "C:\Users\Public\Music\Sample Music\Amanda.wma"

# Create the Media Player object:
$Player = New-Object -ComObject "WMPlayer.OCX"

# Play the music:
$Player.openPlayer($AudioFile)
