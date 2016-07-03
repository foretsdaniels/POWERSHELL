function global:INVOKE-MEDIA ( $Filename ) { 
 
# Windows Powershell Function 
# to play Files in Media Player in Hidden mode 
# 
# "ssssshhhhh....." 
# 
# Create Reference to COMObject for MediaPlayer 
# Found name in Registry in HKLM\Software\Classes 
# 
# Fixed typo in script (Missing - before Comobject) 2/20/2015 
 
$PLAYER=NEW-OBJECT -ComObject 'Mediaplayer.Mediaplayer' 
$PLAYER.Filename=$Filename 
$Player.Play() 
 
} 
