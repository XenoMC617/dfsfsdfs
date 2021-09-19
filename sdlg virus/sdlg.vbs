do
Mensaje=msgbox("SDLG :V" ,16, "Mensaje")
Set Sound = CreateObject("WMPlayer.OCX.7")
Sound.URL = "Himno de la Grasa  -v.wav"
Sound.Controls.play
do while Sound.currentmedia.duration = 0
wscript.sleep 200
loop
wscript.sleep (int(Sound.currentmedia.duration)+1)*1000
loop
