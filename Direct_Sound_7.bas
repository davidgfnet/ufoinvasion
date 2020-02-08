Attribute VB_Name = "Direct_Sound_7"
'#####################################################################################
'################################# DIRECT SOUND ENGINE (by dvd) ######################
'#####################################################################################

Global DSound As DirectSound

Public Sub InitDS()
Set DSound = Dx.DirectSoundCreate("")
DSound.SetCooperativeLevel Juego.hWnd, DSSCL_NORMAL
End Sub

Public Function CargarSonido(ByVal File As String) As DirectSoundBuffer
Dim bdesc As DSBUFFERDESC
Dim formato As WAVEFORMATEX
Set CargarSonido = DSound.CreateSoundBufferFromFile(File, bdesc, formato)
End Function

Public Sub ReproducirSonido(Sonido As DirectSoundBuffer)
Sonido.Play DSBPLAY_DEFAULT
End Sub

