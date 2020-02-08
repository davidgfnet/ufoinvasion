Attribute VB_Name = "Direct_Music_7"
'#####################################################################################
'################################# DIRECT MUSIC ENGINE (by dvd) ######################
'#####################################################################################

Global DMusicLoader As DirectMusicLoader
Global DMusicPerformance As DirectMusicPerformance
Global estado As DirectMusicSegmentState

Public Sub InitDM()
Set DMusicLoader = Dx.DirectMusicLoaderCreate()
Set DMusicPerformance = Dx.DirectMusicPerformanceCreate()
Call DMusicPerformance.Init(Nothing, 0)

DMusicPerformance.SetPort -1, 80
Call DMusicPerformance.SetMasterAutoDownload(True)

DMusicPerformance.SetMasterVolume (85 * 42 - 3000)
End Sub

Public Function CargarMidi(ByVal archivo As String) As DirectMusicSegment
Set CargarMidi = DMusicLoader.LoadSegment(archivo)
CargarMidi.SetStandardMidiFile
End Function

Public Sub ReproducirMidi(segmento As DirectMusicSegment, Optional ByVal Inicio As Long)
segmento.SetStartPoint Inicio
Set estado = DMusicPerformance.PlaySegment(segmento, 0, 0)
End Sub

Public Sub PararMidi(segmento As DirectMusicSegment)
DMusicPerformance.Stop segmento, estado, 0, 0
End Sub

Public Function ReproduciendoMidi(segmento As DirectMusicSegment) As Boolean
ReproduciendoMidi = DMusicPerformance.IsPlaying(segmento, estado)
End Function
