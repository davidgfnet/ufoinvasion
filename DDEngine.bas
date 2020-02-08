Attribute VB_Name = "Direct_Draw_7"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global Dx As DirectX7
Global DDraw As DirectDraw7

Public ViewDesc As DDSURFACEDESC2
Public BackBufferCaps As DDSCAPS2
Public BackBuffer As DirectDrawSurface7
Public View As DirectDrawSurface7

Public Background As DirectDrawSurface7
Public BackgroundDesc As DDSURFACEDESC2

Public Sub InitDx()
Set Dx = New DirectX7
End Sub

Public Sub InitDD()
Set DDraw = Dx.DirectDrawCreate("")
DDraw.SetCooperativeLevel Juego.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
End Sub


'------------------------------------------ DAVID -------------------------------------
Public Sub LoadS(ByVal File As String, out As DirectDrawSurface7, ancho As Integer, alto As Integer)
Dim objdesc As DDSURFACEDESC2
CreateSurfaceFromFile out, objdesc, File, ancho, alto
End Sub

Public Sub Dibuja(x As Integer, y As Integer, ancho As Integer, alto As Integer, obj As DirectDrawSurface7, Optional ByVal SinT As Boolean = False)
Dim SrcRect As RECT
SetRect SrcRect, 0, 0, ancho, alto
If SinT = False Then
    BackBuffer.BltFast x, y, obj, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
Else
    BackBuffer.BltFast x, y, obj, SrcRect, DDBLTFAST_WAIT
End If
End Sub

Public Sub DibujaAdv(x As Integer, y As Integer, ancho As Integer, alto As Integer, obj As DirectDrawSurface7, Optional ByVal SinT As Boolean = False)
Dim recorte As Integer, recorte2 As Integer
Dim SrcRect As RECT, anchorecortado As Integer
anchorecortado = ancho
If (x + ancho) > PAncho Then
    anchorecortado = (x + ancho) - PAncho
End If
If y < 0 Then
    recorte = -y
    recorte = alto + y
    SetRect SrcRect, 0, -y, anchorecortado, recorte
    BackBuffer.BltFast x, 0, obj, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
ElseIf (y + alto) > PAlto Then
    recorte = alto - ((y + alto) - PAlto)
    SetRect SrcRect, 0, 0, anchorecortado, recorte
    BackBuffer.BltFast x, y, obj, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
Else
    Dibuja x, y, anchorecortado, alto, obj, SinT
End If
End Sub

Public Sub Mostrar()
On Local Error Resume Next
Err.Clear: Err.Number = 0
View.Flip Nothing, DDFLIP_WAIT
If Err.Number <> 0 Then SalidaCritica = True
End Sub
'----------------------------------------------------------------------


'------------------------------------- FUNCIONES DDRAW -------------------------------------------
'=================================================================================================
'-------------------------------------------------------------------------------------------------
Sub CreatePrimaryAndBackBuffer()
On Local Error Resume Next
Set View = Nothing
Set BackBuffer = Nothing

ViewDesc.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
ViewDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
ViewDesc.lBackBufferCount = 1
Set View = DDraw.CreateSurface(ViewDesc)

BackBufferCaps.lCaps = DDSCAPS_BACKBUFFER
Set BackBuffer = View.GetAttachedSurface(BackBufferCaps)
BackBuffer.GetSurfaceDesc ViewDesc

BackBuffer.SetFontTransparency True
End Sub

Sub CreateSurfaceFromFile(Surface As DirectDrawSurface7, Surfdesc As DDSURFACEDESC2, filename As String, Width As Integer, Height As Integer)
On Error GoTo LostFile
     Surfdesc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
     Surfdesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
     Surfdesc.lWidth = Width
     Surfdesc.lHeight = Height
     
     Set Surface = DDraw.CreateSurfaceFromFile(filename, Surfdesc)
Exit Sub
LostFile:
MsgBox "Error, archivo no encontrado"
End Sub

Sub SetRect(Box As RECT, Left As Integer, Top As Integer, Width As Integer, Height As Integer)
Box.Left = Left
Box.Top = Top
Box.Right = Left + Width
Box.Bottom = Top + Height
End Sub

Sub SetDisplayMode(Width As Integer, Height As Integer, Colors As Byte)
DDraw.SetDisplayMode Width, Height, Colors, 0, DDSDM_DEFAULT
End Sub

Sub RestoreDisplayMode()
DDraw.RestoreDisplayMode
End Sub
'-------------------------------------------------------------------------------------------------
'=================================================================================================
'-------------------------------------------------------------------------------------------------
