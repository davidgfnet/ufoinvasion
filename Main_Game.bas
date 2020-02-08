Attribute VB_Name = "Main_Game"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

' ------------------ Teclas -----------------------------
Global KArriba As Boolean, KAbajo As Boolean, KDerecha As Boolean, KIzquierda As Boolean
Global KEscape As Boolean, KIntro As Boolean
'--------------------------------------------------------

Global TempoPrivado As Single
Global FondoA As Integer
Global PotenciaActual As Integer
Global MusicaActual As Integer
Global MostrandoFrase As Boolean
Global TimerMF As Single
Global PasoFondo As Single
Global Framejefe As Single
Global EntrandoE As Boolean
Global MaxPasoFondo As Single
Global OsciladorArma As Boolean
Global TempDir As String
Global tiempoMusica As Single
Global JPausado As Boolean
Global Reentrar As Boolean
Global FrameABala As Integer

'   Variables rápidas!!!
Global z As Integer, w As Integer, k As Integer, ww As Integer

'------------- Variables Estaticas declaradas como globales para poder borrarlas ----
Global S_recorte As Single
Global S1_fraseanterior As String
Global S1_wr As Single, S1_cursor As Integer
Global S2_fraseanterior As String
Global S2_wr As Single, S2_cursor As Integer
'------------------------------------------------------------------------------------

Global DibujandoFrase As Boolean

Public Type tObjeto
    x As Integer
    y As Integer
    tag As String
    tagint2 As Single
    tagint As Single
End Type
Public Type tNivel
    key As String
    frase As String
    frase2 As String
    vida As Boolean
    energia As Boolean
    upgrade As Boolean
    especial2 As Boolean
    TiempoFrase As Integer
    especial As Boolean
    TiempoA As Boolean
    
    Jefe As Boolean
    frasesecundaria As String
End Type
Public Type tUFO
    x As Integer
    y As Integer
    vida As Integer
    vidamax As Integer
    tipo As Integer
End Type
Public Type tDisparoEnemigo
    x As Integer
    y As Integer
End Type
Public Type tMet
    x As Integer
    y As Integer
    a As Integer
    frame As Single
    vida As Integer
End Type
Public Type tNave
    normal As DirectDrawSurface7
    derecha As DirectDrawSurface7
    izquierda As DirectDrawSurface7
    x As Integer
    y As Integer
End Type
Public Type tExp
    x As Integer
    y As Integer
    frame As Single
End Type

Global unaMejora As tObjeto
Global Jefe As tUFO
Global Fondos(1) As DirectDrawSurface7
Global MatrizJefes(1 To 1, 1 To 3) As DirectDrawSurface7
Global ExpActivas() As tExp
Global ExpActivasP() As tExp
Global unaVida As tObjeto
Global unaEnergia As tObjeto
Global Letras As DirectDrawSurface7
Global AnimVidas(1 To 5) As DirectDrawSurface7
Global AnimMejoras(1 To 12) As DirectDrawSurface7
Global AnimEnergias(1 To 2) As DirectDrawSurface7
Global AnimExplosion(1 To 11) As DirectDrawSurface7
Global Meteoritos(1 To 8) As DirectDrawSurface7
Global Menus(1 To 4, 1 To 2) As DirectDrawSurface7
Global Mini As DirectDrawSurface7
Global Autor As DirectDrawSurface7
Global Web As DirectDrawSurface7
Global DisparosE() As tDisparoEnemigo
Global UFOs() As tUFO
Global Niveles() As tNivel
Global NivelesB() As tNivel
Global Disparar As Boolean
Global Fondo As DirectDrawSurface7
Global Nave As tNave
Global FPS As Long
Global TimePerFrame As Single
Global EXE As String

Global SalidaCritica As Boolean
Global METs() As tMet
Global Inmunidad As Integer
Global VaivenInmunidad As Single
Global VIAumentar As Boolean
Global TiempoDDM As Integer

Global DisparoE As DirectDrawSurface7
Global Balas(1 To 30) As cBala
Global SurBalas() As DirectDrawSurface7
Global Enemigos(1 To 7, 1 To 7) As DirectDrawSurface7
Global TituloUI As DirectDrawSurface7
Global DxPowered As DirectDrawSurface7
Global CreditsTitle(1 To 2) As DirectDrawSurface7
Global Davidgf_logo As DirectDrawSurface7
Global Linkinpark_logo As DirectDrawSurface7
Global Visita_web As DirectDrawSurface7

Global Musica(1 To 4) As DirectMusicSegment
Global Disparos(1 To 5) As DirectSoundBuffer
Global Muertes(1 To 5) As DirectSoundBuffer
Global Explosion(1 To 5) As DirectSoundBuffer
Global VidasCogidas As DirectSoundBuffer
Global EnergiaCogida As DirectSoundBuffer
Global DisparosEnemigos(1 To 5) As DirectSoundBuffer
Global JefeExp As DirectSoundBuffer
Global Vaiven As Byte
Global DisparoID As Integer
Global MuerteVaiven As Integer
Global ExpVaiven As Integer
Global VaivenEnemigo As Integer

Global VidaNave As Integer
Global Vidas As Integer
Global NivelActual As Integer
Global VaivenDE As Integer

Global Seleccion As Integer

Public Const PasosBarra = 43

' ------------------ Datos Constantes ----------------------
Public Const AnchoVida = 102
Public Const AltoVida = 15
Public Const AnMini = 10
Public Const AlMini = 15
Public Const AnchoDisp = 3
Public Const AltoDisp = 8
Public Const DespNave = 10
Public Const PAlto = 600
Public Const PAncho = 800
Public Const AnchoBala = 5
Public Const AltoBala = 15
Public Const AnchoNave = 40
Public Const AltoNave = 61
Public Const AnchoEnemigo = 65
Public Const AltoEnemigo = 47
Public Const AnchoUFO = 65
Public Const AltoUFO = 47
Public Const margenSeguridad = 200
Public Const AnchoJefe = 211
Public Const AltoJefe = 145
'-----------------------------------------------------------

Sub Main()
If App.PrevInstance = True Then End
'----------------------------------------------
EXE = App.Path
If Right(EXE, 1) <> "\" Then EXE = EXE & "\"
TempDir = GetTempDir()

Cargador.Show
Cargador.info = "Cargando ficheros ..."
Cargador.Refresh: DoEvents
DescompactarFichero EXE & "game.dat", TempDir
Cargador.Refresh: DoEvents
'------------------ Niveles ---------
Cargador.info = "Cargando escenarios ...": DoEvents
CargarNiveles
NivelesB = Niveles
Cargador.IncrementarBarra
'---------------- Iniciar DX -----------
Cargador.info = "Iniciando DirectX 7 ...": DoEvents
InitDx
Cargador.IncrementarBarra
'---------------- Iniciar DD -----------
Cargador.info = "Iniciando DirectDraw ...": DoEvents
InitDD
Cargador.IncrementarBarra
'--------------- Iniciar DS -------------
Cargador.info = "Iniciando DirectSound ...": DoEvents
InitDS
Cargador.IncrementarBarra
'-------------------- Cargar Sonido -------
Cargador.info = "Cargando musica ...": DoEvents
LoadMusic
Cargador.info = "Cargando sonido ...": DoEvents
LoadSnd
Cargador.info = "Iniciando ...": DoEvents
'---------------Iniciar DD 2 ---------------
On Local Error Resume Next
Err.Number = 0
SetDisplayMode PAncho, PAlto, 32
If Err.Number <> 0 Then
    Err.Number = 0
    SetDisplayMode PAncho, PAlto, 16
    If Err.Number <> 0 Then
        BorrarTemporales
        Unload Cargador
        MsgBox "No se puede cargar el juego debido a que su pantalla" & vbCrLf & _
        "no soporta la resolución de 800 por 600 píxeles", vbCritical, "UFO Invasion"
        End
    End If
End If
Call CreatePrimaryAndBackBuffer
'---------------- Cargar Imágenes --------
LoadSupers
'---------------Formulario principal------
Unload Cargador
Load Juego
'-------------------------------------------
'CerrarTodosSonidosWin
BorrarTemporales

OtraVez:
Call LimpiarVEstaticas
Call InitGlobals

Call MenuJuego
Call Jugar
If Reentrar = True Then GoTo OtraVez
BorrarTemporales
ShowCursor 1: DoEvents: ShowCursor 1
End
End Sub

Public Sub Jugar()
Dim ready As Boolean
GetAsyncKeyState vbKeyEscape
GetAsyncKeyState vbKeySpace
GetAsyncKeyState vbKeyReturn
Dim unaBala As cBala, z As Integer, Tempo As Single, w As Integer, k As Integer
Dim unaBala2 As cBala
Dim Vx As Integer, Vy As Integer, Iguala As Integer
Dim Vx2 As Integer, Vy2 As Integer

FPS = 50  'y fps, por tanto cada frame dura ...
TimePerFrame = 1000 / FPS
PotenciaActual = 1: MusicaActual = 0
NivelActual = 0
VidaNave = 4: Vidas = 3
MuerteVaiven = 1: OsciladorArma = True: FrameABala = 0
BorrarTodo
TiempoDDM = -1    'valor k indica k la nave esta mostrandose en la pantalla
'-------------------------------------------------------
Dim DestRect As RECT
Dim SrcRect As RECT
Dim subTempo As Single, subTimerMF As Single, tempoDif As Single, CLMusica As Long
Nave.x = (PAncho - AnchoNave) / 2
Nave.y = PAlto - AltoNave - 20
DoEvents: ShowCursor 0: DoEvents
Do While 1
    'DoEvents
    Tempo = GetTickCount()
    
    If ready = True Then
        Call ChequearInmunidad
        
        Call ControlFondo
        If PasoFondo < 1 Or FondoA < 0 Then PasoFondo = 1
        ' ----------------- Cargar nivel si es necesario ---------
        Call JCargarNivel
        ' ------------------------- Mostrar frases si existe --------------
        If NivelActual > 0 Then
        If Niveles(NivelActual).frase <> "" Then
            MostrandoFrase = True
            TimerMF = GetTickCount()
            Niveles(NivelActual).frase = ""
        End If
        End If
        
        Call Objetos_Nave

        If MostrandoFrase = False Then
            ' ----------- Procesar movimientos jefe -----------
            If Niveles(NivelActual).Jefe = True Then Call Jefe_Nave
            If UBound(METs) <> 0 Then
                Call Nave_Meteoro   'Choque Nave-meteorito
                Call Mover_Met   'Procesar meteoritos
            End If
            ' ----------- Procesar UFOs parar que se muevan-------
            If UBound(UFOs) <> 0 Then Call Procesa_UFOS
        End If
    
        '----------------- Procesar Coords. Nave ----------
        If TiempoDDM <= 0 Then
            Call Mover_Nave
            If MostrandoFrase = False Then Call Nave_UFO
        End If
        ' --------------- Disparos enemigos -----------------
        If UBound(DisparosE) <> 0 Then Call DispEnem_Nave_y_Mover
        If UBound(UFOs) <> 0 And MostrandoFrase = False Then Call Procesar_DispEnem

    '----------------- Procesar coords. bala -------------------
    If MostrandoFrase = False Then
        If UBound(UFOs) <> 0 Then Call Bala_UFO
        If UBound(UFOs) = 0 Then Call Bala_Met
    End If
    
    If FrameABala > 0 Then FrameABala = FrameABala - 1
    If GetAsyncKeyState(vbKeySpace) <> 0 And FrameABala = 0 And TiempoDDM = -1 And _
                                                            PasoFondo = 1 Then
        Call NuevoDisparo(unaBala, unaBala2)
    End If

    If NivelActual > 0 And MostrandoFrase = False Then
    If Niveles(NivelActual).Jefe Then
        Call Bala_Jefe
    End If
    End If
    ' ---------------------------- Musica --------------------
    If MusicaActual = 0 Then
        MusicaActual = Int(Rnd * 4) + 1
        'ReproducirSonidoWin "ufo_invasion_10_music_" & Trim(Str(MusicaActual))
        ReproducirMidi Musica(MusicaActual)
        tiempoMusica = GetTickCount() + 5000
    Else
        'If Parado("ufo_invasion_10_music_" & Trim(Str(MusicaActual))) = True Then
        If ReproduciendoMidi(Musica(MusicaActual)) = False And _
        GetTickCount() > tiempoMusica Then
            MusicaActual = MusicaActual + 1
            If MusicaActual > 4 Then MusicaActual = 1
            'ReproducirSonidoWin "ufo_invasion_10_music_" & Trim(Str(MusicaActual))
            ReproducirMidi Musica(MusicaActual)
            tiempoMusica = GetTickCount() + 5000
        End If
    End If
    ' ---------------- Comprovar frase (desaparecer) ---------------
    If MostrandoFrase = True Then
        If (TimerMF + Niveles(NivelActual).TiempoFrase) < GetTickCount() Then
            MostrandoFrase = False
        End If
        If Niveles(NivelActual).especial2 = True And Niveles(NivelActual).TiempoA = False Then
            Niveles(NivelActual).TiempoA = True
            TimerMF = TimerMF + 2000
        End If
        If DibujandoFrase = True Then MostrandoFrase = True
        If DibujandoFrase = True And (Niveles(NivelActual).especial = True Or _
        Niveles(NivelActual).especial2 = True) Then MostrandoFrase = True
        If (Niveles(NivelActual).especial = False And _
        Niveles(NivelActual).especial2 = False) And DibujandoFrase = True Then TimerMF = GetTickCount()
    End If
End If    'se cierra el if del ready PULSE INTRO

    '-------------------- DirecDraw Dibuja ----------------------
    SetRect DestRect, 0, 0, PAncho, PAlto
    SetRect SrcRect, 0, 0, 1, 1
    BackBuffer.Blt DestRect, Fondo, SrcRect, DDBLT_WAIT
    
    If ready = True Then DibujaFondo Fondos(1)
    
If ready = True Then   'segunda condicion del ready
    DibujarExplosiones
    If unaVida.tag = "1" Then
        If unaVida.tagint >= 5 - (30 / FPS) Then
            unaVida.tagint2 = 1
        ElseIf unaVida.tagint <= 1 Then
            unaVida.tagint2 = 0
        End If
        If unaVida.tagint2 = 0 Then
            unaVida.tagint = unaVida.tagint + 1 * (30 / FPS)
        Else
            unaVida.tagint = unaVida.tagint - 1 * (30 / FPS)
        End If

        DibujaAdv unaVida.x, unaVida.y, 60, 60, AnimVidas(CLng(unaVida.tagint + 0.5))
    End If
    If unaEnergia.tag = "1" Then
        unaEnergia.tagint = unaEnergia.tagint + 1
        If unaEnergia.tagint >= (10 * (FPS / 30)) Then unaEnergia.tagint = 1
        If unaEnergia.tagint <= (5 * (FPS / 30)) Then
            unaEnergia.tagint2 = 1
        Else
            unaEnergia.tagint2 = 2
        End If
        DibujaAdv unaEnergia.x, unaEnergia.y, 45, 45, AnimEnergias(unaEnergia.tagint2)
    End If
    If unaMejora.tag = "1" Then
        unaMejora.tagint = unaMejora.tagint + 1 * (30 / FPS)
        If unaMejora.tagint > 12 Then unaMejora.tagint = 1 * (30 / FPS)
        DibujaAdv unaMejora.x, unaMejora.y, 33, 35, AnimMejoras(CLng(unaMejora.tagint + 0.5))
    End If
    For z = 1 To UBound(DisparosE)
        Dibuja DisparosE(z).x, DisparosE(z).y, AnchoDisp, AltoDisp, DisparoE
    Next
    If TiempoDDM = -1 Then
        If KDerecha = True Then
            Dibuja Nave.x, Nave.y, AnchoNave, AltoNave, Nave.derecha
        ElseIf KIzquierda = True Then
            Dibuja Nave.x, Nave.y, AnchoNave, AltoNave, Nave.izquierda
        Else
            Dibuja Nave.x, Nave.y, AnchoNave, AltoNave, Nave.normal
        End If
        If Inmunidad <> 0 Then
            BackBuffer.setDrawWidth 1
            BackBuffer.SetForeColor RGB(Int(Rnd * 50), Int(Rnd * 50), 255)
            If VaivenInmunidad <= 0 Then
                VaivenInmunidad = 10 * (FPS / 30)
            Else
                If VIAumentar = True Then
                    VaivenInmunidad = VaivenInmunidad + 1
                    If VaivenInmunidad > 10 Then VIAumentar = False
                Else
                    VaivenInmunidad = VaivenInmunidad - 1
                    If VaivenInmunidad < 5 Then VIAumentar = True
                End If
            End If
            BackBuffer.SetFillStyle 1
            BackBuffer.DrawCircle (Nave.x + AnchoNave / 2), (Nave.y + AltoNave / 2), AltoNave / 2 + (VaivenInmunidad / (FPS / 30)) - 1
            BackBuffer.DrawCircle (Nave.x + AnchoNave / 2), (Nave.y + AltoNave / 2), AltoNave / 2 + VaivenInmunidad / (FPS / 30)
            BackBuffer.SetForeColor 16744508
            BackBuffer.DrawCircle (Nave.x + AnchoNave / 2), (Nave.y + AltoNave / 2), AltoNave / 2 + VaivenInmunidad / (FPS / 30) + 2
        End If
    End If
    
    If MostrandoFrase = False And NivelActual > 0 Then
        If Niveles(NivelActual).Jefe Then
            Framejefe = Framejefe + 0.1 * (30 / FPS)
            If Framejefe >= 5 Then Framejefe = 0.25 * (30 / FPS)
            If CLng(Framejefe + 0.5) > 3 Then
                DibujaJefe Jefe.x, Jefe.y, AnchoJefe, AltoJefe, MatrizJefes(Val(Niveles(NivelActual).key), CLng(Framejefe + 0.5) - 2)
            Else
                DibujaJefe Jefe.x, Jefe.y, AnchoJefe, AltoJefe, MatrizJefes(Val(Niveles(NivelActual).key), CLng(Framejefe + 0.5))
            End If
        End If
        For z = 1 To UBound(UFOs)
            DibujaAdv UFOs(z).x, UFOs(z).y, AnchoUFO, AltoUFO, Enemigos(UFOs(z).tipo, (UFOs(z).tipo - UFOs(z).vida + 1))
        Next
        For z = 1 To UBound(METs)
            DibujaAdv METs(z).x, METs(z).y, 50, 50, Meteoritos(Int(METs(z).frame))
        Next
    End If
    
    For z = 1 To 30
        If Balas(z) Is Nothing Then
        Else
            Balas(z).y = Balas(z).y - (15 * (30 / FPS))    'Aprovechamos para eliminar las balas k salen
            DibujaAdv Balas(z).x, Balas(z).y, AnchoBala, AltoBala, SurBalas(Balas(z).potencia)
            If Balas(z).y < -30 Then
                Set Balas(z) = Nothing
            End If
        End If
    Next
    
    If NivelActual <> 0 Then
    If Niveles(NivelActual).Jefe Then
        BackBuffer.SetForeColor vbWhite
        BackBuffer.SetFillStyle 1
        BackBuffer.DrawRoundedBox PAncho - 25 - 5 - 1, 5 - 1, PAncho - 5 + 1 + 1, 5 + 300 + 2, 8, 8
        BackBuffer.SetFillColor RGB(255, 128, 0)  'naranja
        w = 3 + ((300) / Jefe.vidamax) * Jefe.vida
        BackBuffer.SetFillStyle 0
        BackBuffer.SetForeColor RGB(255, 128, 0)
        If Jefe.vida > 0 Then BackBuffer.DrawRoundedBox PAncho - 25 - 5 + 3, 5 + 3, PAncho - 5 - 2, w, 8, 8
    End If
    End If
    
    DibujarExplosionesP
    'Dibuja 5, 5, AnchoVida, AltoVida, BarraVida(VidaNave)
    
    BackBuffer.SetForeColor RGB(0, 75, 0)
    BackBuffer.SetFillStyle 1
    BackBuffer.DrawRoundedBox 6, 6, 104, 19, 3, 3
    BackBuffer.SetFillColor RGB(0, 150, 0)  'verde
    BackBuffer.SetFillStyle 0
    BackBuffer.SetForeColor RGB(0, 150, 0)
    BackBuffer.DrawRoundedBox 8, 8, (25.4 * VidaNave), 17, 0, 0
    
    For z = 1 To Vidas
        Dibuja 117 + (z - 1) * (5 + AnMini), 5, AnMini, AlMini, Mini
    Next
    If MostrandoFrase = True And NivelActual <> 0 Then
        If Niveles(NivelActual).frase2 <> "" Then
            DibujarFrase 50, 600 - 50, Niveles(NivelActual).frase2
        Else
            DibujarFraseRC 50, 600 - 50, Niveles(NivelActual).frasesecundaria
        End If
    End If
End If    'cierre del segundo ready
    
    ' ---------- Inicio -----------------
    If ready = False Then
        DibujarFrase -1, -1, "Pulse intro"
    End If
    '----------------------------------
    
    Mostrar
    DoEvents
    
    If (TimePerFrame - (GetTickCount() - Tempo)) > 0 Then
    If (TimePerFrame - (GetTickCount() - Tempo)) <= TimePerFrame Then
        'Esperar (TimePerFrame - (GetTickCount() - Tempo))
        Sleep (TimePerFrame - (GetTickCount() - Tempo))   'Se "duerme" el tiempo sobrante para regular la velocidad
    End If
    End If
    
    'Tiempo empleado para el frame=GetTickCount() - Tempo)
    'Tiempo que ha de durar cada frametimeperframe
    'por tanto el resto será el tiempo "muerto"
    If KEscape = True Or SalidaCritica = True Then GoTo SALIDA
    If KIntro = True Then ready = True
    If JPausado Then
        subTimerMF = TimerMF - GetTickCount()
        subTempo = Tempo - GetTickCount()
        PararMidi Musica(MusicaActual)
        CLMusica = DMusicPerformance.GetMusicTime()
    End If
    Do While JPausado
        DoEvents
        Sleep TimePerFrame
        If GetAsyncKeyState(vbKeyEscape) Then GoTo SALIDA
    Loop
    If subTempo <> 0 Then
        tiempoMusica = GetTickCount() + 5000
        TimerMF = subTimerMF + GetTickCount()
        Tempo = subTempo + GetTickCount()
        subTempo = 0
        ReproducirMidi Musica(MusicaActual), CLMusica
    End If
Loop
SALIDA:
    'PararSonidoWin "ufo_invasion_10_music_" & Trim(Str(MusicaActual))
    On Local Error Resume Next
    PararMidi Musica(MusicaActual)
    'RestoreDisplayMode
    'DestruirDS
    'Set View = Nothing
    'Set BackBuffer = Nothing
    'DDraw.SetCooperativeLevel Juego.hWnd, DDSCL_NORMAL
    'Set DDraw = Nothing: Set Dx = Nothing
    'Unload Juego
    'ShowCursor 1
    DoEvents
    Reentrar = True
    If SalidaCritica = True Then End
End Sub

Public Sub LoadSupers()
Dim Color As DDCOLORKEY, jj As Integer
Color.high = vbBlack
Color.low = vbBlack
ReDim SurBalas(1 To 7)
LoadS TempDir & "NR.bmp", Nave.normal, AnchoNave, AltoNave
LoadS TempDir & "NI.bmp", Nave.izquierda, AnchoNave, AltoNave
LoadS TempDir & "ND.bmp", Nave.derecha, AnchoNave, AltoNave
LoadS TempDir & "black.bmp", Fondo, PAncho, PAlto
LoadS TempDir & "disp.bmp", DisparoE, AnchoDisp, AltoDisp
LoadS TempDir & "mini.bmp", Mini, AnMini, AlMini
Mini.SetColorKey DDCKEY_SRCBLT, Color
DisparoE.SetColorKey DDCKEY_SRCBLT, Color
Nave.normal.SetColorKey DDCKEY_SRCBLT, Color
Nave.izquierda.SetColorKey DDCKEY_SRCBLT, Color
Nave.derecha.SetColorKey DDCKEY_SRCBLT, Color
Dim j As Integer, j2 As Integer
For j = 1 To 7
    LoadS TempDir & "bala" & Trim(Str(j)) & ".bmp", SurBalas(j), AnchoBala, AltoBala
    SurBalas(j).SetColorKey DDCKEY_SRCBLT, Color
Next
On Local Error Resume Next
For j = 1 To 7
For jj = 1 To 7
    If Existe(TempDir & "enemigo" & Trim(Str(j)) & "-" & Trim(Str(jj)) & ".bmp") Then
        LoadS TempDir & "enemigo" & Trim(Str(j)) & "-" & Trim(Str(jj)) & ".bmp", Enemigos(j, jj), AnchoEnemigo, AltoEnemigo
        Enemigos(j, jj).SetColorKey DDCKEY_SRCBLT, Color
    End If
Next
Next
Err.Clear: Err.Number = 0
On Local Error GoTo 0
For j = 1 To 11
    LoadS TempDir & "e" & Trim(Str(j)) & ".bmp", AnimExplosion(j), 64, 64
    AnimExplosion(j).SetColorKey DDCKEY_SRCBLT, Color
Next
LoadS TempDir & "fondo1.bmp", Fondos(1), 800, 3200
Fondos(1).SetColorKey DDCKEY_SRCBLT, Color

LoadS TempDir & "davidgf_logo.bmp", Davidgf_logo, 309, 71
Davidgf_logo.SetColorKey DDCKEY_SRCBLT, Color
LoadS TempDir & "linkin.bmp", Linkinpark_logo, 312, 73
Linkinpark_logo.SetColorKey DDCKEY_SRCBLT, Color
LoadS TempDir & "credits1.bmp", CreditsTitle(1), 367, 36
CreditsTitle(1).SetColorKey DDCKEY_SRCBLT, Color
LoadS TempDir & "credits2.bmp", CreditsTitle(2), 96, 31
CreditsTitle(2).SetColorKey DDCKEY_SRCBLT, Color
LoadS TempDir & "dx7logo.bmp", DxPowered, 191, 104
DxPowered.SetColorKey DDCKEY_SRCBLT, Color
LoadS TempDir & "visita.bmp", Visita_web, 279, 37
Visita_web.SetColorKey DDCKEY_SRCBLT, Color

For j = 1 To 8
    LoadS TempDir & "met" & Trim(Str(j)) & ".bmp", Meteoritos(j), 50, 50
    Meteoritos(j).SetColorKey DDCKEY_SRCBLT, Color
Next
LoadS TempDir & "letras.bmp", Letras, 600, 20
Letras.SetColorKey DDCKEY_SRCBLT, Color
For j = 1 To 5
    LoadS TempDir & "vida" & Trim(Str(j)) & ".bmp", AnimVidas(j), 60, 60
    AnimVidas(j).SetColorKey DDCKEY_SRCBLT, Color
Next
For j = 1 To 12
    LoadS TempDir & "u" & Trim(Str(j)) & ".bmp", AnimMejoras(j), 33, 35
    AnimMejoras(j).SetColorKey DDCKEY_SRCBLT, Color
Next
For j = 1 To 2
    LoadS TempDir & "energia" & Trim(Str(j)) & ".bmp", AnimEnergias(j), 45, 45
    AnimEnergias(j).SetColorKey DDCKEY_SRCBLT, Color
Next
For j = 1 To 1
  For j2 = 1 To 3
    LoadS TempDir & "jefe" & Trim(Str(j)) & Trim(Str(j2)) & ".bmp", MatrizJefes(j, j2), AnchoJefe, AltoJefe
    MatrizJefes(j, j2).SetColorKey DDCKEY_SRCBLT, Color
  Next
Next
LoadS TempDir & "titulo.bmp", TituloUI, 402, 86
TituloUI.SetColorKey DDCKEY_SRCBLT, Color
LoadS TempDir & "davidgf.bmp", Autor, 206, 20
Autor.SetColorKey DDCKEY_SRCBLT, Color
LoadS TempDir & "web.bmp", Web, 156, 20
Web.SetColorKey DDCKEY_SRCBLT, Color
For j = 1 To 4
For jj = 1 To 2
    LoadS TempDir & "menu" & Trim(Str(j)) & "-" & Trim(Str(jj)) & ".bmp", Menus(j, jj), 251, 61
    Menus(j, jj).SetColorKey DDCKEY_SRCBLT, Color
Next
Next
End Sub

Public Sub LoadSnd()
Dim d As Integer
For d = 1 To 5
    Set Disparos(d) = CargarSonido(TempDir & "snd1.wav")
    Cargador.IncrementarBarra
Next
For d = 1 To 5
    Set Muertes(d) = CargarSonido(TempDir & "snd2.wav")
    Cargador.IncrementarBarra
Next
For d = 1 To 5
    Set Explosion(d) = CargarSonido(TempDir & "snd3.wav")
    Cargador.IncrementarBarra
Next
For d = 1 To 5
    Set DisparosEnemigos(d) = CargarSonido(TempDir & "snd4.wav")
    Cargador.IncrementarBarra
Next
Set VidasCogidas = CargarSonido(TempDir & "snd5.wav")
Set EnergiaCogida = CargarSonido(TempDir & "snd6.wav")
Set JefeExp = CargarSonido(TempDir & "snd7.wav")
End Sub

Public Sub LoadMusic()
InitDM
Dim d As Integer, j As Integer
For d = 1 To 4
    Set Musica(d) = CargarMidi(TempDir & Trim(Str(d)) & ".mid")
    For j = 1 To 5
        Cargador.IncrementarBarra
    Next
Next
End Sub

Public Sub DestruirDS()
Dim d As Integer
For d = 1 To 5
    Set Disparos(d) = Nothing
Next
For d = 1 To 5
    Set Muertes(d) = Nothing
Next
For d = 1 To 5
    Set Explosion(d) = Nothing
Next
Set DSound = Nothing
PararMidi Musica(MusicaActual)
DMusicPerformance.CloseDown
Set DMusicLoader = Nothing
Set DMusicPerformance = Nothing
Set estado = Nothing
End Sub

Public Sub CargarNiveles()
Dim libre As Integer, lin As String, str1 As String, str2 As String, int1 As Integer
Dim f As Integer
libre = FreeFile
Open EXE & "Data.dat" For Input As #libre
Do While Not EOF(libre)
    f = f + 1
    Line Input #libre, lin
    ReDim Preserve Niveles(1 To f)
    Niveles(f).key = lin
    Line Input #libre, lin
    lin = UCase(lin)
    If InStr(1, lin, "U") <> 0 Then Niveles(f).upgrade = True
    If InStr(1, lin, "E") <> 0 Then Niveles(f).energia = True
    If InStr(1, lin, "V") <> 0 Then Niveles(f).vida = True
    If InStr(1, lin, "S") <> 0 Then Niveles(f).especial = True
    If InStr(1, lin, "J") <> 0 Then Niveles(f).Jefe = True
    If InStr(1, lin, "Z") <> 0 Then Niveles(f).especial2 = True
    Line Input #libre, lin
    If lin <> "" Then
        int1 = InStr(1, lin, ":")
        str1 = Left(lin, int1 - 1)
        str2 = Mid(lin, int1 + 1)
        Niveles(f).frase = str1
        Niveles(f).frase2 = str1
        Niveles(f).TiempoFrase = Val(str2)
        int1 = InStr(1, lin, "\")
        If int1 <> 0 Then
            str2 = Mid(lin, int1 + 1)
            Niveles(f).frasesecundaria = str2
        End If
    End If
Loop
Close #libre
End Sub

Public Sub DibujarExplosiones()
Dim g As Integer, h As Integer
Inicio:
For g = 1 To UBound(ExpActivas)
    If ExpActivas(g).frame >= (11 - (30 / FPS)) Then
        If UBound(ExpActivas) = 1 Then
            ReDim ExpActivas(0)
        Else
            If g <> UBound(ExpActivas) Then
                For h = g To UBound(ExpActivas) - 1
                    ExpActivas(h).x = ExpActivas(h + 1).x
                    ExpActivas(h).y = ExpActivas(h + 1).y
                    ExpActivas(h).frame = ExpActivas(h + 1).frame
                Next
            End If
            ReDim Preserve ExpActivas(UBound(ExpActivas) - 1)
            GoTo Inicio
        End If
    End If
Next
For g = 1 To UBound(ExpActivas)
    ExpActivas(g).frame = ExpActivas(g).frame + (1 * (30 / FPS))
    If ExpActivas(g).frame > 0 Then DibujaAdv ExpActivas(g).x, ExpActivas(g).y, 64, 64, AnimExplosion(CLng(ExpActivas(g).frame + 0.5))
Next
End Sub

Public Sub DibujarExplosionesP()
Dim g As Integer, h As Integer
Inicio:
For g = 1 To UBound(ExpActivasP)
    If ExpActivasP(g).frame >= (10 - (30 / FPS)) Then
        If UBound(ExpActivasP) = 1 Then
            ReDim ExpActivasP(0)
        Else
            If g <> UBound(ExpActivasP) Then
                For h = g To UBound(ExpActivasP) - 1
                    ExpActivasP(h).x = ExpActivasP(h + 1).x
                    ExpActivasP(h).y = ExpActivasP(h + 1).y
                    ExpActivasP(h).frame = ExpActivasP(h + 1).frame
                Next
            End If
            ReDim Preserve ExpActivasP(UBound(ExpActivasP) - 1)
            GoTo Inicio
        End If
    End If
Next
For g = 1 To UBound(ExpActivasP)
    ExpActivasP(g).frame = ExpActivasP(g).frame + (2 * (30 / FPS))
    If ExpActivasP(g).frame > 0 Then Dibuja ExpActivasP(g).x, ExpActivasP(g).y, 24, 24, AnimExplosion(CLng(ExpActivasP(g).frame + 0.5))
Next
End Sub

Public Sub NuevaExplosion(ByVal x As Integer, ByVal y As Integer, Optional ByVal FrameInicial As Integer = 0)
ReDim Preserve ExpActivas(UBound(ExpActivas) + 1)
ExpActivas(UBound(ExpActivas)).x = x
ExpActivas(UBound(ExpActivas)).y = y
ExpActivas(UBound(ExpActivas)).frame = FrameInicial
End Sub

Public Sub NuevaExplosionP(ByVal x As Integer, ByVal y As Integer, Optional ByVal FrameInicial As Integer = 0)
ReDim Preserve ExpActivasP(UBound(ExpActivasP) + 1)
ExpActivasP(UBound(ExpActivasP)).x = x
ExpActivasP(UBound(ExpActivasP)).y = y
ExpActivasP(UBound(ExpActivasP)).frame = FrameInicial
End Sub

Public Sub DibujaFondo(obj As DirectDrawSurface7)
Dim recorteR As Integer
If CLng(S_recorte - PasoFondo * (FPS / 30)) <= 0 Then
    S_recorte = 2600
Else
    S_recorte = S_recorte - PasoFondo * (30 / FPS)
End If
recorteR = Int(S_recorte)

Dim SrcRect As RECT
SetRect SrcRect, 0, recorteR, PAncho, PAlto
BackBuffer.BltFast 0, 0, obj, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub

Public Sub BorrarTodo()
Dim z As Integer
ReDim UFOs(0): ReDim DisparosE(0): ReDim ExpActivas(0): ReDim METs(0): ReDim ExpActivasP(0)
Inmunidad = 0: TiempoDDM = 0: VaivenEnemigo = 0
KArriba = False: KAbajo = False: KDerecha = False: KIzquierda = False
For z = 1 To 30
    Set Balas(z) = Nothing
Next
End Sub

Public Sub DibujarCaracter(x As Integer, y As Integer, ByVal caracter As String)
'On Local Error Resume Next
caracter = UCase(caracter)
Dim SrcRect As RECT
Dim z As Integer
If Asc(caracter) = 33 Then
    z = 27
ElseIf Asc(caracter) = 161 Then
    z = 28
ElseIf Asc(caracter) = 32 Then
    z = 29
ElseIf Asc(caracter) = 45 Then
    z = 30
ElseIf Asc(caracter) >= 65 And Asc(caracter) <= 90 Then
    z = Asc(caracter) - 64
Else
    z = 27
End If
SetRect SrcRect, (z - 1) * 20, 0, 20, 20
BackBuffer.BltFast x, y, Letras, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
End Sub

Public Sub DibujarFrase(x As Integer, y As Integer, ByVal frase As String)
'On Local Error Resume Next
Dim hl As Integer
If S1_fraseanterior = "" Then S1_fraseanterior = frase: S1_wr = 0
If S1_fraseanterior <> frase Then S1_fraseanterior = frase: S1_wr = 0
S1_wr = S1_wr + 0.2 * (50 / FPS)
S1_cursor = S1_cursor + 1
If S1_cursor >= (10 * (FPS / 50)) Then
    S1_cursor = -10 * (FPS / 50)
End If
Dim h As Integer, x2 As Integer, y2 As Integer
If x < 0 Then
    x2 = ((PAncho - (Len(frase) * 20)) / 2)
Else
    x2 = x
End If
If y < 0 Then
    y2 = (PAlto - 20) / 2
Else
    y2 = y
End If
If S1_wr > Len(frase) Then
    S1_wr = Len(frase)
    DibujandoFrase = False
Else
    DibujandoFrase = True
End If
For h = 0 To (S1_wr - 1)
    hl = Int(h) + 1
    DibujarCaracter x2 + ((hl - 1) * 20), y2, Mid(frase, hl, 1)
Next
If S1_cursor < 0 Then
    h = S1_wr
    hl = Int(h) + 1
    DibujarCaracter x2 + ((hl - 1) * 20), y2, "-"
End If
End Sub

Public Function ComSepUFOs() As Boolean
Dim d As Integer
For d = 1 To UBound(UFOs) - 1
    If (UFOs(d + 1).x - UFOs(d).x) > (65 * 3 / 2) Then
        ComSepUFOs = True
        Exit Function
    End If
Next
End Function

Public Function ComSep2UFOs(ByVal ufo1 As Integer, ufo2 As Integer) As Boolean
If (UFOs(ufo2).x - UFOs(ufo1).x) > (65 * 3 / 2) Then
    ComSep2UFOs = True
End If
End Function

Public Function ComSep2UFOsLargo(ByVal ufo1 As Integer, ufo2 As Integer) As Boolean
If (UFOs(ufo2).x - UFOs(ufo1).x) > (65 * 2) Then
    ComSep2UFOsLargo = True
End If
End Function

Public Function ComSepUFOsLargo() As Boolean
'Comprueba las separacion entre UFOS, la utiliza
'la funcion Juntar UFOS
Dim d As Integer
For d = 1 To UBound(UFOs) - 1
    If (UFOs(d + 1).x - UFOs(d).x) > (65 * 6 / 5) Then
        ComSepUFOsLargo = True
        Exit Function
    End If
Next
End Function

Public Sub JuntarUFOS()
' La funcion 'magica' k junta los ufos si kedan wecos entre ellos.
'Muy larga pork admite todas las posibilidades
Select Case UBound(UFOs)
Case 6
    If ComSepUFOsLargo() = True Then
        If ComSep2UFOsLargo(5, 6) Then
            UFOs(6).x = UFOs(6).x - Distancia2UFOs(5, 6)
        End If
        If ComSep2UFOs(1, 2) Then
            UFOs(1).x = UFOs(1).x + Distancia2UFOs(1, 2)
        End If
        If ComSep2UFOs(2, 3) Then
            UFOs(1).x = UFOs(1).x + Distancia2UFOs(2, 3)
            UFOs(2).x = UFOs(2).x + Distancia2UFOs(2, 3)
        End If
        If ComSep2UFOs(4, 5) Then
            UFOs(6).x = UFOs(6).x - Distancia2UFOs(4, 5)
            UFOs(5).x = UFOs(5).x - Distancia2UFOs(4, 5)
        End If
        If ComSep2UFOs(3, 4) Then
            UFOs(1).x = UFOs(1).x + Distancia2UFOs(3, 4)
            UFOs(2).x = UFOs(2).x + Distancia2UFOs(3, 4)
            UFOs(3).x = UFOs(3).x + Distancia2UFOs(3, 4)
            UFOs(6).x = UFOs(6).x - Distancia2UFOs(3, 4)
            UFOs(5).x = UFOs(5).x - Distancia2UFOs(3, 4)
            UFOs(4).x = UFOs(4).x - Distancia2UFOs(3, 4)
        End If
    End If
Case 5
    If ComSepUFOsLargo() = True Then
        If ComSep2UFOsLargo(4, 5) Then
            UFOs(5).x = UFOs(5).x - Distancia2UFOs(4, 5)
        End If
        If ComSep2UFOs(1, 2) Then
            UFOs(1).x = UFOs(1).x + Distancia2UFOs(1, 2)
        End If
        If ComSep2UFOs(3, 4) Then
            UFOs(4).x = UFOs(4).x - Distancia2UFOs(3, 4)
            UFOs(5).x = UFOs(5).x - Distancia2UFOs(3, 4)
        End If
        If ComSep2UFOs(2, 3) Then
            UFOs(1).x = UFOs(1).x + Distancia2UFOs(2, 3)
            UFOs(2).x = UFOs(2).x + Distancia2UFOs(2, 3)
        End If
    End If
Case 4
    If ComSepUFOs() = True Then
        If ComSep2UFOs(2, 3) Then
            UFOs(1).x = UFOs(1).x + Distancia2UFOs(2, 3)
            UFOs(2).x = UFOs(2).x + Distancia2UFOs(2, 3)
            UFOs(3).x = UFOs(3).x - Distancia2UFOs(2, 3)
            UFOs(4).x = UFOs(4).x - Distancia2UFOs(2, 3)
        End If
        If ComSep2UFOs(3, 4) Then
            UFOs(4).x = UFOs(4).x - Distancia2UFOs(3, 4)
        End If
        If ComSep2UFOs(1, 2) Then
            UFOs(1).x = UFOs(1).x + Distancia2UFOs(1, 2)
        End If
    End If
Case 3
    If ComSepUFOs() = True Then
        If ComSep2UFOs(2, 3) Then
            UFOs(3).x = UFOs(3).x - Distancia2UFOs(2, 3)
        End If
        If ComSep2UFOs(1, 2) Then
            UFOs(1).x = UFOs(1).x + Distancia2UFOs(1, 2)
        End If
    End If
Case 2
    If ComSep2UFOs(1, 2) Then
        UFOs(1).x = UFOs(1).x + Distancia2UFOs(1, 2)
    End If
End Select
End Sub

Public Function Distancia2UFOs(ByVal ufo1 As Integer, ufo2 As Integer) As Integer
Dim integer1 As Integer
integer1 = UFOs(ufo2).x - UFOs(ufo1).x - (65 * 2)
If integer1 <= 0 Then integer1 = UFOs(ufo2).x - UFOs(ufo1).x - 65
Distancia2UFOs = ((8 * integer1) / 165) * (30 / FPS)
If Distancia2UFOs < 1 Then Distancia2UFOs = 1
'Regla de proporción para calcular la velcidad
End Function

Public Sub DibujaJefe(x As Integer, y As Integer, ancho As Integer, alto As Integer, obj As DirectDrawSurface7, Optional ByVal SinT As Boolean = False)
Dim recorte As Integer
Dim SrcRect As RECT
If y < 0 Then
    recorte = -y
    recorte = AltoJefe + y
    SetRect SrcRect, 0, -y, ancho, recorte
    BackBuffer.BltFast x, 0, obj, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
Else
    Dibuja x, y, ancho, alto, obj, SinT
End If
End Sub


Public Sub BorrarTemporales()
On Local Error Resume Next
If TempDir <> EXE Then
    Kill TempDir & "*.bmp"
    Kill TempDir & "*.mid"
    Kill TempDir & "*.wav"
    RmDir TempDir
End If
End Sub

Public Function GetTempDir() As String
Dim cadena As String
cadena = Space(512)
GetTempPath 512, cadena
cadena = Left(cadena, InStr(1, cadena, Chr(0)) - 1)
cadena = cadena & "ui"
10
cadena = cadena & Int(Rnd * 9)
If ExisteDir(cadena) = True Then GoTo 10
cadena = cadena & "\"
GetTempDir = cadena
End Function

Public Function ExisteDir(ByVal Fic As String) As Boolean
Dim cadena As String
cadena = Dir(Fic, vbDirectory)
If cadena <> "" Then ExisteDir = True
End Function

Public Sub DibujarFraseRapido(x As Integer, y As Integer, ByVal frase As String)
'On Local Error Resume Next
Dim h As Integer, x2 As Integer, y2 As Integer
If x < 0 Then
    x2 = ((PAncho - (Len(frase) * 20)) / 2)
Else
    x2 = x
End If
If y < 0 Then
    y2 = (PAlto - 20) / 2
Else
    y2 = y
End If
For h = 0 To Len(frase)
    DibujarCaracter x2 + ((h - 1) * 20), y2, Mid(frase, h, 1)
Next
End Sub

Public Sub DibujarFraseRC(x As Integer, y As Integer, ByVal frase As String)
'On Local Error Resume Next
Dim hl As Integer
If S2_fraseanterior = "" Then S2_fraseanterior = frase: S2_wr = 0
If S2_fraseanterior <> frase Then S2_fraseanterior = frase: S2_wr = 0
S2_wr = S2_wr + 0.8 * (50 / FPS)
S2_cursor = S2_cursor + 1
If S2_cursor >= (10 * (FPS / 50)) Then
    S2_cursor = -10 * (FPS / 50)
End If
Dim h As Integer, x2 As Integer, y2 As Integer
If x < 0 Then
    x2 = ((PAncho - (Len(frase) * 20)) / 2)
Else
    x2 = x
End If
If y < 0 Then
    y2 = (PAlto - 20) / 2
Else
    y2 = y
End If
If S2_wr > Len(frase) Then
    S2_wr = Len(frase)
    DibujandoFrase = False
Else
    DibujandoFrase = True
End If
For h = 0 To (S2_wr - 1)
    hl = Int(h) + 1
    DibujarCaracter x2 + ((hl - 1) * 20), y2, Mid(frase, hl, 1)
Next
If S2_cursor < 0 Then
    h = S2_wr
    hl = Int(h) + 1
    DibujarCaracter x2 + ((hl - 1) * 20), y2, "-"
End If
End Sub

Public Sub MenuJuego()
InicioMenu:
KEscape = False
Randomize Timer
On Local Error Resume Next
Dim OpcionElegida As Byte
FPS = 25  'y fps, por tanto cada frame dura ...
TimePerFrame = 1000 / FPS
Dim SalirM As Boolean, Tempo As Long
Dim DestRect As RECT
Dim SrcRect As RECT, x As Integer
OpcionElegida = 1
ShowCursor 0
Do While SalirM = False
    Tempo = GetTickCount()
    If KEscape = True Then GoTo SALIR5
    If KAbajo = True Then
        KAbajo = False
        OpcionElegida = OpcionElegida + 1
        If OpcionElegida > 4 Then OpcionElegida = 1
    End If
    If KArriba = True Then
        KArriba = False
        OpcionElegida = OpcionElegida - 1
        If OpcionElegida = 0 Then OpcionElegida = 4
    End If
    If KIntro = True Then
        KIntro = False
        Select Case OpcionElegida
        Case 1
            Exit Sub
        Case 3
            GoTo SalidaCreditos
        Case 4
            KEscape = True
        End Select
    End If
    DoEvents
    SetRect DestRect, 0, 0, PAncho, PAlto
    SetRect SrcRect, 0, 0, 1, 1
    BackBuffer.Blt DestRect, Fondo, SrcRect, DDBLT_WAIT
    
    Dibuja (PAncho - 402) / 2, 65, 402, 86, TituloUI
    Dibuja PAncho - 206 - 5, PAlto - 25, 206, 20, Autor
    Dibuja 5, PAlto - 25, 156, 20, Web
    
    For x = 1 To 4
        If x = OpcionElegida Then
            Dibuja (PAncho - 236) / 2, 200 + (x - 1) * 61, 251, 61, Menus(x, 2)
        Else
            Dibuja (PAncho - 236) / 2, 200 + (x - 1) * 61, 251, 61, Menus(x, 1)
        End If
    Next
    
    Mostrar
    DoEvents
    
    If (TimePerFrame - (GetTickCount() - Tempo)) > 0 Then
        'Esperar (TimePerFrame - (GetTickCount() - Tempo))
        Sleep (TimePerFrame - (GetTickCount() - Tempo))   'Se "duerme" el tiempo sobrante para regular la velocidad
    End If
Loop

Exit Sub
SALIR5:
    On Local Error Resume Next
    PararMidi Musica(MusicaActual)
    RestoreDisplayMode
    DestruirDS
    Set View = Nothing
    Set BackBuffer = Nothing
    DDraw.SetCooperativeLevel Juego.hWnd, DDSCL_NORMAL
    Set DDraw = Nothing: Set Dx = Nothing
    Unload Juego
    ShowCursor 1
    DoEvents
    End
    
Exit Sub
SalidaCreditos:
    Call MostrarCreditos
GoTo InicioMenu
End Sub

Public Sub MostrarCreditos()
On Local Error Resume Next
Dim Tempo As Single
Dim DestRect As RECT
Dim SrcRect As RECT
FPS = 25  'y fps, por tanto cada frame dura ...
TimePerFrame = 1000 / FPS
Do While 1
    Tempo = GetTickCount()
    If KEscape = True Or KIntro = True Then
        KEscape = False
        KIntro = False
        Exit Sub
    End If
    DoEvents
    SetRect DestRect, 0, 0, PAncho, PAlto
    SetRect SrcRect, 0, 0, 1, 1
    BackBuffer.Blt DestRect, Fondo, SrcRect, DDBLT_WAIT
    
    Dibuja (PAncho - 402) / 2, 65, 402, 86, TituloUI
    
    Dibuja (PAncho - 367) / 2, 190, 367, 36, CreditsTitle(1)
    Dibuja (PAncho - 309) / 2, 240, 309, 71, Davidgf_logo
    
    Dibuja (PAncho - 96) / 2, 340, 96, 31, CreditsTitle(2)
    Dibuja (PAncho - 312) / 2, 380, 312, 73, Linkinpark_logo
    
    Dibuja (PAncho - 191) - 5, (PAlto - 104) - 5, 191, 104, DxPowered
    Dibuja 5, (PAlto - 37) - 5, 279, 37, Visita_web
    
    Mostrar
    DoEvents
    
    If (TimePerFrame - (GetTickCount() - Tempo)) > 0 Then
        'Esperar (TimePerFrame - (GetTickCount() - Tempo))
        Sleep (TimePerFrame - (GetTickCount() - Tempo))   'Se "duerme" el tiempo sobrante para regular la velocidad
    End If
Loop
End Sub


Public Sub LimpiarVEstaticas()
S_recorte = 0
S1_fraseanterior = 0
S1_wr = 0
S1_cursor = 0
S2_fraseanterior = 0
S2_wr = 0
S2_cursor = 0
End Sub

Public Sub InitGlobals()
TempoPrivado = 0
FondoA = 0
PotenciaActual = 0
MusicaActual = 0
MostrandoFrase = False
TimerMF = 0
PasoFondo = 0
Framejefe = 0
EntrandoE = False
MaxPasoFondo = 0
OsciladorArma = False
TempDir = ""
tiempoMusica = 0
JPausado = False
Reentrar = False
Niveles = NivelesB
End Sub


Public Sub ChequearInmunidad()
'  La variable inmunidad contiene un numero de frames (ce acuerdo con el FPS i el tiempo)
'  que indica el tiempo k estara la nave con el escudo de inmunidad
If Inmunidad > 0 Then Inmunidad = Inmunidad - 1    'se disminuye tras cada frame
If TiempoDDM > 0 Then
    TiempoDDM = TiempoDDM - 1    'el Tiempo DDM indica los frames a esperar asta k aparezca
                                 'la nave una vez destruida
ElseIf TiempoDDM = 0 Then
    Nave.x = (PAncho - AnchoNave) / 2   'se situa la nave
    Nave.y = PAlto - AltoNave - 20
    TiempoDDM = -1    'se asinga a -1 porque luego comprovaremos mas adelante
    Inmunidad = (90 * (FPS / 30))   'se asignan unos 3 segundos 150 frames a v. normal
End If
End Sub

Public Sub JCargarNivel()
If UBound(UFOs) = 0 And UBound(METs) = 0 And Jefe.vida <= 0 Then
    NivelActual = NivelActual + 1
    If Niveles(NivelActual).Jefe = True Then
        Jefe.tipo = Val(Niveles(NivelActual).key)
        Jefe.vida = 100 * Jefe.tipo
        Jefe.vidamax = Jefe.vida
        Jefe.x = (PAncho - AnchoJefe) / 2
        Jefe.y = -AltoJefe
        EntrandoE = True
    Else
        For z = 1 To 16
            If Mid(Niveles(NivelActual).key, z, 1) <> "-" Then
                If Mid(Niveles(NivelActual).key, z, 1) = "M" Then
                    ReDim Preserve METs(UBound(METs) + 1)
                    METs(UBound(METs)).x = 50 * (z - 1)
                    METs(UBound(METs)).y = -100
                    METs(UBound(METs)).frame = Int(Rnd * 7) + 1
                    METs(UBound(METs)).a = 14 + Int(Rnd * 6)
                    METs(UBound(METs)).vida = Int(Rnd * 4) + 3
                Else
                    ReDim Preserve UFOs(UBound(UFOs) + 1)
                    UFOs(UBound(UFOs)).tipo = Mid(Niveles(NivelActual).key, z, 1)
                    UFOs(UBound(UFOs)).vida = UFOs(UBound(UFOs)).tipo
                    UFOs(UBound(UFOs)).x = 10 + 65 * (z - 1)
                    UFOs(UBound(UFOs)).y = -AltoUFO
                    'esta salida se produce porque sóo caben 12 ovnis a lo largo
                    'o 16 meteoros. Por eso hacemos el bucle de 16 y forzamos la
                    'salida en caso de ser ovnis lo que viene
                End If
            End If
            If UBound(UFOs) <> 0 And UBound(METs) = 0 And z = 12 Then Exit For
        Next
    End If
End If
End Sub


Public Sub Objetos_Nave()
'---- Miraremos las colisiones de la nave y los objetos beneficiosos -------
' ------------------------- Objeto: Vida ---------------------------
If unaVida.tag = "1" Then
    unaVida.y = unaVida.y + (Int(Rnd * 5) + 7) * (30 / FPS)
    If Nave.x + 20 > unaVida.x And (Nave.x + AnchoNave) < (unaVida.x + 60) + 20 Then
    If Nave.y - 20 < unaVida.y And (Nave.y + AltoNave) > (unaVida.y + 60) - 20 Then
        Vidas = Vidas + 1
        unaVida.x = 0: unaVida.y = 0: unaVida.tag = "": unaVida.tagint = 0: unaVida.tagint2 = 0
        ReproducirSonido VidasCogidas
    End If
    End If
    If unaVida.y > PAlto Then unaVida.tag = ""
End If
' ------------------------- Objeto: Energia ---------------------------
If unaEnergia.tag = "1" Then
    unaEnergia.y = unaEnergia.y + (Int(Rnd * 5) + 7) * (30 / FPS)
    If Nave.x + 20 > unaEnergia.x And (Nave.x + AnchoNave) < (unaEnergia.x + 45) + 20 Then
    If Nave.y - 20 < unaEnergia.y And (Nave.y + AltoNave) > (unaEnergia.y + 45) - 20 Then
        If VidaNave < 4 Then VidaNave = VidaNave + 1
        unaEnergia.x = 0: unaEnergia.y = 0: unaEnergia.tag = "": unaEnergia.tagint = 0: unaEnergia.tagint2 = 0
        ReproducirSonido EnergiaCogida
    End If
    End If
    If unaEnergia.y > PAlto Then unaEnergia.tag = ""
End If
'--------------------------- Objeto: mejora ------------------------
If unaMejora.tag = "1" Then
    unaMejora.y = unaMejora.y + (Int(Rnd * 5) + 7) * (30 / FPS)
    If Nave.x + 25 > unaMejora.x And (Nave.x + AnchoNave) < (unaMejora.x + 45) + 25 Then
    If Nave.y - 25 < unaMejora.y And (Nave.y + AltoNave) > (unaMejora.y + 45) - 25 Then
        If OsciladorArma = True Then
            OsciladorArma = False
        Else
            OsciladorArma = True
            PotenciaActual = PotenciaActual + 1
        End If
        If VidaNave < 4 Then VidaNave = VidaNave + 1
        unaMejora.x = 0: unaMejora.y = 0: unaMejora.tag = "": unaMejora.tagint = 0: unaMejora.tagint2 = 0
        'ReproducirSonido EnergiaCogida
    End If
    End If
    If unaMejora.y > PAlto Then unaMejora.tag = ""
End If
End Sub


Public Sub Jefe_Nave()
Static Iguala As Integer
Static Vx2 As Integer, Vy2 As Integer

Iguala = Iguala + 1
If Iguala = 60 Or Iguala = 1 Then
    Iguala = (Int(Rnd * 15) + 10) * (30 / FPS)
    Vx2 = (Int(Rnd * 5) + 3) * (30 / FPS)
    Vy2 = (Int(Rnd * 5) + 3) * (30 / FPS)
    If Int(Rnd * 2) = 1 Then Vx2 = -Vx2
    If Int(Rnd * 2) = 1 Then Vy2 = -Vy2
End If
If (Jefe.x + Vx2) < 0 Then Vx2 = -Vx2
If (Jefe.x + Vx2 + AnchoJefe) > PAncho Then Vx2 = -Vx2
If Jefe.y < 0 Then
    If Vy2 < 0 Then Vy2 = -Vy2
Else
    If EntrandoE = True Then EntrandoE = False
    If (Jefe.y + Vy2) < 0 Then Vy2 = -Vy2
    If (Jefe.y + Vy2 + AltoJefe) > (PAlto - margenSeguridad) Then Vy2 = -Vy2
End If
Jefe.x = Jefe.x + Vx2
Jefe.y = Jefe.y + Vy2
'-------------------------- Choque Nave-Jefe -----------------------
If (Jefe.x - 25) < Nave.x And (Jefe.x + AnchoJefe + 25) > (Nave.x + AnchoNave) Then
If (Jefe.y - 25) < Nave.y And (Jefe.y + AltoJefe + 25) > (Nave.y + AltoNave) Then
    If Inmunidad = 0 Then
        VidaNave = 4
        Vidas = Vidas - 1
        NuevaExplosion Nave.x, Nave.y
        Nave.x = -100: Nave.y = -100
        ExpVaiven = ExpVaiven + 1
        If ExpVaiven = 6 Then ExpVaiven = 1
        ReproducirSonido Explosion(ExpVaiven)
        TiempoDDM = 60
    End If
End If
End If
End Sub

Public Sub Nave_Meteoro()
If TiempoDDM = -1 Then
    If Inmunidad = 0 Then 'explotar
        For z = 1 To UBound(METs)
            If (METs(z).x - 5) < Nave.x And (METs(z).x + 50 + 5) > Nave.x Then
            If Nave.y > (METs(z).y - 10) And Nave.y < (METs(z).y + 50 + 10) Then
                VidaNave = 4
                Vidas = Vidas - 1
                NuevaExplosion Nave.x, Nave.y
                Nave.x = -100: Nave.y = -100
                Call SonidoExplosion
                TiempoDDM = 60
            End If
            End If
        Next
    Else   'explota el meteoro
ResetX2:
        For z = 1 To UBound(METs)
            If (METs(z).x - 25) < Nave.x And (METs(z).x + 50 + 25) > Nave.x Then
            If (METs(z).y - 25) < Nave.y And (METs(z).y + 50 + 25) > Nave.y Then
                ExplotarMeteorito z
                GoTo ResetX2
            End If
            End If
        Next
    End If
End If
End Sub

Public Sub ControlFondo()
'Inicial, el fondo desacelera
If NivelActual > 0 Then
If Niveles(NivelActual).especial = True And Niveles(NivelActual).especial2 = False Then
    FondoA = (TimerMF + Niveles(NivelActual).TiempoFrase - GetTickCount()) / 750
    PasoFondo = 0.5 * FondoA ^ 2
End If
End If

'Tres etapas: acelera, constante i desacelera
If NivelActual > 0 Then
If Niveles(NivelActual).especial2 = True Then
    If -(TimerMF - GetTickCount()) > (Niveles(NivelActual).TiempoFrase / 6 * 4) Then
        Niveles(NivelActual).frase = ""
        Niveles(NivelActual).frase2 = ""
        If MaxPasoFondo = 0 Then MaxPasoFondo = PasoFondo
        PasoFondo = MaxPasoFondo / ((Niveles(NivelActual).TiempoFrase) / 6 * 3) * (TimerMF + (Niveles(NivelActual).TiempoFrase) - GetTickCount())
    ElseIf -(TimerMF - GetTickCount()) > (Niveles(NivelActual).TiempoFrase / 3 * 4) Then
        'nada
    Else
        FondoA = -(TimerMF - GetTickCount()) / 400
        PasoFondo = 0.5 * FondoA ^ 2
    End If
End If
End If
End Sub

Public Sub Procesa_UFOS()
Static Vx As Integer, Vy As Integer
Static Iguala As Integer

Iguala = Iguala + 1
If Iguala = 40 Or Iguala = 1 Then
    Iguala = Int(Rnd * 20) + 2
    Vx = (Int(Rnd * 7) + 3) * (30 / FPS)
    Vy = (Int(Rnd * 7) + 3) * (30 / FPS)
    If Int(Rnd * 2) = 1 Then Vx = -Vx
    If Int(Rnd * 2) = 1 Then Vy = -Vy
End If
            
'El Ufo choca contra la 'pared' i cambia de sentido
'Por eso Iguala = Int(Rnd * 20) + 2 asi el movimiento no cambia
' i es mas uniforme
If ((UFOs(1).x + Vx) < 0) Or ((UFOs(UBound(UFOs)).x + Vx + AnchoUFO) > PAncho) Then
    Vx = -Vx
    Iguala = Int(Rnd * 20) + 2
End If
If UFOs(1).y < 0 Then
    If Vy < 0 Then Vy = -Vy
Else
    If (UFOs(1).y + Vy) < 0 Then Vy = -Vy
    If (UFOs(1).y + Vy + AltoUFO) > (PAlto - margenSeguridad) Then Vy = -Vy
End If
' ------------- Mover ufos -----------
For z = 1 To UBound(UFOs)
    UFOs(z).x = UFOs(z).x + Vx
    UFOs(z).y = UFOs(z).y + Vy
Next
    ' --------------- Juntar ufos ----------------
JuntarUFOS
End Sub


Public Sub Mover_Met()
Dim z As Integer, w As Integer
If UBound(METs) <> 0 Then
InicioMet:
    For z = 1 To UBound(METs)
        If METs(z).y > PAlto Then
            For w = z To UBound(METs) - 1
                METs(w) = METs(w + 1)
            Next
            ReDim Preserve METs(UBound(METs) - 1)
            GoTo InicioMet
        End If
    Next
    For z = 1 To UBound(METs)
        METs(z).y = METs(z).y + (METs(z).a * (30 / FPS))
        METs(z).frame = METs(z).frame + (1 * (30 / FPS))
        If METs(z).frame >= 9 Then METs(z).frame = 1
    Next
End If
End Sub

Public Sub Mover_Nave()
If KArriba = True Then Nave.y = Nave.y - DespNave * (30 / FPS)
If KAbajo = True Then Nave.y = Nave.y + DespNave * (30 / FPS)
If KDerecha = True Then Nave.x = Nave.x + DespNave * (30 / FPS)
If KIzquierda = True Then Nave.x = Nave.x - DespNave * (30 / FPS)
If Nave.y < 0 Then Nave.y = 0
If Nave.y > (PAlto - AltoNave) Then Nave.y = PAlto - AltoNave
If Nave.x < 0 Then Nave.x = 0
If Nave.x > (PAncho - AnchoNave) Then Nave.x = PAncho - AnchoNave
End Sub


Public Sub Nave_UFO()
Dim z As Integer, k As Integer
If Inmunidad = 0 Then 'explotar
    For z = 1 To UBound(UFOs)
        If (UFOs(z).x - 25) < Nave.x And (UFOs(z).x + AnchoUFO + 25) > (Nave.x + AnchoNave) Then
        If (UFOs(z).y - 35) < Nave.y And (UFOs(z).y + AltoUFO + 35) > (Nave.y + AltoNave) Then
            VidaNave = 4
            Vidas = Vidas - 1
            NuevaExplosion Nave.x, Nave.y
            Nave.x = -100: Nave.y = -100
            ExpVaiven = ExpVaiven + 1
            If ExpVaiven = 6 Then ExpVaiven = 1
            ReproducirSonido Explosion(ExpVaiven)
            TiempoDDM = 60
        End If
        End If
    Next
Else   'explota el ufo
ResetX:
    For z = 1 To UBound(UFOs)
        If (UFOs(z).x - 30) < Nave.x And (UFOs(z).x + AnchoUFO + 30) > (Nave.x + AnchoNave) Then
        If (UFOs(z).y - 40) < Nave.y And (UFOs(z).y + AltoUFO + 40) > (Nave.y + AltoNave) Then
            Call ExplotarUFO(z)
            GoTo ResetX
        End If
        End If
    Next
End If
End Sub


Public Sub DispEnem_Nave_y_Mover()
Dim z As Integer
'Disparos xokan contra la nave explota
ResetE:
For z = 1 To UBound(DisparosE)
    DisparosE(z).y = DisparosE(z).y + 18 * (30 / FPS)
    If DisparosE(z).y >= PAlto Then
        For w = z To UBound(DisparosE) - 1
            DisparosE(w).x = DisparosE(w + 1).x
            DisparosE(w).y = DisparosE(w + 1).y
        Next
        ReDim Preserve DisparosE(UBound(DisparosE) - 1)
        GoTo ResetE
    ElseIf DisparosE(z).y > Nave.y And DisparosE(z).y < (Nave.y + AltoNave) Then
        If DisparosE(z).x > Nave.x And DisparosE(z).x < (Nave.x + AnchoNave) Then
        If Inmunidad = 0 Then
            VidaNave = VidaNave - 1
            If VidaNave = 0 Then
                VidaNave = 4
                Vidas = Vidas - 1
                NuevaExplosion Nave.x, Nave.y
                Nave.x = -100: Nave.y = -100
                ExpVaiven = ExpVaiven + 1
                If ExpVaiven = 6 Then ExpVaiven = 1
                ReproducirSonido Explosion(ExpVaiven)
                TiempoDDM = 60
            End If
            ReproducirSonido Muertes(MuerteVaiven)
            MuerteVaiven = MuerteVaiven + 1
            If MuerteVaiven = 6 Then MuerteVaiven = 1
            For w = z To UBound(DisparosE) - 1
                DisparosE(w).x = DisparosE(w + 1).x
                DisparosE(w).y = DisparosE(w + 1).y
            Next
            ReDim Preserve DisparosE(UBound(DisparosE) - 1)
            GoTo ResetE
        End If
        End If
    End If
Next
End Sub

Public Sub Procesar_DispEnem()
If VaivenEnemigo = 0 Then
    VaivenEnemigo = (Int(Rnd * 7) + 2) * (FPS / 30)
    ReDim Preserve DisparosE(UBound(DisparosE) + 1)
    w = Int(Rnd * UBound(UFOs)) + 1
    VaivenDE = VaivenDE + 1
    If VaivenDE = 6 Then VaivenDE = 1
    ReproducirSonido DisparosEnemigos(VaivenDE)
    If Int(Rnd * 2) = 0 Then
        DisparosE(UBound(DisparosE)).x = UFOs(w).x + 11
    Else
        DisparosE(UBound(DisparosE)).x = UFOs(w).x + AnchoUFO - 11
    End If
    DisparosE(UBound(DisparosE)).y = UFOs(w).y + AltoUFO
End If
VaivenEnemigo = VaivenEnemigo - 1
End Sub

Public Sub Bala_UFO()
ResetB:
For z = 1 To UBound(Balas)
If Balas(z) Is Nothing Then
Else
    For w = 1 To UBound(UFOs)
        If Balas(z).x > UFOs(w).x And (Balas(z).x + AnchoBala) < (UFOs(w).x + AnchoUFO) Then
        If Balas(z).y > UFOs(w).y And (Balas(z).y + AltoBala) < (UFOs(w).y + AltoUFO) Then
            UFOs(w).vida = UFOs(w).vida - Balas(z).potencia
            If UFOs(w).vida <= 0 Then
                If Niveles(NivelActual).vida = True Then
                    If UBound(UFOs) = 1 Or Int(Rnd * 3) = 1 Then
                        Niveles(NivelActual).vida = False
                        unaVida.x = UFOs(w).x
                        unaVida.y = UFOs(w).y
                        unaVida.tag = "1"
                    End If
                End If
                If Niveles(NivelActual).energia = True Then
                    If UBound(UFOs) = 1 Or Int(Rnd * 3) = 1 Then
                        Niveles(NivelActual).energia = False
                        unaEnergia.x = UFOs(w).x
                        unaEnergia.y = UFOs(w).y
                        unaEnergia.tag = "1"
                    End If
                End If
                If Niveles(NivelActual).upgrade = True Then
                    If UBound(UFOs) = 1 Or Int(Rnd * 3) = 1 Then
                        Niveles(NivelActual).upgrade = False
                        unaMejora.x = UFOs(w).x
                        unaMejora.y = UFOs(w).y
                        unaMejora.tag = "1"
                    End If
                End If
                        
                Call ExplotarUFO(w)
                Set Balas(z) = Nothing
                GoTo ResetB
            Else
                'NuevaExplosionP UFOs(w).x + (AnchoUFO - 24) / 2, UFOs(w).y + (AltoUFO - 24) / 2
                Set Balas(z) = Nothing
                GoTo ResetB
            End If
        End If
        End If
    Next
End If
Next
End Sub

Public Sub Bala_Met()
ResetB2:
For z = 1 To UBound(Balas)
If Balas(z) Is Nothing Then
Else
    For w = 1 To UBound(METs)
        If Balas(z).x + 5 > METs(w).x And (Balas(z).x + AnchoBala) - 5 < (METs(w).x + 50) Then
        If Balas(z).y + 5 > METs(w).y And (Balas(z).y + AltoBala) - 5 < (METs(w).y + 20) Then
            METs(w).vida = METs(w).vida - Balas(z).potencia
            If METs(w).vida <= 0 Then
                Call ExplotarMeteorito(w)
                Set Balas(z) = Nothing
                GoTo ResetB2
            Else
                Set Balas(z) = Nothing
                GoTo ResetB2
            End If
        End If
        End If
    Next
End If
Next
End Sub

Public Sub ExplotarMeteorito(ByVal index As Integer)
'fuego
NuevaExplosion METs(index).x, METs(index).y

'eliminar el objeto
If index = UBound(METs) Then
    ReDim Preserve METs(UBound(METs) - 1)
Else
    For k = index To UBound(METs) - 1
        METs(k).frame = METs(k + 1).frame
        METs(k).vida = METs(k + 1).vida
        METs(k).x = METs(k + 1).x
        METs(k).y = METs(k + 1).y
        METs(k).a = METs(k + 1).a
    Next
    ReDim Preserve METs(UBound(METs) - 1)
End If

Call SonidoExplosion
End Sub

Public Sub SonidoExplosion()
ExpVaiven = ExpVaiven + 1
If ExpVaiven = 6 Then ExpVaiven = 1
ReproducirSonido Explosion(ExpVaiven)
End Sub

Public Sub ExplotarUFO(ByVal index As Integer)
'se hace explotar el ufo pasado como INDEX
NuevaExplosion UFOs(index).x, UFOs(index).y
If index = UBound(UFOs) Then
    ReDim Preserve UFOs(UBound(UFOs) - 1)
Else
    For k = index To UBound(UFOs) - 1
        UFOs(k).tipo = UFOs(k + 1).tipo
        UFOs(k).vida = UFOs(k + 1).vida
        UFOs(k).x = UFOs(k + 1).x
        UFOs(k).y = UFOs(k + 1).y
    Next
    ReDim Preserve UFOs(UBound(UFOs) - 1)
End If
Call SonidoExplosion
End Sub

Public Sub NuevoDisparo(ByRef unaBala As cBala, ByRef unaBala2 As cBala)
'al invocar esta funcion se dispara una nueva bala desde la nave
DisparoID = DisparoID + 1
If DisparoID = 6 Then DisparoID = 1
ReproducirSonido Disparos(DisparoID)
FrameABala = CLng((4 * (FPS / 30)) + 0.5)
Disparar = False
Set unaBala = New cBala
Set unaBala2 = New cBala
unaBala.potencia = PotenciaActual
unaBala2.potencia = PotenciaActual
If OsciladorArma = True Then
    If Vaiven = 0 Then
        Vaiven = 1
        unaBala.x = Nave.x + 3
    Else
        Vaiven = 0
        unaBala.x = Nave.x + AnchoNave - 8
    End If
Else
    unaBala.x = Nave.x + 3
    unaBala2.x = Nave.x + AnchoNave - 8
End If
unaBala.y = Nave.y - AltoBala / 5
unaBala2.y = Nave.y - AltoBala / 5
For z = 1 To 30
    If Balas(z) Is Nothing Then
        Set Balas(z) = unaBala
        Exit For
    End If
Next
If OsciladorArma = False Then
    For z = 1 To 30
        If Balas(z) Is Nothing Then
            Set Balas(z) = unaBala2
            Exit For
        End If
    Next
End If
End Sub


Public Sub Bala_Jefe()
For z = 1 To UBound(Balas)
    If Balas(z) Is Nothing Then
    Else
        If Balas(z).x > Jefe.x And (Balas(z).x + AnchoBala) < (Jefe.x + AnchoJefe) Then
        If Balas(z).y > Jefe.y And (Balas(z).y + AltoBala) < (Jefe.y + AltoJefe) Then
            Jefe.vida = Jefe.vida - Balas(z).potencia
            Set Balas(z) = Nothing
            If Jefe.vida <= 0 Then
                For k = -128 To 128 Step 64
                    For ww = -128 To 128 Step 64
                        NuevaExplosion k + (Jefe.x + AnchoJefe / 2) + Int(Rnd * 50), ww + (Jefe.y + AltoJefe / 2) + Int(Rnd * 50), -Int(Rnd * 15 * w)
                        NuevaExplosion k + (Jefe.x + AnchoJefe / 2) + Int(Rnd * 50), ww + (Jefe.y + AltoJefe / 2) - Int(Rnd * 50), Int(-Int(Rnd * 15 * w) * 1.5)
                        NuevaExplosion k + (Jefe.x + AnchoJefe / 2) - Int(Rnd * 50), ww + (Jefe.y + AltoJefe / 2) + Int(Rnd * 50), Int(-Int(Rnd * 15 * w) * 2)
                        NuevaExplosion k + (Jefe.x + AnchoJefe / 2) - Int(Rnd * 50), ww + (Jefe.y + AltoJefe / 2) - Int(Rnd * 50), Int(-Int(Rnd * 15 * w) * 2.5)
                        ExpVaiven = ExpVaiven + 1
                        If ExpVaiven = 6 Then ExpVaiven = 1
                        ReproducirSonido JefeExp
                    Next
                Next
                Exit For
            End If
        End If
        End If
    End If
Next
End Sub
