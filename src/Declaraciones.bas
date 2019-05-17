Attribute VB_Name = "Declaraciones"

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' DECLARACION DE TODAS LAS CONSTANTES Y VARIABLES GLOBALES
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit


' * * * * * * * * * * * * * * * *
' DECLARACION DE FUNCIONES DE API
' * * * * * * * * * * * * * * * *

Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, _
                                        lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Public Declare Function QueryPerformanceCounter Lib "Kernel32" _
                                 (X As Currency) As Boolean

Public Declare Function QueryPerformanceFrequency Lib "Kernel32" _
                                 (X As Currency) As Boolean

Public Declare Function GetTickCount Lib "Kernel32" () As Long
      
Public Declare Function timeGetTime Lib "winmm.dll" () As Long


' * * * * * * * * * * * * *
' DECLARACION DE CONSTANTES
' * * * * * * * * * * * * *

'Total filas del mapa
Public Const MAX_ROWS_MAP = 255

' Los diferentes niveles
Public Const NIVEL0 = 0
Public Const NIVEL1 = 1
Public Const NIVEL2 = 2
Public Const NIVEL3 = 3
Public Const NIVEL4 = 4
Public Const NIVEL5 = 5
Public Const NIVEL6 = 6

Public Const MAX_NIVEL = 6

' Total de charcos
Public Const MAX_CHARCOS = 2

' Total de manchas de aceite
Public Const MAX_MANCHAS_ACEITE = 2

' Total de rocas
Public Const MAX_ROCAS = 2

Public Const ASFALTO_LISO = "3"
Public Const MANCHA_ACEITE = "&"
Public Const MANCHA_AGUA = "$"
Public Const ROCA = "%"


' Posiciones de los elementos gráficos del marcador
Public Const POS_RECORRIDO_INICIAL = 143
Public Const POS_RECORRIDO_FINAL = 82

Public Const DESP_POSY_CAR_FROM_TOP = 11

' SONIDOS
Public Const MAX_SOUNDS = 18
Public Const marcha = 0
Public Const crash1 = 1
Public Const crash2 = 2
Public Const adelantada = 3
Public Const semaforo1 = 4
Public Const semaforo2 = 5
Public Const freno1 = 6
Public Const turbo = 7
Public Const carstart = 8
Public Const derrape = 9
Public Const freno2 = 10
Public Const bocinazo = 11
Public Const carga_combustible = 12
Public Const gameover = 13
Public Const intro = 14
Public Const nivel = 15
Public Const outoffuel = 16
Public Const agua = 17

' KM que tiene el recorrido del juego
Public Const TRAMO = 2040

' Total de vueltas completadas necesarias para cubrir un tramo
Public Const VUELTAS_COMPLETADAS = 2

' Litros de combustible iniciales
Public Const COMBUSTIBLE_INICIAL = 50#

' Resolucion de la pantalla
Public Const MAX_WIDTH = 320
Public Const MAX_HEIGHT = 240

' Dimensiones de los coches
'Const MAX_CAR_WIDTH = 10
'Const MAX_CAR_HEIGHT = 16

Public Const MAX_CAR_WIDTH = 14
Public Const MAX_CAR_HEIGHT = 24

'Const MAX_CAR_WIDTH = 19
'Const MAX_CAR_HEIGHT = 32


' Posisiones de los coches
Public Const POSY_INI_CAR = 180
Public Const POSX_INI_CAR = 48
Public Const POSY_INI_BAD_CAR = -50
Public Const POSY_FIN_BAD_CAR = 300
Public Const POSX_INI_BAD_CAR = 63 ' El centro de la carretera - 1/2 * ancho del coche

' Cantidad de filas y columnas del mapa que conforman el paisaje

Public Const MAX_COLS_MAP = 15
Public Const MAX_COLS_ROAD = 10

Public Const MAX_ROWS_SCREEN = 15 ' Numero de filas que caben en pantalla en horizontal

' El objeto DirectX
Public DX As DirectX7

' El objeto DirectDraw
Public DD As DirectDraw7


' * * * * * * * * * * * * * *
' Variables para Direct Music
' * * * * * * * * * * * * * *

' Un perfomance controla y organiza los eventos de la musica. Comprueba cambios despues
' de enviar los datos al puerto de musica, establecer opciones, convertir señales, etc.
Public perf As DirectMusicPerformance

' Matriz de segmentos->almacena los datos, igual que una superficie en DirectDraw
Public seg(1 To MAX_NIVEL) As DirectMusicSegment

' Indica el estado de la musica: reproduciendo, parada, etc.
Public segstate As DirectMusicSegmentState

' Un objeto cargador traspasa el fichero en disco hacia una area de la memoria
Public loader As DirectMusicLoader

' * * * * * * * * * * * * * *
' Variables para Direct Sound
' * * * * * * * * * * * * * *

' Variables de Direct Sound, Direct Sound Buffer y Direct Sound Buffer Description.
Public DS As DirectSound
Public sb(MAX_SOUNDS) As DirectSoundBuffer
Public sbd As DSBUFFERDESC
Public wf As WAVEFORMATEX

'
'

' Los límites horizontales en tiempo real de la carretera
Public limite_borde_izquierdo As Integer
Public limite_borde_derecho As Integer

' MAPAS
Public Mapa(MAX_ROWS_MAP) As String
Public MapaPaisaje(MAX_ROWS_MAP) As String




' Las superficies

' La Superficie Primaria
Public supPrimary As DirectDrawSurface7

' La Superficie Secundaria
Public SupBackBuffer As DirectDrawSurface7

' La superficie para la imagen de fondo
Public SupFondo As DirectDrawSurface7

' La superficie para representar graficamente la velocidad
Public SupVelocidad As DirectDrawSurface7
Public SupAntiVelocidad As DirectDrawSurface7

' La superficie para representar graficamente el combustible disponible
Public SupCombustible As DirectDrawSurface7

' La superficie para representar graficamente el trayecto recorrido
Public SupPunteroCoche As DirectDrawSurface7

' Pantalla que muestra que se empieza el nivel 2
Public SupNivel As DirectDrawSurface7

' Superficie que almacena la carretera
Public PaisajeCarretera As DirectDrawSurface7

' Superficie que almacena el bitmap con los tiles del paisaje del nivel 1
Public Paisaje As DirectDrawSurface7

' Superficie que almacena el nivel 1 tal y como debe aparecer en pantalla
Public PaisajePantalla As DirectDrawSurface7

' Array de Superficies para la explosión del coche
Public SupExplosion(13) As DirectDrawSurface7


' Mapa y mapa auxiliar
Public Carretera1 As DirectDrawSurface7
Public Carretera2 As DirectDrawSurface7

' Estructura de los Sprites; superficies con atributos extras
Public Type t_Sprite
    posX As Single
    posY As Single
    Ancho As Single
    Alto As Single
    Velocidad As Double
    Aceleracion As Double
    Superficie As DirectDrawSurface7
End Type

' Los coches
Public coche As t_Sprite, CocheMalo1 As t_Sprite, CocheMalo2 As t_Sprite
Public Camion As t_Sprite, Camion2 As t_Sprite, CocheMalo3 As t_Sprite
Public CocheMalo4 As t_Sprite, HumoTuboEscape(1) As t_Sprite, RecargaCombustible As t_Sprite

Public Type CoordenadasObstaculos
  PosXMapa As Integer
  PosYMapa As Integer
  PosXPantalla As Integer
  PosYPantalla As Integer
End Type

Public ObjetoCharco(MAX_CHARCOS) As CoordenadasObstaculos
Public ObjetoManchaAceite(MAX_MANCHAS_ACEITE) As CoordenadasObstaculos
Public ObjetoRoca(MAX_ROCAS) As CoordenadasObstaculos

Public Sonidos(MAX_SOUNDS) As String

' Para los cuadros por segundo
Public FPS_tUltimo As Long
Public FPS_Suma As Single
Public FPS_Actual As Single

Public RowIndex As Long
Public MapDspRow As Double

Public NivelActual As Integer

' Descripcion para la fuente
Public DescFuente As New StdFont

' Las descripciones de las superficies
Public DescPri As DDSURFACEDESC2
Public DescSec As DDSURFACEDESC2
Public DescFondo As DDSURFACEDESC2
Public DescVelocidad As DDSURFACEDESC2
Public DescCoche As DDSURFACEDESC2
Public DescCamion As DDSURFACEDESC2
Public DescCamion2 As DDSURFACEDESC2
Public DescHumoTuboEscape As DDSURFACEDESC2
Public DescPunteroCoche As DDSURFACEDESC2
Public DescCarretera1 As DDSURFACEDESC2
Public DescRecargaCombustible As DDSURFACEDESC2
Public DescNivel As DDSURFACEDESC2
Public DescPaisajeCarretera As DDSURFACEDESC2
Public DescPaisaje As DDSURFACEDESC2
Public DescPaisaje1Pantalla As DDSURFACEDESC2
Public DescExplosion As DDSURFACEDESC2

Public puntos As Long

Public Combustible As Single

Public SinCombustible As Boolean
Public CamionActivo As Boolean ' Si el camion ha de estar visible o no
Public Camion2Activo As Boolean ' Si el camion ha de estar visible o no

Public CocheActivo As Boolean
Public CocheMalo1Activo As Boolean
Public CocheMalo2Activo As Boolean
Public CocheMalo3Activo As Boolean
Public CocheMalo4Activo As Boolean
Public RecargaCombustibleActiva As Boolean

Public BarraEspaciadora As Boolean
Public TeclaIzquierda As Boolean
Public TeclaDerecha As Boolean

' Control de choque entre camiones y vehiculos enemigos
Public CocheMalo1ContraCamion As Boolean
Public CocheMalo2ContraCamion As Boolean
Public CocheMalo3ContraCamion As Boolean
Public CocheMalo4ContraCamion As Boolean

' Flag indicativo de que el jugador ha chocado
Public Choque_en_Curso As Boolean

' Velocidad del jugador
Public vel As Currency

' Guardar la velocidad a la que el jugador se da el porrazo
Public VelocidadInicialColision As Currency

' Segundos que deben transcurrir para mostrar el camion
Public TiempoNecesarioParaMostrarCamion As Integer

' Acumulador de segundos sin aparecer el camion
Public SegundosAcumuladosSinMostrarCamion As Integer

' Segundos que deben transcurrir para mostrar el segundo modelo de camion
Public TiempoNecesarioParaMostrarCamion2 As Integer

' Acumulador de segundos sin aparecer el segundo modelo de camion
Public SegundosAcumuladosSinMostrarCamion2 As Integer

' Segundos que deben transcurrir para mostrar el cochemalo1
Public TiempoNecesarioParaMostrarCocheMalo1 As Integer

' Acumulador de segundos sin aparecer el coche
Public SegundosAcumuladosSinMostrarCochemalo1 As Integer

' Segundos que deben transcurrir para mostrar el cochemalo1
Public TiempoNecesarioParaMostrarCocheMalo2 As Integer

' Acumulador de segundos sin aparecer el coche
Public SegundosAcumuladosSinMostrarCochemalo2 As Integer

' Segundos que deben transcurrir para mostrar el cochemalo1
Public TiempoNecesarioParaMostrarCocheMalo3 As Integer

' Acumulador de segundos sin aparecer el coche
Public SegundosAcumuladosSinMostrarCochemalo3 As Integer

' Segundos que deben transcurrir para mostrar el cochemalo1
Public TiempoNecesarioParaMostrarCocheMalo4 As Integer

' Acumulador de segundos sin aparecer el coche
Public SegundosAcumuladosSinMostrarCochemalo4 As Integer


' Matriz que almacena flags para indicar que perspectivas del coche del jugador se han
' mostrado ya

Public Posiciones_Choque(0 To 12) As Boolean

Public perf_flag As Boolean        ' Timer Selection Flag
Public time_factor As Currency     ' Time Scaling Factor
Public time_span As Double    ' time elapsed since last frame
Public last_time As Currency           ' Previous timer value
Public last_time2 As Currency

' Maxima puntuacion conseguida
Public Record As String

' Velocidad Maxima del jugador
Public max_velocidad As Currency

' Indicador de que está activo el turbo
Public FlagTurbo As Boolean

' Vector de superficies para cada posicion del coche del jugador al colisionar
Public CocheColision(7) As t_Sprite

' Ir almacenando temporalmente las vueltas completas al mapa
Public VueltasCompletadas As Integer

' CuentaKilometros
Public Recorrido As Long

' Posicion del puntero al coche en en la representacion gráfica del recorrido
Public PosicionRecorrido As Integer

' Vector de superficies para cada posicion de los coches contrincantes al colisionar
Public CocheMalo1Colision(7) As t_Sprite
Public CocheMalo2Colision(7) As t_Sprite
Public CocheMalo3Colision(7) As t_Sprite
Public CocheMalo4Colision(7) As t_Sprite

' Indicador de vehiculo enemigo en pleno choque
Public ChoqueCocheMalo1EnCurso As Boolean
Public ChoqueCocheMalo2EnCurso As Boolean
Public ChoqueCocheMalo3EnCurso As Boolean
Public ChoqueCocheMalo4EnCurso As Boolean

' Vector de indicadores de la posicion de los coches contrincantes al colisionar
Public Posiciones_Choque_CocheMalo1(0 To 12) As Boolean
Public Posiciones_Choque_CocheMalo2(0 To 12) As Boolean
Public Posiciones_Choque_CocheMalo3(0 To 12) As Boolean
Public Posiciones_Choque_CocheMalo4(0 To 12) As Boolean

Public bEjecutandose   As Boolean          'Indica si el bucle está ejecutándose


Public Sub LeerMapaDisco(ByVal NomFichero As String, ByRef MapaDestino() As String)

Dim LineaVector As String
Dim i As Integer

On Error GoTo errores

Open App.Path & "\dat\" & NomFichero For Input As #1   ' Abre el archivo.

' Cargar cada línea del fichero de cada vector
Do While Not EOF(1)
  Line Input #1, LineaVector
  MapaDestino(i) = LineaVector
  i = i + 1
Loop

Close #1   ' Cierra el archivo.

Exit Sub

errores:

CerrarDD "LeerMapaDisco"

End Sub
Public Sub CerrarDD(Optional ByVal ProcedimientoConError As String)
    
    DError.HandleAnyErrors Err.Number, DirectDraw, ProcedimientoConError
    bEjecutandose = False
    
    End
    
End Sub

Private Sub InicializarVariables()

  Dim i As Integer
    
  ' Inicializacion generica para cada nivel
  max_velocidad = 10#
  Combustible = COMBUSTIBLE_INICIAL
  FPS_tUltimo = 0
  FPS_Suma = 0
  FPS_Actual = 0
  limite_borde_izquierdo = 59
  limite_borde_derecho = 59 + 160 - 11
  VueltasCompletadas = 0
  Recorrido = 0
  PosicionRecorrido = POS_RECORRIDO_INICIAL
  SinCombustible = False
  CocheActivo = True
  CamionActivo = False
  Camion2Activo = False
  CocheMalo1Activo = False
  CocheMalo2Activo = False
  CocheMalo3Activo = False
  CocheMalo4Activo = False
  RecargaCombustibleActiva = False
  BarraEspaciadora = False
  TeclaIzquierda = False
  TeclaDerecha = False
  Choque_en_Curso = False
  CocheMalo1ContraCamion = False
  CocheMalo2ContraCamion = False
  CocheMalo3ContraCamion = False
  CocheMalo4ContraCamion = False
  
  vel = 0#
  
  TiempoNecesarioParaMostrarCamion = 0
  TiempoNecesarioParaMostrarCamion2 = 0
  TiempoNecesarioParaMostrarCocheMalo1 = 0
  TiempoNecesarioParaMostrarCocheMalo2 = 0
  TiempoNecesarioParaMostrarCocheMalo3 = 0
  TiempoNecesarioParaMostrarCocheMalo4 = 0
  
  SegundosAcumuladosSinMostrarCamion = 0
  SegundosAcumuladosSinMostrarCamion2 = 0
  SegundosAcumuladosSinMostrarCochemalo1 = 0
  SegundosAcumuladosSinMostrarCochemalo2 = 0
  SegundosAcumuladosSinMostrarCochemalo3 = 0
  SegundosAcumuladosSinMostrarCochemalo4 = 0
  
  ' Limpiar la matriz de flags de posiciones del coche del jugador cuando colisiona
  For i = 0 To 7
    Posiciones_Choque(i) = False
    Posiciones_Choque_CocheMalo1(i) = False
    Posiciones_Choque_CocheMalo2(i) = False
    Posiciones_Choque_CocheMalo3(i) = False
    Posiciones_Choque_CocheMalo4(i) = False
  Next i
  
  ' Velocidad Maxima del jugador
  max_velocidad = 10#
  
  FlagTurbo = False
  
  ChoqueCocheMalo1EnCurso = False
  ChoqueCocheMalo2EnCurso = False
  ChoqueCocheMalo3EnCurso = False
  ChoqueCocheMalo4EnCurso = False
  
  ' Inicializacion del desplazamiento para el scroll vertical 'suave'
  MapDspRow = 0
  
  With coche
    .posX = POSX_INI_CAR + (160 / 2) - (.Ancho / 2)
    .posY = POSY_INI_CAR - (MAX_CAR_HEIGHT / 2)
    .Ancho = MAX_CAR_WIDTH
    .Alto = MAX_CAR_HEIGHT
    .Velocidad = 1
  End With
  
  With CocheMalo1
    .posX = InicializarCoordenadasX
    .posY = POSY_INI_BAD_CAR
    .Ancho = MAX_CAR_WIDTH
    .Alto = MAX_CAR_HEIGHT
    .Velocidad = 20
    .Aceleracion = 9.8
  End With
  
  With CocheMalo2
    .posX = InicializarCoordenadasX
    .posY = POSY_INI_BAD_CAR * Rnd(-420)
    .Ancho = MAX_CAR_WIDTH
    .Alto = MAX_CAR_HEIGHT
    .Velocidad = 65
    .Aceleracion = 9.5
  End With
  
  With CocheMalo3
    .posX = InicializarCoordenadasX
    .posY = POSY_INI_BAD_CAR * Rnd(-260)
    .Ancho = MAX_CAR_WIDTH
    .Alto = MAX_CAR_HEIGHT
    .Velocidad = 55
    .Aceleracion = 8.5
  End With
    
  With CocheMalo4
    .posX = InicializarCoordenadasX
    .posY = POSY_INI_BAD_CAR * Rnd(-1350)
    .Ancho = MAX_CAR_WIDTH
    .Alto = MAX_CAR_HEIGHT
    .Velocidad = 77
    .Aceleracion = 5.5
  End With
  
  With Camion
    .posX = InicializarCoordenadasCamionX
    .posY = POSY_FIN_BAD_CAR + 62
    .Ancho = DescCamion.lWidth
    .Alto = DescCamion.lHeight
    .Velocidad = 20
    .Aceleracion = 2.5
  End With
  
  With Camion2
    .posX = InicializarCoordenadasCamionX
    .posY = POSY_FIN_BAD_CAR + 92
    .Ancho = DescCamion2.lWidth
    .Alto = DescCamion2.lHeight
    .Velocidad = 10
    .Aceleracion = 1.5
  End With

  With HumoTuboEscape(0)
    .posX = 0
    .posY = 0
    .Ancho = DescHumoTuboEscape.lWidth
    .Alto = DescHumoTuboEscape.lHeight
    .Velocidad = 0
    .Aceleracion = 0
  End With
    
  With RecargaCombustible
    .posX = InicializarCoordenadasX
    .posY = POSY_INI_BAD_CAR
    .Ancho = DescRecargaCombustible.lWidth
    .Alto = DescRecargaCombustible.lHeight
    .Velocidad = 80
    .Aceleracion = 10
  End With
  
  ' Apunta al tramo actual->indica en todo momento donde estamos
  RowIndex = TRAMO - MAX_ROWS_SCREEN
  
End Sub

Private Function InicializarCoordenadasX()

'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)
InicializarCoordenadasX = Int((limite_borde_derecho - limite_borde_izquierdo + 1) * Rnd + limite_borde_izquierdo)

End Function

Private Function InicializarCoordenadasCamionX()
  'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)
  
  InicializarCoordenadasCamionX = Int((limite_borde_derecho - Camion.Ancho - 2 - _
                                    limite_borde_izquierdo) + 1) * _
                                    Rnd + limite_borde_izquierdo

End Function


Private Sub CargarMapaP1()

On Error GoTo errores

 Dim X As Integer, Y As Integer
 Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
 Dim rSprite As RECT
 Dim bloq As String * 1, bloq_anterior As String * 1
  
' Cargar la 1/8
LeerMapaDisco "map1a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), (Y * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y

'-----------------------------------------------------------------------------------------

' Cargar la 2/8
LeerMapaDisco "map1b.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 254) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y


'-----------------------------------------------------------------------------------------

' Cargar la 3/8

LeerMapaDisco "map1c.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 509) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y

'-----------------------------------------------------------------------------------------

' Cargar la 4/8

LeerMapaDisco "map1d.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 764) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y


'-----------------------------------------------------------------------------------------

' Cargar la 5/8

LeerMapaDisco "map1d.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 1019) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y

'-----------------------------------------------------------------------------------------

' Cargar la 6/8

LeerMapaDisco "map1d.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 1274) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y


'-----------------------------------------------------------------------------------------

' Cargar la 7/8

LeerMapaDisco "map1d.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 1529) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y


'-----------------------------------------------------------------------------------------

' Cargar la 8/8

LeerMapaDisco "map1d.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 1784) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y
    
'
' SITUAMOS LA CARRETERA SOBRE EL PAISAJE
'

With rSprite
  .Top = 0: .Left = 0: .Right = 160: .Bottom = 32640
End With
    
PaisajePantalla.BltFast POSX_INI_CAR, 0, PaisajeCarretera, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Exit Sub

errores:
CerrarDD "CargarMapaP1"

End Sub
'
' EN ESTE NIVEL, POR SER EL PRIMERO, NO HAY NINGUN OBSTACULO
'
Private Sub CargarMapaCarretera1()


On Error GoTo errores

 Dim X As Integer, Y As Integer
 Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
 Dim rSprite As RECT
 Dim bloq As String * 1, bloq_anterior As String * 1
  
'
' DEFINICION DEL MAPA DE LA CARRETERA DEL NIVEL 1 Y CARGA DEL MISMO EN UNA SUPERFICIE
'
' Las 10 primeras posiciones representa los 'tiles' de la carretera.
' La undécima posición indica la posición en la que está el tile borde izquierdo
' La duodécima posición indica la posición en la que está el tile borde derecho

LeerMapaDisco "carret1.mmp", Mapa

' Se barre el mapa y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_ROAD - 1
  
     bloq = Mid$(Mapa(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del mapa si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_rectangulo bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1
        .Right = x2
        .Top = y1
        .Bottom = y2
      End With
    End If
          
    PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
  
  Next X

Next Y

Exit Sub

errores:
CerrarDD "CargarMapaCarretera1"

End Sub
Private Sub CargarSonidos()

Set sb(marcha) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(marcha), sbd, wf)
Set sb(crash1) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(crash1), sbd, wf)
Set sb(crash2) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(crash2), sbd, wf)
Set sb(adelantada) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(adelantada), sbd, wf)
Set sb(semaforo1) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(semaforo1), sbd, wf)
Set sb(semaforo2) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(semaforo2), sbd, wf)
Set sb(freno1) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(freno1), sbd, wf)
Set sb(turbo) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(turbo), sbd, wf)
Set sb(carstart) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(carstart), sbd, wf)
Set sb(derrape) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(derrape), sbd, wf)
Set sb(freno2) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(freno2), sbd, wf)
Set sb(bocinazo) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(bocinazo), sbd, wf)
Set sb(carga_combustible) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(carga_combustible), sbd, wf)
Set sb(gameover) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(gameover), sbd, wf)
Set sb(intro) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(intro), sbd, wf)
Set sb(nivel) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(nivel), sbd, wf)
Set sb(outoffuel) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(outoffuel), sbd, wf)
Set sb(agua) = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\" & Sonidos(agua), sbd, wf)

End Sub
Private Sub IniciarMatrizSonidos()
  
Sonidos(marcha) = "marcha.wav"
Sonidos(crash1) = "crash1.wav"
Sonidos(crash2) = "crash2.wav"
Sonidos(adelantada) = "adelantada.wav"
Sonidos(semaforo1) = "semaforo1.wav"
Sonidos(semaforo2) = "semaforo2.wav"
Sonidos(freno1) = "freno1.wav"
Sonidos(turbo) = "turbo.wav"
Sonidos(carstart) = "carstart.wav"
Sonidos(derrape) = "derrape.wav"
Sonidos(freno2) = "freno2.wav"
Sonidos(bocinazo) = "bocinazo.wav"
Sonidos(carga_combustible) = "combustible.wav"
Sonidos(gameover) = "gameover.wav"
Sonidos(intro) = "intro.wav"
Sonidos(nivel) = "nivel.wav"
Sonidos(outoffuel) = "outoffuel.wav"
Sonidos(agua) = "agua.wav"

End Sub
Public Sub ResetInicial()
  
  Dim perf_cnt As Currency
  
  frmPrincipal.Show
  
  Randomize ' Inicializar semilla de numeros aleatorios
  
  ' ¿El ordenador admite mejora de control del tiempo?

  If QueryPerformanceFrequency(perf_cnt) Then
    ' Sí, activar flag de mejora de control del tiempo
    perf_flag = True

    ' Establecer el factor de escala
    time_factor = 1# / perf_cnt
 
    ' Leer el tiempo inicial
    QueryPerformanceCounter last_time
  Else
    ' Si no hay mejora de control del tiempo, leer usando timeGetTime
    last_time = timeGetTime()

    ' Limpiar el flag de seleccion del tiempo
    perf_flag = False

    ' Establecer el factor de escala
    time_factor = 0.001
  End If
  
  
  
  ' Inicializar el objeto DirectX 7
  Set DX = New DirectX7
  
  ' Inicializar objetos de DirectMusic
  InicializarDM
  
  ' Cargar todas las musicas
  InicializarMusica
  
  ' Inicializar objetos de DirectSound
  InicializarDS
  IniciarMatrizSonidos
  CargarSonidos
  
  ' Inicializar objetos de DirectDraw
  InicializarDD
  InicializarSuperficies
  
  ' NIVEL0 sólo se utiliza a modo de inicialización. Una vez mostrada la pantalla que
  ' anuncia que se está en el NIVEL1, NivelActual valdrá NIVEL1 para todo el nivel 1
  NivelActual = NIVEL0
  
  InicializarSuperficiesSprites
    
  ' Inicializar fuentes y asignarla a la superficie secundaria (buffer)
  InicializarFuentes
  
  ' Leer la puntuacion record desde el archivo del directorio 'dat'
  Record = LeerRecord
    
End Sub
Private Sub InicializarDD()
On Error GoTo errores
    
    '### Creación de los objetos principales ###
    
    Set DD = DX.DirectDrawCreate("")
    
    '### Nivel Cooperativo ###
    DD.SetCooperativeLevel frmPrincipal.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE 'a pantalla completa
    
    '### Modo de Pantalla ###
    DD.SetDisplayMode MAX_WIDTH, MAX_HEIGHT, 16, 0, DDSDM_DEFAULT '320x240 a 16 bits de color
    
    Exit Sub
    
errores:
    CerrarDD "InicializarDD"

End Sub

Private Sub InicializarDM()

  On Error GoTo errores
      
  ' Crear objeto cargador
  Set loader = DX.DirectMusicLoaderCreate()

  ' Crear objeto Perfomance
  Set perf = DX.DirectMusicPerformanceCreate()
  
  ' Inicializarlo
  Call perf.Init(Nothing, 0)
    
  ' Establecer sus opciones (casi siempre opciones estandars para todos los proyectos)
  perf.SetPort -1, 80
  Call perf.SetMasterAutoDownload(True)
   
  ' El 75 en la siguiente sentencia puede ser cualquier numero entre 0 y 100.
  ' La fórmula solo utiliza este número para crear un volumen que DM pueda interpretar
  perf.SetMasterVolume (75 * 42 - 3000)
      
  Exit Sub
  
errores:
  
  CerrarDD "InicializarDirectMusic"
End Sub
Private Sub InicializarMusica()
  
  On Error GoTo errores

  Dim i As Integer
  
  ' Usar el objeto cargador para crear un buffer de musica
  Set seg(NIVEL1) = loader.LoadSegment(App.Path & "\musica\ambiente.mid")
  Set seg(NIVEL2) = loader.LoadSegment(App.Path & "\musica\fun.mid")
  Set seg(NIVEL3) = loader.LoadSegment(App.Path & "\musica\fun.mid")
  Set seg(NIVEL4) = loader.LoadSegment(App.Path & "\musica\fun.mid")
  Set seg(NIVEL5) = loader.LoadSegment(App.Path & "\musica\fun.mid")
  Set seg(NIVEL6) = loader.LoadSegment(App.Path & "\musica\fun.mid")
  
  For i = 1 To MAX_NIVEL
    
    'Establecer el formato a MIDI
    seg(i).SetStandardMidiFile
   
    ' Establecer el punto de loop
    seg(i).SetLoopPoints 0, 0
    seg(i).SetRepeats 10
  
  Next i
  
  Exit Sub
  
errores:
  CerrarDD "InicializarMusica"
End Sub
Private Sub ReproducirMusica(ByVal nivel As Integer)
  
  ' Reproducir la musica para el nivel actual
  Select Case nivel
    Case NIVEL1
      Set segstate = perf.PlaySegment(seg(NIVEL1), 0, 0)
    Case NIVEL2
      Set segstate = perf.PlaySegment(seg(NIVEL2), 0, 0)
    Case NIVEL3
      Set segstate = perf.PlaySegment(seg(NIVEL3), 0, 0)
    Case NIVEL4
      Set segstate = perf.PlaySegment(seg(NIVEL4), 0, 0)
    Case NIVEL5
      Set segstate = perf.PlaySegment(seg(NIVEL5), 0, 0)
    Case NIVEL6
      Set segstate = perf.PlaySegment(seg(NIVEL6), 0, 0)
  End Select
  
End Sub
Private Sub DecidirParCochesIniciales()

Dim Valor As Integer
Valor = Int((4 - 1 + 1) * Rnd + 1)


Select Case Valor
  Case 1
    CocheMalo1Activo = True
    CocheMalo2Activo = True
  Case 2
    CocheMalo3Activo = True
    CocheMalo4Activo = True
  Case 3
    CocheMalo2Activo = True
    CocheMalo3Activo = True
  Case 4
    CocheMalo1Activo = True
    CocheMalo4Activo = True
End Select

End Sub
Private Sub InicializarDS()

On Error GoTo errores


Set DS = DX.DirectSoundCreate("")
DS.SetCooperativeLevel frmPrincipal.hWnd, DSSCL_NORMAL

' Establecer el descriptor del Buffer
sbd.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
'wf.nSize = LenB(WAVEFORMATEX)
wf.nSize = 0
wf.nFormatTag = WAVE_FORMAT_PCM
wf.nChannels = 2
wf.lSamplesPerSec = 44000
wf.nBitsPerSample = 16
wf.nBlockAlign = wf.nBitsPerSample / 8 * wf.nChannels
wf.lAvgBytesPerSec = wf.lSamplesPerSec * wf.nBlockAlign

Exit Sub

errores:
  CerrarDD "InicializarDirectSound"
End Sub
Private Sub InicializarFuentes()
  
  ' Descripcion de la fuente
  With DescFuente
    .Name = "Tahoma"
    .Size = 7
    .Bold = True
    .Italic = False
    .Underline = False
    .Strikethrough = False
  End With
  
  ' Asignacion de la fuente a la superficie secundaria
  With SupBackBuffer
    .SetForeColor RGB(255, 255, 255)
    .SetFontTransparency True
    .SetFont DescFuente
  End With
  
End Sub
Private Sub InicializarSuperficies()

  Dim i As Integer
  
  On Error GoTo errores

  ' Contiene la descripción de la Superficie Primaria
  DescPri.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
  DescPri.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
  DescPri.lBackBufferCount = 1
  Set supPrimary = DD.CreateSurface(DescPri)
    
  ' Descripcion de la superficie secundaria
    
  Dim Caps As DDSCAPS2 ' Contiene las "capacidades" de la Superficie Secundaria
  
  Caps.lCaps = DDSCAPS_BACKBUFFER
  Set SupBackBuffer = supPrimary.GetAttachedSurface(Caps)
  SupBackBuffer.GetSurfaceDesc DescSec
 
  Exit Sub

errores:
    CerrarDD "InicializarSuperficies"

End Sub


Private Sub InicializarSuperficiesSprites()
  
  Dim Key As DDCOLORKEY
  Dim i As Integer
  
  On Error GoTo errores
  
  Key.low = 0
  Key.high = 0
  
  ' Creación de la superficies extras y cargas de los BMP
  
  
  ' Creación de la superficie para la explosión del coche
  With DescExplosion
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
    .lWidth = 28
    .lHeight = 21
  End With
    
  Set SupExplosion(0) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo1.bmp", DescExplosion)
  Set SupExplosion(1) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo1.bmp", DescExplosion)
  Set SupExplosion(2) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo2.bmp", DescExplosion)
  Set SupExplosion(3) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo3.bmp", DescExplosion)
  Set SupExplosion(4) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo4.bmp", DescExplosion)
  Set SupExplosion(5) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo5.bmp", DescExplosion)
  Set SupExplosion(6) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo6.bmp", DescExplosion)
  Set SupExplosion(7) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo7.bmp", DescExplosion)
  Set SupExplosion(8) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo8.bmp", DescExplosion)
  Set SupExplosion(9) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo9.bmp", DescExplosion)
  Set SupExplosion(10) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo10.bmp", DescExplosion)
  Set SupExplosion(11) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo11.bmp", DescExplosion)
  Set SupExplosion(12) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo12.bmp", DescExplosion)
  Set SupExplosion(13) = DD.CreateSurfaceFromFile(App.Path & "\imagenes\explo13.bmp", DescExplosion)
  
  SupExplosion(0).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(1).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(2).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(3).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(4).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(5).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(6).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(7).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(8).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(9).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(10).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(11).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(12).SetColorKey DDCKEY_SRCBLT, Key
  SupExplosion(13).SetColorKey DDCKEY_SRCBLT, Key
      
  '_________________________________________________________________________________________
  
  ' Creación de la superficie para los tiles de la carretera del nivel 1
  With DescCarretera1
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
  End With
    
  Set Carretera1 = DD.CreateSurfaceFromFile(App.Path & "\imagenes\carretera1.bmp", DescCarretera1)
  'Set Carretera2 = DD.CreateSurfaceFromFile(App.Path & "\imagenes\carretera2.bmp", DescCarretera1)

  ' Hacemos que el Paisaje tenga color transparente (key color)
  
  Carretera1.SetColorKey DDCKEY_SRCBLT, Key
      
  '_________________________________________________________________________________________
  
  
  ' Creación de la superficie para los tiles del paisaje del nivel x
  With DescPaisaje
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    
  End With
  
  
  
  '_________________________________________________________________________________________
  
  ' El paisaje del nivel 1 que se muestra en pantalla
  With DescPaisaje1Pantalla
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    .lWidth = 240
    .lHeight = 32640
  End With
  
  Set PaisajePantalla = DD.CreateSurface(DescPaisaje1Pantalla)
  
  '_________________________________________________________________________________________
  
  ' La carretera del nivel 1 que se muestra en pantalla
   
  With DescPaisajeCarretera
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    .lWidth = 160
    '.lWidth = 240
    .lHeight = 32640 '(255 filas * 128 pixels de alto)
  End With
  
  Set PaisajeCarretera = DD.CreateSurface(DescPaisajeCarretera)
  
  PaisajeCarretera.SetColorKey DDCKEY_SRCBLT, Key
    
  
  '_________________________________________________________________________________________
  
  ' Presentacion de inicio del nivel 2
  With DescNivel
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    .lWidth = 218
    .lHeight = 47
  End With
  Set SupNivel = DD.CreateSurface(DescNivel)
  
  '_________________________________________________________________________________________
  
  ' Creación de la superficie que presenta la informacion del estado: puntos, record, meta
  With DescFondo
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    .lWidth = 80
    .lHeight = DescSec.lHeight
  End With
  
  Set SupFondo = DD.CreateSurfaceFromFile(App.Path & "\imagenes\fondo1.bmp", DescFondo)
      
  '_________________________________________________________________________________________

  ' Creacion de la superficie que representa el trayecto recorrido
  With DescPunteroCoche
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CKSRCBLT
    .ddckCKSrcBlt.low = 0
    .ddckCKSrcBlt.high = 0
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
    .lWidth = 4
    .lHeight = 4
  End With
  
  Set SupPunteroCoche = DD.CreateSurfaceFromFile(App.Path & "\imagenes\puntero_car.bmp", DescPunteroCoche)
  
  SupPunteroCoche.SetColorKey DDCKEY_SRCBLT, Key
  
  
  '_________________________________________________________________________________________
    
  ' Creacion de la superficie que representa la velocidad
  With DescVelocidad
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
    .lWidth = 17
    .lHeight = 70
  End With
  
  Set SupVelocidad = DD.CreateSurfaceFromFile(App.Path & "\imagenes\velocidad.bmp", DescVelocidad)
  Set SupAntiVelocidad = DD.CreateSurfaceFromFile(App.Path & "\imagenes\antivelocidad.bmp", DescVelocidad)
  Set SupCombustible = DD.CreateSurfaceFromFile(App.Path & "\imagenes\combustible.bmp", DescVelocidad)
      
      
  '_________________________________________________________________________________________
      
      
  ' Creación de la superficie para los coches
  With DescCoche
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CKSRCBLT
    .ddckCKSrcBlt.low = 0
    .ddckCKSrcBlt.high = 0
    
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
    .lWidth = MAX_CAR_WIDTH
    .lHeight = MAX_CAR_HEIGHT
  End With

  Set coche.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\coche1.bmp", DescCoche)
  Set CocheMalo1.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\cochemalo1.bmp", DescCoche)
  Set CocheMalo2.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\cochemalo2.bmp", DescCoche)
  Set CocheMalo3.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\cochemalo3.bmp", DescCoche)
  Set CocheMalo4.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\cochemalo4.bmp", DescCoche)
  
  coche.Superficie.SetColorKey DDCKEY_SRCBLT, Key
  
  For i = 0 To 6
    Set CocheColision(i).Superficie = DD.CreateSurfaceFromFile(App.Path & _
                                   "\imagenes\coche" & Chr(i + 65) & ".bmp", DescCoche)
    
    Set CocheMalo1Colision(i).Superficie = DD.CreateSurfaceFromFile(App.Path & _
                                   "\imagenes\cochemalo1" & Chr(i + 65) & ".bmp", DescCoche)
  
    Set CocheMalo2Colision(i).Superficie = DD.CreateSurfaceFromFile(App.Path & _
                                   "\imagenes\cochemalo2" & Chr(i + 65) & ".bmp", DescCoche)
    
    Set CocheMalo3Colision(i).Superficie = DD.CreateSurfaceFromFile(App.Path & _
                                   "\imagenes\cochemalo3" & Chr(i + 65) & ".bmp", DescCoche)
  
    Set CocheMalo4Colision(i).Superficie = DD.CreateSurfaceFromFile(App.Path & _
                                   "\imagenes\cochemalo4" & Chr(i + 65) & ".bmp", DescCoche)
  
  Next i
  
  Set CocheColision(7).Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\coche1.bmp", DescCoche)
  Set CocheMalo1Colision(7).Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\cochemalo1.bmp", DescCoche)
  Set CocheMalo2Colision(7).Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\cochemalo2.bmp", DescCoche)
  Set CocheMalo3Colision(7).Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\cochemalo3.bmp", DescCoche)
  Set CocheMalo4Colision(7).Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\cochemalo4.bmp", DescCoche)
  

  '_________________________________________________________________________________________
  
  ' Creación de la superficie para el humo del tubo de escape
  With DescHumoTuboEscape
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CKSRCBLT
    .ddckCKSrcBlt.low = 0
    .ddckCKSrcBlt.high = 0
    
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
    .lWidth = 7
    .lHeight = 6
    
  End With
 
  Set HumoTuboEscape(0).Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\humotuboescape.bmp", DescHumoTuboEscape)
  Set HumoTuboEscape(1).Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\humoturbo.bmp", DescHumoTuboEscape)
  
  
  HumoTuboEscape(0).Superficie.SetColorKey DDCKEY_SRCBLT, Key
   

  '_________________________________________________________________________________________
  
  
  ' Creación de la superficie para el camion
  With DescCamion
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CKSRCBLT
    .ddckCKSrcBlt.low = 0
    .ddckCKSrcBlt.high = 0
    
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
    .lWidth = 20
    .lHeight = 62
  End With
  
  Set Camion.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\camion.bmp", DescCamion)
  
  Camion.Superficie.SetColorKey DDCKEY_SRCBLT, Key
  
  '_________________________________________________________________________________________
  
  
  With DescCamion2
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CKSRCBLT
    .ddckCKSrcBlt.low = 0
    .ddckCKSrcBlt.high = 0
    
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
    .lWidth = 20
    .lHeight = 92
  End With
  
  Set Camion2.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\camion2.bmp", DescCamion2)
  
  Camion2.Superficie.SetColorKey DDCKEY_SRCBLT, Key
  
  '_________________________________________________________________________________________
  
  
  ' Creación de la superficie para la recarga de combustible
  With DescRecargaCombustible
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CKSRCBLT
    .ddckCKSrcBlt.low = 0
    .ddckCKSrcBlt.high = 0
    
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN + DDSCAPS_VIDEOMEMORY
    .lWidth = 19
    .lHeight = 17
  End With

  Set RecargaCombustible.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\fuel.bmp", DescRecargaCombustible)
  
  RecargaCombustible.Superficie.SetColorKey DDCKEY_SRCBLT, Key
     
  Exit Sub

errores:
    CerrarDD "InicializarSuperficiesSprites"


End Sub

Public Sub FinDelJuego()
  
  Dim rectOrig As RECT, rectDest As RECT, inicio As Variant
 
  ' Parar la musica del nivel en el que se esté
  perf.Stop seg(NivelActual), segstate, 0, 0
  sb(marcha).Stop
  
  ' Reproducir melodia de fin del juego
  frmPrincipal.ReproducirSonido gameover
  
    With rectOrig
      .Left = 0:  .Top = 0:  .Right = DescNivel.lWidth: .Bottom = DescNivel.lHeight
    End With
    
    With rectDest
      .Left = 0:  .Top = 0:  .Right = MAX_WIDTH: .Bottom = MAX_HEIGHT
    End With
    
    Set SupNivel = DD.CreateSurfaceFromFile(App.Path & "\imagenes\fin.bmp", DescNivel)
        
    ' Pintar el fondo de negro
    SupBackBuffer.BltColorFill rectDest, 0
    
    SupBackBuffer.BltFast MAX_WIDTH / 2 - DescNivel.lWidth / 2, MAX_HEIGHT / 2 - DescNivel.lHeight / 2, SupNivel, rectOrig, DDBLTFAST_WAIT
        
    supPrimary.Flip SupBackBuffer, DDFLIP_WAIT
    
         
    ' Mostrar la informacion de cambio de nivel durante 5 segundos
    inicio = Timer
    
    Do While Timer < inicio + 5
      DoEvents
    Loop
   
  CerrarDD
  
End Sub
Public Sub CelebrarFinNivel()

  Dim segX As DirectSoundBuffer

  ' Parar la musica del nivel1
  perf.Stop seg(NIVEL1), segstate, 0, 0
  
  ' Parar el sonido de la marcha
  sb(marcha).Stop
  
  ' Reproducir el sonido de trayecto completado
  Set segX = DS.CreateSoundBufferFromFile(App.Path & "\Sonido\nivelcompletado.wav", sbd, wf)
  segX.Play DSBPLAY_DEFAULT
  
  ' Parar la ejecucion del juego hasta que la melodia finalize
  Do
    DoEvents
  Loop Until segX.GetStatus <> DSBSTATUS_PLAYING
    
  ' Liberar el objeto temporal de la melodia
  Set segX = Nothing

End Sub
Public Sub DibujarCambioNivel()
  
  Dim rectOrig As RECT, rectDest As RECT, inicio As Variant
 
  With rectOrig
    .Left = 0:  .Top = 0:  .Right = DescNivel.lWidth: .Bottom = DescNivel.lHeight
  End With
    
  With rectDest
    .Left = 0:  .Top = 0:  .Right = MAX_WIDTH: .Bottom = MAX_HEIGHT
  End With
    
  ' Pintar el fondo de negro
  SupBackBuffer.BltColorFill rectDest, 0
    
  Select Case NivelActual
    Case NIVEL1
      Set SupNivel = DD.CreateSurfaceFromFile(App.Path & "\imagenes\nivel1.bmp", DescNivel)
    Case NIVEL2
      Set SupNivel = DD.CreateSurfaceFromFile(App.Path & "\imagenes\nivel2.bmp", DescNivel)
    Case NIVEL3
      Set SupNivel = DD.CreateSurfaceFromFile(App.Path & "\imagenes\nivel3.bmp", DescNivel)
    Case NIVEL4
      Set SupNivel = DD.CreateSurfaceFromFile(App.Path & "\imagenes\nivel4.bmp", DescNivel)
    Case NIVEL5
      Set SupNivel = DD.CreateSurfaceFromFile(App.Path & "\imagenes\nivel5.bmp", DescNivel)
    Case NIVEL6
      Set SupNivel = DD.CreateSurfaceFromFile(App.Path & "\imagenes\nivel6.bmp", DescNivel)
  End Select
    
  SupBackBuffer.BltFast MAX_WIDTH / 2 - DescNivel.lWidth / 2, MAX_HEIGHT / 2 - DescNivel.lHeight / 2, SupNivel, rectOrig, DDBLTFAST_WAIT
  supPrimary.Flip SupBackBuffer, DDFLIP_WAIT
    
  ' Mostrar la informacion de cambio de nivel durante 5 segundos
  'inicio = Timer
   
  'Do While Timer < inicio + 3
  '  DoEvents
  'Loop
    
End Sub
Private Sub get_Paisaje(ByVal bloq As String, ByRef x1 As Integer, ByRef y1 As Integer, ByRef x2 As Integer, ByRef y2 As Integer)

  ' Devolver las coordenadas de un rectángulo de 16x16

Select Case NivelActual

Case NIVEL1
  
  Select Case bloq
  
    Case "1" ' árbol
      x1 = 0: y1 = 0: x2 = 16: y2 = 16
    
    Case "2" ' hierba
      x1 = 16: y1 = 0: x2 = 32: y2 = 16

    Case "3" ' lago; esquina izquierda superior
      x1 = 32: y1 = 0: x2 = 48: y2 = 16

    Case "4" ' lago; parte superior central
      x1 = 48: y1 = 0: x2 = 64: y2 = 16
      
    Case "5" ' lago; parte izquierda central
      x1 = 64: y1 = 0: x2 = 80: y2 = 16
      
    Case "6" ' lago; parte media central
      x1 = 80: y1 = 0: x2 = 96: y2 = 16
      
    Case "7" ' lago; esquina izquierda inferior
      x1 = 96: y1 = 0: x2 = 112: y2 = 16
      
    Case "8" ' lago; parte inferior central
      x1 = 112: y1 = 0: x2 = 128: y2 = 16
      
    Case "9" ' casa azul; esquina izquierda superior
      x1 = 128: y1 = 0: x2 = 144: y2 = 16
      
    Case "A" ' casa azul; esquina derecha superior
      x1 = 144: y1 = 0: x2 = 160: y2 = 16
      
    Case "B" ' casa azul; parte central izquierda
      x1 = 160: y1 = 0: x2 = 176: y2 = 16
      
    Case "C" ' casa azul; parte central derecha
      x1 = 176: y1 = 0: x2 = 192: y2 = 16
      
    Case "D" ' casa azul; esquina inferior izquierda
      x1 = 192: y1 = 0: x2 = 208: y2 = 16
      
    Case "E" ' casa azul; esquina inferior derecha
      x1 = 208: y1 = 0: x2 = 224: y2 = 16
      
    Case "F" ' casa roja; parte superior (rojo oscuro)
      x1 = 0: y1 = 16: x2 = 16: y2 = 32
      
    Case "G" ' casa roja; parte central
      x1 = 16: y1 = 16: x2 = 32: y2 = 32
      
    Case "H" ' casa roja; parte inferior
      x1 = 32: y1 = 16: x2 = 48: y2 = 32
      
    Case "I" ' lago; esquina superior derecha
      x1 = 48: y1 = 16: x2 = 64: y2 = 32
      
    Case "J" ' puente tapado; segunda parte central
      x1 = 64: y1 = 16: x2 = 80: y2 = 32
      
    Case "K" ' puente tapado; parte con columna derecha
      x1 = 80: y1 = 16: x2 = 96: y2 = 32
      
    Case "L" ' puente tapado; parte derecha
      x1 = 96: y1 = 16: x2 = 112: y2 = 32
      
    Case "M" ' puente; segunda parte central
      x1 = 112: y1 = 16: x2 = 128: y2 = 32
      
    Case "N" ' puente; parte derecha
      x1 = 128: y1 = 16: x2 = 144: y2 = 32
      
    Case "O" ' puente; parte derecha
      x1 = 144: y1 = 16: x2 = 160: y2 = 32
      
    Case "P" ' lago; parte lateral derecha
      x1 = 160: y1 = 16: x2 = 176: y2 = 32
      
    Case "R" ' lago; esquina inferior derecha
      x1 = 176: y1 = 16: x2 = 192: y2 = 32
      
 End Select

Case NIVEL2
 
  Select Case bloq
  
    Case "1"
      x1 = 0: y1 = 0: x2 = 16: y2 = 16
    
    Case "2"
      x1 = 16: y1 = 0: x2 = 32: y2 = 16

    Case "3"
      x1 = 32: y1 = 0: x2 = 48: y2 = 16

    Case "4"
      x1 = 48: y1 = 0: x2 = 64: y2 = 16
      
    Case "5"
      x1 = 64: y1 = 0: x2 = 80: y2 = 16
      
    Case "6"
      x1 = 80: y1 = 0: x2 = 96: y2 = 16
      
    Case "7"
      x1 = 96: y1 = 0: x2 = 112: y2 = 16
      
    Case "8"
      x1 = 112: y1 = 0: x2 = 128: y2 = 16
      
    Case "9"
      x1 = 128: y1 = 0: x2 = 144: y2 = 16
      
    Case "A"
      x1 = 144: y1 = 0: x2 = 160: y2 = 16
      
    Case "B" ' Ultimo bloque de la fila 1
      x1 = 160: y1 = 0: x2 = 176: y2 = 16
      
    Case "C"
      x1 = 0: y1 = 16: x2 = 16: y2 = 32
      
    Case "D"
      x1 = 16: y1 = 16: x2 = 32: y2 = 32
      
    Case "E"
      x1 = 32: y1 = 16: x2 = 48: y2 = 32
      
    Case "F"
      x1 = 48: y1 = 16: x2 = 64: y2 = 32
      
    Case "G"
      x1 = 64: y1 = 16: x2 = 80: y2 = 32
      
    Case "H"
      x1 = 80: y1 = 16: x2 = 96: y2 = 32
      
    Case "I"
      x1 = 96: y1 = 16: x2 = 112: y2 = 32
      
    Case "J"
      x1 = 112: y1 = 16: x2 = 128: y2 = 32
      
    Case "K"
      x1 = 128: y1 = 16: x2 = 144: y2 = 32

  End Select
  
End Select

End Sub
Private Sub get_rectangulo(ByVal bloq As String, ByRef x1 As Integer, ByRef y1 As Integer, ByRef x2 As Integer, ByRef y2 As Integer)

  ' Devolver las coordenadas de un rectángulo de 16x128

  Select Case bloq
  
    Case "1" ' no carretera (negro)
      x1 = 0: y1 = 0: x2 = 16: y2 = 128
    
    Case "2" ' borde izquierdo recto
      x1 = 16: y1 = 0: x2 = 32: y2 = 128
      
    Case "3" ' asfalto limpio
      x1 = 32: y1 = 0: x2 = 48: y2 = 128

    Case "4" ' asfalto raya recta
      x1 = 48: y1 = 0: x2 = 64: y2 = 128
      
    Case "5" ' borde derecho recto
      x1 = 64: y1 = 0: x2 = 80: y2 = 128
      
    Case "6" ' borde izquierdo girando a izquierda (1/2)
      x1 = 80: y1 = 0: x2 = 96: y2 = 128
      
    Case "7" ' borde izquierdo girando a izquierda (2/2)
      x1 = 96: y1 = 0: x2 = 112: y2 = 128
      
    Case "8" ' asfalto girando a izquierda (1/2)
      x1 = 112: y1 = 0: x2 = 128: y2 = 128
      
    Case "9" ' asfalto girando a izquierda (2/2)
      x1 = 128: y1 = 0: x2 = 144: y2 = 128
      
    Case "A" ' borde derecho girando a izquierda (1/2)
      x1 = 144: y1 = 0: x2 = 160: y2 = 128
      
    Case "B" ' borde derecho girando a izquierda (2/2)
      x1 = 160: y1 = 0: x2 = 176: y2 = 128
      
    Case "C" ' borde izquierdo girando a derecha (1/2)
      x1 = 176: y1 = 0: x2 = 192: y2 = 128
      
    Case "D" ' borde izquierdo girando a derecha (2/2)
      x1 = 192: y1 = 0: x2 = 208: y2 = 128
      
    Case "E" ' asfalto girando a derecha (1/2)
      x1 = 208: y1 = 0: x2 = 224: y2 = 128
      
    Case "F" ' asfalto girando a derecha (2/2)
      x1 = 224: y1 = 0: x2 = 240: y2 = 128
      
    Case "G" ' borde derecho girando a derecha (1/2)
      x1 = 240: y1 = 0: x2 = 256: y2 = 128
      
    Case "H" ' borde derecho girando a derecha (2/2)
      x1 = 256: y1 = 0: x2 = 272: y2 = 128
      
    Case "&" ' Mancha de aceite
      x1 = 272: y1 = 0: x2 = 288: y2 = 24
      
    Case "$" ' Mancha de agua
      x1 = 272: y1 = 24: x2 = 288: y2 = 48
      
    Case "%" ' Roca
      x1 = 272: y1 = 48: x2 = 288: y2 = 72
      
 End Select

End Sub
Private Function LeerRecord() As String
  
On Error GoTo errores

  Dim buffer As String, calculos As Currency
   
  Open App.Path & "\dat\dat" For Input As #1
  Line Input #1, buffer
  
  Close #1
  
  If buffer <> "" Then
    calculos = CDbl(buffer) + 23
    calculos = calculos / 5
    calculos = calculos * calculos
    buffer = CStr(CLng(calculos))
    LeerRecord = buffer
  Else
    LeerRecord = "0"
  End If
  
  Exit Function
  
errores:
  Record = 0
End Function
Public Sub PrepararNuevoNivel()
  
  ' Promoción al siguiente nivel
  Select Case NivelActual
    Case NIVEL0
      NivelActual = NIVEL1
    Case NIVEL1
      NivelActual = NIVEL2
    Case NIVEL2
      NivelActual = NIVEL3
    Case NIVEL3
      NivelActual = NIVEL4
    Case NIVEL4
      NivelActual = NIVEL5
    Case NIVEL5
      NivelActual = NIVEL6
  End Select
     
  DibujarCambioNivel
  PrepararJuego

End Sub
Public Sub GenerarFrecuenciaMostrarCamion()

' Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)

' Obtener un numero aleatorio entre 1 minuto y 2 minutos
TiempoNecesarioParaMostrarCamion = Int((45 - 30 + 1) * Rnd + 30)
'TiempoNecesarioParaMostrarCamion = Int((7 - 3 + 1) * Rnd + 3)

End Sub

Public Function GenerarFrecuenciaMostrarCocheX(ByVal Vehiculo As Integer)

' Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)

' De mas breve a menos breve
If Vehiculo = 1 Then
  GenerarFrecuenciaMostrarCocheX = Int((4 - 2 + 1) * Rnd + 2)
ElseIf Vehiculo = 2 Then
  GenerarFrecuenciaMostrarCocheX = Int((5 - 3 + 1) * Rnd + 3)
ElseIf Vehiculo = 3 Then
  GenerarFrecuenciaMostrarCocheX = Int((4 - 2 + 1) * Rnd + 2)
Else
  GenerarFrecuenciaMostrarCocheX = Int((5 - 2 + 1) * Rnd + 2)
End If

End Function

Public Sub GuardarRecord()

On Error GoTo errores

  Dim buffer As String, calculos As Currency
  
  calculos = ((Sqr(puntos)) * 5) - 23
  
  buffer = CStr(calculos)
  
  Open App.Path & "\dat\dat" For Output As #1
  Print #1, buffer
  Close #1

Exit Sub

errores:

CerrarDD "GuardarRecord"

End Sub
Public Sub PararMusica(ByVal nivel As Integer)
  If Not (perf Is Nothing) Then
    Call perf.Stop(seg(nivel), segstate, 0, 0) ' Parar la musica
  End If
End Sub
Public Sub PrepararJuego()
  
  InicializarVariables
  
  ' Llamar a las funciones que obtendran aleatoriamente, la frecuencia con la que deben
  ' aparecer los camiones
  GenerarFrecuenciaMostrarCamion
  GenerarFrecuenciaMostrarCamion2

  ' Dos coches enemigos deben aparecer al principio. Aqui, aleatoriamente se escojen dos
  DecidirParCochesIniciales

  Select Case NivelActual
    Case NIVEL1
      CargarMapaCarretera1
      DescPaisaje.lWidth = 224: DescPaisaje.lHeight = 32
      Set Paisaje = DD.CreateSurfaceFromFile(App.Path & "\imagenes\Paisaje1.bmp", DescPaisaje)
      CargarMapaP1
    Case NIVEL2
      CargarMapaCarretera2
      DescPaisaje.lWidth = 176: DescPaisaje.lHeight = 32
      Set Paisaje = DD.CreateSurfaceFromFile(App.Path & "\imagenes\Paisaje2.bmp", DescPaisaje)
      CargarMapaP2
    Case NIVEL3
      CargarMapaCarretera3
      CargarMapaP1
    Case NIVEL4
      CargarMapaCarretera4
      CargarMapaP1
    Case NIVEL5
      'CargarMapaCarretera5
      CargarMapaP1
    Case NIVEL6
      'CargarMapaCarretera6
      CargarMapaP1
  End Select

  ReproducirMusica NivelActual ' Reproducir la melodia del nivel 2
  
  ' El jugador arranca el motor
  frmPrincipal.ReproducirSonido carstart
    
   
End Sub
Public Sub GenerarFrecuenciaMostrarCamion2()

' Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)

' Obtener un numero aleatorio entre 1 minuto y 2 minutos
TiempoNecesarioParaMostrarCamion2 = Int((60 - 45 + 1) * Rnd + 45)
'TiempoNecesarioParaMostrarCamion2 = Int((10 - 7 + 1) * Rnd + 7)
End Sub
'
' EN ESTE NIVEL, APARECE EL OBSTACULO ACEITE
'
Private Sub CargarMapaCarretera2()


On Error GoTo errores

 Dim X As Integer, Y As Integer
 Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
 Dim rSprite As RECT
 Dim bloq As String * 1, bloq_anterior As String * 1

' Una vez leido el mapa desde el disco, se sustituyen dos 'tiles' de asfalto por sendas
' manchas de aceite
ReemplazarAsfaltoPorObstaculo ASFALTO_LISO, MANCHA_ACEITE

' Se barre el mapa y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_ROAD - 1
  
    bloq = Mid$(Mapa(Y), X + 1, 1)
     
    ' Si el tile es la mancha de aceite, hay que pegarla sobre un trozo de asfalto liso
    If bloq = MANCHA_ACEITE Then
      
      ' Primero pinto un trozo de asfalto liso en la parte de carretera que toca
      get_rectangulo ASFALTO_LISO, x1, y1, x2, y2
      With rSprite
        .Left = x1: .Right = x2: .Top = y1: .Bottom = y2
      End With
                
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
        
      ' Luego, sobre el trozo de asfalto pinto la macha de aceite
      get_rectangulo MANCHA_ACEITE, x1, y1, x2, y2
      rSprite.Left = x1: rSprite.Right = x2: rSprite.Top = y1: rSprite.Bottom = y2
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
      bloq_anterior = MANCHA_ACEITE
    Else
     
      ' Solo obtener las coordenadas de la imagen del mapa si es diferente a la anterior
      If bloq <> bloq_anterior Then
        bloq_anterior = bloq
        get_rectangulo bloq, x1, y1, x2, y2
        With rSprite
          .Left = x1:  .Right = x2:   .Top = y1:   .Bottom = y2
        End With
      End If
          
      ' Ir pintando la carretera con las extracciones que hacemos de los tiles
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
    End If
    
  Next X

Next Y

Exit Sub

errores:
CerrarDD "CargarMapaCarretera2"

End Sub
Private Sub CoordenadasObstaculosAleatorias(ByRef X As Integer, ByRef Y As Integer)

  ' Un obstáculo no puede estar fuera de la carretera ni enlos bordes, sólo en el asfalto
  
  ' La Y deber estar entre 16 y la fila mas alta del mapa menos 16
  ' NOTA de la Y: menos 16 para que no aparezcan de golpe, ni en la primera pantalla,
  ' ni en la ultima
  
  'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)
  
  X = Int((MAX_COLS_ROAD - 2 + 1) * Rnd + 2)
  Y = Int((MAX_ROWS_MAP - 16 - 16 + 1) * Rnd + 16)
  
End Sub

Private Sub ReemplazarAsfaltoPorObstaculo(ByVal Asfalto As String, ByVal Obstaculo As String)

' Este procedimiento busca aleatoriamente un trozo de asfalto de la carretera que se
' ajuste al primer parametro que recibe y lo sustituye en la matriz

Dim col As Integer, fila As Integer, tile As String * 1, nueva_cadena As String
Dim fin_bucle As Boolean

' Buscar el primer trozo de asfalto para la sustitución

' Repetir el bucle hasta que se consiga un trozo de asfalto válido
Do
   ' Obtener una posicion aleatoria del Mapa
   CoordenadasObstaculosAleatorias col, fila
   
   ' Evaluar la posición obtenida.
   tile = Mid$(Mapa(fila), col, 1)
   
   ' Si se ha conseguido el objetivo, reemplazar el tile y salir del bucle
   If tile = Asfalto Then
     nueva_cadena = ReemplazarCaracterEnMatriz(Mapa(fila), col, Obstaculo)
     Mapa(fila) = nueva_cadena
     
     If Obstaculo = MANCHA_AGUA Then
       With ObjetoCharco(0)
         .PosXMapa = col + 3 - 1 ' El +3 es necesario porque los 3 primeros tiles son del paisaje
         .PosYMapa = fila
     
         ' Ahora que se tienen las coordenas del mapa se obtienen las coordenadas que
         ' corresponden a la pantalla (en la deteccion de colision seran muy utiles)
         .PosXPantalla = .PosXMapa * 16
         .PosYPantalla = .PosYMapa * 8
       End With
     ElseIf Obstaculo = MANCHA_ACEITE Then
       With ObjetoManchaAceite(0)
         .PosXMapa = col + 3 - 1 ' El +3 es necesario porque los 3 primeros tiles son del paisaje
         .PosYMapa = fila
     
         ' Ahora que se tienen las coordenas del mapa se obtienen las coordenadas que
         ' corresponden a la pantalla (en la deteccion de colision seran muy utiles)
         .PosXPantalla = .PosXMapa * 16
         .PosYPantalla = .PosYMapa * 8
       End With
     ElseIf Obstaculo = ROCA Then
       With ObjetoRoca(0)
         .PosXMapa = col + 3 - 1 ' El +3 es necesario porque los 3 primeros tiles son del paisaje
         .PosYMapa = fila
     
         ' Ahora que se tienen las coordenas del mapa se obtienen las coordenadas que
         ' corresponden a la pantalla (en la deteccion de colision seran muy utiles)
         .PosXPantalla = .PosXMapa * 16
         .PosYPantalla = .PosYMapa * 8
       End With
     End If
     
     fin_bucle = True
   End If
   
Loop Until fin_bucle

fin_bucle = False

' Buscar el SEGUNDO trozo de asfalto para la sustitución

' Repetir el bucle hasta que se consiga un trozo de asfalto válido
Do
   ' Obtener una posicion aleatoria del Mapa
   CoordenadasObstaculosAleatorias col, fila
   
   ' Evaluar la posición obtenida.
   tile = Mid$(Mapa(fila), col, 1)
   
   ' Si se ha conseguido el objetivo, reemplazar el tile y salir del bucle
   If tile = Asfalto Then
     nueva_cadena = ReemplazarCaracterEnMatriz(Mapa(fila), col, Obstaculo)
     Mapa(fila) = nueva_cadena
     
     If Obstaculo = MANCHA_AGUA Then
       With ObjetoCharco(1)
         .PosXMapa = col + 3 - 1 ' El +3 es necesario porque los 3 primeros tiles son del paisaje
         .PosYMapa = fila
     
         ' Ahora que se tienen las coordenas del mapa se obtienen las coordenadas que
         ' corresponden a la pantalla (en la deteccion de colision seran muy utiles)
         .PosXPantalla = .PosXMapa * 16
         .PosYPantalla = .PosYMapa * 8
       End With
     ElseIf Obstaculo = MANCHA_ACEITE Then
       With ObjetoManchaAceite(1)
         .PosXMapa = col + 3 - 1 ' El +3 es necesario porque los 3 primeros tiles son del paisaje
         .PosYMapa = fila
     
         ' Ahora que se tienen las coordenas del mapa se obtienen las coordenadas que
         ' corresponden a la pantalla (en la deteccion de colision seran muy utiles)
         .PosXPantalla = .PosXMapa * 16
         .PosYPantalla = .PosYMapa * 8
       End With
     ElseIf Obstaculo = ROCA Then
       With ObjetoRoca(1)
         .PosXMapa = col + 3 - 1 ' El +3 es necesario porque los 3 primeros tiles son del paisaje
         .PosYMapa = fila
     
         ' Ahora que se tienen las coordenas del mapa se obtienen las coordenadas que
         ' corresponden a la pantalla (en la deteccion de colision seran muy utiles)
         .PosXPantalla = .PosXMapa * 16
         .PosYPantalla = .PosYMapa * 8
       End With
     End If
     
     fin_bucle = True
   End If
   
Loop Until fin_bucle


End Sub

Private Function ReemplazarCaracterEnMatriz(ByVal Origen As String, ByVal pos As Integer, ByVal Valor As String)

' Esta funcion reemplaza el parámetro 'Valor', en la posicion del parametro 'pos' del
' parámetro 'Origen'

Dim i As Integer, car As String * 1, temp As String

For i = 1 To Len(Origen)
  If i = pos Then
    car = Valor
  Else
    car = Mid$(Origen, i, 1)
  End If
  temp = temp & car
Next i

ReemplazarCaracterEnMatriz = temp

End Function
'
' EN ESTE NIVEL, APARECE EL OBSTACULO ACEITE Y EL AGUA
'
Private Sub CargarMapaCarretera3()


On Error GoTo errores

 Dim X As Integer, Y As Integer
 Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
 Dim rSprite As RECT
 Dim bloq As String * 1, bloq_anterior As String * 1
  
' Se sustituyen dos 'tiles' de asfalto por sendas manchas de agua
ReemplazarAsfaltoPorObstaculo ASFALTO_LISO, MANCHA_AGUA

' Se barre el mapa y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_ROAD - 1
  
    bloq = Mid$(Mapa(Y), X + 1, 1)
     
    ' Si el tile es la mancha de aceite, hay que pegarla sobre un trozo de asfalto liso
    If bloq = MANCHA_ACEITE Then
      
      ' Primero pinto un trozo de asfalto liso en la parte de carretera que toca
      get_rectangulo ASFALTO_LISO, x1, y1, x2, y2
      With rSprite
        .Left = x1: .Right = x2: .Top = y1: .Bottom = y2
      End With
                
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
        
      ' Luego, sobre el trozo de asfalto pinto la macha de aceite
      get_rectangulo MANCHA_ACEITE, x1, y1, x2, y2
      rSprite.Left = x1: rSprite.Right = x2: rSprite.Top = y1: rSprite.Bottom = y2
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
      bloq_anterior = MANCHA_ACEITE
    
    ' Si el tile es la mancha de agua, hay que pegarla sobre un trozo de asfalto liso
    ElseIf bloq = MANCHA_AGUA Then
      
      ' Primero pinto un trozo de asfalto liso en la parte de carretera que toca
      get_rectangulo ASFALTO_LISO, x1, y1, x2, y2
      With rSprite
        .Left = x1: .Right = x2: .Top = y1: .Bottom = y2
      End With
                
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
        
      ' Luego, sobre el trozo de asfalto pinto la macha de aceite
      get_rectangulo MANCHA_AGUA, x1, y1, x2, y2
      rSprite.Left = x1: rSprite.Right = x2: rSprite.Top = y1: rSprite.Bottom = y2
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
      bloq_anterior = MANCHA_AGUA
    
    Else
     
      ' Solo obtener las coordenadas de la imagen del mapa si es diferente a la anterior
      If bloq <> bloq_anterior Then
        bloq_anterior = bloq
        get_rectangulo bloq, x1, y1, x2, y2
        With rSprite
          .Left = x1:  .Right = x2:   .Top = y1:   .Bottom = y2
        End With
      End If
          
      ' Ir pintando la carretera con las extracciones que hacemos de los tiles
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
    End If
    
  Next X

Next Y

Exit Sub

errores:
CerrarDD "CargarMapaCarretera3"

End Sub
'
' EN ESTE NIVEL, APARECE EL OBSTACULO ACEITE, EL AGUA y LA ROCA
'
Private Sub CargarMapaCarretera4()


On Error GoTo errores

 Dim X As Integer, Y As Integer
 Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
 Dim rSprite As RECT
 Dim bloq As String * 1, bloq_anterior As String * 1
  
' Se sustituyen dos 'tiles' de asfalto por sendas manchas de agua
ReemplazarAsfaltoPorObstaculo ASFALTO_LISO, ROCA

' Se barre el mapa y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_ROAD - 1
  
    bloq = Mid$(Mapa(Y), X + 1, 1)
     
    ' Si el tile es la mancha de aceite, hay que pegarla sobre un trozo de asfalto liso
    If bloq = ROCA Then
      
      ' Primero pinto un trozo de asfalto liso en la parte de carretera que toca
      get_rectangulo ASFALTO_LISO, x1, y1, x2, y2
      With rSprite
        .Left = x1: .Right = x2: .Top = y1: .Bottom = y2
      End With
                
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
        
      ' Luego, sobre el trozo de asfalto pinto la macha de aceite
      get_rectangulo ROCA, x1, y1, x2, y2
      rSprite.Left = x1: rSprite.Right = x2: rSprite.Top = y1: rSprite.Bottom = y2
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
      bloq_anterior = ROCA
      
    ' Si el tile es la mancha de aceite, hay que pegarla sobre un trozo de asfalto liso
    ElseIf bloq = MANCHA_ACEITE Then
      
      ' Primero pinto un trozo de asfalto liso en la parte de carretera que toca
      get_rectangulo ASFALTO_LISO, x1, y1, x2, y2
      With rSprite
        .Left = x1: .Right = x2: .Top = y1: .Bottom = y2
      End With
                
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
        
      ' Luego, sobre el trozo de asfalto pinto la macha de aceite
      get_rectangulo MANCHA_ACEITE, x1, y1, x2, y2
      rSprite.Left = x1: rSprite.Right = x2: rSprite.Top = y1: rSprite.Bottom = y2
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
      bloq_anterior = MANCHA_ACEITE
    
    ' Si el tile es la mancha de agua, hay que pegarla sobre un trozo de asfalto liso
    ElseIf bloq = MANCHA_AGUA Then
      
      ' Primero pinto un trozo de asfalto liso en la parte de carretera que toca
      get_rectangulo ASFALTO_LISO, x1, y1, x2, y2
      With rSprite
        .Left = x1: .Right = x2: .Top = y1: .Bottom = y2
      End With
                
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
        
      ' Luego, sobre el trozo de asfalto pinto la macha de aceite
      get_rectangulo MANCHA_AGUA, x1, y1, x2, y2
      rSprite.Left = x1: rSprite.Right = x2: rSprite.Top = y1: rSprite.Bottom = y2
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
      bloq_anterior = MANCHA_AGUA
    
    Else
     
      ' Solo obtener las coordenadas de la imagen del mapa si es diferente a la anterior
      If bloq <> bloq_anterior Then
        bloq_anterior = bloq
        get_rectangulo bloq, x1, y1, x2, y2
        With rSprite
          .Left = x1:  .Right = x2:   .Top = y1:   .Bottom = y2
        End With
      End If
          
      ' Ir pintando la carretera con las extracciones que hacemos de los tiles
      PaisajeCarretera.BltFast (X * 16), (Y * 128), Carretera1, rSprite, DDBLTFAST_WAIT
    End If
    
  Next X

Next Y

Exit Sub

errores:
CerrarDD "CargarMapaCarretera4"

End Sub
Private Sub CargarMapaP2()

On Error GoTo errores

 Dim X As Integer, Y As Integer
 Dim x1 As Integer, x2 As Integer, y1 As Integer, y2 As Integer
 Dim rSprite As RECT
 Dim bloq As String * 1, bloq_anterior As String * 1
  
' Cargar la 1/8
LeerMapaDisco "map2a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), (Y * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y

'-----------------------------------------------------------------------------------------

' Cargar la 2/8
LeerMapaDisco "map2a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 254) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y


'-----------------------------------------------------------------------------------------

' Cargar la 3/8

LeerMapaDisco "map2a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 509) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y

'-----------------------------------------------------------------------------------------

' Cargar la 4/8

LeerMapaDisco "map2a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 764) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y


'-----------------------------------------------------------------------------------------

' Cargar la 5/8

LeerMapaDisco "map2a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 1019) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y

'-----------------------------------------------------------------------------------------

' Cargar la 6/8

LeerMapaDisco "map2a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 1274) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y


'-----------------------------------------------------------------------------------------

' Cargar la 7/8

LeerMapaDisco "map2a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 1529) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y


'-----------------------------------------------------------------------------------------

' Cargar la 8/8

LeerMapaDisco "map2a.mmp", MapaPaisaje
 
' Se barre el MapaPaisaje y se pasa a la superficie auxiliar
For Y = 0 To MAX_ROWS_MAP - 1
  
  For X = 0 To MAX_COLS_MAP - 1
  
     bloq = Mid$(MapaPaisaje(Y), X + 1, 1)
 
    ' Solo obtener las coordenadas de la imagen del MapaPaisaje si es diferente a la anterior
    If bloq <> bloq_anterior Then
      bloq_anterior = bloq
      get_Paisaje bloq, x1, y1, x2, y2
      
      With rSprite
        .Left = x1: .Right = x2:   .Top = y1: .Bottom = y2
      End With
    End If
          
    PaisajePantalla.BltFast (X * 16), ((Y + 1784) * 16), Paisaje, rSprite, DDBLTFAST_WAIT
    
  Next X

Next Y
    
'
' SITUAMOS LA CARRETERA SOBRE EL PAISAJE
'

With rSprite
  .Top = 0: .Left = 0: .Right = 160: .Bottom = 32640
End With
    
PaisajePantalla.BltFast POSX_INI_CAR, 0, PaisajeCarretera, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

Exit Sub

errores:
CerrarDD "CargarMapaP2"

End Sub
