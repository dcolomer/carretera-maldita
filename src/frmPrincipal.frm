VERSION 5.00
Begin VB.Form frmPrincipal 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then
  
  If Record <= puntos Then
    GuardarRecord
  End If
        
  CerrarDD "Form_KeyPress"
End If

End Sub

Private Sub Form_Load()

  ' Acciones que se realizan sólo una vez para todo el juego
  ResetInicial
  
  ' Rutina de inicilización general (se invocará en cada cambio de nivel)

' MANIPULACION
'NivelActual = NIVEL1
'LeerMapaDisco "carret1.mmp", Mapa
' FIN MANIPULACION

PrepararNuevoNivel

' MANIPULACION
'PrepararJuego
' FIN MANIPULACION

  ' Condición para la ejecución continua del bucle principal
  bEjecutandose = True
  
  ' Bucle Principal
  Do While bEjecutandose = True
    ActualizarJuego
    Blt
    DoEvents
  Loop
    
  CerrarDD "Form_Load"
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyRight Then TeclaDerecha = True
    If KeyCode = vbKeyLeft Then TeclaIzquierda = True
    If KeyCode = vbKeySpace Then BarraEspaciadora = True
   
    If LCase(Chr(KeyCode)) = "m" And CInt(vel) >= 10 Then FlagTurbo = True
    
    
    
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyRight Then TeclaDerecha = False
    If KeyCode = vbKeyLeft Then TeclaIzquierda = False
    If KeyCode = vbKeySpace Then BarraEspaciadora = False
    
    If LCase(Chr(KeyCode)) = "m" Then FlagTurbo = False
    
End Sub
Private Sub ActualizarJuego()
  
  ControlarFinTramo                 ' Controlar si se ha llegado al final del tramo
  ControlarTiempos                  ' Comprobaciones que deben hacerse cada 1 segundo
  ControlarColisiones               ' Monitorizar las colisiones entre vehículos
  ControlarBordesCarretera          ' Controlar que no se salga el coche
  
  ControlarMovimientos              ' Comprobar movimientos del jugador
  ControlarCochesMalos              ' Controlar IA vehículos enemigos
  
  ControlarAceleracion              ' Acciones derivadas de la aceleracion/desaceleracion
  MoverPaisaje                      ' Actualizar punteros al paisaje

  ControlarTurbo                    ' Controlar si usuario usa el turbo
  ControlarCombustible              ' Cuando se ha agotado o está a punto de agotarse
  ControlarRecargaCombustible       ' Comprobar si usuario merece recarga combustible
  
  If NivelActual > NIVEL1 Then
    ControlarObstaculos                 ' Controlar aceite, agua y rocas
  End If
  
End Sub

Private Sub Blt()

On Error GoTo errores
    
Dim rectOrig As RECT, rectDest As RECT
Dim Y As Long, X As Integer
Dim posX As Double, pos1X As Double, posY As Double, pos1Y As Double
      

' Pintamos los sprites en Back Buffer.
With SupBackBuffer
    
    ' -----------------------------------------------------------------------------
    ' CONTROL DEL REFRESCO DEL MARCADOR
    ' -----------------------------------------------------------------------------
    
    With rectOrig
      .Left = 0:  .Top = 0:  .Right = DescFondo.lWidth: .Bottom = DescFondo.lHeight
    End With
    
    .BltFast 240, 0, SupFondo, rectOrig, DDBLTFAST_WAIT
        
    ' Representar graficamente el trayecto recorrido
     
    With rectOrig
      .Left = 0:  .Top = 0:  .Right = 4:  .Bottom = 4
    End With
    
    With rectDest
      .Left = 36: .Top = POS_RECORRIDO_INICIAL - PosicionRecorrido: .Right = 40: .Bottom = .Top + 4
    End With
        
    SupFondo.Blt rectDest, SupPunteroCoche, rectOrig, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
                
    ' Representar graficamente la velocidad
  
    ' Primero pintar la antivelocidad
    With rectOrig
      .Left = 0:  .Top = 0:  .Right = 17:  .Bottom = 70
    End With
    
    With rectDest
      .Left = 16:  .Top = 160:   .Right = 33:   .Bottom = 230
    End With
    
    SupFondo.Blt rectDest, SupAntiVelocidad, rectOrig, DDBLT_WAIT
    
    ' Luego la velocidad
    With rectOrig
      .Left = 0:  .Top = 0:  .Right = 17
      If vel > 10 Then ' Turbo
        .Bottom = CInt(70)
      Else
        .Bottom = CInt(70 * vel / 10) ' 10 como valor de velocidad maxima
      End If
    End With
    
    With rectDest
      .Left = 16:  .Top = 230 - rectOrig.Bottom:  .Right = 33:  .Bottom = 230
    End With
        
    SupFondo.Blt rectDest, SupVelocidad, rectOrig, DDBLT_WAIT

    ' Representar graficamente el combustible
    
    ' Primero pintar el anticombustible
    With rectOrig
      .Left = 0:   .Top = 0:   .Right = 17:  .Bottom = 70
    End With
    
    With rectDest
      .Left = 42:  .Top = 160: .Right = 59:  .Bottom = 230
    End With
        
    ' Aprovechar la descripcion de superficie 'SupAntiVelocidad' para el anticombustible
    SupFondo.Blt rectDest, SupAntiVelocidad, rectOrig, DDBLT_WAIT
    
    ' Luego el combustible
    With rectOrig
      .Left = 0: .Top = 0: .Right = 17: .Bottom = CInt(70 * Combustible / COMBUSTIBLE_INICIAL)
    End With
    
    With rectDest
      .Left = 42: .Top = 230 - rectOrig.Bottom: .Right = 59: .Bottom = 230
    End With
        
    SupFondo.Blt rectDest, SupCombustible, rectOrig, DDBLT_WAIT
    
    ' -----------------------------------------------------------------------------
    ' FIN DEL CONTROL DEL REFRESCO DEL MARCADOR
    ' -----------------------------------------------------------------------------
    
    
    ' * * * *
    ' Dibujamos el Paisaje
    
     For Y = MAX_ROWS_SCREEN + RowIndex To 0 + RowIndex Step -1

       ' Cuando se acaba el tramo se pasa al siguiente nivel, o no.
       If RowIndex <= 0 Then
         Exit Sub
       End If

       With rectOrig
         .Top = (Y - 1) * 16: .Left = 0:  .Right = 240:  .Bottom = .Top + 16
       End With
       
       posY = (Y * 16) - (RowIndex * 16) - MapDspRow
       pos1X = posX + 16
       pos1Y = posY + 16

       '/////////////////// clipping inferior /////////////////
       If pos1Y >= 240 Then
         rectOrig.Bottom = rectOrig.Bottom - (pos1Y - 240)
       End If
       
       '/////////////////// clipping superior /////////////////
       If posY <= 0 Then
         rectOrig.Top = rectOrig.Top + (Abs(posY))
         posY = 0
       End If
       
       ' Pintar el paisaje
       .BltFast 0, posY, PaisajePantalla, rectOrig, DDBLTFAST_WAIT
                     
   Next Y
   
    
  ' Si no se está en una explosión, se pinta el coche del jugador
  If CocheActivo Then

    ' Después el coche
    With rectOrig
      .Left = 0:  .Top = 0:  .Right = coche.Ancho:  .Bottom = coche.Alto
    End With
        
    With rectDest
      .Left = coche.posX
      .Top = coche.posY
      .Right = coche.posX + coche.Ancho
      .Bottom = coche.posY + coche.Alto
    End With
    
    ' Dibujar el coche
    .Blt rectDest, coche.Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  
  End If

  ' Ahora el coche malo1
  If CocheMalo1Activo Then
    PintarEnemigos CocheMalo1, rectOrig, rectDest
    .Blt rectDest, CocheMalo1.Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If
    
  ' Ahora el coche malo2
  If CocheMalo2Activo Then
    PintarEnemigos CocheMalo2, rectOrig, rectDest
    .Blt rectDest, CocheMalo2.Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If
    
  ' Ahora el coche malo3
  If CocheMalo3Activo Then
    PintarEnemigos CocheMalo3, rectOrig, rectDest
    .Blt rectDest, CocheMalo3.Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If

  ' Ahora el coche malo4
  If CocheMalo4Activo Then
    PintarEnemigos CocheMalo4, rectOrig, rectDest
    .Blt rectDest, CocheMalo4.Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If

  ' Ahora el camion
  If CamionActivo Then
    PintarEnemigos Camion, rectOrig, rectDest
    .Blt rectDest, Camion.Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If

  ' Ahora el camion 2
  If Camion2Activo Then
    PintarEnemigos Camion2, rectOrig, rectDest
    .Blt rectDest, Camion2.Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If

  ' Dibujar la recarga de combustible (corazón)
  If RecargaCombustibleActiva Then
    PintarEnemigos RecargaCombustible, rectOrig, rectDest
    .Blt rectDest, RecargaCombustible.Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If

  ' Ahora el humo del tubo de escape
  If (vel < 1.5 And BarraEspaciadora And CocheActivo) Then
    PintarHumos HumoTuboEscape(0), rectOrig, rectDest
    .Blt rectDest, HumoTuboEscape(0).Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If

  ' Ahora el Turbo
  If FlagTurbo Then
    PintarHumos HumoTuboEscape(0), rectOrig, rectDest
    .Blt rectDest, HumoTuboEscape(1).Superficie, rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If


  ' Si toca la explosion del coche
  If Choque_en_Curso And VelocidadInicialColision >= 10 Then
    PintarExplosion coche, rectOrig, rectDest
    .Blt rectDest, SupExplosion(0), rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If

  If ChoqueCocheMalo1EnCurso And CocheMalo1ContraCamion Then
    PintarExplosion CocheMalo1, rectOrig, rectDest
    .Blt rectDest, SupExplosion(0), rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If
  
  If ChoqueCocheMalo2EnCurso And CocheMalo2ContraCamion Then
    PintarExplosion CocheMalo2, rectOrig, rectDest
    .Blt rectDest, SupExplosion(0), rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If
  
  If ChoqueCocheMalo3EnCurso And CocheMalo3ContraCamion Then
    PintarExplosion CocheMalo3, rectOrig, rectDest
    .Blt rectDest, SupExplosion(0), rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If
  
  If ChoqueCocheMalo4EnCurso And CocheMalo4ContraCamion Then
    PintarExplosion CocheMalo4, rectOrig, rectDest
    .Blt rectDest, SupExplosion(0), rectOrig, DDBLT_WAIT Or DDBLT_KEYSRC
  End If
  
  ' Cuadros por segundo
  .DrawText 5, 5, "FPS: " & FPS_Actual, False
  .DrawText 5, 25, "RI: " & RowIndex, False

  If SinCombustible Then ' FIN DEL JUEGO
    .DrawText 74, 60, "SIN COMBUSTIBLE", False
  End If
        
  ' Mostrar Record y Puntos en curso
    
  .DrawText 260, 12, Record, False
  .DrawText 260, 36, puntos, False
    
    
End With ' SupBackBuffer


  'Intercambio las superficies Primaria por Secundaria
  supPrimary.Flip SupBackBuffer, DDFLIP_WAIT
     
  Exit Sub
    
errores:
      CerrarDD "Blt"
    
End Sub


Private Sub Form_Paint()
  Blt
End Sub


Private Function HayInterseccion(ObjetoA As t_Sprite, ObjetoB As t_Sprite) As Boolean

Dim TempRect As RECT, rect1 As RECT, rect2 As RECT

With rect1
  .Top = ObjetoA.posY:  .Left = ObjetoA.posX
  .Right = .Left + ObjetoA.Ancho:  .Bottom = .Top + ObjetoA.Alto
End With

With rect2
  .Top = ObjetoB.posY:  .Left = ObjetoB.posX
  .Right = .Left + ObjetoB.Ancho:  .Bottom = .Top + ObjetoB.Alto
End With

HayInterseccion = IntersectRect(TempRect, rect1, rect2)

End Function

Private Function CoordenadaXAleatoria() As Integer
  'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)

CoordenadaXAleatoria = Int((limite_borde_derecho - limite_borde_izquierdo + 1) * Rnd + limite_borde_izquierdo)

End Function


Private Sub ControlarTiempos()

  ' Valor del tiempo actual
  Dim cur_time As Currency
  
  ' Controlar si ha pasado un segundo
  If (GetTickCount() - FPS_tUltimo) >= 1000 Then
      
    ' Ir quitando combustible
    'Combustible = Combustible - 0.5
    Combustible = Combustible - 0.1
    
    If BarraEspaciadora Then AcumularPuntos
    
    If Not CamionActivo Then ComprobarFrecuencias 0 'parámetro 0 = Camion
    If Not Camion2Activo Then ComprobarFrecuencias 1 'parámetro 1 = Camion2
    If Not CocheMalo1Activo Then ComprobarFrecuencias 2 'parámetro 2 = CocheMalo1Activo
    If Not CocheMalo2Activo Then ComprobarFrecuencias 3 'parámetro 3 = CocheMalo2Activo
    If Not CocheMalo3Activo Then ComprobarFrecuencias 4 'parámetro 4 = CocheMalo3Activo
    If Not CocheMalo4Activo Then ComprobarFrecuencias 5 'parámetro 5 = CocheMalo4Activo

    ' Si ya se ha acabado la explosion hay que activar la presencia del coche
    If Not CocheActivo And Not Choque_en_Curso Then
      CocheActivo = True
      ReproducirSonido carstart
    End If
    

    FPS_Actual = FPS_Suma
    FPS_Suma = 0
    
    'If TeclaDerecha Or TeclaIzquierda Then
    '    Set coche.Superficie = DD.CreateSurfaceFromFile(App.Path & "\imagenes\coche1.bmp", DescCoche)
    'End If

    FPS_tUltimo = GetTickCount()
  End If
  
  FPS_Suma = FPS_Suma + 1

' Establecer el tiempo base

' Leer el contador apropiado

If perf_flag Then
  QueryPerformanceCounter cur_time
Else
  cur_time = timeGetTime()
End If

time_span = (cur_time - last_time) * time_factor
last_time = cur_time

  ' Cada 25 unidades de tiempo se comprueba si se deben controlar los efectos de la
  ' colision de vehiculos
  If (cur_time - last_time2) >= 25 Then
    ControlarTiempoChoque
    last_time2 = cur_time
  End If

End Sub

Private Sub ControlarCombustible()
  If Combustible <= 0# Then
    sb(outoffuel).Stop
    FinDelJuego
  ElseIf Combustible <= 3# Then
    SinCombustible = True
    ReproducirSonido outoffuel, True
  End If
End Sub

Private Sub AcumularPuntos()
    
  If FlagTurbo Then puntos = puntos + 20 Else puntos = puntos + 10
    
  ' Controlar si la puntuacion actual supera el record
  If Record < puntos Then Record = puntos
  
End Sub

Private Sub ControlarCochesMalos()

' MANIPULACION
'CocheMalo1Activo = False
'CocheMalo2Activo = False
'CocheMalo3Activo = False
'CocheMalo4Activo = False
'CamionActivo = False
'Camion2Activo = False
' FIN MANIPULACION

If BarraEspaciadora And vel > 7# Then

  If CocheMalo1Activo Then
    ' Si aun no ha llegado al fondo de la pantalla
    If CocheMalo1.posY <= POSY_FIN_BAD_CAR Then
      
      With CocheMalo1
      
       ' Controlar que no se salga de la carretera
       If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
       If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
 
 
        ' Cuando el jugador adelanta tambien se debe reproducir sonido de adelantamiento
        If vel > 9# And (.posY > coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        
        ' Si aun no ha llegado a su velocidad máxima->incrementar su velocidad negativa
        If CInt(.Velocidad) <= 80 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        End If
      
        .posY = .posY + .Velocidad * time_span + .Aceleracion * time_span * time_span * 0.5
      
      End With
    ' Cuando ya ha llegado al final de la pantalla
    Else
      CocheMalo1Activo = False
      With CocheMalo1
        .posY = POSY_INI_BAD_CAR
        '.posY = POSY_INI_BAD_CAR * Rnd(-180)
        .posX = CoordenadaXAleatoria ' Con una X aleatoria
      End With
    
    End If
  End If
  
  If CocheMalo2Activo Then
     ' Si aun no ha llegado al fondo de la pantalla
    If CocheMalo2.posY <= POSY_FIN_BAD_CAR Then
      With CocheMalo2
        
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
       
        ' Cuando el juagador adelanta tambien se debe reproducir sonido de adelantamiento
        If vel > 9# And (.posY > coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        
        If CInt(.Velocidad) <= 80 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        End If
      
       .posY = .posY + .Velocidad * time_span + .Aceleracion * time_span * time_span * 0.5
      
      End With
    ' Cuando ya ha llegado al final de la pantalla
    Else
      CocheMalo2Activo = False
      With CocheMalo2
        '.posY = POSY_INI_BAD_CAR
        .posY = POSY_INI_BAD_CAR * Rnd(-420)
        .posX = CoordenadaXAleatoria 'CoordenadaXAleatoriaLateral ' Con una X aleatoria de lateral
      End With
    
    End If
  End If
  
  If CocheMalo3Activo Then
     ' Si aun no ha llegado al fondo de la pantalla
     If CocheMalo3.posY <= POSY_FIN_BAD_CAR Then
       With CocheMalo3
       
         ' Controlar que no se salga de la carretera
         If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
         If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
              
         ' Cuando el juagador adelanta tambien se debe reproducir sonido de adelantamiento
        If vel > 9# And (.posY > coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
         
         If CInt(.Velocidad) <= 80 Then
           .Velocidad = .Velocidad + .Aceleracion * time_span
         End If
         .posY = .posY + .Velocidad * time_span + .Aceleracion * time_span * time_span * 0.5
       End With
     ' Cuando ya ha llegado al final de la pantalla
     Else
       CocheMalo3Activo = False
       With CocheMalo3
        '.posY = POSY_INI_BAD_CAR
        .posY = POSY_INI_BAD_CAR * Rnd(-260)
        .posX = CoordenadaXAleatoria ' Con una X aleatoria
       End With
     End If
  End If
  
  If CocheMalo4Activo Then
    ' Si aun no ha llegado al fondo de la pantalla
    If CocheMalo4.posY <= POSY_FIN_BAD_CAR Then
      With CocheMalo4
      
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
      
        ' Cuando el juagador adelanta tambien se debe reproducir sonido de adelantamiento
        If vel > 9# And (.posY > coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        
        If CInt(.Velocidad) <= 80 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        End If
       
        .posY = .posY + .Velocidad * time_span + .Aceleracion * time_span * time_span * 0.5
      
      End With
    ' Cuando ya ha llegado al final de la pantalla
    Else
      CocheMalo4Activo = False
      With CocheMalo4
        '.posY = POSY_INI_BAD_CAR * Rnd(-40) 'Lo mando por arriba de la pantalla, aleatoriamente
        .posY = POSY_INI_BAD_CAR * Rnd(-750)
        .posX = CoordenadaXAleatoria ' Con una X aleatoria
      End With
    End If
  End If
  

  
  If CamionActivo Then
    ' Si el camion aun no ha llegado al final de la pantalla
    If Camion.posY <= POSY_FIN_BAD_CAR Then
      
      With Camion
      
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
      
        ' Cuando el juagador adelanta tambien se debe reproducir sonido de adelantamiento
        If vel > 9# And (.posY > coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        
        If CInt(.Velocidad) <= 30 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        End If
        .posY = .posY + .Velocidad * time_span + .Aceleracion * time_span * time_span * 0.5
      End With
    Else
      ' Se desactiva la presencia del camion hasta nueva orden
      CamionActivo = False
      With Camion
        .posY = POSY_INI_BAD_CAR - 62
        .posX = CoordenadaXAleatoriaCamion ' Con una X aleatoria
      End With
            
    End If
  End If
  
  If Camion2Activo Then
    ' Si el camion aun no ha llegado al final de la pantalla
    If Camion2.posY <= POSY_FIN_BAD_CAR Then
      
      With Camion2
      
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
      
        ' Cuando el juagador adelanta tambien se debe reproducir sonido de adelantamiento
        If vel > 9# And (.posY > coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        
        If CInt(.Velocidad) <= 30 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        End If
        .posY = .posY + .Velocidad * time_span + .Aceleracion * time_span * time_span * 0.5
      End With
    Else
      ' Se desactiva la presencia del camion hasta nueva orden
      Camion2Activo = False
      With Camion2
        .posY = POSY_INI_BAD_CAR - 92
        .posX = CoordenadaXAleatoriaCamion ' Con una X aleatoria
      End With
            
    End If
  End If
  
Else

  If CocheMalo1Activo Then
    ' Si aun no ha llegado al inicio de la pantalla
    If CocheMalo1.posY > POSY_INI_BAD_CAR Then
      With CocheMalo1
      
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
         If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
      
        .posY = .posY - .Velocidad * time_span
        If CInt(.Velocidad) <= 50 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        Else
          If (.posY <= coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        End If
      End With
  
    ' Si ya ha llegado al inicio de la pantalla
    Else
      CocheMalo1Activo = False
      CocheMalo1.posY = POSY_FIN_BAD_CAR
      CocheMalo1.posX = CoordenadaXAleatoria
    End If
  End If
  
  If CocheMalo2Activo Then
    ' Si aun no ha llegado al inicio de la pantalla
    If CocheMalo2.posY > POSY_INI_BAD_CAR Then
      With CocheMalo2
      
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
      
        .posY = .posY - .Velocidad * time_span
        If CInt(.Velocidad) <= 52 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        Else
          If (.posY <= coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        End If
      End With
    ' Si ya ha llegado al inicio de la pantalla
    Else
      CocheMalo2Activo = False
      CocheMalo2.posY = POSY_FIN_BAD_CAR + Rnd(100)
      CocheMalo2.posX = CoordenadaXAleatoria
    End If
  End If
  
  If CocheMalo3Activo Then
    ' Si aun no ha llegado al inicio de la pantalla
    If CocheMalo3.posY > POSY_INI_BAD_CAR Then
      With CocheMalo3
        
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
                
        .posY = .posY - .Velocidad * time_span
        If CInt(.Velocidad) <= 54 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        Else
          If (.posY <= coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        End If
      End With
    ' Si ya ha llegado al inicio de la pantalla
    Else
      CocheMalo3Activo = False
      CocheMalo3.posY = POSY_FIN_BAD_CAR + Rnd(100)
      CocheMalo3.posX = CoordenadaXAleatoria 'CoordenadaXAleatoriaLateral
    End If
  End If

  If CocheMalo4Activo Then
    ' Si aun no ha llegado al inicio de la pantalla
    If CocheMalo4.posY > POSY_INI_BAD_CAR Then
      With CocheMalo4
                
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
        
        .posY = .posY - .Velocidad * time_span
        If CInt(.Velocidad) <= 48 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        Else
          If (.posY <= coche.posY) And (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido adelantada
        End If
      End With
      ' Si ya ha llegado al inicio de la pantalla
    Else
      CocheMalo4Activo = False
      CocheMalo4.posY = POSY_FIN_BAD_CAR + Rnd(100)
      CocheMalo4.posX = CoordenadaXAleatoria
    End If
  End If
  
  If CamionActivo Then
    ' Si aun no ha llegado al inicio de la pantalla
    If Camion.posY >= POSY_INI_BAD_CAR - Camion.Alto Then
      With Camion
        
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
        
        .posY = .posY - .Velocidad * time_span
        If CInt(.Velocidad) <= 30 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        Else
          ' Cuando el camion adelanta al jugador en la misma carretera->bocinazo
          If (.posY <= coche.posY) And _
               (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido bocinazo
        End If
      End With
  
    ' Si ya ha llegado al inicio de la pantalla
    Else
      ' Se desactiva la presencia del camion
      CamionActivo = False
      Camion.posX = CoordenadaXAleatoriaCamion
      Camion.posY = POSY_FIN_BAD_CAR + 62 ' Volver a establecer su posicion vertical
    End If
  End If
  
  If Camion2Activo Then
    ' Si aun no ha llegado al inicio de la pantalla
    If Camion2.posY >= POSY_INI_BAD_CAR - Camion2.Alto Then
      With Camion2
        
        ' Controlar que no se salga de la carretera
        If .posX <= limite_borde_izquierdo + 3 Then .posX = .posX + 1
        If .posX >= limite_borde_derecho - 3 - .Ancho Then .posX = .posX - 1
        
        .posY = .posY - .Velocidad * time_span
        If CInt(.Velocidad) <= 30 Then
          .Velocidad = .Velocidad + .Aceleracion * time_span
        Else
          ' Cuando el camion adelanta al jugador en la misma carretera->bocinazo
          If (.posY <= coche.posY) And _
             (Abs(.posY - coche.posY) < coche.Alto) Then ReproducirSonido bocinazo
        End If
      End With
  
    ' Si ya ha llegado al inicio de la pantalla
    Else
      ' Se desactiva la presencia del camion
      Camion2Activo = False
      Camion2.posX = CoordenadaXAleatoriaCamion
      Camion2.posY = POSY_FIN_BAD_CAR + 92 ' Volver a establecer su posicion vertical
    End If
  End If

End If
  



End Sub



Private Sub ControlarMovimientos()
  
If Choque_en_Curso Then Exit Sub
  
  With coche
  
    If TeclaIzquierda Then
      .posX = .posX - .Velocidad
    ElseIf TeclaDerecha Then
      .posX = .posX + .Velocidad
    End If
  
  End With
    
End Sub

Private Sub EfectosColisionCoche()
       
  If FlagTurbo Then
    If vel >= 5# Then vel = vel - 5#
    ReproducirSonido crash2
  Else
    If vel >= 1# Then vel = vel - 0.5
  End If
  
  
If VelocidadInicialColision >= 10 Then
  EfectoExplosion
Else

  ReproducirSonido freno2
  
  ' Producir efecto de vueltas sobre si mismo
  If Posiciones_Choque(0) = False Then
    Set coche.Superficie = CocheColision(0).Superficie
    Posiciones_Choque(0) = True
  ElseIf Posiciones_Choque(1) = False Then
    Set coche.Superficie = CocheColision(1).Superficie
    Posiciones_Choque(1) = True
  ElseIf Posiciones_Choque(2) = False Then
    Set coche.Superficie = CocheColision(2).Superficie
    Posiciones_Choque(2) = True
  ElseIf Posiciones_Choque(3) = False Then
    Set coche.Superficie = CocheColision(3).Superficie
    Posiciones_Choque(3) = True
  ElseIf Posiciones_Choque(4) = False Then
    Set coche.Superficie = CocheColision(4).Superficie
    Posiciones_Choque(4) = True
  ElseIf Posiciones_Choque(5) = False Then
    Set coche.Superficie = CocheColision(5).Superficie
    Posiciones_Choque(5) = True
  ElseIf Posiciones_Choque(6) = False Then
    Set coche.Superficie = CocheColision(6).Superficie
    Posiciones_Choque(6) = True
  ElseIf Posiciones_Choque(7) = False Then
    ' Restaurar coche normal
    Set coche.Superficie = CocheColision(7).Superficie
    Posiciones_Choque(7) = True
  Else
    ' Volver a desactivar el flag y resetear el vector de posiciones
    Choque_en_Curso = False
    
    Dim i As Integer
    For i = 0 To 12: Posiciones_Choque(i) = False: Next
    ' Se restan puntos
    If puntos >= 1 Then puntos = puntos - 1
  End If


End If

End Sub

Private Sub ControlarTiempoChoque()

    If Choque_en_Curso Then EfectosColisionCoche
       
    If ChoqueCocheMalo1EnCurso Then
      ReproducirSonido crash1
      EfectosColisionCochesEnemigos 1 ' El coche malo 1
    End If
    
    If ChoqueCocheMalo2EnCurso Then
      ReproducirSonido crash1
      EfectosColisionCochesEnemigos 2 ' El coche malo 2
    End If
    
    If ChoqueCocheMalo3EnCurso Then
      ReproducirSonido crash1
      EfectosColisionCochesEnemigos 3 ' El coche malo 3
    End If
    
    If ChoqueCocheMalo4EnCurso Then
      ReproducirSonido crash1
      EfectosColisionCochesEnemigos 4 ' El coche malo 4
    End If
    


End Sub

Private Sub ComprobarProximidad()

' Solo se hae si la velocidad del jugador es alta y estamos en el nivel 2

If vel > 9# Then

  ' Esto solo lo hago con el cochemalo3 (coche rosa)
  ' Si estan casi a la misma altura, intenta colisionar con el coche del jugador
  
  With CocheMalo3
    If Abs(coche.posY - .posY) <= 10 Then
      ' Si la distancia horizontal entre ellos es de 3 pixels
      If Abs(coche.posX - .posX) <= coche.Ancho + 3 Then
        ' Intimidar-> Aproximar 1 pixel el coche malo al jugador
        If coche.posX < .posX Then .posX = .posX - 1 Else .posX = .posX + 1
      End If
    End If
  End With
  
  ' Esto solo lo hago con el cochemaloaux2 (coche amarillo)
  ' Este coche, intenta ponerse en la misma vertical que el coche del jugador,
  ' siempre y cuando esten entre 100 y 30 pixeles verticales de distancia

    Dim distancia As Integer
    
    With CocheMalo4
      distancia = Abs(coche.posY - .posY + .Alto)
      If distancia <= 100 And distancia >= 30 Then
        If .posX < coche.posX - 15 Then
          .posX = .posX + 1
        ElseIf .posX > coche.posX + 15 Then
          .posX = .posX - 1
        End If
      End If
    End With
    

End If

End Sub

Public Sub ReproducirSonido(Sonido As Integer, Optional SonidoContinuo As Boolean)

' Reproducit el Sound Buffer

If SonidoContinuo Then
  sb(Sonido).Play DSBPLAY_LOOPING
Else
  sb(Sonido).Play DSBPLAY_DEFAULT
End If

End Sub

Private Sub ControlarObstaculos()

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Si el coche atraviesa una mancha de aceite-> zigzaguearlo
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

If ChoqueConObstaculos(ObjetoManchaAceite(0), ObjetoManchaAceite(1)) Then
  ReproducirSonido agua
  coche.posX = coche.posX + 15
End If

' * * * * * * * * * * * * * * * * * * * * * *
' Si el coche atraviesa un charco-> frenarlo
' * * * * * * * * * * * * * * * * * * * * * *

If ChoqueConObstaculos(ObjetoCharco(0), ObjetoCharco(1)) Then
  ReproducirSonido agua
  If vel > 1# Then vel = 0#
End If

' * * * * * * * * * * * * * * * * * * * * * * * * *
' Si el coche topa contra una roca->dejarlo clavado
' * * * * * * * * * * * * * * * * * * * * * * * * *

If ChoqueConObstaculos(ObjetoRoca(0), ObjetoRoca(1)) Then '
  vel = 10 ' Para que el porrazo sea con explosion
  PrepararColision False
  coche.posX = POSX_INI_CAR
End If

' La siguiente funcion controla que los vehiculos enemigos eviten las rocas
IARoca

End Sub



Private Sub IARoca()

' Evitar que cualquier vehiculo esté en la vertical de las rocas mientras éstas esten
' en la pantalla

  
  ' Si roca presente...
  
  If ObjetoRoca(0).PosYMapa >= RowIndex - MAX_ROWS_SCREEN And _
     ObjetoRoca(0).PosYMapa <= RowIndex + MAX_ROWS_SCREEN + 1 Then

    If CocheMalo1Activo Then
      With CocheMalo1
        If Abs(.posX - ObjetoRoca(0).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + .Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - .Ancho - 5
          End If
        End If
      End With
    End If
    
    If CocheMalo2Activo Then
      With CocheMalo2
        If Abs(.posX - ObjetoRoca(0).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + .Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - .Ancho - 5
          End If
        End If
      End With
    End If
    
    If CocheMalo3Activo Then
      With CocheMalo3
        If Abs(.posX - ObjetoRoca(0).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + .Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - .Ancho - 5
          End If
        End If
      End With
    End If
    
    If CocheMalo4Activo Then
      With CocheMalo4
        If Abs(.posX - ObjetoRoca(0).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + .Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - .Ancho - 5
          End If
        End If
      End With
    End If
    
    If CamionActivo Then
      With Camion
        If Abs(.posX - ObjetoRoca(0).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + CocheMalo1.Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - CocheMalo1.Ancho - 5
          End If
        End If
      End With
    End If
    
    If Camion2Activo Then
      With Camion2
        If Abs(.posX - ObjetoRoca(0).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + CocheMalo1.Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - CocheMalo1.Ancho - 5
          End If
        End If
      End With
    End If
End If

  
If ObjetoRoca(1).PosYMapa >= RowIndex - MAX_ROWS_SCREEN And _
     ObjetoRoca(1).PosYMapa <= RowIndex + MAX_ROWS_SCREEN + 1 Then

    If CocheMalo1Activo Then
      With CocheMalo1
        If Abs(.posX - ObjetoRoca(1).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + .Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - .Ancho - 5
          End If
        End If
      End With
    End If
    
    If CocheMalo2Activo Then
      With CocheMalo2
        If Abs(.posX - ObjetoRoca(1).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + .Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - .Ancho - 5
          End If
        End If
      End With
    End If
    
    If CocheMalo3Activo Then
      With CocheMalo3
        If Abs(.posX - ObjetoRoca(1).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + .Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - .Ancho - 5
          End If
        End If
      End With
    End If
    
    If CocheMalo4Activo Then
      With CocheMalo4
        If Abs(.posX - ObjetoRoca(1).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + .Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - .Ancho - 5
          End If
        End If
      End With
    End If
    
    If CamionActivo Then
      With Camion
        If Abs(.posX - ObjetoRoca(1).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + CocheMalo1.Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - CocheMalo1.Ancho - 5
          End If
        End If
      End With
    End If
    
    If Camion2Activo Then
      With Camion2
        If Abs(.posX - ObjetoRoca(1).PosXPantalla) <= .Ancho Then
          If .posX > 120 Then ' Si la roca le queda por la izquierda, apartar el coche a la derecha
            .posX = .posX + CocheMalo1.Ancho + 5
          Else ' Si la roca le queda por la derecha, apartar el coche a la izquierda
            .posX = .posX - CocheMalo1.Ancho - 5
          End If
        End If
      End With
    End If
  End If

End Sub

Private Sub ControlarRecargaCombustible()

' 1ª condicion - Tiene que estar cubierto, por lo menos, el 50% del trayecto
If RowIndex < TRAMO / 2 Then

  ' 2ª condicion - Tiene que quedar menos de 1/4 de depósito
  If Combustible <= COMBUSTIBLE_INICIAL / 4 Then

    RecargaCombustibleActiva = True
    
    With RecargaCombustible
      
      ' Si aun no es visible por la parte superior, ir bajandola
      If .posY <= POSY_INI_BAD_CAR Then
        
        ' Descenderla
        .posY = .posY + 1
        
      ' Cuando ya es visible hacemos descender la recarga
      ElseIf .posY > POSY_INI_BAD_CAR And .posY < POSY_FIN_BAD_CAR Then
         ' La recarga debe descender dentro de los límites de la carretera, sino el
         ' jugador no la puede cojer.
         If .posX < limite_borde_izquierdo - 1 Then .posX = .posX + 1
         If .posX > limite_borde_derecho + 1 Then .posX = .posX - 1
         
         .Velocidad = .Velocidad + .Aceleracion * time_span
         .posY = .posY + .Velocidad * time_span + .Aceleracion * time_span * time_span * 0.5
         
         ' Comprobar si el jugador recoje la recarga
         If HayInterseccion(coche, RecargaCombustible) And CocheActivo Then
           'Combustible = Combustible + (COMBUSTIBLE_INICIAL / 4) ' Aumentar el combustible
           Combustible = COMBUSTIBLE_INICIAL
           sb(outoffuel).Stop
           SinCombustible = False
           ReproducirSonido carga_combustible
           .posY = POSY_FIN_BAD_CAR + 1
         End If
      
      Else 'If .posY >= POSY_FIN_BAD_CAR Then
        
        RecargaCombustibleActiva = False  ' Cuando llega al final desaparece
        .posY = POSY_FIN_BAD_CAR
        '.posX = CoordenadaXAleatoria
        .Velocidad = 0
      End If
    End With
    
  End If
   
End If

End Sub
Private Function CoordenadaXAleatoriaCamion() As Integer
  'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)
  
  CoordenadaXAleatoriaCamion = Int((limite_borde_derecho - Camion.Ancho - 2 - _
                                    limite_borde_izquierdo) + 1) * _
                                    Rnd + limite_borde_izquierdo
                                    
                                    
End Function

Private Sub EfectosColisionCochesEnemigos(ByVal CocheEnemigo As Integer)
   
Dim i As Integer
  
    
Select Case CocheEnemigo
  
  Case 1

    If CocheMalo1ContraCamion Then
      EfectoExplosionCochesEnemigos 1
    Else
      ' Producir efecto de vueltas sobre si mismo
      If Posiciones_Choque_CocheMalo1(0) = False Then
        Set CocheMalo1.Superficie = CocheMalo1Colision(0).Superficie
        Posiciones_Choque_CocheMalo1(0) = True
      ElseIf Posiciones_Choque_CocheMalo1(1) = False Then
        Set CocheMalo1.Superficie = CocheMalo1Colision(1).Superficie
        Posiciones_Choque_CocheMalo1(1) = True
      ElseIf Posiciones_Choque_CocheMalo1(2) = False Then
        Set CocheMalo1.Superficie = CocheMalo1Colision(2).Superficie
        Posiciones_Choque_CocheMalo1(2) = True
      ElseIf Posiciones_Choque_CocheMalo1(3) = False Then
        Set CocheMalo1.Superficie = CocheMalo1Colision(3).Superficie
        Posiciones_Choque_CocheMalo1(3) = True
      ElseIf Posiciones_Choque_CocheMalo1(4) = False Then
        Set CocheMalo1.Superficie = CocheMalo1Colision(4).Superficie
        Posiciones_Choque_CocheMalo1(4) = True
      ElseIf Posiciones_Choque_CocheMalo1(5) = False Then
        Set CocheMalo1.Superficie = CocheMalo1Colision(5).Superficie
        Posiciones_Choque_CocheMalo1(5) = True
      ElseIf Posiciones_Choque_CocheMalo1(6) = False Then
        Set CocheMalo1.Superficie = CocheMalo1Colision(6).Superficie
        Posiciones_Choque_CocheMalo1(6) = True
      ElseIf Posiciones_Choque_CocheMalo1(7) = False Then
        ' Restaurar coche normal
        Set CocheMalo1.Superficie = CocheMalo1Colision(7).Superficie
        Posiciones_Choque_CocheMalo1(7) = True
      Else
        ' Volver a desactivar el flag y resetear el vector de posiciones
        ChoqueCocheMalo1EnCurso = False
        For i = 0 To 12: Posiciones_Choque_CocheMalo1(i) = False: Next
      End If
    End If
  
  Case 2
  
    If CocheMalo2ContraCamion Then
      EfectoExplosionCochesEnemigos 2
    Else
      ' Producir efecto de vueltas sobre si mismo
      If Posiciones_Choque_CocheMalo2(0) = False Then
        Set CocheMalo2.Superficie = CocheMalo2Colision(0).Superficie
        Posiciones_Choque_CocheMalo2(0) = True
      ElseIf Posiciones_Choque_CocheMalo2(1) = False Then
        Set CocheMalo2.Superficie = CocheMalo2Colision(1).Superficie
        Posiciones_Choque_CocheMalo2(1) = True
      ElseIf Posiciones_Choque_CocheMalo2(2) = False Then
        Set CocheMalo2.Superficie = CocheMalo2Colision(2).Superficie
        Posiciones_Choque_CocheMalo2(2) = True
      ElseIf Posiciones_Choque_CocheMalo2(3) = False Then
        Set CocheMalo2.Superficie = CocheMalo2Colision(3).Superficie
        Posiciones_Choque_CocheMalo2(3) = True
      ElseIf Posiciones_Choque_CocheMalo2(4) = False Then
        Set CocheMalo2.Superficie = CocheMalo2Colision(4).Superficie
        Posiciones_Choque_CocheMalo2(4) = True
      ElseIf Posiciones_Choque_CocheMalo2(5) = False Then
        Set CocheMalo2.Superficie = CocheMalo2Colision(5).Superficie
        Posiciones_Choque_CocheMalo2(5) = True
      ElseIf Posiciones_Choque_CocheMalo2(6) = False Then
        Set CocheMalo2.Superficie = CocheMalo2Colision(6).Superficie
        Posiciones_Choque_CocheMalo2(6) = True
      ElseIf Posiciones_Choque_CocheMalo2(7) = False Then
        ' Restaurar coche normal
        Set CocheMalo2.Superficie = CocheMalo2Colision(7).Superficie
        Posiciones_Choque_CocheMalo2(7) = True
      Else
        ' Volver a desactivar el flag y resetear el vector de posiciones
        ChoqueCocheMalo2EnCurso = False
        For i = 0 To 12: Posiciones_Choque_CocheMalo2(i) = False: Next
      End If
    End If
    
  Case 3
  
    If CocheMalo3ContraCamion Then
      EfectoExplosionCochesEnemigos 3
    Else
      ' Producir efecto de vueltas sobre si mismo
      If Posiciones_Choque_CocheMalo3(0) = False Then
        Set CocheMalo3.Superficie = CocheMalo3Colision(0).Superficie
        Posiciones_Choque_CocheMalo3(0) = True
      ElseIf Posiciones_Choque_CocheMalo3(1) = False Then
        Set CocheMalo3.Superficie = CocheMalo3Colision(1).Superficie
        Posiciones_Choque_CocheMalo3(1) = True
      ElseIf Posiciones_Choque_CocheMalo3(2) = False Then
        Set CocheMalo3.Superficie = CocheMalo3Colision(2).Superficie
        Posiciones_Choque_CocheMalo3(2) = True
      ElseIf Posiciones_Choque_CocheMalo3(3) = False Then
        Set CocheMalo3.Superficie = CocheMalo3Colision(3).Superficie
        Posiciones_Choque_CocheMalo3(3) = True
      ElseIf Posiciones_Choque_CocheMalo3(4) = False Then
        Set CocheMalo3.Superficie = CocheMalo3Colision(4).Superficie
        Posiciones_Choque_CocheMalo3(4) = True
      ElseIf Posiciones_Choque_CocheMalo3(5) = False Then
        Set CocheMalo3.Superficie = CocheMalo3Colision(5).Superficie
        Posiciones_Choque_CocheMalo3(5) = True
      ElseIf Posiciones_Choque_CocheMalo3(6) = False Then
        Set CocheMalo3.Superficie = CocheMalo3Colision(6).Superficie
        Posiciones_Choque_CocheMalo3(6) = True
      ElseIf Posiciones_Choque_CocheMalo3(7) = False Then
        ' Restaurar coche normal
        Set CocheMalo3.Superficie = CocheMalo3Colision(7).Superficie
        Posiciones_Choque_CocheMalo3(7) = True
      Else
        ' Volver a desactivar el flag y resetear el vector de posiciones
        ChoqueCocheMalo3EnCurso = False
        For i = 0 To 12: Posiciones_Choque_CocheMalo3(i) = False: Next
      End If
    
    End If
    
  Case 4
  
    If CocheMalo4ContraCamion Then
      EfectoExplosionCochesEnemigos 4
    Else
      ' Producir efecto de vueltas sobre si mismo
      If Posiciones_Choque_CocheMalo4(0) = False Then
        Set CocheMalo4.Superficie = CocheMalo4Colision(0).Superficie
        Posiciones_Choque_CocheMalo4(0) = True
      ElseIf Posiciones_Choque_CocheMalo4(1) = False Then
        Set CocheMalo4.Superficie = CocheMalo4Colision(1).Superficie
        Posiciones_Choque_CocheMalo4(1) = True
      ElseIf Posiciones_Choque_CocheMalo4(2) = False Then
        Set CocheMalo4.Superficie = CocheMalo4Colision(2).Superficie
        Posiciones_Choque_CocheMalo4(2) = True
      ElseIf Posiciones_Choque_CocheMalo4(3) = False Then
        Set CocheMalo4.Superficie = CocheMalo4Colision(3).Superficie
        Posiciones_Choque_CocheMalo4(3) = True
      ElseIf Posiciones_Choque_CocheMalo4(4) = False Then
        Set CocheMalo4.Superficie = CocheMalo4Colision(4).Superficie
        Posiciones_Choque_CocheMalo4(4) = True
      ElseIf Posiciones_Choque_CocheMalo4(5) = False Then
        Set CocheMalo4.Superficie = CocheMalo4Colision(5).Superficie
        Posiciones_Choque_CocheMalo4(5) = True
      ElseIf Posiciones_Choque_CocheMalo4(6) = False Then
        Set CocheMalo4.Superficie = CocheMalo4Colision(6).Superficie
        Posiciones_Choque_CocheMalo4(6) = True
      ElseIf Posiciones_Choque_CocheMalo4(7) = False Then
        ' Restaurar coche normal
        Set CocheMalo4.Superficie = CocheMalo4Colision(7).Superficie
        Posiciones_Choque_CocheMalo4(7) = True
      Else
        ' Volver a desactivar el flag y resetear el vector de posiciones
        ChoqueCocheMalo4EnCurso = False
        For i = 0 To 12: Posiciones_Choque_CocheMalo4(i) = False: Next
      End If
    End If
End Select
    
End Sub


Private Sub ControlarTurbo()
  
  ' Si esta pulsado el turbo
  If FlagTurbo Then
    ' Si la velocidad es 10 ' La maxima sin turbo
    If max_velocidad = 10# Then
      ' La nueva velocidad maxima es 15.0
      max_velocidad = 15#
      ' Reproducir sonido de turbo
      ReproducirSonido turbo
    
    ElseIf vel > 10# Then
      coche.Velocidad = 2
      CocheMalo1.Velocidad = 300
      CocheMalo2.Velocidad = 300
      CocheMalo3.Velocidad = 300
      CocheMalo4.Velocidad = 300
      Camion.Velocidad = 100
      Camion2.Velocidad = 60
    Else
      FlagTurbo = False
    End If
  ' Si no está pulsado el turbo
  Else
    
    If max_velocidad = 15# Then
      max_velocidad = 10#
      coche.Velocidad = 1
      CocheMalo1.Velocidad = 150
      CocheMalo2.Velocidad = 150
      CocheMalo3.Velocidad = 150
      CocheMalo4.Velocidad = 150
      Camion.Velocidad = 20
      Camion2.Velocidad = 10
    End If
    
    ' 10 debe ser la velocidad máxima
    
    ' Deter sonido de turbo (si lo hubiera)
    sb(turbo).Stop
    ' Reducir la velocidad a 10 si es superior a 10
    If vel > 10# Then vel = 10#
  End If

End Sub



Private Sub ControlarColisiones()
  
Dim flagColision As Boolean
Dim flagColisionCamion As Boolean

' Sólo evaluar si NO se está en proceso de choque
If CocheActivo Then
  ' Comprobar si hay choque entre el coche del jugador y los coches enemigos
  If HayInterseccion(coche, CocheMalo1) Then
    ChoqueCocheMalo1EnCurso = True
    flagColision = True
  ElseIf HayInterseccion(coche, CocheMalo2) Then
    ChoqueCocheMalo2EnCurso = True
    flagColision = True
  ElseIf HayInterseccion(coche, CocheMalo3) Then
    ChoqueCocheMalo3EnCurso = True
    flagColision = True
  ElseIf HayInterseccion(coche, CocheMalo4) Then
    ChoqueCocheMalo4EnCurso = True
    flagColision = True
  ElseIf HayInterseccion(coche, Camion) Or HayInterseccion(coche, Camion2) Then
    flagColision = True
    flagColisionCamion = True
  End If

  If flagColision Then
    PrepararColision flagColisionCamion
  End If
End If

' Comprobar si hay choque entre los coches enemigos
  
If CocheMalo1Activo Then
  With CocheMalo1
    If .posY > 0 And .posY < 240 Then ' Si el vehiculo está en pantalla
                
      If CocheMalo2Activo And HayInterseccion(CocheMalo1, CocheMalo2) Then
        ChoqueCocheMalo1EnCurso = True
        .Velocidad = 20 ' Reset
        .Aceleracion = 9.8
        ChoqueCocheMalo2EnCurso = True
        CocheMalo2.Velocidad = 65 ' Reset
        CocheMalo2.Aceleracion = 9.5
        Exit Sub
      End If
        
      If CocheMalo3Activo And HayInterseccion(CocheMalo1, CocheMalo3) Then
        ChoqueCocheMalo1EnCurso = True
        .Velocidad = 20 ' Reset
        .Aceleracion = 9.8
        ChoqueCocheMalo3EnCurso = True
        CocheMalo3.Velocidad = 55 ' Reset
        CocheMalo3.Aceleracion = 8.5
        Exit Sub
      End If
        
      If CocheMalo4Activo And HayInterseccion(CocheMalo1, CocheMalo4) Then
        ChoqueCocheMalo1EnCurso = True
        .Velocidad = 20 ' Reset
        .Aceleracion = 9.8
        ChoqueCocheMalo4EnCurso = True
        CocheMalo4.Velocidad = 77 ' Reset
        CocheMalo4.Aceleracion = 5.5
        Exit Sub
      End If
        
      If CamionActivo And HayInterseccion(CocheMalo1, Camion) And _
         Not CocheMalo1ContraCamion Then
         
        CocheMalo1ContraCamion = True
        CocheMalo1Activo = False
        ChoqueCocheMalo1EnCurso = True
    
        .Velocidad = 0
        .Aceleracion = 9.8
        'ReproducirSonido crash2
        Exit Sub
      End If
        
      If Camion2Activo And HayInterseccion(CocheMalo1, Camion2) And _
         Not CocheMalo1ContraCamion Then
        
        CocheMalo1ContraCamion = True
        CocheMalo1Activo = False
        ChoqueCocheMalo1EnCurso = True
        
        .Velocidad = 0
        .Aceleracion = 9.8
        'ReproducirSonido crash2
        Exit Sub
      End If
                    
    End If
  End With
End If
  
    
If CocheMalo2Activo Then
  With CocheMalo2
    If .posY > 0 And .posY < 240 Then ' Si el vehiculo está en pantalla
        
      If CocheMalo3Activo And HayInterseccion(CocheMalo2, CocheMalo3) Then
        ChoqueCocheMalo2EnCurso = True
        .Velocidad = 65 ' Reset
        .Aceleracion = 9.5
        ChoqueCocheMalo3EnCurso = True
        CocheMalo3.Velocidad = 55 ' Reset
        CocheMalo3.Aceleracion = 8.5
        Exit Sub
      End If
       
      If CocheMalo4Activo And HayInterseccion(CocheMalo2, CocheMalo4) Then
        ChoqueCocheMalo2EnCurso = True
        .Velocidad = 65 ' Reset
        .Aceleracion = 9.5
        ChoqueCocheMalo4EnCurso = True
        CocheMalo4.Velocidad = 77 ' Reset
        CocheMalo4.Aceleracion = 5.5
        Exit Sub
      End If
        
      If CamionActivo And HayInterseccion(CocheMalo2, Camion) And _
         Not CocheMalo2ContraCamion Then
         
        CocheMalo2ContraCamion = True
        CocheMalo2Activo = False
        ChoqueCocheMalo2EnCurso = True
                    
        .Velocidad = 0
        .Aceleracion = 9.5
          
        Exit Sub
      End If
        
      If Camion2Activo And HayInterseccion(CocheMalo2, Camion2) And _
         Not CocheMalo2ContraCamion Then
           
        CocheMalo2ContraCamion = True
        CocheMalo2Activo = False
        ChoqueCocheMalo2EnCurso = True
          
        .Velocidad = 0
        .Aceleracion = 9.5
        Exit Sub
      End If
    End If
  End With
End If
 
If CocheMalo3Activo Then
  With CocheMalo3
    If .posY > 0 And .posY < 240 Then ' Si el vehiculo está en pantalla
        
      If CocheMalo4Activo And HayInterseccion(CocheMalo3, CocheMalo4) Then
        ChoqueCocheMalo3EnCurso = True
        .Velocidad = 55 ' Reset
        .Aceleracion = 8.5
        ChoqueCocheMalo4EnCurso = True
        CocheMalo4.Velocidad = 77 ' Reset
        CocheMalo4.Aceleracion = 5.5
        Exit Sub
      End If
        
      If CamionActivo And HayInterseccion(CocheMalo3, Camion) And _
         Not CocheMalo3ContraCamion Then
         
        ChoqueCocheMalo3EnCurso = True
        CocheMalo3ContraCamion = True
        CocheMalo3Activo = False
        .Velocidad = 0
        .Aceleracion = 8.5
         Exit Sub
      End If
        
      If Camion2Activo And HayInterseccion(CocheMalo3, Camion2) And _
         Not CocheMalo3ContraCamion Then
         
        CocheMalo3ContraCamion = True
        CocheMalo3Activo = False
        ChoqueCocheMalo3EnCurso = True
        .Velocidad = 0
        .Aceleracion = 8.5
 
        Exit Sub
      End If
    End If
  End With
End If
  
  
If CocheMalo4Activo Then
  With CocheMalo4
    If .posY > 0 And .posY < 240 Then ' Si el vehiculo está en pantalla
        
      If CamionActivo And HayInterseccion(CocheMalo4, Camion) And _
         Not CocheMalo4ContraCamion Then
         
        CocheMalo4ContraCamion = True
        CocheMalo4Activo = False
        ChoqueCocheMalo4EnCurso = True
        .Velocidad = 0
        .Aceleracion = 5.5
        Exit Sub
      End If
                    
      If Camion2Activo And HayInterseccion(CocheMalo4, Camion2) And _
         Not CocheMalo4ContraCamion Then
         
        CocheMalo4ContraCamion = True
        CocheMalo4Activo = False
        ChoqueCocheMalo4EnCurso = True
        .Velocidad = 0
        .Aceleracion = 5.5
        Exit Sub
      End If
        
    End If
  End With
End If
  
  
  If CamionActivo Then
    With Camion
      If .posY > POSY_INI_BAD_CAR And .posY < POSY_FIN_BAD_CAR Then
    
        If HayInterseccion(Camion, Camion2) Then
          
          ' Si el camion2 le da por detras al camion->mover la y del camion tanto como
          ' sea necesario para que no se solapen
          If .posY < Camion2.posY Then
            .posY = Camion2.posY - Camion.posY - Camion.Alto
            Camion2.Velocidad = Camion2.Velocidad - 5
          ' Si el camion le da por detras al camion2->frenar en seco la velocidad del camion
          Else
            .Velocidad = .Velocidad - 5
          End If
        End If
      End If
    End With
  End If
  
End Sub



Private Sub ControlarAceleracion()

' Si barra espaciadora pulsada
  
  If BarraEspaciadora Then
    
    If vel < max_velocidad Then vel = vel + 1# * time_span
    If vel < 0.5 Then ReproducirSonido derrape

    If vel >= 3# Then
      'ReproducirSonido marcha, True
      If Not (sb(marcha).GetStatus = DSBSTATUS_PLAYING) Then ReproducirSonido marcha, True
    End If
            
    ComprobarProximidad ' Hacer que algunos vehiculos enemigos sean hostiles
          
  Else
        
    Dim decel As Currency ' desaceleracion
    
    decel = 1# * time_span
                    
    If Abs(decel) > Abs(vel) Then
      vel = 0#
    Else
      ' Control para que desaceleraciones cortas no produzcan un efecto feo
      ' de aceleracion-desaceleracion en los vehiculos enemigos
      If vel <= 7 Then
         vel = vel - decel
      Else
         vel = 6.9
      End If
    End If
    
    'Quitar el sonido de la marcha
    If vel < 3# Then sb(marcha).Stop
    
  End If

End Sub


Private Sub MoverPaisaje()
  
  MapDspRow = MapDspRow - vel
        
  If MapDspRow <= 0 Then
    MapDspRow = 15
    RowIndex = RowIndex - 1
    Recorrido = Recorrido + 1
    PosicionRecorrido = (POS_RECORRIDO_FINAL * Recorrido \ TRAMO * VUELTAS_COMPLETADAS)
  End If
  
End Sub

Private Sub ControlarBordesCarretera()

Dim TipoBordeIzquierdo As String * 1, TipoBordeDerecho As String * 1
Dim posInMatrix As Integer, diferencial As Integer
Dim flagColision As Boolean

If Not CocheActivo Then Exit Sub

' Obtener el tipo de borde izquierdo y su posicion en la matriz
TipoBordeIzquierdo = GetTipoBorde(0, posInMatrix)

Select Case TipoBordeIzquierdo

  Case "2" ' borde izquierdo recto
    diferencial = 11
    
  Case "C" ' borde izquierdo girando a derecha (1/2)
    
    Select Case RowIndex
      
      Case 1717, 1653, 1644, 1589, 1580, 1141, 1077, 1068, 1005, 997, 669, 661, 652, 437, 428, 301, 237, 229, 172
        diferencial = 13
      Case 1716, 1652, 1643, 1588, 1579, 1140, 1076, 1067, 1004, 996, 668, 660, 651, 436, 427, 300, 236, 228, 171
        diferencial = 15
      Case 1715, 1651, 1642, 1587, 1578, 1139, 1075, 1066, 1003, 995, 667, 659, 650, 435, 426, 299, 235, 227, 170
        diferencial = 17
      Case 1714, 1650, 1641, 1586, 1577, 1138, 1074, 1065, 1002, 994, 666, 658, 649, 434, 425, 298, 234, 226, 169
        diferencial = 19
      Case 1713, 1649, 1640, 1585, 1576, 1137, 1073, 1064, 1001, 993, 665, 657, 648, 433, 424, 297, 233, 225, 168
        diferencial = 21
      Case 1712, 1648, 1639, 1584, 1575, 1136, 1072, 1063, 1000, 992, 664, 656, 647, 432, 423, 296, 232, 224, 167
        diferencial = 23
      Case 1711, 1647, 1638, 1583, 1574, 1135, 1071, 1062, 999, 991, 663, 655, 646, 431, 422, 295, 231, 223, 166
        diferencial = 25
      Case 1710, 1646, 1637, 1582, 1573, 1134, 1070, 1061, 998, 990, 662, 654, 645, 430, 421, 294, 230, 222, 165
        diferencial = 27
      Case 1709, 1645, 1636, 1581, 1572, 1133, 1069, 1060, 999, 989, 662, 653, 644, 429, 420, 293, 230, 221, 164
        diferencial = 27 ' sí, también 27
        
    End Select
    
  Case "7" ' borde izquierdo girando a izquierda (2/2)
    
    Select Case RowIndex
      
      Case 1669, 1637, 1660, 1628, 1517, 1093, 1084, 1061, 1052, 933, 782, 773, 469, 453, 444, 253, 245, 221, 213
        diferencial = 10
      Case 1668, 1636, 1659, 1627, 1516, 1092, 1083, 1060, 1051, 932, 781, 772, 468, 452, 443, 252, 244, 220, 212
        diferencial = 8
      Case 1667, 1635, 1658, 1626, 1515, 1091, 1082, 1059, 1050, 931, 780, 771, 467, 451, 442, 251, 243, 219, 211
        diferencial = 6
      Case 1666, 1634, 1657, 1625, 1514, 1090, 1081, 1058, 1049, 930, 779, 770, 466, 450, 441, 250, 242, 218, 210
        diferencial = 4
      Case 1665, 1633, 1656, 1624, 1513, 1089, 1080, 1057, 1048, 929, 778, 769, 465, 449, 440, 249, 241, 217, 209
        diferencial = 2
      Case 1664, 1632, 1655, 1623, 1512, 1088, 1079, 1056, 1047, 928, 777, 768, 464, 448, 439, 248, 240, 216, 208
        diferencial = 0
      Case 1663, 1631, 1654, 1622, 1511, 1087, 1078, 1055, 1046, 927, 776, 767, 463, 447, 438, 247, 239, 215, 207
        diferencial = -2
      Case 1662, 1630, 1653, 1621, 1510, 1086, 1077, 1054, 1045, 926, 775, 766, 462, 446, 437, 246, 238, 214, 206
        diferencial = -4
      Case 1661, 1629, 1652, 1620, 1509, 1085, 1076, 1053, 1044, 925, 774, 765, 461, 445, 436, 246, 237, 214, 205
        diferencial = -6

    End Select
    
End Select



limite_borde_izquierdo = POSX_INI_CAR + (16 * (posInMatrix - 1)) + diferencial



' Obtener el tipo de borde derecho y su posicion en la matriz
TipoBordeDerecho = GetTipoBorde(1, posInMatrix)

Select Case TipoBordeDerecho

  Case "5" ' borde derecho recto
    diferencial = 4
    
  Case "G" ' borde derecho girando a derecha (1/2)
    
    Select Case RowIndex
      
      Case 1717, 1653, 1644, 1589, 1580, 1141, 1077, 1068, 1005, 997, 669, 661, 652, 437, 428, 301, 237, 229, 172
        diferencial = 5
      Case 1716, 1652, 1643, 1588, 1579, 1140, 1076, 1067, 1004, 996, 668, 660, 651, 436, 427, 300, 236, 228, 171
        diferencial = 7
      Case 1715, 1651, 1642, 1587, 1578, 1139, 1075, 1066, 1003, 995, 667, 659, 650, 435, 426, 299, 235, 227, 170
        diferencial = 9
      Case 1714, 1650, 1641, 1586, 1577, 1138, 1074, 1065, 1002, 994, 666, 658, 649, 434, 425, 298, 234, 226, 169
        diferencial = 11
      Case 1713, 1649, 1640, 1585, 1576, 1137, 1073, 1064, 1001, 993, 665, 657, 648, 433, 424, 297, 233, 225, 168
        diferencial = 13
      Case 1712, 1648, 1639, 1584, 1575, 1136, 1072, 1063, 1000, 992, 664, 656, 647, 432, 423, 296, 232, 224, 167
        diferencial = 15
      Case 1711, 1647, 1638, 1583, 1574, 1135, 1071, 1062, 999, 991, 663, 655, 646, 431, 422, 295, 231, 223, 166
        diferencial = 17
      Case 1710, 1646, 1637, 1582, 1573, 1134, 1070, 1061, 998, 990, 662, 654, 645, 430, 421, 294, 230, 222, 165
        diferencial = 19
      Case 1709, 1645, 1636, 1581, 1572, 1133, 1069, 1060, 999, 989, 662, 653, 644, 429, 420, 293, 230, 221, 164
        diferencial = 21 ' sí, también 21
        
    End Select
    
  Case "B" ' borde derecho girando a izquierda (2/2)
    
    Select Case RowIndex
      
      Case 1669, 1637, 1660, 1628, 1517, 1093, 1084, 1061, 1052, 933, 781, 773, 469, 453, 444, 253, 245, 221, 213
        diferencial = 3
      Case 1668, 1636, 1659, 1627, 1516, 1092, 1083, 1060, 1051, 932, 780, 772, 468, 452, 443, 252, 244, 220, 212
        diferencial = 1
      Case 1667, 1635, 1658, 1626, 1515, 1091, 1082, 1059, 1050, 931, 779, 771, 467, 451, 442, 251, 243, 219, 211
        diferencial = -1
      Case 1666, 1634, 1657, 1625, 1514, 1090, 1081, 1058, 1049, 930, 778, 770, 466, 450, 441, 250, 242, 218, 210
        diferencial = -3
      Case 1665, 1633, 1656, 1624, 1513, 1089, 1080, 1057, 1048, 929, 777, 769, 465, 449, 440, 249, 241, 217, 209
        diferencial = -4
      Case 1664, 1632, 1655, 1623, 1512, 1088, 1079, 1056, 1047, 928, 776, 768, 464, 448, 439, 248, 240, 216, 208
        diferencial = -7
      Case 1663, 1631, 1654, 1622, 1511, 1087, 1078, 1055, 1046, 927, 775, 767, 463, 447, 438, 247, 239, 215, 207
        diferencial = -9
      Case 1662, 1630, 1653, 1621, 1510, 1086, 1077, 1054, 1045, 926, 774, 766, 462, 446, 437, 246, 238, 214, 206
        diferencial = -11
      Case 1661, 1629, 1652, 1620, 1509, 1085, 1076, 1053, 1044, 925, 774, 765, 461, 445, 436, 246, 237, 214, 205
        diferencial = -13

    End Select
    
End Select

limite_borde_derecho = POSX_INI_CAR + (16 * (posInMatrix - 1)) + diferencial - coche.Ancho


With coche
'Limitaciones del coche respecto al ancho de la carretera
  If .posX >= limite_borde_derecho Then
    flagColision = True
    If Not (sb(freno1) Is Nothing) Then
      If Not (sb(freno1).GetStatus = DSBSTATUS_PLAYING) Then ReproducirSonido freno1
    End If
    .posX = limite_borde_derecho - 1
  End If
  
  If .posX <= limite_borde_izquierdo Then
    flagColision = True
    If Not (sb(freno1) Is Nothing) Then
      If Not (sb(freno1).GetStatus = DSBSTATUS_PLAYING) Then ReproducirSonido freno1
    End If
    .posX = limite_borde_izquierdo + 1
  End If
End With

If flagColision Then

  VelocidadInicialColision = vel
  
  If VelocidadInicialColision >= 10 Then
    'vel = 0
    'ReproducirSonido crash2
    'Choque_en_Curso = True
    'CocheActivo = False
    'EfectosColisionCoche
    PrepararColision False
  Else
    ReproducirSonido crash1
    If vel > 1# Then vel = vel - 1#
  End If
    
End If

End Sub

Private Function GetTipoBorde(ByVal LadoDelBorde As Integer, ByRef posInMatrix As Integer) As String

Const POSICION_BORDE_IZQUIERDO = 11
Const POSICION_BORDE_DERECHO = 12

Dim IndexMapa As Integer


' Obtener la correspondencia entre la posicion en pantalla y la posicion en el mapa
IndexMapa = CInt((RowIndex + 15) / 8) - 1


' Obtener la posicion del borde izquierdo en el mapa
If LadoDelBorde = 0 Then
  posInMatrix = CInt(Mid$(Mapa(IndexMapa), POSICION_BORDE_IZQUIERDO, 1))
Else
  posInMatrix = CInt(Mid$(Mapa(IndexMapa), POSICION_BORDE_DERECHO, 1))
  ' Si se obtiene un 0, en realidad es un 10
  If posInMatrix = 0 Then posInMatrix = 10
End If

' Obtener el borde propiamente
GetTipoBorde = Mid$(Mapa(IndexMapa), posInMatrix, 1)

' Ya sé como es y su posicion en el mapa

End Function

Private Sub PintarEnemigos(ByRef Vehiculo As t_Sprite, ByRef rectOrig As RECT, ByRef rectDest As RECT)
    
    With rectOrig
      .Left = 0:  .Top = 0:  .Right = Vehiculo.Ancho:  .Bottom = Vehiculo.Alto
    End With
    
    With rectDest
      .Left = Vehiculo.posX
      .Top = Vehiculo.posY
      .Right = Vehiculo.posX + Vehiculo.Ancho
      .Bottom = Vehiculo.posY + Vehiculo.Alto
    
      ' Para que el coche malo1 desaparezca gradualmente
      ' clipping superior (Si el coche malo1 se sale por arriba)
      If .Top <= 0 Then
        rectOrig.Top = Abs(Vehiculo.posY)
        rectOrig.Bottom = Vehiculo.Alto
        .Top = 0
        .Bottom = rectOrig.Bottom - rectOrig.Top
      End If
      
      ' clipping inferior (Si el Vehiculo no cabe entero por abajo)
      If .Bottom >= 240 Then
        rectOrig.Bottom = Vehiculo.Alto - (.Bottom - 240)
        .Bottom = 240
      End If
    
    End With

End Sub

Private Sub ComprobarFrecuencias(ByVal Vehiculo As Integer)

Select Case Vehiculo

  Case 0 ' Camion
  
    ' Comprobar si ha transcurrido el suficiente tiempo como para mostrar el camion
    If SegundosAcumuladosSinMostrarCamion = TiempoNecesarioParaMostrarCamion Then
      CamionActivo = True ' Se debe mostrar el camion en la carretera
      SegundosAcumuladosSinMostrarCamion = 0 ' Resetear el acumulador
      GenerarFrecuenciaMostrarCamion ' Generar una nuevo tiempo para mostrar el camion
    Else
      ' Ir acumulando los segundos que lleva sin aparecer el camion
      SegundosAcumuladosSinMostrarCamion = SegundosAcumuladosSinMostrarCamion + 1
    End If
    
  Case 1 ' Camion2
  
    ' Comprobar si ha transcurrido el suficiente tiempo como para mostrar el camion
    If SegundosAcumuladosSinMostrarCamion2 = TiempoNecesarioParaMostrarCamion2 Then
      Camion2Activo = True ' Se debe mostrar el camion en la carretera
      SegundosAcumuladosSinMostrarCamion2 = 0 ' Resetear el acumulador
      GenerarFrecuenciaMostrarCamion2 ' Generar una nuevo tiempo para mostrar el camion
    Else
      ' Ir acumulando los segundos que lleva sin aparecer el camion
      SegundosAcumuladosSinMostrarCamion2 = SegundosAcumuladosSinMostrarCamion2 + 1
    End If

  Case 2 ' CocheMalo1

    ' Comprobar si ha transcurrido el suficiente tiempo como para mostrar el cochemalo1
    If SegundosAcumuladosSinMostrarCochemalo1 = TiempoNecesarioParaMostrarCocheMalo1 Then
      CocheMalo1Activo = True ' Se debe mostrar el coche
      SegundosAcumuladosSinMostrarCochemalo1 = 0 ' Resetear el acumulador
      ' Generar una nuevo tiempo para mostrar el cochemalo1
      TiempoNecesarioParaMostrarCocheMalo1 = GenerarFrecuenciaMostrarCocheX(1)
    Else
      ' Ir acumulando los segundos que lleva sin aparecer el coche
      SegundosAcumuladosSinMostrarCochemalo1 = SegundosAcumuladosSinMostrarCochemalo1 + 1
    End If
  
  Case 3 ' CocheMalo2
      
    ' Comprobar si ha transcurrido el suficiente tiempo como para mostrar el cochemalo1
    If SegundosAcumuladosSinMostrarCochemalo2 = TiempoNecesarioParaMostrarCocheMalo2 Then
      CocheMalo2Activo = True ' Se debe mostrar el coche
      SegundosAcumuladosSinMostrarCochemalo2 = 0 ' Resetear el acumulador
      ' Generar una nuevo tiempo para mostrar el cochemalo1
      TiempoNecesarioParaMostrarCocheMalo2 = GenerarFrecuenciaMostrarCocheX(2)
    Else
      ' Ir acumulando los segundos que lleva sin aparecer el coche
      SegundosAcumuladosSinMostrarCochemalo2 = SegundosAcumuladosSinMostrarCochemalo2 + 1
    End If
    
  Case 4 ' CocheMalo3
  
    ' Comprobar si ha transcurrido el suficiente tiempo como para mostrar el cochemalo1
    If SegundosAcumuladosSinMostrarCochemalo3 = TiempoNecesarioParaMostrarCocheMalo3 Then
      CocheMalo3Activo = True ' Se debe mostrar el coche
      SegundosAcumuladosSinMostrarCochemalo3 = 0 ' Resetear el acumulador
      ' Generar una nuevo tiempo para mostrar el cochemalo1
      TiempoNecesarioParaMostrarCocheMalo3 = GenerarFrecuenciaMostrarCocheX(3)
    Else
      ' Ir acumulando los segundos que lleva sin aparecer el coche
      SegundosAcumuladosSinMostrarCochemalo3 = SegundosAcumuladosSinMostrarCochemalo3 + 1
    End If
  
  Case 5 ' CocheMalo4
  
    ' Comprobar si ha transcurrido el suficiente tiempo como para mostrar el cochemalo1
    If SegundosAcumuladosSinMostrarCochemalo4 = TiempoNecesarioParaMostrarCocheMalo4 Then
      CocheMalo4Activo = True ' Se debe mostrar el coche
      SegundosAcumuladosSinMostrarCochemalo4 = 0 ' Resetear el acumulador
      ' Generar una nuevo tiempo para mostrar el cochemalo1
      TiempoNecesarioParaMostrarCocheMalo4 = GenerarFrecuenciaMostrarCocheX(4)
    Else
      ' Ir acumulando los segundos que lleva sin aparecer el coche
      SegundosAcumuladosSinMostrarCochemalo4 = SegundosAcumuladosSinMostrarCochemalo4 + 1
    End If
    
End Select

End Sub

Private Sub ControlarFinTramo()
    
  ' Comprobar si se ha llegado a la meta
  If RowIndex <= 0 Then
  
    ' Doy por hecho que un tramo consiste en realizar dos vueltas al mapa
    VueltasCompletadas = VueltasCompletadas + 1
        
    If VueltasCompletadas < VUELTAS_COMPLETADAS Then
      RowIndex = TRAMO - MAX_ROWS_SCREEN
      
    Else
      CelebrarFinNivel
      PrepararNuevoNivel
    End If
    
  End If
  
End Sub

Private Sub PintarHumos(ByRef Obj As t_Sprite, ByRef rectOrig As RECT, ByRef rectDest As RECT)

  With rectOrig
    .Left = 0: .Top = 0: .Right = Obj.Ancho: .Bottom = Obj.Alto
  End With
    
  With rectDest
    .Left = coche.posX + 5
    .Top = coche.posY + coche.Alto
    .Right = .Left + Obj.Ancho
    .Bottom = .Top + Obj.Alto
  End With

End Sub
Private Sub PintarExplosion(ByRef Vehiculo As t_Sprite, ByRef rectOrig As RECT, ByRef rectDest As RECT)

  With rectOrig
    .Left = 0: .Top = 0: .Right = DescExplosion.lWidth: .Bottom = DescExplosion.lHeight
  End With
    
  With rectDest
    .Left = Vehiculo.posX
    .Top = Vehiculo.posY
    .Right = .Left + rectOrig.Right
    .Bottom = .Top + rectOrig.Bottom
  End With

End Sub

Private Sub EfectoExplosion()
  
  ' Producir efecto de explosion
  If Posiciones_Choque(0) = False Then
    Set SupExplosion(0) = SupExplosion(1)
    Posiciones_Choque(0) = True
  ElseIf Posiciones_Choque(1) = False Then
    Set SupExplosion(0) = SupExplosion(2)
    Posiciones_Choque(1) = True
  ElseIf Posiciones_Choque(2) = False Then
    Set SupExplosion(0) = SupExplosion(3)
    Posiciones_Choque(2) = True
  ElseIf Posiciones_Choque(3) = False Then
    Set SupExplosion(0) = SupExplosion(4)
    Posiciones_Choque(3) = True
  ElseIf Posiciones_Choque(4) = False Then
    Set SupExplosion(0) = SupExplosion(5)
    Posiciones_Choque(4) = True
  ElseIf Posiciones_Choque(5) = False Then
    Set SupExplosion(0) = SupExplosion(6)
    Posiciones_Choque(5) = True
  ElseIf Posiciones_Choque(6) = False Then
    Set SupExplosion(0) = SupExplosion(7)
    Posiciones_Choque(6) = True
  ElseIf Posiciones_Choque(7) = False Then
    Set SupExplosion(0) = SupExplosion(8)
    Posiciones_Choque(7) = True
  ElseIf Posiciones_Choque(8) = False Then
    Set SupExplosion(0) = SupExplosion(9)
    Posiciones_Choque(8) = True
  ElseIf Posiciones_Choque(9) = False Then
    Set SupExplosion(0) = SupExplosion(10)
    Posiciones_Choque(9) = True
  ElseIf Posiciones_Choque(10) = False Then
    Set SupExplosion(0) = SupExplosion(11)
    Posiciones_Choque(10) = True
  ElseIf Posiciones_Choque(11) = False Then
    Set SupExplosion(0) = SupExplosion(12)
    Posiciones_Choque(11) = True
  ElseIf Posiciones_Choque(12) = False Then
    Set SupExplosion(0) = SupExplosion(13)
    Posiciones_Choque(12) = True
  Else
    ' Volver a desactivar el flag y resetear el vector de posiciones
    Choque_en_Curso = False
    VelocidadInicialColision = 0
    Dim i As Integer
    For i = 0 To 12: Posiciones_Choque(i) = False: Next
    
    ' Se restan puntos
    If puntos >= 1 Then puntos = puntos - 1
      
  End If
  
End Sub

Private Sub PrepararColision(flagColisionCamion As Boolean)

  Choque_en_Curso = True
  VelocidadInicialColision = vel
    
  ' Si el coche choca contra un camion, también debe explotar. Es el mismo caso que
  ' si fuera máxima velocidad
  If flagColisionCamion Then VelocidadInicialColision = 10
    
  If VelocidadInicialColision >= 10 Then
    vel = 0
    ReproducirSonido crash2
    CocheActivo = False
  Else
    ReproducirSonido crash1
  End If
    
  EfectosColisionCoche
    
End Sub
Private Sub EfectoExplosionCochesEnemigos(ByVal CocheEnemigo As Integer)
  
Dim i As Integer

Select Case CocheEnemigo
  
  Case 1
  
    ' Producir efecto de explosion
    If Posiciones_Choque_CocheMalo1(0) = False Then
      Set SupExplosion(0) = SupExplosion(1)
      Posiciones_Choque_CocheMalo1(0) = True
    ElseIf Posiciones_Choque_CocheMalo1(1) = False Then
      Set SupExplosion(0) = SupExplosion(2)
      Posiciones_Choque_CocheMalo1(1) = True
    ElseIf Posiciones_Choque_CocheMalo1(2) = False Then
      Set SupExplosion(0) = SupExplosion(3)
      Posiciones_Choque_CocheMalo1(2) = True
    ElseIf Posiciones_Choque_CocheMalo1(3) = False Then
      Set SupExplosion(0) = SupExplosion(4)
      Posiciones_Choque_CocheMalo1(3) = True
    ElseIf Posiciones_Choque_CocheMalo1(4) = False Then
      Set SupExplosion(0) = SupExplosion(5)
      Posiciones_Choque_CocheMalo1(4) = True
    ElseIf Posiciones_Choque_CocheMalo1(5) = False Then
      Set SupExplosion(0) = SupExplosion(6)
      Posiciones_Choque_CocheMalo1(5) = True
    ElseIf Posiciones_Choque_CocheMalo1(6) = False Then
      Set SupExplosion(0) = SupExplosion(7)
      Posiciones_Choque_CocheMalo1(6) = True
    ElseIf Posiciones_Choque_CocheMalo1(7) = False Then
      Set SupExplosion(0) = SupExplosion(8)
      Posiciones_Choque_CocheMalo1(7) = True
    ElseIf Posiciones_Choque_CocheMalo1(8) = False Then
      Set SupExplosion(0) = SupExplosion(9)
      Posiciones_Choque_CocheMalo1(8) = True
    ElseIf Posiciones_Choque_CocheMalo1(9) = False Then
      Set SupExplosion(0) = SupExplosion(10)
      Posiciones_Choque_CocheMalo1(9) = True
    ElseIf Posiciones_Choque_CocheMalo1(10) = False Then
      Set SupExplosion(0) = SupExplosion(11)
      Posiciones_Choque_CocheMalo1(10) = True
    ElseIf Posiciones_Choque_CocheMalo1(11) = False Then
      Set SupExplosion(0) = SupExplosion(12)
      Posiciones_Choque_CocheMalo1(11) = True
    ElseIf Posiciones_Choque_CocheMalo1(12) = False Then
      Set SupExplosion(0) = SupExplosion(13)
      Posiciones_Choque_CocheMalo1(12) = True
    Else
      ' Volver a desactivar el flag y resetear el vector de posiciones
      CocheMalo1.posY = POSY_INI_BAD_CAR
      ChoqueCocheMalo1EnCurso = False
      CocheMalo1ContraCamion = False
      
      For i = 0 To 12: Posiciones_Choque_CocheMalo1(i) = False: Next
      
    End If
  
  Case 2
  
    ' Producir efecto de explosion
    If Posiciones_Choque_CocheMalo2(0) = False Then
      Set SupExplosion(0) = SupExplosion(1)
      Posiciones_Choque_CocheMalo2(0) = True
    ElseIf Posiciones_Choque_CocheMalo2(1) = False Then
      Set SupExplosion(0) = SupExplosion(2)
      Posiciones_Choque_CocheMalo2(1) = True
    ElseIf Posiciones_Choque_CocheMalo2(2) = False Then
      Set SupExplosion(0) = SupExplosion(3)
      Posiciones_Choque_CocheMalo2(2) = True
    ElseIf Posiciones_Choque_CocheMalo2(3) = False Then
      Set SupExplosion(0) = SupExplosion(4)
      Posiciones_Choque_CocheMalo2(3) = True
    ElseIf Posiciones_Choque_CocheMalo2(4) = False Then
      Set SupExplosion(0) = SupExplosion(5)
      Posiciones_Choque_CocheMalo2(4) = True
    ElseIf Posiciones_Choque_CocheMalo2(5) = False Then
      Set SupExplosion(0) = SupExplosion(6)
      Posiciones_Choque_CocheMalo2(5) = True
    ElseIf Posiciones_Choque_CocheMalo2(6) = False Then
      Set SupExplosion(0) = SupExplosion(7)
      Posiciones_Choque_CocheMalo2(6) = True
    ElseIf Posiciones_Choque_CocheMalo2(7) = False Then
      Set SupExplosion(0) = SupExplosion(8)
      Posiciones_Choque_CocheMalo2(7) = True
    ElseIf Posiciones_Choque_CocheMalo2(8) = False Then
      Set SupExplosion(0) = SupExplosion(9)
      Posiciones_Choque_CocheMalo2(8) = True
    ElseIf Posiciones_Choque_CocheMalo2(9) = False Then
      Set SupExplosion(0) = SupExplosion(10)
      Posiciones_Choque_CocheMalo2(9) = True
    ElseIf Posiciones_Choque_CocheMalo2(10) = False Then
      Set SupExplosion(0) = SupExplosion(11)
      Posiciones_Choque_CocheMalo2(10) = True
    ElseIf Posiciones_Choque_CocheMalo2(11) = False Then
      Set SupExplosion(0) = SupExplosion(12)
      Posiciones_Choque_CocheMalo2(11) = True
    ElseIf Posiciones_Choque_CocheMalo2(12) = False Then
      Set SupExplosion(0) = SupExplosion(13)
      Posiciones_Choque_CocheMalo2(12) = True
    Else
      ' Volver a desactivar el flag y resetear el vector de posiciones
      CocheMalo2.posY = POSY_INI_BAD_CAR
      ChoqueCocheMalo2EnCurso = False
      CocheMalo2ContraCamion = False
    
      For i = 0 To 12: Posiciones_Choque_CocheMalo2(i) = False: Next
      
    End If
  
  Case 3
  
    ' Producir efecto de explosion
    If Posiciones_Choque_CocheMalo3(0) = False Then
      Set SupExplosion(0) = SupExplosion(1)
      Posiciones_Choque_CocheMalo3(0) = True
    ElseIf Posiciones_Choque_CocheMalo3(1) = False Then
      Set SupExplosion(0) = SupExplosion(2)
      Posiciones_Choque_CocheMalo3(1) = True
    ElseIf Posiciones_Choque_CocheMalo3(2) = False Then
      Set SupExplosion(0) = SupExplosion(3)
      Posiciones_Choque_CocheMalo3(2) = True
    ElseIf Posiciones_Choque_CocheMalo3(3) = False Then
      Set SupExplosion(0) = SupExplosion(4)
      Posiciones_Choque_CocheMalo3(3) = True
    ElseIf Posiciones_Choque_CocheMalo3(4) = False Then
      Set SupExplosion(0) = SupExplosion(5)
      Posiciones_Choque_CocheMalo3(4) = True
    ElseIf Posiciones_Choque_CocheMalo3(5) = False Then
      Set SupExplosion(0) = SupExplosion(6)
      Posiciones_Choque_CocheMalo3(5) = True
    ElseIf Posiciones_Choque_CocheMalo3(6) = False Then
      Set SupExplosion(0) = SupExplosion(7)
      Posiciones_Choque_CocheMalo3(6) = True
    ElseIf Posiciones_Choque_CocheMalo3(7) = False Then
      Set SupExplosion(0) = SupExplosion(8)
      Posiciones_Choque_CocheMalo3(7) = True
    ElseIf Posiciones_Choque_CocheMalo3(8) = False Then
      Set SupExplosion(0) = SupExplosion(9)
      Posiciones_Choque_CocheMalo3(8) = True
    ElseIf Posiciones_Choque_CocheMalo3(9) = False Then
      Set SupExplosion(0) = SupExplosion(10)
      Posiciones_Choque_CocheMalo3(9) = True
    ElseIf Posiciones_Choque_CocheMalo3(10) = False Then
      Set SupExplosion(0) = SupExplosion(11)
      Posiciones_Choque_CocheMalo3(10) = True
    ElseIf Posiciones_Choque_CocheMalo3(11) = False Then
      Set SupExplosion(0) = SupExplosion(12)
      Posiciones_Choque_CocheMalo3(11) = True
    ElseIf Posiciones_Choque_CocheMalo3(12) = False Then
      Set SupExplosion(0) = SupExplosion(13)
      Posiciones_Choque_CocheMalo3(12) = True
    Else
      ' Volver a desactivar el flag y resetear el vector de posiciones
      CocheMalo3.posY = POSY_INI_BAD_CAR
      ChoqueCocheMalo3EnCurso = False
      CocheMalo3ContraCamion = False
    
      For i = 0 To 12: Posiciones_Choque_CocheMalo3(i) = False: Next
      
    End If
  
  Case 4
  
    ' Producir efecto de explosion
    If Posiciones_Choque_CocheMalo4(0) = False Then
      Set SupExplosion(0) = SupExplosion(1)
      Posiciones_Choque_CocheMalo4(0) = True
    ElseIf Posiciones_Choque_CocheMalo4(1) = False Then
      Set SupExplosion(0) = SupExplosion(2)
      Posiciones_Choque_CocheMalo4(1) = True
    ElseIf Posiciones_Choque_CocheMalo4(2) = False Then
      Set SupExplosion(0) = SupExplosion(3)
      Posiciones_Choque_CocheMalo4(2) = True
    ElseIf Posiciones_Choque_CocheMalo4(3) = False Then
      Set SupExplosion(0) = SupExplosion(4)
      Posiciones_Choque_CocheMalo4(3) = True
    ElseIf Posiciones_Choque_CocheMalo4(4) = False Then
      Set SupExplosion(0) = SupExplosion(5)
      Posiciones_Choque_CocheMalo4(4) = True
    ElseIf Posiciones_Choque_CocheMalo4(5) = False Then
      Set SupExplosion(0) = SupExplosion(6)
      Posiciones_Choque_CocheMalo4(5) = True
    ElseIf Posiciones_Choque_CocheMalo4(6) = False Then
      Set SupExplosion(0) = SupExplosion(7)
      Posiciones_Choque_CocheMalo4(6) = True
    ElseIf Posiciones_Choque_CocheMalo4(7) = False Then
      Set SupExplosion(0) = SupExplosion(8)
      Posiciones_Choque_CocheMalo4(7) = True
    ElseIf Posiciones_Choque_CocheMalo4(8) = False Then
      Set SupExplosion(0) = SupExplosion(9)
      Posiciones_Choque_CocheMalo4(8) = True
    ElseIf Posiciones_Choque_CocheMalo4(9) = False Then
      Set SupExplosion(0) = SupExplosion(10)
      Posiciones_Choque_CocheMalo4(9) = True
    ElseIf Posiciones_Choque_CocheMalo4(10) = False Then
      Set SupExplosion(0) = SupExplosion(11)
      Posiciones_Choque_CocheMalo4(10) = True
    ElseIf Posiciones_Choque_CocheMalo4(11) = False Then
      Set SupExplosion(0) = SupExplosion(12)
      Posiciones_Choque_CocheMalo4(11) = True
    ElseIf Posiciones_Choque_CocheMalo4(12) = False Then
      Set SupExplosion(0) = SupExplosion(13)
      Posiciones_Choque_CocheMalo4(12) = True
    Else
      ' Volver a desactivar el flag y resetear el vector de posiciones
      CocheMalo4.posY = POSY_INI_BAD_CAR
      ChoqueCocheMalo4EnCurso = False
      CocheMalo4ContraCamion = False
    
      For i = 0 To 12: Posiciones_Choque_CocheMalo4(i) = False: Next
      
    End If
    
End Select

End Sub

Private Function ChoqueConObstaculos(ByRef Obj1 As CoordenadasObstaculos, ByRef Obj2 As CoordenadasObstaculos)

Dim rect1 As RECT, rect2 As RECT, TempRect As RECT, ret As Boolean

If Abs(RowIndex + DESP_POSY_CAR_FROM_TOP - Obj1.PosYPantalla - 1) < 2 Then
  ' Recta que determina el charco
  With rect1
  .Top = DESP_POSY_CAR_FROM_TOP - 16: .Left = Obj1.PosXPantalla
  .Right = .Left + 16:  .Bottom = .Top + 32
  End With

  ' Recta que determina el coche
  With rect2
    .Top = DESP_POSY_CAR_FROM_TOP:  .Left = coche.posX
    .Right = .Left + coche.Ancho:  .Bottom = .Top + coche.Alto
  End With

  If IntersectRect(TempRect, rect1, rect2) Then ret = True
  
ElseIf Abs(RowIndex + DESP_POSY_CAR_FROM_TOP - Obj2.PosYPantalla - 1) < 2 Then
  
  ' Recta que determina el charco
  With rect1
  .Top = DESP_POSY_CAR_FROM_TOP: .Left = Obj2.PosXPantalla
  .Right = .Left + 16:  .Bottom = .Top + 16
  End With

  ' Recta que determina el coche
  With rect2
    .Top = DESP_POSY_CAR_FROM_TOP:  .Left = coche.posX
    .Right = .Left + coche.Ancho:  .Bottom = .Top + coche.Alto
  End With

  If IntersectRect(TempRect, rect1, rect2) Then ret = True
  
End If

ChoqueConObstaculos = ret

End Function
