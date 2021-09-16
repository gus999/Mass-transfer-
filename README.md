# Mass-transfer-
This Program iterates a solution concentration to find a wide array of results. Plotting of results and Video representation is also shown
Option Explicit
Dim M As Double, Csat As Double, Vd As Double, Dp As Double, Cdes As Double
Dim Dens As Double, PM As Double, mT As Double, Cmax As Double, mNec As Double
Dim KgNec As Double, Vpart As Double, mPP As Double, NumPart As Double
Dim MTrans As Double, Dif As Double, Cao As Double, Mquedan As Double
Dim MquedanPP As Double, DnPart As Double, Vmf As Double, KgMax As Double
Dim Conta As Double, IMTrans As Double, KgTrans As Double, Iti As Double
Dim Cfin As Double, Error As Double, a As Integer
Private Sub CmbResultado1_Click()
CmbResultado2.ListIndex = CmbResultado1.ListIndex
End Sub
Private Sub Cmbsistemas_Click()
CmbDifusividad.ListIndex = Cmbsistemas.ListIndex
CmbCsat.ListIndex = Cmbsistemas.ListIndex
CmbDensidad.ListIndex = Cmbsistemas.ListIndex
CmbPM.ListIndex = Cmbsistemas.ListIndex
End Sub

Private Sub CmdCalcular_Click()
'Limpiando combos y gráficas
CmbResultado1.Clear
CmbResultado2.Clear
Picture1.Cls
Picture2.Cls
Picture3.Cls
Picture4.Cls
'Verificando que se haya escogido un sistema
If Cmbsistemas.ListIndex = -1 Then
    MsgBox "Debes de escoger un sistema para realizar los cálculos", vbInformation, "REVISAR"
    Exit Sub
End If
'Tomando valores de las entradas, dependiendo de la computadora y el visual basic usado puede tomar
'los valores erroneamente, ya que dependiendo de estos sistemas, se convierten los puntos en comas,
'y los valores cambian por este detalle "PONERLE ATENCIÓN Y CORREGIR ESTO DEPENDIENDO DE TU SISTEMA"
Dif = CmbDifusividad.Text
M = TxtMasa.Text
Csat = CmbCsat.Text
Cao = TxtCao.Text
Vd = TxtVd.Text
Dp = TxtDp.Text
Cdes = TxtCd.Text
Dens = CmbDensidad.Text
PM = CmbPM.Text
Iti = TxtIt.Text
'Haciendo cálculos iniciales
mT = M / PM 'Moles totales de entrada
Cmax = mT / Vd 'Concentración máxima alcanzable con la masa que entra
KgMax = Csat * Vd * PM 'Kg máximos de soluto que pueden entrar al sistema
mNec = Cdes * Vd 'Moles necesarios para alcanzar la concentración deseada y volumen deseado
KgNec = mNec * PM 'Masa necesaria de entrada para alcanzar la concentración y volumen deseado
Vpart = 3.141593 / 6 * Dp ^ 3 'Volumen inicial de la partícula con su diámetro
mPP = Dens * Vpart / PM 'Moles por partícula
NumPart = mT / mPP 'Número de partículas presentes en el sistema
MTrans = 2 * 3.141593 * Dif * Dp * (Csat - Cao) * NumPart * Iti 'Moles transferidos totales en el primer intervalo
Mquedan = mT 'Igualando los moles que quedan sin disolver a los moles totales
'Verificando que la concentración máxima alcanzable no sea mayor a la de saturación para evitar
'un sistema sobresaturado, y se le imprime cual es la cantidad máxima de masa que puede ingresar
If Cmax > Csat Then
    MsgBox "Cantidad excesiva de soluto, la cantidad necesaria debería ser igual o menor a" + Str(KgMax) + " por la Csat", vbExclamation, "Error"
    TxtMasa.Text = KgMax
Exit Sub
End If
'Verificando que la concentración deseada no sea mayor que la máxima alcanzable, en este caso tambien se le imprimen los KgMax
If Cdes > Cmax Then
    MsgBox "La concentración deseada es mayor a la concentración máxima, el máximo de concentración, con la masa que entra es " + Str(Cmax) + " (Kmol/m^3) y la puedes alcanzar con" + Str(KgMax) + "(Kg)", vbExclamation, "Error"
    TxtMasa.Text = KgMax: TxtCd.Text = Csat
Exit Sub
End If
'Iniciando los contadores en cero
Conta = 0: Vmf = 0: IMTrans = 0
'Empezando el bucle para los cálculos
Do
    Conta = Conta + 1
    'Verificando que los moles que quedan no sean menores a los transferidos, para evitar cálculos con moléculas que ya se consumieron y ya no existen
    If Mquedan >= MTrans Then
        Mquedan = Mquedan - MTrans 'Como los moles que quedan son menores a los transferidos se calculan los nuevos moles que quedan
        Else 'Los moles transferidos son mayores a los existentes
        MsgBox "Sucedió que los moles transferidos son mas que los moles que quedan, esto puede deberse al intervalo de tiempo usado,INTÉNTALO CON OTRO Y DEFINELO DE ACUERDO AL ULTIMO DATO DE CONCENTRACIÓN OBTENIDO", vbCritical, "Cambiar ITI (Verificar)"
        TxtIt.Text = Iti / 2.5 'Como los moles transferidos dependen del tiempo se intenta reducir este dividiendo entre un factor arbitarrio
        'aunque se recomienda que el usuario defina el intervalo de tiempo nuevo y mas pequeño, partiendo del hecho de que se le imprimen
        'las concentaciones finales en un tiempo determinado y si esta cerca de la concentración deseada o muy lejos defina con respecto a esto
        'Colocandose en la ultima iteración realizada para que el usuario tenga una idea de que tan pequeño deba hacer el intervalo que ya usó
        CmbResultado1.ListIndex = CmbResultado1.ListCount - 1
        CmdGraficar.Enabled = True
        CmbResultado1.Visible = True: CmbResultado2.Visible = True
        Label21.Visible = True: Label22.Visible = True: Label23.Visible = True: Label24.Visible = True
        Label25.Visible = True: Label26.Visible = True
        FormPmasa.Height = 8790: FormPmasa.Width = 11700
        Exit Sub
    End If
    'Ya calculados los moles que quedan se procede a calcular los que quedan pero por partícula
    'Esto para poder calcular el nuevo volumen y con este el diámetro nuevo de la partícula, con este su velocidad mínima de fluidización
    'y los nuevos moles transferidos y se hace un acumulado de los moles tansferidos, con el cual se calcula la concentración que ya se alcanzó
    'y por último la cantidad de masa transferida
    MquedanPP = Mquedan / NumPart 'Moles que quedan por partícula
    Vpart = MquedanPP / Dens * PM 'Volumen nuevo de la partícula
    DnPart = (Vpart * 6 / 3.141593) ^ (1 / 3) 'Diámetro nuevo de la partícula
    Vmf = ((0.001 / (DnPart * 1000))) * (Sqr((33.7 ^ 2) + ((0.0408 * (DnPart ^ 3) * 1000 * (Dens - 1000) * 9.81) / (0.001 ^ 2))) - 33.7) 'Velocidad mínima de fluidización
    MTrans = 2 * 3.141593 * Dif * DnPart * (Csat - Cao) * NumPart * Iti 'Nuevos moles transferidos
    IMTrans = IMTrans + MTrans 'Acumulado de moles transferidos
    Cfin = IMTrans / Vd 'concentracion alcanzada
    KgTrans = IMTrans * PM 'Masa transferida total
    'Colocación de los resultados de la iteración en los combos
    CmbResultado1.AddItem ((Conta * Iti) & "   " & Vmf & "   " & DnPart)
    CmbResultado2.AddItem ((Conta * Iti) & "       " & (M - KgTrans) & "              " & Cfin)
    If Conta > 32765 Then '32765 es el numero de datos m'aximos que soporta el combo
        MsgBox "Debes de utilizar un intervalo de tiempo mayor para poder hallar tu concentracion deseada pues son muchas iteraciones y no pueden mostrarse todas, SE RECOMIENDA QUE TU DEFINAS EL NUEVO INTERVALO, DE ACUERO A LA CONCENTRACIÓN FINAL ALCANZADA", vbCritical, "CAMBIAR ITI (Verificar)"
        TxtIt.Text = Iti * (1.5) 'se multiplica por un factor arbitrario para aumentar el intervalo
        'Mostrando los resultados obtenidos
        CmbResultado1.ListIndex = CmbResultado1.ListCount - 1
        CmdGraficar.Enabled = True
        CmbResultado1.Visible = True: CmbResultado2.Visible = True
        Label21.Visible = True: Label22.Visible = True: Label23.Visible = True: Label24.Visible = True
        Label25.Visible = True: Label26.Visible = True
        FormPmasa.Height = 8790: FormPmasa.Width = 11700
        Exit Sub
    End If
    Error = (Cdes - Cfin) / Cdes
Loop Until Error <= 0.0000001 Or Conta > 32765 'si se puede con este pero es lento
'Haciendo visibles los resultados
Label21.Visible = True: Label22.Visible = True: Label23.Visible = True: Label24.Visible = True
Label25.Visible = True: Label26.Visible = True
FormPmasa.Height = 8790: FormPmasa.Width = 15700
CmbResultado1.Visible = True: CmbResultado2.Visible = True
CmbResultado1.ListIndex = CmbResultado1.ListCount - 1
'Habilitando la posibilidad de graficar resultados
CmdGraficar.Enabled = True
End Sub

Private Sub CmdGraficar_Click()
CmdGraficar.Enabled = False
CmdCalcular.Enabled = False
Xwin.Visible = True
Dif = CmbDifusividad.Text
M = TxtMasa.Text
Csat = CmbCsat.Text
Cao = TxtCao.Text
Vd = TxtVd.Text
Dp = TxtDp.Text
Cdes = TxtCd.Text
Dens = CmbDensidad.Text
PM = CmbPM.Text
mT = M / PM
Cmax = mT / Vd
KgMax = Csat * Vd * PM
mNec = Cdes * Vd
KgNec = mNec * PM
Vpart = 3.141593 / 6 * Dp ^ 3
mPP = Dens * Vpart / PM
NumPart = mT / mPP
MTrans = 2 * 3.141593 * Dif * Dp * (Csat - Cao) * NumPart * Iti
Mquedan = mT
Vmf = ((0.001 * Dp) / 1000) * (Sqr((33.7 ^ 2) + ((0.0408 * (Dp ^ 3) * 1000 * (Dens - 1000) * 9.81) / (0.001 ^ 2))) - 33.7)
Label21.Visible = True: Label22.Visible = True: Label23.Visible = True: Label24.Visible = True
Label25.Visible = True: Label26.Visible = True
Label32.Visible = True: Label35.Visible = True: Label2.Visible = True: Label14.Visible = True
Label33.Visible = True: Label31.Visible = True: Label29.Visible = True: Label15.Visible = True: Label12.Visible = True
Lblx.Visible = True: LblX2.Visible = True: LblX3.Visible = True: LblX4.Visible = True
Lbly.Visible = True: LblY2.Visible = True: LblY3.Visible = True: LblY4.Visible = True
Picture1.Visible = True: Picture2.Visible = True: Picture3.Visible = True: Picture4.Visible = True

'Activando el ShockWave
FlashVid.Visible = True
FormPmasa.ScaleMode = 3
FlashVid.Movie = App.Path & "\esfera.swf"
FlashVid.Left = 700
FlashVid.Top = 120
FlashVid.Width = 300
FlashVid.Height = 380
FlashVid.Play

'GRAFICANDO
'ajustamos la escala,para graficar necesitamos los datos finales de las iteraciones para dimensionar los limites de las graficas
'grafica de velocidad
Picture1.ScaleMode = 3
Picture1.Scale (-0.1, Vmf)-(Conta * Iti, 0)
Picture1.Cls
'grafica de Cfin
Picture2.ScaleMode = 3
Picture2.Scale (-0.1, Cfin)-(Conta * Iti, Cao)
Picture2.Cls
'grafica de Dpart
Picture3.ScaleMode = 3
Picture3.Scale (-0.1, Dp)-(Conta * Iti, DnPart)
Picture3.Cls
'grafica de Acumulado de masa transferida
Picture4.ScaleMode = 3
Picture4.Scale (-0.1, IMTrans * PM)-(Conta * Iti, 0)
Picture4.Cls

'graficaremos los ejes de velocidad
Picture1.Line (0, 0)-(Conta * Iti, 0), vbBlack
Picture1.Line (0, 0)-(0, Vmf), vbBlack
'graficaremos los ejes de Cfin
Picture2.Line (0, 0)-(Conta * Iti, 0), vbBlack
Picture2.Line (0, Cao)-(0, Cfin), vbBlack
'graficaremos los ejes de dpart
Picture3.Line (0, 0)-(Conta * Iti, 0), vbBlack
Picture3.Line (0, DnPart)-(0, Dp), vbBlack
'graficaremos los ejes del acumulado de masa transferida
Picture4.Line (0, 0)-(Conta * Iti, 0), vbBlack
Picture4.Line (0, 0)-(0, IMTrans / PM), vbBlack
'Para poder graficar se necesitan de nuevo las datos de las iteraciones asi que
'SE TIENEN QUE HACER LOS CÁLCULOS DE NUEVO
Conta = 0: Vmf = 0: IMTrans = 0
Do
    Conta = Conta + 1
    If Mquedan >= MTrans Then
        Mquedan = Mquedan - MTrans
        Else
        Exit Sub
    End If
    MquedanPP = Mquedan / NumPart
    Vpart = MquedanPP / Dens * PM
    DnPart = (Vpart * 6 / 3.141593) ^ (1 / 3)
    Vmf = ((0.001 * DnPart) / 1000) * (Sqr((33.7 ^ 2) + ((0.0408 * (DnPart ^ 3) * 1000 * (1587.7 - 1000) * 9.81) / (0.001 ^ 2))) - 33.7)
    MTrans = 2 * 3.141593 * Dif * DnPart * (Csat - Cao) * NumPart * Iti
    IMTrans = IMTrans + MTrans
    Cfin = IMTrans / Vd
    KgTrans = IMTrans * PM
    Error = (Cdes - Cfin) / Cdes
    'graficaremos los puntos
    Picture1.PSet (Conta * Iti, Vmf), vbRed
    Picture2.PSet (Conta * Iti, Cfin), vbBlue
    Picture3.PSet (Conta * Iti, DnPart), vbBlack
    Picture4.PSet (Conta * Iti, IMTrans * PM), vbYellow
    'TERMINO DE LA GRAFICACIÓN
Loop Until Error <= 0.0000001 Or Conta > 32765
CmbResultado1.ListIndex = CmbResultado1.ListCount - 1
End Sub

Private Sub CmdReset_Click()
Call Form_Load
CmdCalcular.Enabled = True
FlashVid.Rewind
FlashVid.Visible = False

End Sub

Private Sub CmdSalir_Click()
End
End Sub


Private Sub Form_Load()

CmbResultado1.Visible = False
CmbResultado2.Visible = False
Label21.Visible = False: Label22.Visible = False: Label23.Visible = False
Label24.Visible = False: Label25.Visible = False: Label26.Visible = False
FormPmasa.Height = 9000: FormPmasa.Width = 8000
Xwin.Visible = False
Label32.Visible = False: Label35.Visible = False: Label2.Visible = False: Label14.Visible = False:
Label33.Visible = False: Label31.Visible = False: Label29.Visible = False: Label15.Visible = False: Label12.Visible = False:
Lblx.Visible = False: LblX2.Visible = False: LblX3.Visible = False: LblX4.Visible = False
Lbly.Visible = False: LblY2.Visible = False: LblY3.Visible = False: LblY4.Visible = False
Picture1.Visible = False: Picture2.Visible = False: Picture3.Visible = False: Picture4.Visible = False:

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Lblx = X
Lbly = Y
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblX2 = X
LblY2 = Y
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblX3 = X
LblY3 = Y
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblX4 = X
LblY4 = Y
End Sub

Private Sub TxtCao_Validate(Cancel As Boolean)
If IsNumeric(TxtCao.Text) = True Then
    If TxtCao.Text >= 0 Then
        Cancel = False
        Else
        MsgBox "Debes utilizar alguna concentración inicial", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para la concentración inicial", vbCritical, "ERROR"
    Cancel = True
End If
End Sub

Private Sub TxtCd_Validate(Cancel As Boolean)
If IsNumeric(TxtCd.Text) = True Then
    If TxtCd.Text > 0 Then
        Cancel = False
        Else
        MsgBox "Debes utilizar alguna concentración", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para la concentración deseada", vbCritical, "ERROR"
    Cancel = True
End If
End Sub

Private Sub TxtCsat_Validate(Cancel As Boolean)
If IsNumeric(TxtCsat.Text) = True Then
    If TxtCsat.Text > 0 Then
        Cancel = False
        Else
        MsgBox "Debes utilizar alguna concentración de saturación", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para la concentración de saturación", vbCritical, "ERROR"
    Cancel = True
End If
End Sub
Private Sub TxtDens_Validate(Cancel As Boolean)
If IsNumeric(TxtDens.Text) = True Then
    If TxtDens.Text > 0 Then
        Cancel = False
        Else
        MsgBox "Debes utilizar alguna densidad para el soluto", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para la densidad", vbCritical, "ERROR"
    Cancel = True
End If
End Sub

Private Sub TxtDif_Validate(Cancel As Boolean)
If IsNumeric(TxtDif.Text) = True Then
    If TxtDif.Text > 0 Then
        Cancel = False
        Else
        MsgBox "La difusividad de una partícula debe ser positiva", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para la difusividad", vbCritical, "ERROR"
    Cancel = True
End If
End Sub
Private Sub TxtDp_Validate(Cancel As Boolean)
If IsNumeric(TxtDp.Text) = True Then
    If TxtDp.Text > 0 Then
        Cancel = False
        Else
        MsgBox "Debes utilizar algún diámetro de partícula", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para el diámetro de la partícula", vbCritical, "ERROR"
    Cancel = True
End If
End Sub
Private Sub TxtIt_Validate(Cancel As Boolean)
If IsNumeric(TxtIt.Text) = True Then
    If TxtIt.Text > 0 Then
        Cancel = False
        Else
        MsgBox "Debes utilizar algun intervalo de tiempo, este no puede ser negativo", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para la concentración de saturación", vbCritical, "ERROR"
    Cancel = True
End If

End Sub

Private Sub TxtMasa_Validate(Cancel As Boolean)
If IsNumeric(TxtMasa.Text) = True Then
    If TxtMasa.Text > 0 Then
        Cancel = False
        Else
        MsgBox "Debes introducir algo de soluto para concentrar", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para la masa de entrada", vbCritical, "ERROR"
    Cancel = True
End If

End Sub
Private Sub TxtPM_Validate(Cancel As Boolean)
If IsNumeric(TxtPM.Text) = True Then
    If TxtPM.Text > 0 Then
        Cancel = False
        Else
        MsgBox "Debes poner el peso molecular del soluto", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para el peso molecular", vbCritical, "ERROR"
    Cancel = True
End If
End Sub

Private Sub TxtVd_Validate(Cancel As Boolean)
If IsNumeric(TxtVd.Text) = True Then
    If TxtVd.Text > 0 Then
        Cancel = False
        Else
        MsgBox "El volumen deseado debe ser mayor de cero", vbCritical, "ERROR"
        Cancel = True
    End If
    Else
    MsgBox "Debes de utilizar un número, no una letra o símbolo, para el volumen deseado", vbCritical, "ERROR"
    Cancel = True
End If

End Sub
