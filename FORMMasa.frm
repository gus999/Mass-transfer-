VERSION 5.00
Begin VB.Form FormPmasa 
   Caption         =   "Sistemas Agua-Soluto"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Xwin 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   63
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "Reset"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   62
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   1800
      ScaleHeight     =   1755
      ScaleWidth      =   2835
      TabIndex        =   56
      Top             =   4920
      Width           =   2895
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   7200
      ScaleHeight     =   1755
      ScaleWidth      =   2835
      TabIndex        =   51
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton CmdGraficar 
      Caption         =   "Graficar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      TabIndex        =   42
      Top             =   7440
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   7200
      ScaleHeight     =   1755
      ScaleWidth      =   2835
      TabIndex        =   41
      Top             =   3840
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   7200
      ScaleHeight     =   1755
      ScaleWidth      =   2835
      TabIndex        =   40
      Top             =   1560
      Width           =   2895
   End
   Begin VB.ComboBox CmbPM 
      Height          =   315
      ItemData        =   "FORMMasa.frx":0000
      Left            =   2040
      List            =   "FORMMasa.frx":001C
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "PM"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox CmbDensidad 
      Height          =   315
      ItemData        =   "FORMMasa.frx":005B
      Left            =   480
      List            =   "FORMMasa.frx":0077
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Densidad"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ComboBox CmbCsat 
      Height          =   315
      ItemData        =   "FORMMasa.frx":00BB
      Left            =   2040
      List            =   "FORMMasa.frx":00D7
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "Csat"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox CmbDifusividad 
      Height          =   315
      ItemData        =   "FORMMasa.frx":0135
      Left            =   480
      List            =   "FORMMasa.frx":0151
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "Difusividad"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.ComboBox Cmbsistemas 
      Height          =   315
      ItemData        =   "FORMMasa.frx":01A1
      Left            =   120
      List            =   "FORMMasa.frx":01BD
      TabIndex        =   32
      Text            =   "Sistemas"
      Top             =   360
      Width           =   2535
   End
   Begin VB.ComboBox CmbResultado2 
      Height          =   315
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "(t, (KgEnt-KgTrans),Cfin)"
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox TxtIt 
      Height          =   285
      Left            =   1440
      TabIndex        =   21
      Text            =   "2"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox TxtCao 
      Height          =   285
      Left            =   3960
      TabIndex        =   18
      Text            =   "0"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ComboBox CmbResultado1 
      Height          =   315
      Left            =   6240
      TabIndex        =   16
      Text            =   "(t, Vmf, Dpart)"
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   5280
      TabIndex        =   15
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton CmdCalcular 
      Caption         =   "Calcular"
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox TxtCd 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Text            =   "0.5"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox TxtDp 
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Text            =   "0.001"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox TxtMasa 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "12.8796"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox TxtVd 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "0.1"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label LblY4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   600
      TabIndex        =   61
      Top             =   5880
      Width           =   45
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   " transferida (Kg)"
      Height          =   195
      Left            =   360
      TabIndex        =   60
      Top             =   5520
      Width           =   1110
   End
   Begin VB.Label LblX4 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3360
      TabIndex        =   59
      Top             =   6840
      Width           =   45
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "Acumulado de masa"
      Height          =   195
      Left            =   240
      TabIndex        =   58
      Top             =   5280
      Width           =   1440
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "t(s)"
      Height          =   195
      Left            =   3000
      TabIndex        =   57
      Top             =   6840
      Width           =   210
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "Dpart (m)"
      Height          =   195
      Left            =   6360
      TabIndex        =   55
      Top             =   6840
      Width           =   645
   End
   Begin VB.Label LblY3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6120
      TabIndex        =   54
      Top             =   7200
      Width           =   45
   End
   Begin VB.Label LblX3 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8760
      TabIndex        =   53
      Top             =   7920
      Width           =   45
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "t(s)"
      Height          =   195
      Left            =   8400
      TabIndex        =   52
      Top             =   7920
      Width           =   210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      X1              =   3000
      X2              =   3000
      Y1              =   3240
      Y2              =   4560
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "t(s)"
      Height          =   195
      Left            =   8400
      TabIndex        =   50
      Top             =   5760
      Width           =   210
   End
   Begin VB.Label LblX2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8760
      TabIndex        =   49
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label LblY2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6120
      TabIndex        =   48
      Top             =   5040
      Width           =   45
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Cfin (Kmol/m^3)"
      Height          =   195
      Left            =   6000
      TabIndex        =   47
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "t(s)"
      Height          =   195
      Left            =   8280
      TabIndex        =   46
      Top             =   3600
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vmf (m/s)"
      Height          =   195
      Left            =   6360
      TabIndex        =   45
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label Lblx 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8640
      TabIndex        =   44
      Top             =   3600
      Width           =   45
   End
   Begin VB.Label Lbly 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   6120
      TabIndex        =   43
      Top             =   2640
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Variables definidas por el Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   39
      Top             =   2760
      Width           =   3600
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "Csat(Kmol/m^3)"
      Height          =   195
      Left            =   2040
      TabIndex        =   35
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   33
      Top             =   0
      Width           =   870
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Variables que dependen del sistema:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   31
      Top             =   840
      Width           =   3915
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Cfin(Kmol/m^3)"
      Height          =   195
      Left            =   9360
      TabIndex        =   30
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Dif.-entre KgEnt y KgTrans (Kg)"
      Height          =   195
      Left            =   6720
      TabIndex        =   29
      Top             =   840
      Width           =   2220
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "t(s)"
      Height          =   195
      Left            =   6360
      TabIndex        =   28
      Top             =   840
      Width           =   210
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Vmf(m/s)"
      Height          =   195
      Left            =   7440
      TabIndex        =   27
      Top             =   0
      Width           =   630
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Dpart(m)"
      Height          =   195
      Left            =   9000
      TabIndex        =   26
      Top             =   0
      Width           =   600
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "t(s)"
      Height          =   195
      Left            =   6360
      TabIndex        =   25
      Top             =   0
      Width           =   210
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Intervalo tiempo"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "s"
      Height          =   195
      Left            =   2640
      TabIndex        =   22
      Top             =   4200
      Width           =   75
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Cinicial"
      Height          =   195
      Left            =   3360
      TabIndex        =   20
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Kmol/m^3"
      Height          =   195
      Left            =   5160
      TabIndex        =   19
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Difusividad m^2/s"
      Height          =   195
      Left            =   480
      TabIndex        =   17
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Peso Molecular(Kg/Kmol)"
      Height          =   195
      Left            =   2040
      TabIndex        =   13
      Top             =   2040
      Width           =   1800
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Densidad Kg/m^3"
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Kmol/m^3"
      Height          =   195
      Left            =   5160
      TabIndex        =   11
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cdeseada"
      Height          =   195
      Left            =   3120
      TabIndex        =   10
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "m"
      Height          =   195
      Left            =   5160
      TabIndex        =   8
      Top             =   4200
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Dpartícula"
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Top             =   4200
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Kg"
      Height          =   195
      Left            =   2640
      TabIndex        =   5
      Top             =   3720
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Masa de entrada"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "m^3"
      Height          =   195
      Left            =   2640
      TabIndex        =   2
      Top             =   3240
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vol. Deseado"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   960
   End
End
Attribute VB_Name = "FormPmasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'remove comment below, if Shockwave flash player will be actived
'FlashVid.Visible = True
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
'FormPmasa.ScaleMode = 3
'FlashVid.Movie = App.Path & "\esfera.swf"
'FlashVid.Left = 700
'FlashVid.Top = 100
'FlashVid.Width = 300
'FlashVid.Height = 380
'FlashVid.Play

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
