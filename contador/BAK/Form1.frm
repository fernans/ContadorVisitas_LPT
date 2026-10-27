VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tAcumulado 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7080
      TabIndex        =   88
      Text            =   "000000"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   74
      Top             =   4200
      Width           =   135
   End
   Begin VB.TextBox BCD 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4320
      TabIndex        =   71
      Text            =   "8"
      Top             =   4035
      Width           =   255
   End
   Begin VB.TextBox digito 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      TabIndex        =   70
      Text            =   "0"
      Top             =   4170
      Width           =   375
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   69
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   68
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   67
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   66
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   65
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   3840
      TabIndex        =   64
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   63
      Top             =   4200
      Width           =   135
   End
   Begin VB.TextBox velocidad2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   62
      Text            =   "100"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CheckBox tick2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1560
      TabIndex        =   61
      Top             =   4200
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1560
      Top             =   3720
   End
   Begin VB.TextBox Tiempos 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   48
      Text            =   "2"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox velocidad 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   47
      Text            =   "100"
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox chpasoC 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2520
      TabIndex        =   46
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chpasoB 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2520
      TabIndex        =   45
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox chpasoA 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2520
      TabIndex        =   44
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox reloj 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2880
      TabIndex        =   43
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox activo 
      Alignment       =   1  'Right Justify
      Caption         =   "puertaC"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   42
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox activo 
      Alignment       =   1  'Right Justify
      Caption         =   "puertaB"
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   41
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox activo 
      Alignment       =   1  'Right Justify
      Caption         =   "puertaA"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   40
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox tick 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4080
      TabIndex        =   39
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox contador 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   38
      Text            =   "0"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox tiempo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   37
      Text            =   "--:--:--"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox direc 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   36
      Text            =   "---"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox contador 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   35
      Text            =   "0"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox tiempo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   34
      Text            =   "--:--:--"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox direc 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   33
      Text            =   "---"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox contador 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   32
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox tiempo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2880
      TabIndex        =   31
      Text            =   "--:--:--"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox direc 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   30
      Text            =   "---"
      Top             =   840
      Width           =   495
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   26
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   7
      Left            =   2160
      TabIndex        =   25
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   24
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   23
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   22
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   21
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   20
      Top             =   1200
      Width           =   255
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   19
      Top             =   960
      Width           =   255
   End
   Begin VB.CheckBox value 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   18
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   17
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   16
      Top             =   2760
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   15
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   14
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   13
      Top             =   1800
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   12
      Top             =   1560
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   11
      Top             =   1200
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   960
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox negar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox chInp 
      Caption         =   "c0"
      Height          =   255
      Index           =   8
      Left            =   840
      TabIndex        =   8
      Top             =   3000
      Width           =   495
   End
   Begin VB.CheckBox chInp 
      Caption         =   "c1"
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   495
   End
   Begin VB.CheckBox chInp 
      Caption         =   "c2"
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   6
      Top             =   2520
      Width           =   495
   End
   Begin VB.CheckBox chInp 
      Caption         =   "c3"
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.CheckBox chInp 
      Caption         =   "s3"
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.CheckBox chInp 
      Caption         =   "s4"
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   3
      Top             =   1560
      Width           =   495
   End
   Begin VB.CheckBox chInp 
      Caption         =   "s5"
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.CheckBox chInp 
      Caption         =   "s6"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   495
   End
   Begin VB.CheckBox chInp 
      Caption         =   "s7"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2400
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "SALIDA (BCD):"
      Height          =   855
      Left            =   120
      TabIndex        =   59
      Top             =   3720
      Width           =   4695
      Begin VB.TextBox suma 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Text            =   "-"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "suma A+B"
         Height          =   255
         Left            =   360
         TabIndex        =   73
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "data  7 6 5 4   3 2 1 0"
         Height          =   255
         Left            =   2520
         TabIndex        =   72
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Line Line7 
      X1              =   5280
      X2              =   9960
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line6 
      X1              =   5280
      X2              =   9960
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label23 
      Caption         =   "Acumulado anterior:"
      Height          =   255
      Left            =   5520
      TabIndex        =   87
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lAcumulado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6360
      TabIndex        =   86
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label22 
      Caption         =   "Acumulado Total:"
      Height          =   255
      Left            =   6360
      TabIndex        =   85
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lAforo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8280
      TabIndex        =   84
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "AFORO:"
      Height          =   255
      Left            =   8400
      TabIndex        =   83
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lSalidas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7920
      TabIndex        =   82
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Salidas Hoy:"
      Height          =   255
      Left            =   7920
      TabIndex        =   81
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lLibres 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      TabIndex        =   80
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "Entradas Libres:"
      Height          =   255
      Left            =   5520
      TabIndex        =   79
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lGrupos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      TabIndex        =   78
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "Entradas Grupos:"
      Height          =   255
      Left            =   5520
      TabIndex        =   77
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lHoy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5520
      TabIndex        =   76
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Entradas Hoy:"
      Height          =   255
      Left            =   5520
      TabIndex        =   75
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Line Line5 
      X1              =   5160
      X2              =   5160
      Y1              =   0
      Y2              =   4680
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "limite tiempo:"
      Height          =   255
      Left            =   1920
      TabIndex        =   58
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4920
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Label Label12 
      Caption         =   "pin   1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   57
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "pin 14"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "pin 16"
      Height          =   255
      Left            =   240
      TabIndex        =   55
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "pin 17"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   54
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "pin 15"
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "pin 13"
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "pin 12"
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "pin 10"
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "pin 11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   49
      Top             =   720
      Width           =   495
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   120
      X2              =   4920
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   4920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   4920
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Label Label3 
      Caption         =   "Valor"
      Height          =   255
      Left            =   2040
      TabIndex        =   29
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Negar"
      Height          =   255
      Left            =   1440
      TabIndex        =   28
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Inp"
      Height          =   255
      Left            =   840
      TabIndex        =   27
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Dim valores(8)
Dim port
Dim tiempoA, pasoA, limiteA, dirA
Dim tiempoB, pasoB, limiteB, dirB
Dim tiempoC, pasoC, limiteC, dirC
Dim estado


Private Sub contador_Change(Index As Integer)
    RefrescaPanel
End Sub

Sub Form_Load()
    port = &H378
    For I = 0 To 8
        valores(I) = 0
    Next I
    For I = 0 To 2
        tiempo.Item(I).Text = "00:00:00"
        direc.Item(I).Text = "none"
        contador.Item(I).Text = "0"
    Next I
    dirA = 0
    dirB = 0
    dirC = 0
    
    v = Val(Tiempos.Text)
    limiteA = v
    limiteB = v
    limiteC = v
    
    estado = 0
    
    Out port + 2, 0
End Sub

Private Sub tAcumulado_Change()
    RefrescaPanel
End Sub

Private Sub Tiempos_Change()
    v = Val(Tiempos.Text)
    limiteA = v
    limiteB = v
    limiteC = v
End Sub

Private Sub Timer1_Timer()
    tick.value = 1 - tick.value
    ' status port
    b = Inp(port + 1)
    For I = 3 To 7
        v = (b And (2 ^ I)) / (2 ^ I)
        chInp.Item(7 - I).value = v
        v = v Xor (negar.Item(7 - I).value)
        value.Item(7 - I).value = v
    Next I
    
    ' control port
    b = Inp(port + 2)
    For I = 0 To 3
        v = (b And (2 ^ I)) / (2 ^ I)
        chInp.Item(8 - I).value = v
        v = v Xor (negar.Item(8 - I).value)
        value.Item(8 - I).value = v
    Next I
    
    
    reloj.Text = Int(Timer)
    
    'CONTROL DE PUERTAS
    ' PUERTA A (value 0, 1 & 2)
    If (value.Item(2).value = 1) Then pasoA = True
    If (pasoA) Then chpasoA.value = 1 Else chpasoA.value = 0
    If (activo.Item(0).value = 0) Then
        If (value.Item(0).value > value.Item(1).value) Then dirA = 1
        If (value.Item(0).value < value.Item(1).value) Then dirA = -1
        direc.Item(0).Text = dirA
    End If
    If (value.Item(0).value = 1 And value.Item(1).value = 1) Then
        If (activo.Item(0).value = 0) Then
            tiempoA = Timer
            activo.Item(0).value = 1
            pasoA = False
        End If
    End If
    If (value.Item(0).value = 0 And value.Item(1).value = 0) Then
        If (activo.Item(0).value = 1) Then
            tiempoA = Timer - tiempoA
            tiempo.Item(0).Text = tiempoA
            
            If ((pasoA) And (tiempoA > limiteA)) Then contador.Item(0).Text = Val(contador.Item(0).Text) + 1
            activo.Item(0).value = 0
            dirA = 0
        End If
    End If
    
    ' PUERTA B (value 3, 4 & 5)
    If (value.Item(5).value = 1) Then pasoB = True
    If (pasoB) Then chpasoB.value = 1 Else chpasoB.value = 0
    If (activo.Item(1).value = 0) Then
        If (value.Item(3).value > value.Item(4).value) Then dirB = 1
        If (value.Item(3).value < value.Item(4).value) Then dirB = -1
        direc.Item(1).Text = dirB
    End If
    If (value.Item(3).value = 1 And value.Item(4).value = 1) Then
        If (activo.Item(1).value = 0) Then
            tiempoB = Timer
            activo.Item(1).value = 1
            pasoB = False
        End If
    End If
    If (value.Item(3).value = 0 And value.Item(4).value = 0) Then
        If (activo.Item(1).value = 1) Then
            tiempoB = Timer - tiempoB
            tiempo.Item(1).Text = tiempoB
            
            If ((pasoB) And (tiempoB > limiteB)) Then contador.Item(1).Text = Val(contador.Item(1).Text) + 1
            activo.Item(1).value = 0
            dirB = 0
        End If
    End If
    
    
    ' PUERTA C (value 6, 7 & 8)
    If (value.Item(8).value = 1) Then pasoC = True
    If (pasoC) Then chpasoC.value = 1 Else chpasoC.value = 0
    If (activo.Item(2).value = 0) Then
        If (value.Item(6).value > value.Item(7).value) Then dirC = 1
        If (value.Item(6).value < value.Item(7).value) Then dirC = -1
        direc.Item(2).Text = dirC
    End If
    If (value.Item(6).value = 1 And value.Item(7).value = 1) Then
        If (activo.Item(2).value = 0) Then
            tiempoC = Timer
            activo.Item(2).value = 1
            pasoC = False
        End If
    End If
    If (value.Item(6).value = 0 And value.Item(7).value = 0) Then
        If (activo.Item(2).value = 1) Then
            tiempoC = Timer - tiempoC
            tiempo.Item(2).Text = tiempoC
            
            If ((pasoC) And (tiempoC > limiteC)) Then contador.Item(2).Text = Val(contador.Item(2).Text) + 1
            activo.Item(2).value = 0
            dirC = 0
        End If
    End If
    
End Sub


Private Sub Timer2_Timer()
    tick2.value = 1 - tick2.value
     
    suma = Val(contador.Item(0).Text) + Val(contador.Item(1).Text) + Val(tAcumulado.Text)
    '6 digitos
    If (Len(suma.Text) < 6) Then
        suma.Text = String(6 - Len(suma.Text), "0") + suma.Text
    End If
    
    If (estado = 0) Then
        vd = Val(digito.Text)
        vd = vd + 1
        If (vd >= Len(suma.Text)) Then vd = 0
        digito.Text = vd
        dig = Mid(suma.Text, (Len(suma.Text) - vd), 1)
        BCD = dig
    End If
    If (estado = 1) Then
        dig = Val(BCD.Text)
        vd = &HFF
    End If
    
    
    data.Item(4).value = (vd And 1)
    data.Item(5).value = ((vd And 2) / (2))
    data.Item(6).value = ((vd And 4) / (4))
    data.Item(7).value = ((vd And 8) / (8))
    
    data.Item(0).value = (dig And 1)
    data.Item(1).value = ((dig And 2) / (2))
    data.Item(2).value = ((dig And 4) / (4))
    data.Item(3).value = ((dig And 8) / (8))
    
    'salida al port
    dsalida = (vd * 16) + dig
    Out port, dsalida
    
    estado = estado + 1
    If (estado = 2) Then estado = 0
End Sub

Private Sub RefrescaPanel()
    lHoy.Caption = Val(contador(0).Text) + Val(contador(1).Text)
    lGrupos.Caption = Val(contador(0).Text)
    lLibres.Caption = Val(contador(1).Text)
    lSalidas.Caption = Val(contador(2).Text)
    lAforo.Caption = Val(lHoy.Caption) - Val(lSalidas.Caption)
    lAcumulado.Caption = Val(tAcumulado.Text) + Val(lHoy.Caption)
End Sub

Private Sub velocidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then velocidad_Validate (True)
End Sub

Private Sub velocidad_Validate(Cancel As Boolean)
    Timer1.Interval = Val(velocidad.Text)
End Sub

Private Sub velocidad2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then velocidad2_Validate (True)
End Sub

Private Sub velocidad2_Validate(Cancel As Boolean)
    Timer2.Interval = Val(velocidad2.Text)
End Sub

Private Sub velocidad3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then velocidad3_Validate (True)
End Sub

Private Sub velocidad3_Validate(Cancel As Boolean)
    Timer3.Interval = Val(velocidad3.Text)
End Sub
