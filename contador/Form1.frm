VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FlDbg9c.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Contador Digital"
   ClientHeight    =   5940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10290
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   615
      Left            =   9240
      TabIndex        =   105
      Top             =   5160
      Width           =   735
      _cx             =   1296
      _cy             =   1085
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.CommandButton bResetPuerto 
      Caption         =   "Reset Puerto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   104
      Top             =   5160
      Width           =   975
   End
   Begin VB.CheckBox chBeep 
      Caption         =   "beep"
      Height          =   375
      Left            =   5520
      TabIndex        =   103
      Top             =   5280
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.Timer TimerPort 
      Interval        =   30000
      Left            =   3840
      Top             =   5280
   End
   Begin VB.CommandButton Interfaz 
      Caption         =   "Ver Interfaz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   102
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   101
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton Restar 
      Caption         =   "-"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   100
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Sumar 
      Caption         =   "+"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   99
      Top             =   2520
      Width           =   255
   End
   Begin VB.CommandButton Restar 
      Caption         =   "-"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   98
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Sumar 
      Caption         =   "+"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   97
      Top             =   1560
      Width           =   255
   End
   Begin VB.CommandButton Restar 
      Caption         =   "-"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   96
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Sumar 
      Caption         =   "+"
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   95
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox direcRestar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   94
      Top             =   2520
      Width           =   255
   End
   Begin VB.CheckBox direcRestar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   93
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox direcRestar 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   92
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox direcContar 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   91
      Text            =   "0"
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox direcContar 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   90
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox direcContar 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   89
      Text            =   "0"
      Top             =   720
      Width           =   375
   End
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
      Left            =   7320
      TabIndex        =   88
      Text            =   "000000"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   7
      Left            =   3360
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
      Left            =   4560
      TabIndex        =   71
      Text            =   "8"
      Top             =   4035
      Width           =   255
   End
   Begin VB.TextBox digito 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2880
      TabIndex        =   70
      Text            =   "0"
      Top             =   4170
      Width           =   375
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   69
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   68
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   67
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   66
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   3960
      TabIndex        =   65
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   64
      Top             =   4200
      Width           =   135
   End
   Begin VB.CheckBox data 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   63
      Top             =   4200
      Width           =   135
   End
   Begin VB.TextBox velocidad2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      TabIndex        =   62
      Text            =   "100"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CheckBox tick2 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1800
      TabIndex        =   61
      Top             =   4200
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1800
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
      Left            =   2760
      TabIndex        =   37
      Text            =   "--:--:--"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox direc 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   36
      Text            =   "---"
      Top             =   2520
      Width           =   375
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
      Left            =   2760
      TabIndex        =   34
      Text            =   "--:--:--"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox direc 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   33
      Text            =   "---"
      Top             =   1560
      Width           =   375
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
      Left            =   2760
      TabIndex        =   31
      Text            =   "--:--:--"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox direc 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   30
      Text            =   "---"
      Top             =   720
      Width           =   375
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
      Left            =   360
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
   Begin VB.Line Line8 
      X1              =   120
      X2              =   10200
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line7 
      X1              =   5520
      X2              =   10200
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line6 
      X1              =   5520
      X2              =   10200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label23 
      Caption         =   "Acumulado anterior:"
      Height          =   255
      Left            =   5760
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
      Left            =   6600
      TabIndex        =   86
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label22 
      Caption         =   "Acumulado Total:"
      Height          =   255
      Left            =   6600
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
      Left            =   8520
      TabIndex        =   84
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "AFORO:"
      Height          =   255
      Left            =   8640
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
      Left            =   8160
      TabIndex        =   82
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Salidas Hoy:"
      Height          =   255
      Left            =   8160
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
      Left            =   5760
      TabIndex        =   80
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "Entradas Libres:"
      Height          =   255
      Left            =   5760
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
      Left            =   5760
      TabIndex        =   78
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "Entradas Grupos:"
      Height          =   255
      Left            =   5760
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
      Left            =   5760
      TabIndex        =   76
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Entradas Hoy:"
      Height          =   255
      Left            =   5760
      TabIndex        =   75
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Line Line5 
      X1              =   5400
      X2              =   5400
      Y1              =   0
      Y2              =   4800
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
      X2              =   5280
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
      X2              =   5280
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   5280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   5280
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
Dim estado, refreshPort, refreshPortLimit
Dim fRegistro, fAcumulado, fConfig
Const puertaSwap = True

Public Function FileExists(sFullPath As String) As Boolean
    Dim oFile As New Scripting.FileSystemObject
    FileExists = oFile.FileExists(sFullPath)
End Function

Private Sub bResetPuerto_Click()
    'pasar a modo SPP
    confPuerto port
End Sub

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
       
    estado = 0
    refreshPort = 0
    refreshPortLimit = 5    'Rrefrescar Modo del Puerto a SPP cada 5 minutos.
    
    'pasar a modo SPP
    confPuerto port
    
    'Prepara Id para Abrir archivos de registro y configurción
    fRegistro = 1
    fAcumulado = 2
    fConfig = 3
    
    If FileExists("config.txt") Then
        Open "config.txt" For Input Access Read As #fConfig
            'puertas negadas
            Input #fConfig, n0a, n0b, n0c, n1a, n1b, n1c, n2a, n2b, n2c
            negar(0) = n0a: negar(1) = n0b: negar(2) = n0c
            negar(3) = n1a: negar(4) = n1b: negar(5) = n1c
            negar(6) = n2a: negar(7) = n2b: negar(8) = n2c
            ' direcciones y retar puertas
            Input #fConfig, dir0, negar0, dir1, negar1, dir2, negar2
            direcContar(0) = dir0: direcRestar(0) = negar0
            direcContar(1) = dir1: direcRestar(1) = negar1
            direcContar(2) = dir2: direcRestar(2) = negar2
            ' tiempo limite paso, y refresco de relojes
            Input #fConfig, tlimite, refrescoIn, refrescoOut
            Tiempos = tlimite
            velocidad = refrescoIn
            Timer1.Interval = refrescoIn
            velocidad2 = refrescoOut
            Timer2.Interval = refrescoOut
        'fin config
        Close fConfig
    End If
       
    'usar el mismo tiempo.limite para todas las puertas
    v = Val(Tiempos.Text)
    limiteA = v
    limiteB = v
    limiteC = v
   
    'Leer acumulado anterior
    If FileExists("acumulado.txt") Then
        Open "acumulado.txt" For Input Access Read As #fAcumulado
        Input #fAcumulado, acum
        Close #fAcumulado
        tAcumulado.Text = acum
    End If
   
   'CARGAR INTERFAZ FLASH
   Me.Flash.Movie = App.Path + "\interfaz.swf"
   'Me.Flash.Visible = False
   Me.Flash.WMode = 0
   
    ' tamaño de la ventana... (para multiples del monitores)
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    ' posicion... respecto al (0,0) del monitor principal
    Me.Top = 0
    Me.Left = 0
    
    'que el panel flash... esté correcto
    RefrescaPanel
End Sub
Private Sub confPuerto(port)
    'en el puerto Base+H402 esta para configurar el modo
    'los 3 MSB XXXxxxxx: 000-SPP, 001-SPPBiDir, 011-ECP(normal)
    
    'pasar a modo SPP
    Out port + &H402, 0
    'poner control a cero (excepto c2, que es negado)
    Out port + 2, 4
    
    'señalar el hecho, con el speaker
    If (chBeep.value = 1) Then Beep
End Sub
Private Sub Form_Resize()
    With Flash
    .Top = 0
    .Left = 0
    .Width = Me.Width
    .Height = Me.Height
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close fRegistro
    'Close fAcumulado
End Sub

Private Sub Interfaz_Click()
    Me.Flash.Visible = True
End Sub

Private Sub Restar_Click(Index As Integer)
    EventoPasar Index, -1
End Sub

Private Sub SALIR_Click()
    PreguntaCerrar
End Sub

Private Sub Sumar_Click(Index As Integer)
    EventoPasar Index, 1
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
            
            If ((pasoA) And (tiempoA > limiteA)) Then 'paso
                ' original 0 (swap 1)
                If (puertaSwap) Then numPuerta = 1 Else numPuerta = 0
                If (direcContar(0) = 0) Then 'siempre cuenta
                    EventoPasar numPuerta, 1
                Else
                    If (dirA = direcContar(0)) Then
                        EventoPasar numPuerta, 1
                    Else
                        If (direcRestar(0).value = 1) Then EventoPasar numPuerta, -1
                    End If
                End If
            End If
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
            
            If ((pasoB) And (tiempoB > limiteB)) Then 'paso
                ' original 1 (swap 0)
                If (puertaSwap) Then numPuerta = 0 Else numPuerta = 1
                If (direcContar(1) = 0) Then 'siempre cuenta
                    EventoPasar numPuerta, 1
                Else
                    If (dirB = direcContar(1)) Then
                        EventoPasar numPuerta, 1
                    Else
                        If (direcRestar(1).value = 1) Then EventoPasar numPuerta, -1
                    End If
                End If
            End If
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
            
            If ((pasoC) And (tiempoC > limiteC)) Then
                If (direcContar(2) = 0) Then 'siempre cuenta
                    EventoPasar 2, 1
                Else
                    If (dirC = direcContar(2)) Then
                        EventoPasar 2, 1
                    Else
                        If (direcRestar(2).value = 1) Then EventoPasar 2, -1
                    End If
                End If
            End If
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
        'guardar acumulado
        Open "acumulado.txt" For Output Access Write As #fAcumulado
        Write #fAcumulado, Val(lAcumulado.Caption)
        Close #fAcumulado
        
    'Refrescar panel flash
    Flash.SetVariable "anterior", tAcumulado.Text
    Flash.SetVariable "hoy", lHoy.Caption
    Flash.SetVariable "grupos", lGrupos.Caption
    Flash.SetVariable "libres", lLibres.Caption
    Flash.SetVariable "salidas", lSalidas.Caption
    Flash.SetVariable "aforo", lAforo.Caption
    Flash.SetVariable "total", lAcumulado.Caption
    
End Sub

Private Sub TimerPort_Timer()
    'If (refreshPort > refreshPortLimit) Then
    '    refreshPort = 0
        confPuerto port
    'End If
    'refreshPort = refreshPort + 1
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

Private Sub EventoPasar(puerta, valor)
    v = Val(contador.Item(puerta).Text)
    r = v + valor
    If (r >= 0) Then 'no permitir negativos
        contador.Item(puerta).Text = r
        'registrar evento
        nn = "regs/registro_" & Format(Date, "yyyymmdd") & ".txt"
        Open nn For Append Access Write As #fRegistro
        Write #fRegistro, Now, puerta, valor
        Close #fRegistro
    End If
End Sub
Private Sub Flash_FSCommand(ByVal command As String, ByVal args As String)
    If command = "alert" Then MsgBox args
    If command = "ESC" Then PreguntaCerrar
    If command = "ocultar" Then Me.Flash.Visible = False
    
    If command = "grupos" Then
        v = Val(args)
        EventoPasar 0, v
    End If
    If command = "libres" Then
        v = Val(args)
        EventoPasar 1, v
    End If
    If command = "salidas" Then
        v = Val(args)
        EventoPasar 2, v
    End If
    
        'callFlashEscribirMensaje "Código incorrecto: " & args
        'Flash.SetVariable "_root.codigo", ""

End Sub
Private Sub PreguntaCerrar()
    Msg = "¿Cerrar el contador?"
    If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbYes Then End
End Sub
