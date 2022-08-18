VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAParejas 
   Caption         =   "mml_FRASE0377"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   14115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActEdad 
      Caption         =   "Actualizar Grupos Edad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   78
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CheckBox chkParejasAEBDC 
      Caption         =   "mml_FRASE1233"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8040
      TabIndex        =   77
      Top             =   8400
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.TextBox tbCodAEBDC 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4500
      MaxLength       =   6
      TabIndex        =   0
      Top             =   0
      Width           =   1515
   End
   Begin VB.CommandButton cmdCorregirNombres 
      Caption         =   "mml_FRASE1183"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6300
      TabIndex        =   75
      Top             =   8340
      Width           =   1665
   End
   Begin VB.CommandButton cmdCambiarCateg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13470
      Picture         =   "frmAParejas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "mml_FRASE1182"
      Top             =   3600
      Width           =   465
   End
   Begin VB.CheckBox chkActLista 
      Caption         =   "mml_FRASE1006"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9300
      TabIndex        =   72
      Top             =   4980
      Width           =   2235
   End
   Begin VB.CommandButton cmdSelFNH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Picture         =   "frmAParejas.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   1620
      Width           =   465
   End
   Begin VB.CommandButton cmdSelFNM 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8100
      Picture         =   "frmAParejas.frx":0834
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   1620
      Width           =   465
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "mml_FRASE0029"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   12120
      TabIndex        =   69
      Top             =   8280
      Width           =   1875
   End
   Begin VB.CommandButton CommandButton3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1470
      Picture         =   "frmAParejas.frx":0B76
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4080
      Width           =   465
   End
   Begin VB.CommandButton cmdSelCateg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7830
      Picture         =   "frmAParejas.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   0
      Width           =   465
   End
   Begin VB.ComboBox tbProv 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9255
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   3105
      Width           =   4740
   End
   Begin VB.CheckBox chkParAdicional 
      Caption         =   "mml_FRASE0378"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   6900
      TabIndex        =   18
      Top             =   3600
      Width           =   1515
   End
   Begin VB.CheckBox chkPagado 
      Caption         =   "mml_FRASE0379"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   4680
      TabIndex        =   24
      Top             =   4920
      Width           =   1740
   End
   Begin VB.ComboBox tbGrupoEdad 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmAParejas.frx":144A
      Left            =   11340
      List            =   "frmAParejas.frx":144C
      TabIndex        =   22
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox tbSMSTelef 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1230
      TabIndex        =   21
      Top             =   4980
      Width           =   2505
   End
   Begin VB.CheckBox chkSMS 
      Caption         =   "mml_FRASE0380"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3810
      TabIndex        =   23
      Top             =   4890
      Width           =   855
   End
   Begin VB.CheckBox chkEmailMovil 
      Caption         =   "mml_FRASE0381"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7410
      TabIndex        =   26
      Top             =   4920
      Width           =   1725
   End
   Begin VB.CheckBox chkEmail 
      Caption         =   "mml_FRASE0382"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6420
      TabIndex        =   25
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox tbEmailMovil 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   20
      Top             =   4560
      Width           =   3975
   End
   Begin VB.TextBox tbEmail 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   4560
      Width           =   4695
   End
   Begin VB.ComboBox cbCombinar 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmAParejas.frx":144E
      Left            =   11220
      List            =   "frmAParejas.frx":1458
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2640
      Width           =   2775
   End
   Begin VB.ComboBox cbOrden 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      ItemData        =   "frmAParejas.frx":1473
      Left            =   12660
      List            =   "frmAParejas.frx":147D
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton cmdBorrarCateg 
      Caption         =   "mml_FRASE0383"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   90
      TabIndex        =   31
      Top             =   8310
      Width           =   1815
   End
   Begin VB.TextBox tbNumParejas 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13260
      TabIndex        =   61
      Top             =   4500
      Width           =   615
   End
   Begin VB.TextBox tbCodGrupoEdad 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10860
      TabIndex        =   5
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox tbDescMod 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   59
      Top             =   4080
      Width           =   5355
   End
   Begin VB.TextBox tbCodMod 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox tbEscuelas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2640
      Width           =   7695
   End
   Begin VB.TextBox tbDescComp 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   51
      Top             =   0
      Width           =   4935
   End
   Begin VB.TextBox tbCodComp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   29
      Top             =   0
      Width           =   855
   End
   Begin VB.ComboBox tbCat 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmAParejas.frx":1491
      Left            =   11520
      List            =   "frmAParejas.frx":150A
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "mml_FRASE0384"
      Height          =   2175
      Left            =   6720
      TabIndex        =   28
      Top             =   420
      Width           =   7275
      Begin VB.ComboBox tbMNombre 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1305
         TabIndex        =   3
         Tag             =   "Mujer"
         Top             =   675
         Width           =   5805
      End
      Begin VB.TextBox tbMEdad 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   55
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox tbMNumSoc 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox tbMFecNac 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox tbMNif 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "mml_FRASE0385"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   56
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "mml_FRASE0386"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "mml_FRASE0387"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "mml_FRASE0388"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "mml_FRASE0266"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.TextBox tbDir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   3120
      Width           =   6435
   End
   Begin VB.TextBox tbTlf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1470
      TabIndex        =   10
      Top             =   3600
      Width           =   5355
   End
   Begin VB.TextBox tbCodPareja 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   39
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "mml_FRASE0250"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9840
      TabIndex        =   35
      Top             =   8310
      Width           =   2145
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "mml_FRASE0251"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   4830
      TabIndex        =   34
      Top             =   8310
      Width           =   1395
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "mml_FRASE0252"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   3390
      TabIndex        =   33
      Top             =   8310
      Width           =   1395
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "mml_FRASE0365"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   1950
      TabIndex        =   32
      Top             =   8310
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Caption         =   "mml_FRASE0389"
      Height          =   2175
      Left            =   0
      TabIndex        =   37
      Top             =   420
      Width           =   6675
      Begin VB.ComboBox tbHNombre 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Tag             =   "Hombre"
         Top             =   660
         Width           =   5205
      End
      Begin VB.TextBox tbHEdad 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   53
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox tbHNumSoc 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox tbHFecNac 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox tbHNif 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "mml_FRASE0385"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   54
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "mml_FRASE0386"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "mml_FRASE0387"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "mml_FRASE0388"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "mml_FRASE0266"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0034"
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   5370
      Width           =   14055
      Begin MSDataGridLib.DataGrid dgParejas 
         Bindings        =   "frmAParejas.frx":15FD
         Height          =   2595
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   4577
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc adoParejas 
      Height          =   495
      Left            =   240
      Top             =   7680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Escrutinio"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "mml_FRASE0033"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Parejas"
      Caption         =   "mml_FRASE0034"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label26 
      Caption         =   "mml_FRASE1228"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   76
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label25 
      Caption         =   "mml_FRASE0190"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   11340
      TabIndex        =   73
      Top             =   4980
      Width           =   1245
   End
   Begin VB.Label Label24 
      Caption         =   "mml_FRASE0390"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   67
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "mml_FRASE0391"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   66
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label22 
      Caption         =   "mml_FRASE0392"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   65
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label21 
      Caption         =   "GrupoEdad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   64
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "mml_FRASE0393"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8010
      TabIndex        =   63
      Top             =   3120
      Width           =   1245
   End
   Begin VB.Label Label19 
      Caption         =   "mml_FRASE0394"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11460
      TabIndex        =   62
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "mml_FRASE0187"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   60
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "mml_FRASE0395"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   58
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "mml_FRASE0396"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8700
      TabIndex        =   57
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label13 
      Caption         =   "mml_FRASE0215"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   52
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "mml_FRASE0301"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10170
      TabIndex        =   50
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "mml_FRASE0274"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "mml_FRASE0397"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "mml_FRASE0261"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmAParejas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NIF = 1
Const C_COD_AEBDC = 2

Private Sub chkPagado_Click()
    On Local Error GoTo error
    If chkPagado.Value = 1 Then
        chkPagado.Caption = mml_FRASE0089
        chkPagado.ForeColor = &HFF00&
    Else
        chkPagado.Caption = mml_FRASE0379
        chkPagado.ForeColor = &HFF
    End If
    Exit Sub
error:
    ProcesarError "chkPagado_Click"
End Sub

Private Sub cmdAct_Click()
Dim i As Integer
    On Local Error GoTo error

    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    Sleep 500
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    adoParejas.Refresh
    dgParejas.Refresh

    adoParejas.ConnectionString = "DSN=Escrutinio"
    adoParejas.RecordSource = "SELECT * FROM parejas WHERE nombre_hombre LIKE '%" & tbHNombre.Text & "%' AND nombre_mujer LIKE '%" & tbMNombre.Text & "%' AND cod_competicion LIKE '" & tbCodComp.Text & "%' AND (categoria Like '" & tbCat.Text & "%' or categoria IS NULL) And (cod_grupoedad Like '" & tbCodGrupoEdad.Text & "%' OR cod_grupoedad IS NULL) AND (cod_modalidad Like '" & tbCodMod.Text & "%' OR cod_modalidad IS NULL) ORDER BY " & IIf(cbOrden.ListIndex = 1, "nombre_hombre", "codigo")
    adoParejas.Refresh
    
    Debug.Print adoParejas.RecordSource
    
    If adoParejas.Recordset.EOF Then
        dgParejas.Enabled = False
        tbNumParejas.Text = "0"
    Else
        adoParejas.Recordset.MoveLast
        adoParejas.Recordset.MoveFirst
        tbNumParejas.Text = adoParejas.Recordset.RecordCount
        dgParejas.Enabled = True
    End If
    Exit Sub
error:
    ProcesarError "cmdAct_Click"
    
End Sub

Private Sub cmdActEdad_Click()
Dim rs As Recordset
Dim hedad As Integer, medad As Integer
Dim sGrupoEdad As String
        hedad = 0
        medad = 0
        Set rs = db.OpenRecordset("SELECT * FROM parejas as p, competiciones as c WHERE p.cod_competicion = c.codigo and p.cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
        While Not rs.EOF
                If IsDate(rs.Fields("fecha_nac_hombre")) Then
                    hedad = CalcularEdad(CDate(rs.Fields("fecha_nac_hombre")), rs.Fields("cod_competicion"))
                ElseIf IsDate(rs.Fields("fecha_nac_mujer")) Then
                    medad = CalcularEdad(CDate(rs.Fields("fecha_nac_mujer")), rs.Fields("cod_competicion"))
                End If
                sGrupoEdad = CalcularGrupoEdad(hedad, medad, rs.Fields("cod_modalidad"))
                If sGrupoEdad <> "" Then
                    sSQL = "UPDATE parejas SET grupoedad = '" & sGrupoEdad & "', cod_grupoedad = " & BuscarCodGrupoEdad(sGrupoEdad) & " WHERE codigo = " & rs.Fields("p.codigo")
                    db.Execute sSQL
                End If
            rs.MoveNext
        Wend
        rs.Close
        MsgBox "Grupos de edad actualizados", vbOKOnly Or vbInformation
        
End Sub

Private Sub cmdBorrar_Click()
    On Local Error GoTo error
    
    If tbCodPareja.Text = "" Then
        MsgBox mml_FRASE0398, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    db.Execute ("DELETE FROM parejas WHERE codigo = " & tbCodPareja.Text)
    Call cmdNuevo_Click
    Call cmdAct_Click
    Exit Sub
error:
    ProcesarError "cmdBorrar_Click"
End Sub

Private Sub cmdBorrarCateg_Click()
    On Local Error GoTo error
    
    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If UCase(InputBox(mml_FRASE0399, vbYesNo Or vbQuestion, "")) <> "SI" Then
        Exit Sub
    End If
    db.Execute ("DELETE FROM dorsales WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")")
    db.Execute ("DELETE FROM parejas WHERE cod_competicion = " & tbCodComp.Text)
    Exit Sub
error:
    ProcesarError "cmdBorrarCateg_Click"
End Sub

Private Sub cmdCambiarCateg_Click()
Dim sDesc As String

    If Not C_DEBUG Then On Local Error GoTo error
    sDesc = InputBox(mml_FRASE1182, G_MSG_PREGUNTA, "")
    
    If sDesc <> "" Then
        db.Execute "UPDATE parejas SET categoria = '" & sDesc & "' WHERE cod_competicion = " & tbCodComp.Text & " AND categoria = '" & tbCat.Text & "'"
        
        Call cmdNuevo_Click
        Call cmdAct_Click
        MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    End If
    Exit Sub
error:
    ProcesarError "cmdCambiarCateg_Click"
End Sub

Private Sub cmdCorregirNombres_Click()
Dim rs As Recordset
    
    If Val(tbCodComp.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE1184, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("SELECT codigo, nombre_hombre, nombre_mujer FROM parejas WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    While Not rs.EOF
        db.Execute "UPDATE parejas SET nombre_hombre = '" & CorregirNombre(SinNulos(rs.Fields("nombre_hombre"))) & "', nombre_mujer = '" & CorregirNombre(SinNulos(rs.Fields("nombre_mujer"))) & "' WHERE codigo = " & rs.Fields("codigo")
        rs.MoveNext
    Wend
    rs.Close
    
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
End Sub

Private Sub cmdGrabar_Click()
Dim rs As Recordset
Dim sCad As String
Dim bError As Boolean
Dim rsSocioAnulado As Recordset

    On Local Error GoTo error
    
    If Trim$(tbHNombre.Text) = "" Then tbHNombre.Text = "-"
    If Trim$(tbMNombre.Text) = "" Then tbMNombre.Text = "-"

    If tbCodComp.Text = "" Or tbCodGrupoEdad.Text = "" Or tbCodMod.Text = "" Or tbHNombre.Text = "" Or _
       tbMNombre.Text = "" Or tbCat.Text = "" Then
        MsgBox mml_FRASE0264, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    If tbHNif.Text = "" Or tbHFecNac.Text = "" Or Not IsDate(tbHFecNac.Text) Or tbMNif.Text = "" Or _
       tbMFecNac.Text = "" Or Not IsDate(tbMFecNac.Text) Then
        If MsgBox(mml_FRASE0400, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            Exit Sub
        End If
    End If
    
    If Val(tbHEdad.Text) = 0 Then tbHEdad.Text = "0"
    If Val(tbMEdad.Text) = 0 Then tbMEdad.Text = "0"
    
    'Validaciones de grupos de edad y categorias válidos
    'If tbGrupoEdad.Text = mml_FRASE0135 And (tbCat.Text = "A" Or tbCat.Text = "B") Then
    '    MsgBox mml_FRASE0401, vbOKOnly Or vbInformation, mml_FRASE0096
    '    Exit Sub
    'End If
    'If tbGrupoEdad.Text = mml_FRASE0136 And (tbCat.Text = "A") Then
    '    MsgBox mml_FRASE0402, vbOKOnly Or vbInformation, mml_FRASE0096
    '    Exit Sub
    'End If
    'If tbGrupoEdad.Text = mml_FRASE0139 And (tbCat.Text = "A") Then
    '    MsgBox mml_FRASE0403, vbOKOnly Or vbInformation, mml_FRASE0096
    '    Exit Sub
    'End If
    
    'Comprobamos que si el socio está inhabilitado para competir
    Set rsSocioAnulado = db.OpenRecordset("SELECT * FROM sociosanulados WHERE num_socio = '" & tbHNumSoc.Text & "'", dbOpenSnapshot)
    If Not rsSocioAnulado.EOF Then
        If MsgBox(mml_FRASE0404 & tbHNumSoc.Text & ", " & rsSocioAnulado!Nombre & " " & rsSocioAnulado!apellidos & mml_FRASE0405, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            Exit Sub
        End If
    End If
    rsSocioAnulado.Close
    
    
    ' Solo se comprueba si es un alta
    If tbCodPareja.Text = "" Then
        bError = False
        ' Comprobación del número de socio
        sCad = ""
        Set rs = db.OpenRecordset("SELECT * from parejas WHERE num_socio_hombre <> '' AND num_socio_hombre = '" & tbHNumSoc.Text & "' AND cod_modalidad = " & tbCodMod.Text & " AND cod_competicion = " & tbCodComp.Text & " AND LEN(categoria) = " & Len(tbCat.Text), dbOpenSnapshot)
        While Not rs.EOF
            bError = True
            If sCad <> "" Then sCad = sCad & ","
            sCad = sCad & Str$(rs!codigo) & "-" & rs!categoria
            rs.MoveNext
        Wend
        rs.Close
        If sCad <> "" Then
            If MsgBox(mml_FRASE0406 & sCad & mml_FRASE0407, vbYesNo Or vbQuestion, mml_FRASE0086) = vbNo Then
                Exit Sub
            End If
        End If
        sCad = ""
        Set rs = db.OpenRecordset("SELECT * from parejas WHERE num_socio_mujer <> '' AND num_socio_mujer = '" & tbHNumSoc.Text & "' AND cod_modalidad = " & tbCodMod.Text & " AND cod_competicion = " & tbCodComp.Text & " AND LEN(categoria) = " & Len(tbCat.Text), dbOpenSnapshot)
        While Not rs.EOF
            bError = True
            If sCad <> "" Then sCad = sCad & ","
            sCad = sCad & Str$(rs!codigo) & "-" & rs!categoria
            rs.MoveNext
        Wend
        rs.Close
        If sCad <> "" Then
            If MsgBox(mml_FRASE0408 & sCad & mml_FRASE0407, vbYesNo Or vbQuestion, mml_FRASE0086) = vbNo Then
                Exit Sub
            End If
        End If
        
        'Comprobación de NIFs
        If tbHNif.Text <> "" Then
            sCad = ""
            Set rs = db.OpenRecordset("SELECT * from parejas WHERE nif_hombre = '" & tbHNif.Text & "' AND cod_modalidad = " & tbCodMod.Text & " AND cod_competicion = " & tbCodComp.Text & " AND LEN(categoria) = " & Len(tbCat.Text), dbOpenSnapshot)
            While Not rs.EOF
                bError = True
                If sCad <> "" Then sCad = sCad & ","
                sCad = sCad & Str$(rs!codigo) & "-" & rs!categoria
                rs.MoveNext
            Wend
            rs.Close
            If sCad <> "" Then
                If MsgBox(mml_FRASE0409 & sCad & mml_FRASE0407, vbYesNo Or vbQuestion, mml_FRASE0086) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        If tbMNif.Text <> "" Then
            sCad = ""
            Set rs = db.OpenRecordset("SELECT * from parejas WHERE nif_mujer = '" & tbMNif.Text & "' AND cod_modalidad = " & tbCodMod.Text & " AND cod_competicion = " & tbCodComp.Text & " AND LEN(categoria) = " & Len(tbCat.Text), dbOpenSnapshot)
            While Not rs.EOF
                bError = True
                If sCad <> "" Then sCad = sCad & ","
                sCad = sCad & Str$(rs!codigo) & "-" & rs!categoria
                rs.MoveNext
            Wend
            rs.Close
            If sCad <> "" Then
                If MsgBox(mml_FRASE0410 & sCad & mml_FRASE0407, vbYesNo Or vbQuestion, mml_FRASE0086) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        sCad = ""
        Set rs = db.OpenRecordset("SELECT * from parejas WHERE nombre_hombre = '" & tbHNombre.Text & "' AND cod_modalidad = " & tbCodMod.Text & " AND cod_competicion = " & tbCodComp.Text & " AND LEN(categoria) = " & Len(tbCat.Text), dbOpenSnapshot)
        While Not rs.EOF
            bError = True
            If sCad <> "" Then sCad = sCad & ","
            sCad = sCad & Str$(rs!codigo) & "-" & rs!categoria
            rs.MoveNext
        Wend
        rs.Close
        ' Solo se comprueba si hemos introducido un nombre válido
        If Len(tbHNombre.Text) > 5 Then
            If sCad <> "" Then
                If MsgBox(mml_FRASE0411 & sCad & mml_FRASE0407, vbYesNo Or vbQuestion, mml_FRASE0086) = vbNo Then
                    Exit Sub
                End If
            End If
        Else
            bError = False
        End If
        sCad = ""
        Set rs = db.OpenRecordset("SELECT * from parejas WHERE nombre_mujer = '" & tbMNombre.Text & "' AND cod_modalidad = " & tbCodMod.Text & " AND cod_competicion = " & tbCodComp.Text & " AND LEN(categoria) = " & Len(tbCat.Text), dbOpenSnapshot)
        While Not rs.EOF
            bError = True
            If sCad <> "" Then sCad = sCad & ","
            sCad = sCad & Str$(rs!codigo) & "-" & rs!categoria
            rs.MoveNext
        Wend
        rs.Close
        ' Solo se comprueba si hemos introducido un nombre válido
        If Len(tbMNombre.Text) > 5 Then
            If sCad <> "" Then
                If MsgBox(mml_FRASE0412 & sCad & mml_FRASE0407, vbYesNo Or vbQuestion, mml_FRASE0086) = vbNo Then
                    Exit Sub
                End If
            End If
        Else
            bError = False
        End If
        If bError Then
            If MsgBox(mml_FRASE0413, vbYesNo Or vbQuestion, mml_FRASE0086) = vbNo Then
                Exit Sub
            End If
        End If
        
        Dim lCodPareja As Long
        
        lCodPareja = MaxCod("parejas")
        g_lCodUltimaPareja = lCodPareja
        sExecSQL = "INSERT INTO parejas VALUES(" & lCodPareja & ", '" & tbHNif.Text & "','" & tbHNombre.Text & "','" & tbHFecNac.Text & "','" & tbHNumSoc.Text & "'," & tbHEdad.Text & ",'" & tbMNif.Text & "','" & tbMNombre & "','" & tbMFecNac.Text & "','" & tbMNumSoc.Text & "'," & tbMEdad.Text & ",'" & tbDir.Text & "','" & tbTlf.Text & "'," & tbCodComp.Text & ",'" & tbGrupoEdad.Text & "','" & tbEscuelas.Text & "'," & tbCodGrupoEdad.Text & "," & tbCodMod.Text & ",'" & tbCat.Text & "','" & tbProv.Text & "'," & cbCombinar.ListIndex & ",'" & tbEmail.Text & "','" & tbEmailMovil.Text & "'," & chkEmail.Value & "," & chkEmailMovil.Value & "," & chkSMS.Value & ",'" & tbSMSTelef.Text & "'," & chkPagado.Value & "," & chkParAdicional.Value & ",'" & tbCodAEBDC.Text & "')"
        db.Execute (sExecSQL)
        Debug.Print sExecSQL
    Else
        db.Execute ("UPDATE parejas SET " & _
                    "nif_hombre = '" & tbHNif.Text & _
                    "',nombre_hombre ='" & tbHNombre.Text & _
                    "',fecha_nac_hombre='" & tbHFecNac.Text & _
                    "',num_socio_hombre='" & tbHNumSoc.Text & _
                    "',edad_hombre='" & tbHEdad.Text & _
                    "',nif_mujer = '" & tbMNif.Text & _
                    "',nombre_mujer ='" & tbMNombre.Text & _
                    "',fecha_nac_mujer='" & tbMFecNac.Text & _
                    "',num_socio_mujer='" & tbMNumSoc.Text & _
                    "',edad_mujer='" & tbMEdad.Text & _
                    "',direcciones='" & tbDir.Text & _
                    "',telefonos='" & tbTlf.Text & _
                    "',grupoedad='" & tbGrupoEdad & _
                    "',cod_competicion='" & tbCodComp.Text & _
                    "',escuelas='" & tbEscuelas.Text & _
                    "',cod_grupoedad='" & tbCodGrupoEdad.Text & _
                    "',cod_modalidad='" & tbCodMod.Text & _
                    "',categoria='" & tbCat.Text & _
                    "',provincia='" & tbProv.Text & _
                    "',combinar_edad=" & cbCombinar.ListIndex & _
                    ",pagado=" & chkPagado.Value & ", pareja_adicional = " & chkParAdicional.Value & _
                    ",email='" & tbEmail.Text & "',emailmovil='" & tbEmail.Text & _
                    "',email_selec=" & chkEmail.Value & ",emailmovil_selec=" & chkEmailMovil.Value & ",sms_selec=" & chkSMS.Value & ",sms_telef='" & tbSMSTelef.Text & "',aebdc_codigo = '" & tbCodAEBDC.Text & _
                    "' WHERE codigo = " & tbCodPareja.Text)
        g_lCodUltimaPareja = Val(tbCodPareja.Text)
    End If
    If chkActLista.Value = 1 Then
        Call cmdAct_Click
    End If
    SeleccionarCampo tbCodMod
    tbCodMod.SetFocus
    Exit Sub
error:
    ProcesarError "cmdGrabar_Click"
End Sub

Private Sub cmdNuevo_Click()
    On Local Error GoTo error
    
    tbCodPareja.Text = ""
    tbHNif.Text = ""
    tbHNombre.Text = ""
    tbHFecNac.Text = ""
    tbHNumSoc.Text = ""
    tbMNif.Text = ""
    tbMNombre.Text = ""
    tbMFecNac.Text = ""
    tbMNumSoc.Text = ""
    tbDir.Text = ""
    tbTlf.Text = ""
    tbProv.Text = ""
    tbCat.Text = ""
    tbCodMod.Text = ""
    tbDescMod.Text = ""
    tbCodGrupoEdad.Text = ""
    tbGrupoEdad.Text = ""
    cbCombinar.ListIndex = 0
    tbEmail.Text = ""
    tbEmailMovil.Text = ""
    tbSMSTelef.Text = ""
    chkSMS.Value = 0
    chkEmail.Value = 0
    chkEmailMovil.Value = 0
    tbEscuelas.Text = ""
    tbHEdad.Text = ""
    tbMEdad.Text = ""
    chkPagado.Value = 0
    chkParAdicional.Value = 0
    tbCodAEBDC.Text = ""
    
    tbHNombre.Clear
    tbMNombre.Clear
    
    tbCodAEBDC.SetFocus
    Exit Sub
error:
    ProcesarError "cmdNuevo_Click"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub




Private Sub cmdSelCateg_Click()
    On Local Error GoTo error
    
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    Exit Sub
error:
    ProcesarError "cmdSelCateg_Click"

End Sub

Private Sub cmdSelFNH_Click()
    On Local Error GoTo error
    
    tbHFecNac.Text = frmCalendario.Mostrar
    Call tbHFecNac_LostFocus
    Exit Sub
error:
    ProcesarError "cmdSelFNH_Click"
End Sub

Private Sub cmdSelFNM_Click()
    On Local Error GoTo error
    
    tbMFecNac.Text = frmCalendario.Mostrar
    Call tbMFecNac_LostFocus
    Exit Sub
error:
    ProcesarError "cmdSelFNM_Click"

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CommandButton3_Click()
    On Local Error GoTo error
    
    tbCodMod.Text = sSeleccionar("SELECT * FROM modalidad")
    tbDescMod.Text = sResultado(2)

    Exit Sub
error:
    ProcesarError "CommandButton3_Click"
End Sub

Private Sub dgParejas_Click()
    On Local Error Resume Next
    tbDescMod.Text = ""
    dgParejas.Col = 0
    tbCodPareja.Text = dgParejas.Text
    dgParejas.Col = 1
    tbHNif.Text = dgParejas.Text
    dgParejas.Col = 2
    tbHNombre.Text = dgParejas.Text
    dgParejas.Col = 3
    tbHFecNac.Text = dgParejas.Text
    dgParejas.Col = 4
    tbHNumSoc.Text = dgParejas.Text
    dgParejas.Col = 5
    tbHEdad.Text = dgParejas.Text
    dgParejas.Col = 6
    tbMNif.Text = dgParejas.Text
    dgParejas.Col = 7
    tbMNombre.Text = dgParejas.Text
    dgParejas.Col = 8
    tbMFecNac.Text = dgParejas.Text
    dgParejas.Col = 9
    tbMNumSoc.Text = dgParejas.Text
    dgParejas.Col = 10
    tbMEdad.Text = dgParejas.Text
    dgParejas.Col = 11
    tbDir.Text = dgParejas.Text
    dgParejas.Col = 12
    tbTlf.Text = dgParejas.Text
    dgParejas.Col = 13
    tbCodComp.Text = dgParejas.Text
    tbDescComp.Text = Buscar("competiciones", "descripcion", dgParejas.Text)
    dgParejas.Col = 14
    tbGrupoEdad.Text = dgParejas.Text
    dgParejas.Col = 15
    tbEscuelas.Text = dgParejas.Text
    dgParejas.Col = 17
    tbCodMod.Text = dgParejas.Text
    tbDescMod.Text = Buscar("Modalidad", "nombre", dgParejas.Text)
    dgParejas.Col = 18
    tbCat.Text = dgParejas.Text
    dgParejas.Col = 19
    tbProv.Text = dgParejas.Text
    dgParejas.Col = 20
    cbCombinar.ListIndex = Val(dgParejas.Text)
    dgParejas.Col = 21
    tbEmail.Text = dgParejas.Text
    dgParejas.Col = 22
    tbEmailMovil.Text = dgParejas.Text
    dgParejas.Col = 23
    chkEmail.Value = Val(dgParejas.Text)
    dgParejas.Col = 24
    chkEmailMovil.Value = Val(dgParejas.Text)
    dgParejas.Col = 25
    chkSMS.Value = Val(dgParejas.Text)
    dgParejas.Col = 26
    tbSMSTelef = dgParejas.Text
    dgParejas.Col = 27
    If dgParejas.Text = "" Then
        chkPagado.Value = 0
    Else
        chkPagado.Value = dgParejas.Text
    End If
    dgParejas.Col = 28
    If dgParejas.Text = "" Then
        chkParAdicional.Value = 0
    Else
        chkParAdicional.Value = dgParejas.Text
    End If
    dgParejas.Col = 29
    tbCodAEBDC.Text = dgParejas.Text
    
    
    'tbHFecNac_LostFocus
    'tbMFecNac_LostFocus
End Sub


Private Sub dgParejas_KeyPress(KeyAscii As Integer)
    dgParejas_Click
End Sub

Private Sub Form_Activate()
    If C_COUNTRY Then tbHNumSoc.SetFocus
End Sub

Private Sub Form_Load()
Dim rs As Recordset
Dim sError As String
    
    sError = ""
    
    TraducirCadenas Me
    
    If C_DEBUG Then On Local Error GoTo error
    
    tbCodComp.Text = CodCompActiva
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    sError = "Cargar gruposedad"
    tbGrupoEdad.Clear
    Set rs = db.OpenRecordset("SELECT * FROM gruposedad")
    While Not rs.EOF
        If Left(rs!Nombre, 1) <> "*" Then
            tbGrupoEdad.AddItem rs!Nombre
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    sError = "Cargar comunidad"
    tbProv.Clear
    Set rs = db.OpenRecordset("SELECT * FROM Comunidades ORDER BY 1")
    While Not rs.EOF
        tbProv.AddItem rs!comunidad
        rs.MoveNext
    Wend
    rs.Close
    cbCombinar.ListIndex = 0

    sError = "Actualizando categorías"
    'Actualizamos las categorías
    ActualizarCategorias tbCat, Val(tbCodComp.Text)
    
    If C_COUNTRY Then
        chkParejasAEBDC.Value = 0
    Else
        chkParejasAEBDC.Value = G_BLOQUEAR_PAREJAS_AEBDC
    End If
        
    cmdAct_Click
    Exit Sub
error:
    ProcesarError "Load_Click: " & sError
End Sub

Private Sub tbCat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr$(KeyAscii)))
End Sub



Private Sub tbCat_LostFocus()
    Dim i As Integer
    
    If tbCat.Text <> "" Then
        For i = 0 To tbCat.ListCount - 1
            If tbCat.List(i) = tbCat.Text Then
                Exit Sub
            End If
        Next
        MsgBox mml_FRASE1254, vbOKOnly Or vbCritical, "Warn"
        tbCat.SetFocus
    End If
End Sub

Private Sub tbCodAEBDC_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbCodAEBDC_LostFocus()
    If Val(tbCodAEBDC.Text) > 0 Then BuscarPareja C_COD_AEBDC
End Sub

Private Sub tbCodGrupoEdad_GotFocus()
    SeleccionarCampo tbCodGrupoEdad
End Sub

Private Sub tbCodGrupoEdad_LostFocus()
    tbGrupoEdad.Text = tbGrupoEdad.List(Val(tbCodGrupoEdad.Text) - 1)

End Sub

Private Sub tbCodMod_GotFocus()
    SeleccionarCampo tbCodMod
End Sub

Private Sub tbCodMod_LostFocus()
    tbDescMod.Text = sDescModalidad(Val(tbCodMod.Text))
    If tbDescMod.Text = "" Then tbCodMod.Text = ""

End Sub

Private Sub tbGrupoEdad_Change()
Dim rs As Recordset
Dim sCodGrupo As String

    On Local Error GoTo error
    
    If tbGrupoEdad.Text <> "" Then
        sCodGrupo = BuscarCodGrupoEdad(tbGrupoEdad.Text)
        If sCodGrupo = "" Then
            MsgBox mml_FRASE0415, vbOKOnly Or vbCritical, mml_FRASE0416
            Exit Sub
        Else
            tbCodGrupoEdad.Text = sCodGrupo
        End If
    End If
    Exit Sub
error:
    ProcesarError "tbGrupoEdad_Click"
End Sub

Private Sub tbGrupoEdad_Click()
    tbGrupoEdad_Change
End Sub

Private Sub tbGrupoEdad_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub tbHFecNac_LostFocus()
Dim sCad As String
    If IsDate(tbHFecNac.Text) Then
        tbHEdad.Text = CalcularEdad(CDate(tbHFecNac.Text), tbCodComp.Text)
    End If
    sCad = CalcularGrupoEdad(Val(tbHEdad.Text), Val(tbMEdad.Text), Val(tbCodMod.Text))
    If sCad <> "" Then
        tbGrupoEdad.Text = sCad
        tbCodGrupoEdad.Text = BuscarCodGrupoEdad(sCad)
    End If
End Sub

Private Function CalcularEdad(dHFecNac As Date, sCodComp As String) As String
Dim rs As Recordset
Dim dFecha As Date
    On Local Error GoTo error
    
    If IsDate(dHFecNac) Then
        ' Seleccionamos el año de la competición
        Set rs = db.OpenRecordset("SELECT fecha FROM competiciones WHERE codigo = " & sCodComp, dbOpenSnapshot)
        If rs.EOF Then
            dFecha = Now
        Else
            dFecha = rs!fecha
        End If
        rs.Close
        If G_COUNTRY_MODALID_CALC_GRUPOS_EDAD = "3" Then
        Dim dFechaComp As Date
            ' En country la edad del año se tiene en septiembre
            ' Si la competición es anterior a setiembre la edad es la del año pasado, si no es la de este año
            If Month(dFecha) < 9 Then
                dFechaComp = CDate("01/09/" & Year(dFecha) - 1)
            Else
                dFechaComp = CDate("01/09/" & Year(dFecha))
            End If
            CalcularEdad = DateDiff("yyyy", CDate(dHFecNac), dFechaComp)
        Else
            CalcularEdad = DateDiff("yyyy", CDate(dHFecNac), CDate("01/01/" & Year(dFecha)))
        End If
    'Else
    '    MsgBox mml_FRASE0417, vbOKOnly Or vbInformation, mml_FRASE0096
    End If
    Exit Function
error:
    ProcesarError "tbHFecNac_LostFocus"
End Function
Private Sub tbHNif_LostFocus()
    BuscarPareja C_NIF
End Sub

Private Sub tbHNombre_KeyPress(KeyAscii As Integer)
    EscrituraAnticipada tbHNombre, KeyAscii
End Sub
Private Sub EscrituraAnticipada(tbNombre As ComboBox, KeyAscii As Integer)
Dim rs As Recordset
Dim iPosCursor As Integer
Dim sTabla As String

    ' En cuanto se pulsa espacio comprobamos los artículos que comienzan por la palabra
    On Local Error GoTo error
    
    If chkParejasAEBDC.Value = 1 Then
        sTabla = "parejas_aebdc"
    Else
        sTabla = "parejas"
    End If
    
    If KeyAscii = 8 Then Exit Sub
    
    If Chr$(KeyAscii) = " " Then
        If tbNombre.Tag = "Hombre" Then
            Set rs = db.OpenRecordset("SELECT DISTINCT nombre_hombre FROM " & sTabla & " WHERE nombre_hombre LIKE '" & tbNombre.Text & " *'", dbOpenSnapshot)
        Else
            Set rs = db.OpenRecordset("SELECT DISTINCT nombre_mujer FROM  " & sTabla & "  WHERE nombre_mujer LIKE '" & tbNombre.Text & " *'", dbOpenSnapshot)
        End If
        If Not rs.EOF Then
            iPosCursor = tbNombre.SelStart
            tbNombre.Clear
            If tbNombre.Tag = "Hombre" Then
                tbNombre.Text = rs!nombre_hombre & " "
                While Not rs.EOF
                    tbNombre.AddItem rs!nombre_hombre
                    rs.MoveNext
                Wend
            Else
                tbNombre.Text = rs!nombre_mujer
                While Not rs.EOF
                    tbNombre.AddItem rs!nombre_mujer & " "
                    rs.MoveNext
                Wend
            End If
            tbNombre.SelStart = iPosCursor + 1
            KeyAscii = 0
        End If
    ElseIf tbNombre.SelStart < Len(tbNombre.Text) Then
        Dim i As Integer
        iPosCursor = tbNombre.SelStart
        If iPosCursor > 0 Then
            tbNombre.Text = Mid$(tbNombre.Text, 1, iPosCursor) & Chr$(KeyAscii) & Mid$(tbNombre.Text, iPosCursor + 2)
            KeyAscii = 0
            
            For i = 0 To tbNombre.ListCount - 1
                If UCase(Mid$(tbNombre.List(i), 1, iPosCursor + 1)) = UCase(Mid$(tbNombre.Text, 1, iPosCursor + 1)) Then
                    tbNombre.Text = tbNombre.List(i)
                    tbNombre.SelStart = iPosCursor + 1
                    Exit Sub
                End If
            Next
            tbNombre.Text = Mid$(tbNombre.Text, 1, iPosCursor + 1)
            tbNombre.SelStart = iPosCursor + 1
        End If
    End If
    Exit Sub
error:
    ProcesarError "EscrituraAnticipada"

End Sub

Private Sub tbHNombre_LostFocus()
    tbHNif_LostFocus
End Sub

Private Sub tbHNumSoc_LostFocus()
    tbHNif_LostFocus
End Sub

Private Sub tbMFecNac_LostFocus()
Dim rs As Recordset
Dim dFecha As Date
Dim sCad As String
    On Local Error GoTo error
    
    If IsDate(tbMFecNac.Text) Then
        tbMEdad.Text = CalcularEdad(CDate(tbMFecNac.Text), tbCodComp.Text)
    End If
    sCad = CalcularGrupoEdad(Val(tbHEdad.Text), Val(tbMEdad.Text), Val(tbCodMod.Text))
    If sCad <> "" Then
        tbGrupoEdad.Text = sCad
        tbCodGrupoEdad.Text = BuscarCodGrupoEdad(sCad)
    End If
    
    Exit Sub
error:
    ProcesarError "tbMFecNac_LostFocus"
End Sub


Function CalcularGrupoEdad(hedad As Integer, medad As Integer, iCodMod As Integer)
Dim sCad As String
    On Local Error GoTo error
    
    CalcularGrupoEdad = frmMenu.CalcularGrupoEdad(hedad, medad, iCodMod)
    Exit Function
error:
    ProcesarError "CalcularGrupoEdad"

End Function

Private Sub tbMNif_LostFocus()
Dim rs As Recordset
    On Local Error GoTo error
    
    'Comprobamos si hay un participante en la BD con el mismo NIF y recuperamos la información
    Set rs = db.OpenRecordset("SELECT p.*, c.fecha FROM parejas p, competiciones c WHERE p.cod_competicion = c.codigo AND ((nif_mujer = '" & tbMNif.Text & "' and nif_mujer <> '') OR (num_socio_mujer <> '' AND num_socio_mujer= '" & tbMNumSoc.Text & "')) ORDER BY c.fecha DESC", dbOpenSnapshot)
    If Not rs.EOF Then
        If tbHNombre.Text = "" Then tbHNombre.Text = rs!nombre_hombre
        If tbHFecNac.Text = "" Then tbHFecNac.Text = rs!fecha_nac_hombre
        If tbHEdad.Text = "" Then tbHEdad.Text = rs!edad_hombre
        If tbHNumSoc.Text = "" Then tbHNumSoc.Text = rs!num_socio_hombre
        If tbHNif.Text = "" Then tbHNif.Text = rs!nif_hombre
        If tbMNif.Text = "" Then tbMNif.Text = rs!nif_mujer
        If tbMNombre.Text = "" Then tbMNombre.Text = rs!nombre_mujer
        If tbMFecNac.Text = "" Then tbMFecNac.Text = rs!fecha_nac_mujer
        If tbMEdad.Text = "" Then tbMEdad.Text = rs!edad_mujer
        If tbMNumSoc.Text = "" Then tbMNumSoc.Text = rs!num_socio_mujer
        If tbEscuelas.Text = "" Then tbEscuelas.Text = rs!Escuelas
        If tbDir.Text = "" Then tbDir.Text = rs!direcciones
        If tbProv.Text = "" Then tbProv.Text = rs!provincia
        If tbTlf.Text = "" Then tbTlf.Text = rs!telefonos
        If tbCat.Text = "" Then tbCat.Text = rs!categoria
        If tbCodMod.Text = "" Then tbCodMod.Text = rs!cod_modalidad
        If tbCodGrupoEdad.Text = "" Then tbCodGrupoEdad.Text = rs!cod_grupoedad
        If tbEmail.Text = "" Then tbEmail.Text = rs!email
        If tbEmailMovil.Text = "" Then tbEmailMovil.Text = rs!emailmovil
        If tbSMSTelef.Text = "" Then tbSMSTelef.Text = rs!sms_telef
        If tbGrupoEdad.Text = "" Then tbGrupoEdad.Text = rs!grupoedad
        If tbDescMod.Text = "" Then tbDescMod.Text = Buscar("modalidad", "nombre", tbCodMod.Text)
        
        If cbCombinar.ListIndex = 0 Then cbCombinar.ListIndex = rs!combinar_edad
        
        If chkSMS.Value = 0 Then chkSMS.Value = rs!sms_selec
        If chkEmail.Value = 0 Then chkEmail.Value = rs!email_selec
        If chkEmailMovil.Value = 0 Then chkEmailMovil.Value = rs!emailmovil_selec
        
        'CubrirSiVacio chkParAdicional.Value, Val(SinNulos(rs!pareja_adicional))
        'CubrirSiVacio chkPagado.Value, Val(SinNulos(rs!pagado))
        'If tbCodComp.Text = "" Then tbCodComp.Text = rs!cod_competicion
        'If tbDescComp.Text = "" Then tbDescComp.Text = Buscar(mml_FRASE0049, mml_FRASE0267, dgParejas.Text)
    End If
    rs.Close
    Exit Sub
error:
    ProcesarError "tbMNif_LostFocus"

End Sub


Private Sub tbMNombre_KeyPress(KeyAscii As Integer)
    EscrituraAnticipada tbMNombre, KeyAscii

End Sub

Sub BuscarPareja(iMetodo As Integer)
Dim rs As Recordset
    If Not C_DEBUG Then On Local Error GoTo error
    
    'Comprobamos si hay un participante en la BD con el mismo NIF y recuperamos la información
    If iMetodo = C_NIF Then
        tbHNombre.Text = Trim(tbHNombre.Text)
        tbMNombre.Text = Trim(tbMNombre.Text)
        If chkParejasAEBDC.Value = 1 Then
            Set rs = db.OpenRecordset("SELECT p.* FROM parejas_aebdc p WHERE nombre_hombre Like '" & Trim(tbHNombre.Text) & "'", dbOpenSnapshot)
        Else
            If G_BUSCAR_NOMBRE = "S" Then
                Set rs = db.OpenRecordset("SELECT p.*, c.fecha FROM parejas p, competiciones c WHERE p.cod_competicion = c.codigo AND ((nif_hombre = '" & tbHNif.Text & "' and nif_hombre <> '') OR (num_socio_hombre <> '' AND num_socio_hombre= '" & tbHNumSoc.Text & "') OR (nombre_hombre <> '' AND nombre_hombre Like '" & tbHNombre.Text & "')) ORDER BY c.fecha DESC", dbOpenSnapshot)
            Else
                Set rs = db.OpenRecordset("SELECT p.*, c.fecha FROM parejas p, competiciones c WHERE p.cod_competicion = c.codigo AND ((nif_hombre = '" & tbHNif.Text & "' and nif_hombre <> '') OR (num_socio_hombre <> '' AND num_socio_hombre= '" & tbHNumSoc.Text & "')) ORDER BY c.fecha DESC", dbOpenSnapshot)
            End If
        End If
    ElseIf Val(tbCodAEBDC.Text) <> 0 Then
        If chkParejasAEBDC.Value = 1 Then
            Set rs = db.OpenRecordset("SELECT p.* FROM parejas_aebdc p WHERE p.aebdc_codigo = '" & tbCodAEBDC.Text & "'", dbOpenSnapshot)
        Else
            Set rs = db.OpenRecordset("SELECT p.*, c.fecha FROM parejas p, competiciones c WHERE p.cod_competicion = c.codigo AND p.aebdc_codigo = " & tbCodAEBDC.Text & "  ORDER BY c.fecha DESC", dbOpenSnapshot)
        End If
    End If
    If Not rs.EOF Then
        If tbHNombre.Text = "" Then tbHNombre.Text = rs!nombre_hombre
        If tbHNumSoc.Text = "" Then tbHNumSoc.Text = rs!num_socio_hombre
        If tbMNombre.Text = "" Then tbMNombre.Text = rs!nombre_mujer
        If tbMNumSoc.Text = "" Then tbMNumSoc.Text = rs!num_socio_mujer
        If tbCodAEBDC.Text = "" Then tbCodAEBDC.Text = SinNulos(rs!aebdc_codigo)
        
        If chkParejasAEBDC.Value = 0 Then
            If tbEscuelas.Text = "" Then tbEscuelas.Text = rs!Escuelas
            If tbHFecNac.Text = "" Then tbHFecNac.Text = rs!fecha_nac_hombre
            If tbHEdad.Text = "" Then tbHEdad.Text = rs!edad_hombre
            If tbMNif.Text = "" Then tbMNif.Text = rs!nif_mujer
            If tbMFecNac.Text = "" Then tbMFecNac.Text = rs!fecha_nac_mujer
            If tbMEdad.Text = "" Then tbMEdad.Text = rs!edad_mujer
            If tbDir.Text = "" Then tbDir.Text = rs!direcciones
            If tbProv.Text = "" Then tbProv.Text = rs!provincia
            If tbTlf.Text = "" Then tbTlf.Text = rs!telefonos
            If tbCat.Text = "" Then tbCat.Text = rs!categoria
            If tbCodMod.Text = "" Then tbCodMod.Text = rs!cod_modalidad
            If tbCodGrupoEdad.Text = "" Then tbCodGrupoEdad.Text = rs!cod_grupoedad
            If tbEmail.Text = "" Then tbEmail.Text = rs!email
            If tbEmailMovil.Text = "" Then tbEmailMovil.Text = rs!emailmovil
            If tbSMSTelef.Text = "" Then tbSMSTelef.Text = rs!sms_telef
            If tbGrupoEdad.Text = "" Then tbGrupoEdad.Text = rs!grupoedad
            If tbDescMod.Text = "" Then tbDescMod.Text = Buscar("modalidad", "nombre", tbCodMod.Text)
            
            If cbCombinar.ListIndex = 0 Then cbCombinar.ListIndex = rs!combinar_edad
            
            If chkSMS.Value = 0 Then chkSMS.Value = rs!sms_selec
            If chkEmail.Value = 0 Then chkEmail.Value = rs!email_selec
            If chkEmailMovil.Value = 0 Then chkEmailMovil.Value = rs!emailmovil_selec
            
            'CubrirSiVacio chkParAdicional.Value, Val(SinNulos(rs!pareja_adicional))
            'CubrirSiVacio chkPagado.Value, Val(SinNulos(rs!pagado))
            'If tbCodComp.Text = "" Then tbCodComp.Text = rs!cod_competicion
            'If tbDescComp.Text = "" Then tbDescComp.Text = Buscar(mml_FRASE0049, mml_FRASE0267, dgParejas.Text)
        End If
    End If
    rs.Close
    Exit Sub
error:
    ProcesarError "tbHNif"

End Sub
