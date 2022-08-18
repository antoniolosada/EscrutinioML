VERSION 5.00
Begin VB.Form frmEnlacePPC1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0577"
   ClientHeight    =   11055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEnlacePPC1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   16335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   435
      Left            =   14580
      TabIndex        =   121
      Top             =   0
      Width           =   1725
      Begin VB.Label lblLock 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   60
         TabIndex        =   122
         Top             =   150
         Width           =   1605
      End
   End
   Begin VB.Frame frmPPC 
      Caption         =   "mml_FRASE0577"
      Height          =   11025
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16305
      Begin VB.CommandButton cmdMover 
         Caption         =   "mml_FRASE1007"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   13380
         TabIndex        =   142
         Top             =   2460
         Width           =   1005
      End
      Begin VB.CommandButton cmdBorrarFicherosEntrada 
         Caption         =   "mml_FRASE0584"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   14370
         TabIndex        =   141
         Top             =   2460
         Width           =   1845
      End
      Begin VB.ListBox lstFichRxOrigen 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   10620
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   139
         Top             =   450
         Width           =   5655
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "mml_FRASE1209"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   138
         Top             =   8760
         Width           =   1155
      End
      Begin VB.CommandButton cmdMoverFichero 
         Caption         =   "mml_FRASE1208"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10590
         TabIndex        =   137
         Top             =   5460
         Width           =   2700
      End
      Begin VB.CommandButton cmdActivarFFichSalida 
         Caption         =   "mml_FRASE1197"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10530
         TabIndex        =   136
         Top             =   8730
         Width           =   2040
      End
      Begin VB.ListBox lstError 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   12240
         Sorted          =   -1  'True
         TabIndex        =   134
         Top             =   9390
         Width           =   4005
      End
      Begin VB.CommandButton cmdActivarRx 
         Caption         =   "mml_FRASE1197"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13320
         TabIndex        =   133
         Top             =   5460
         Width           =   2910
      End
      Begin VB.TextBox tbCerrar 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   15240
         TabIndex        =   130
         Top             =   8730
         Width           =   885
      End
      Begin VB.CommandButton cmdCerrarFicheros 
         Caption         =   "mml_FRASE1195"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12630
         TabIndex        =   129
         Top             =   8730
         Width           =   2580
      End
      Begin VB.ListBox lstFichTxTmp 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   13410
         Sorted          =   -1  'True
         TabIndex        =   126
         Top             =   6180
         Width           =   2775
      End
      Begin VB.ListBox lstFichRxTmp 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10590
         Sorted          =   -1  'True
         TabIndex        =   125
         Top             =   4650
         Width           =   5655
      End
      Begin VB.ListBox lstFichTx 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   10560
         Sorted          =   -1  'True
         TabIndex        =   124
         Top             =   6180
         Width           =   2835
      End
      Begin VB.ListBox lstFichRx 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   10590
         Sorted          =   -1  'True
         TabIndex        =   123
         Top             =   2760
         Width           =   5655
      End
      Begin VB.CheckBox chkSoloUnPC 
         Caption         =   "mml_FRASE1187"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   120
         Top             =   2100
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CommandButton cmdPuntuaciones 
         Height          =   435
         Left            =   9540
         Picture         =   "frmEnlacePPC1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "mml_FRASE0040"
         Top             =   780
         Width           =   465
      End
      Begin VB.CheckBox chkCalcAuto 
         Caption         =   "mml_FRASE1172"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5310
         TabIndex        =   118
         Top             =   2310
         Width           =   1605
      End
      Begin VB.CommandButton cmdGenFichHora 
         Caption         =   "mml_FRASE1166"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         TabIndex        =   117
         Top             =   1290
         Width           =   1710
      End
      Begin VB.CommandButton cmdPanelJueces 
         Caption         =   "mml_FRASE1094"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         TabIndex        =   116
         Top             =   8760
         Width           =   2370
      End
      Begin VB.CommandButton cmdControlFicherosTemporales 
         Caption         =   "mml_FRASE1140"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   115
         Top             =   8760
         Width           =   2925
      End
      Begin VB.TextBox tbLog 
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   114
         Top             =   9960
         Width           =   12135
      End
      Begin VB.CheckBox chkRecombinarAlgenerar 
         Caption         =   "mml_FRASE1106"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         TabIndex        =   75
         Top             =   2310
         Width           =   4995
      End
      Begin VB.Frame Frame4 
         Caption         =   "mml_FRASE1041"
         Height          =   1440
         Left            =   420
         TabIndex        =   59
         Top             =   6900
         Width           =   9945
         Begin VB.CommandButton cmdBorrarControlJueces 
            Caption         =   "Limpiar Jueces"
            Height          =   585
            Left            =   9090
            TabIndex        =   76
            Top             =   780
            Width           =   705
         End
         Begin VB.TextBox tbControlJueces 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   180
            Width           =   9735
         End
         Begin VB.TextBox tbJuecesAct 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   60
            Top             =   780
            Width           =   9015
         End
      End
      Begin VB.CheckBox chkRecarga 
         Caption         =   "mml_FRASE1040"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   420
         TabIndex        =   58
         Top             =   8760
         Width           =   1755
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "mml_FRASE0029"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   8850
         TabIndex        =   57
         Top             =   8385
         Width           =   1515
      End
      Begin VB.CommandButton cmdBailes 
         Caption         =   "mml_FRASE0188"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         TabIndex        =   56
         Top             =   8370
         Width           =   2370
      End
      Begin VB.CommandButton cmdBorrarFicheros 
         Caption         =   "mml_FRASE0584"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   55
         Top             =   8385
         Width           =   2925
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "mml_FRASE0058"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   420
         TabIndex        =   54
         Top             =   8370
         Width           =   2865
      End
      Begin VB.CommandButton cmdGenDatos 
         Caption         =   "mml_FRASE0581"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8820
         TabIndex        =   51
         Top             =   1650
         Width           =   1710
      End
      Begin VB.CommandButton cmdSelFase 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1845
         Picture         =   "frmEnlacePPC1.frx":1024
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1305
         Width           =   450
      End
      Begin VB.CommandButton cmdSelCat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1845
         Picture         =   "frmEnlacePPC1.frx":148E
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   855
         Width           =   450
      End
      Begin VB.CommandButton cmdSelComp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1845
         Picture         =   "frmEnlacePPC1.frx":18F8
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   360
         Width           =   450
      End
      Begin VB.ComboBox cbBailes 
         BackColor       =   &H00FFFFC0&
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
         Left            =   6975
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   2070
         Width           =   1800
      End
      Begin VB.Timer tmrCalcular 
         Enabled         =   0   'False
         Left            =   0
         Top             =   6420
      End
      Begin VB.CheckBox chkJPasosGen 
         Caption         =   "mml_FRASE0025"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         TabIndex        =   44
         Top             =   2010
         Width           =   2625
      End
      Begin VB.CommandButton cmdDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9060
         Picture         =   "frmEnlacePPC1.frx":1D62
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "mml_FRASE0425"
         Top             =   780
         Width           =   465
      End
      Begin VB.ComboBox cbPista 
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
         ItemData        =   "frmEnlacePPC1.frx":2044
         Left            =   765
         List            =   "frmEnlacePPC1.frx":2066
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1710
         Width           =   1065
      End
      Begin VB.CheckBox chkUltimos5Bailes 
         Caption         =   "mml_FRASE0578"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4950
         TabIndex        =   36
         Top             =   1980
         Width           =   1845
      End
      Begin VB.CheckBox chkGenSigCat 
         Caption         =   "mml_FRASE0579"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         TabIndex        =   35
         Top             =   1740
         Value           =   1  'Checked
         Width           =   5025
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   0
         Top             =   2730
      End
      Begin VB.ComboBox cbJuez 
         BackColor       =   &H00FFFFC0&
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
         Left            =   6390
         TabIndex        =   30
         Top             =   1305
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "mml_FRASE0580"
         Height          =   1605
         Left            =   435
         TabIndex        =   20
         Top             =   5280
         Width           =   9915
         Begin VB.CommandButton cmdSubirDatosSigFase 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            Picture         =   "frmEnlacePPC1.frx":2099
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   270
            Width           =   405
         End
         Begin VB.CommandButton cmdGenDatosSig 
            Caption         =   "mml_FRASE0581"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   7950
            TabIndex        =   53
            Top             =   270
            Width           =   1875
         End
         Begin VB.CommandButton cmdDatosFaseSig 
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
            Left            =   5250
            Picture         =   "frmEnlacePPC1.frx":2503
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "mml_FRASE0425"
            Top             =   240
            Width           =   420
         End
         Begin VB.TextBox lblJuecesSig 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   1080
            Width           =   9765
         End
         Begin VB.ComboBox cbJuezSig 
            BackColor       =   &H00FFFFC0&
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
            Left            =   6735
            TabIndex        =   28
            Top             =   645
            Width           =   975
         End
         Begin VB.TextBox tbDescFaseSig 
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
            Left            =   915
            TabIndex        =   25
            Top             =   660
            Width           =   4740
         End
         Begin VB.TextBox tbCodFaseSig 
            BackColor       =   &H00C0FFFF&
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
            Left            =   75
            TabIndex        =   24
            Top             =   660
            Width           =   855
         End
         Begin VB.TextBox tbDescCatSig 
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
            Left            =   1365
            TabIndex        =   23
            Top             =   270
            Width           =   3840
         End
         Begin VB.TextBox tbCodCatSig 
            BackColor       =   &H00C0FFFF&
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
            Left            =   510
            TabIndex        =   22
            Top             =   270
            Width           =   855
         End
         Begin VB.CheckBox chkRepSig 
            Caption         =   "mml_FRASE0289"
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
            Left            =   5985
            TabIndex        =   21
            Top             =   225
            Width           =   1770
         End
         Begin VB.Label Label5 
            Caption         =   "mml_FRASE0421"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5925
            TabIndex        =   29
            Top             =   735
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "mml_FRASE0582"
         Height          =   2670
         Left            =   450
         TabIndex        =   14
         Top             =   2640
         Width           =   9915
         Begin VB.CommandButton cmdGenDatosAct 
            Caption         =   "mml_FRASE0581"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   8010
            TabIndex        =   52
            Top             =   120
            Width           =   1830
         End
         Begin VB.CommandButton cmdSubirDatos 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   90
            Picture         =   "frmEnlacePPC1.frx":27E5
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   225
            Width           =   405
         End
         Begin VB.TextBox tbNumJueces 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
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
            Left            =   3570
            TabIndex        =   42
            Top             =   210
            Width           =   525
         End
         Begin VB.CommandButton cmdDatosFaseAct 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6210
            Picture         =   "frmEnlacePPC1.frx":2C4F
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "mml_FRASE0425"
            Top             =   150
            Width           =   420
         End
         Begin VB.ComboBox cbJuezAct 
            BackColor       =   &H00FFFFC0&
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
            Left            =   5070
            TabIndex        =   26
            Top             =   210
            Width           =   975
         End
         Begin VB.CheckBox chkRepAct 
            Caption         =   "mml_FRASE0289"
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
            Left            =   6750
            TabIndex        =   19
            Top             =   270
            Width           =   1320
         End
         Begin VB.TextBox tbCodCatAct 
            BackColor       =   &H00C0FFFF&
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
            Left            =   45
            TabIndex        =   18
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox tbDescCatAct 
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
            Left            =   885
            TabIndex        =   17
            Top             =   720
            Width           =   4440
         End
         Begin VB.TextBox tbCodFaseAct 
            BackColor       =   &H00C0FFFF&
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
            Left            =   5325
            TabIndex        =   16
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox tbDescFaseAct 
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
            Left            =   6165
            TabIndex        =   15
            Top             =   720
            Width           =   3660
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   17
            Left            =   9240
            TabIndex        =   113
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   17
            Left            =   9570
            TabIndex        =   112
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   16
            Left            =   8700
            TabIndex        =   111
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   16
            Left            =   9030
            TabIndex        =   110
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   2
            Left            =   1140
            TabIndex        =   109
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   2
            Left            =   1470
            TabIndex        =   108
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   15
            Left            =   8490
            TabIndex        =   107
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   15
            Left            =   8160
            TabIndex        =   106
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   14
            Left            =   7950
            TabIndex        =   105
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   14
            Left            =   7620
            TabIndex        =   104
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   13
            Left            =   7410
            TabIndex        =   103
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   13
            Left            =   7080
            TabIndex        =   102
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   12
            Left            =   6870
            TabIndex        =   101
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   12
            Left            =   6540
            TabIndex        =   100
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   11
            Left            =   6330
            TabIndex        =   99
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   11
            Left            =   6000
            TabIndex        =   98
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   10
            Left            =   5790
            TabIndex        =   97
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   10
            Left            =   5460
            TabIndex        =   96
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   9
            Left            =   5250
            TabIndex        =   95
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   9
            Left            =   4920
            TabIndex        =   94
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   8
            Left            =   4710
            TabIndex        =   93
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   8
            Left            =   4380
            TabIndex        =   92
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   7
            Left            =   4170
            TabIndex        =   91
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   7
            Left            =   3840
            TabIndex        =   90
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   6
            Left            =   3630
            TabIndex        =   89
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   6
            Left            =   3300
            TabIndex        =   88
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   5
            Left            =   3090
            TabIndex        =   87
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   5
            Left            =   2760
            TabIndex        =   86
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   4
            Left            =   2550
            TabIndex        =   85
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   4
            Left            =   2220
            TabIndex        =   84
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   3
            Left            =   2010
            TabIndex        =   83
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   3
            Left            =   1680
            TabIndex        =   82
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   930
            TabIndex        =   81
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   600
            TabIndex        =   80
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label tbNumBAilesTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   390
            TabIndex        =   79
            Top             =   2220
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label tbJuezTx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   60
            TabIndex        =   78
            Top             =   2220
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label Label7 
            Caption         =   "mml_FRASE0026"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1170
            TabIndex        =   43
            Top             =   300
            Width           =   2385
         End
         Begin VB.Label lblJuecesAct 
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   60
            TabIndex        =   34
            Top             =   1140
            Width           =   9765
         End
         Begin VB.Label Label4 
            Caption         =   "mml_FRASE0050"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4170
            TabIndex        =   27
            Top             =   300
            Width           =   840
         End
      End
      Begin VB.CommandButton cmdDescalif 
         Height          =   435
         Left            =   10020
         Picture         =   "frmEnlacePPC1.frx":2F31
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "mml_FRASE0043"
         Top             =   780
         Width           =   465
      End
      Begin VB.CommandButton cmdDorsales 
         Height          =   375
         Left            =   9840
         Picture         =   "frmEnlacePPC1.frx":3B13
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "mml_FRASE0028"
         Top             =   360
         Width           =   645
      End
      Begin VB.CommandButton cmdCategAct 
         Height          =   375
         Left            =   9060
         Picture         =   "frmEnlacePPC1.frx":4635
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "mml_FRASE0428"
         Top             =   360
         Width           =   675
      End
      Begin VB.TextBox tbCodComp 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox tbDescComp 
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
         Left            =   3120
         TabIndex        =   6
         Top             =   360
         Width           =   5895
      End
      Begin VB.TextBox tbCodCat 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2280
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox tbDescCat 
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
         Left            =   3120
         TabIndex        =   4
         Top             =   840
         Width           =   5895
      End
      Begin VB.TextBox tbCodFase 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2280
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox tbDescFase 
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
         Left            =   3120
         TabIndex        =   2
         Top             =   1320
         Width           =   2610
      End
      Begin VB.CheckBox chkRep 
         Caption         =   "mml_FRASE0289"
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
         Left            =   7425
         TabIndex        =   1
         Top             =   1305
         Width           =   1365
      End
      Begin VB.Label Label11 
         Caption         =   "mml_FRASE1212"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10620
         TabIndex        =   140
         Top             =   120
         Width           =   3885
      End
      Begin VB.Label Label10 
         Caption         =   "mml_FRASE1207"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12240
         TabIndex        =   135
         Top             =   9090
         Width           =   3885
      End
      Begin VB.Label Label9 
         Caption         =   "mml_FRASE1194"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10560
         TabIndex        =   132
         Top             =   5820
         Width           =   3885
      End
      Begin VB.Label lblFichRx 
         Caption         =   "mml_FRASE1193"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10590
         TabIndex        =   131
         Top             =   2430
         Width           =   2865
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   14
         Left            =   11370
         TabIndex        =   128
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   13
         Left            =   10560
         TabIndex        =   127
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   12
         Left            =   9750
         TabIndex        =   74
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   11
         Left            =   8940
         TabIndex        =   73
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   10
         Left            =   8130
         TabIndex        =   72
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   9
         Left            =   7320
         TabIndex        =   71
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   8
         Left            =   6510
         TabIndex        =   70
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   7
         Left            =   5700
         TabIndex        =   69
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   6
         Left            =   4920
         TabIndex        =   68
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   5
         Left            =   4110
         TabIndex        =   67
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   4
         Left            =   3300
         TabIndex        =   66
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   3
         Left            =   2490
         TabIndex        =   65
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   2
         Left            =   1680
         TabIndex        =   64
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A)  100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   1
         Left            =   870
         TabIndex        =   63
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblBat 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(A1)100% 228 min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Index           =   0
         Left            =   60
         TabIndex        =   62
         Top             =   9150
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         Caption         =   "mml_FRASE0011"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   46
         Top             =   1785
         Width           =   1770
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0437"
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
         Left            =   135
         TabIndex        =   38
         Top             =   1755
         Width           =   630
      End
      Begin VB.Label lblActTimer 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mml_FRASE0583"
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
         Left            =   8820
         TabIndex        =   32
         Tag             =   "0"
         Top             =   2070
         Width           =   1740
      End
      Begin VB.Label Label6 
         Caption         =   "mml_FRASE0421"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   31
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblComp 
         Caption         =   "mml_FRASE0299"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmEnlacePPC1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim g_sNoPresentes As String

Const C_NP_NO_INICIADO = 0
Const C_NP_INICIADO = 1
Const C_NP_PREPARADO = 2

Const C_MAX_JUEZ_PANEL = 15

Dim aNoPresentes() As String

Dim g_iNoPresentesIniciado As Integer
Dim g_iCJueces As Integer
Dim g_bNoPresentesProcesados As Boolean
Dim g_sNumNoPresentes As Integer

Dim iCodCat As Integer
Dim iFase As Integer
Dim iRepesca As Integer

Dim m_sFicheroEntrada As String
Dim m_sFicheroSalida As String
Dim m_sFicheroHora As String
Dim m_sFicheroBateria As String
Dim m_sFicheroControlJueces As String
Dim m_sVersiónFicheroJueces As String
Dim m_sPathOrigen As String

Sub BorrarFicheros(Optional bEntrada As Boolean = True, Optional bSoloJueces As Boolean = False)
    On Error Resume Next
    If bSoloJueces Then
        Kill m_sFicheroControlJueces & "*"
    Else
        chkRecarga.Value = False
        Kill m_sFicheroSalida & ".*.TXT"
        Kill m_sFicheroBateria & ".*.TXT"
        Kill m_sFicheroSalida & ".Jueces.TXT"
        Kill m_sFicheroControlJueces & "*.TXT"
        If bEntrada Then
            Kill m_sFicheroEntrada & "*"
        End If
    End If
End Sub
Sub CargarPathFicheros()
Dim sPath As String
Dim sPathCopia As String
Dim sFich As String
Dim sPista As String

    If Not C_DEBUG Then On Local Error GoTo error
    
    ' Si no hemos seleccionado pista o si no utilizamos solo un PC para todas las pistas
    If cbPista.ListIndex <= 0 Or Not chkSoloUnPC.Value = 1 Then
        sPista = ""
    Else
        sPista = "\P" & Trim$(Str$(cbPista.ListIndex))
    End If
    
    m_sPathOrigen = VarCfg("dir_fichas") & sPista
    If G_MOVER_FICHEROS_PDA Then
        sPath = G_RUTA_COPIA_FICH_PDA & sPista
        lstFichRxOrigen.BackColor = vbGreen
    Else
        sPath = m_sPathOrigen
        lstFichRxOrigen.BackColor = lstFichRx.BackColor
    End If
    
    sFich = sExtraerFichero(G_FICHERO_ENTRADA)
    m_sFicheroEntrada = sPath & "\" & sFich
    
    sFich = sExtraerFichero(G_FICHERO_BATERIA)
    m_sFicheroBateria = sPath & "\" & sFich
    
    sFich = sExtraerFichero(G_FICHERO_CONTROL_JUECES)
    m_sFicheroControlJueces = sPath & "\" & sFich

    sFich = sExtraerFichero(G_FICHERO_HORA)
    m_sFicheroHora = m_sPathOrigen & "\" & sFich
    
    'El fichero de salida tiene una ruta distinta
    sPath = sExtraerPath(G_FICHERO_SALIDA)
    sFich = sExtraerFichero(G_FICHERO_SALIDA)
    m_sFicheroSalida = sPath & "\" & sFich

    Exit Sub
error:
    PPCLog ProcesarError("CargarPathFicheros", False)
End Sub

Private Sub cbPista_Click()
    CargarPathFicheros
    Me.Caption = cbPista.Text & " " & mml_FRASE0577
End Sub

Private Sub chkRecarga_Click()
    
    If Not C_DEBUG Then On Local Error GoTo error
    If chkRecarga.Value = 1 Then
        If MsgBox(mml_FRASE1137, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
            Call cmdBorrarControlJueces_Click
        End If
    End If
    GenerarAutorizacionRecarga chkRecarga.Value
    Exit Sub
error:
    ProcesarError "chkRecarga_Click"
End Sub

Private Sub cmdActivarFFichSalida_Click()
Dim i As Integer
Dim sFile As String
Dim sFileDestino As String
Dim sPath As String
    
    If Not PreguntaOperacion Then Exit Sub
    
    If Not C_DEBUG Then On Local Error GoTo error
    If lstFichTxTmp.ListIndex >= 0 Then
        sFile = lstFichTxTmp.List(lstFichTxTmp.ListIndex)
        i = InStr(sFile, ".cTXT")
        If i > 0 Then
            sFileDestino = Mid$(sFile, 1, i - 1) & ".TXT"
            sPath = sExtraerPath(m_sFicheroSalida)
            Name sPath & sFile As sPath & sFileDestino
        End If
    End If
    ActualizarListasFicheros
    Exit Sub
error:
    ProcesarError "cmdActivarFFichSalida_Click"

End Sub

Private Sub cmdActivarRx_Click()
Dim i As Integer
Dim sFile As String
Dim sFileDestino As String
Dim sPath As String
    
    If Not PreguntaOperacion Then Exit Sub
    
    If Not C_DEBUG Then On Local Error GoTo error
    If lstFichRxTmp.ListIndex >= 0 Then
        sFile = lstFichRxTmp.List(lstFichRxTmp.ListIndex)
        i = InStr(sFile, ".cTXT")
        If i > 0 Then
            sFileDestino = Mid$(sFile, 1, i - 1) & ".TXT"
            sPath = m_sPathOrigen & "\"
            Name sPath & sFile As sPath & sFileDestino
        End If
    End If
    ActualizarListasFicheros
    Exit Sub
error:
    ProcesarError "cmdActivarRx_Click"
End Sub

Private Sub cmdBailes_Click()
Dim rs As Recordset, sMsj As String

    If Not C_DEBUG Then On Local Error GoTo error
    Set rs = db.OpenRecordset("SELECT * FROM bailes")
    While Not rs.EOF
        sMsj = sMsj & rs!codigo & " - " & rs!Nombre & Chr$(13) & Chr$(10)
        rs.MoveNext
    Wend
    rs.Close
    
    MsgBox sMsj, vbOKOnly Or vbInformation, mml_FRASE0185
    Exit Sub
error:
    PPCLog ProcesarError("cmdBailes_Click", False)
End Sub

Private Sub cmdBorrarControlJueces_Click()
    On Local Error Resume Next
    Kill m_sFicheroControlJueces & ".*.TXT"
    tbJuecesAct.Text = ""
End Sub

Private Sub cmdBorrarFicheros_Click()
    If MsgBox(mml_FRASE0585, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        On Local Error Resume Next
        
        If MsgBox(mml_FRASE0586, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
            BorrarFicheros True
        Else
            BorrarFicheros False
        End If
        MsgBox mml_FRASE0587, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
    
End Sub

Private Sub cmdBorrarFicherosEntrada_Click()
Dim i As Integer

    If Not PreguntaOperacion Then Exit Sub
    
    If Not C_DEBUG Then On Local Error GoTo error
    For i = 0 To lstFichRxOrigen.ListCount - 1
        If lstFichRxOrigen.Selected(i) Then
            Kill m_sPathOrigen & "\" & lstFichRxOrigen.List(i)
        End If
    Next
    ActualizarListasFicheros
    Exit Sub
error:
    ProcesarError "cmdBorrarFicherosEntrada_Click"
End Sub

Private Sub cmdCalcular_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        cmdSubirDatos_Click
        
        frmCalcular.tbCodComp.Text = tbCodComp.Text
        frmCalcular.tbDescComp.Text = tbDescComp.Text
        frmCalcular.tbCodCat.Text = tbCodCatAct.Text
        frmCalcular.tbDescCat.Text = tbDescCatAct.Text
        frmCalcular.tbCodFase.Text = Val(tbCodFaseAct.Text)
        frmCalcular.tbDescFase.Text = sDescFase(Val(tbCodFaseAct.Text))
        frmCalcular.chkRep.Value = chkRepAct.Value
    
        frmCalcular.Show vbModal
    End If
error:
    PPCLog ProcesarError("cmdCalcular_Click", False)

End Sub

Private Sub cmdCategAct_Click()
Dim rs As Recordset, sMsj As String
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    RecuperarCatActualHorario rs, cbPista.List(cbPista.ListIndex), Val(tbCodComp.Text)
    If Not rs.EOF Then
        tbCodComp.Text = rs!cod_competicion
        tbDescComp.Text = sDescCompeticion(rs!cod_competicion)
        tbCodCat.Text = rs!cod_categoria
        tbDescCat.Text = sDescCategoria(rs!cod_categoria)
        tbCodFase.Text = rs!numfase
        chkRep.Value = rs!repesca
        DescFase
        CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
        
        If rs!cod_Baile > 0 Then
            Dim i As Integer
            
            For i = 0 To cbBailes.ListCount - 1
                If Val(cbBailes.List(i)) = rs!cod_Baile Then
                    cbBailes.ListIndex = i
                End If
            Next
        ElseIf rs!cod_Baile < 0 Then
            chkUltimos5Bailes.Value = 1
        Else
            If C_RESET_ULTIMOS_5_BAILES Then chkUltimos5Bailes.Value = 0
        End If
        
    Else
        MsgBox mml_FRASE0442, vbOKOnly Or vbInformation, mml_FRASE0147
    End If
    rs.Close

    RecuperarJueces Val(tbCodCat.Text), cbJuez
    
    Exit Sub
error:
    PPCLog ProcesarError("cmdCategAct_Click", False)
End Sub

Private Sub cmdCerrarFicheros_Click()
Dim i As Integer
    
    If Not PreguntaOperacion Then Exit Sub
    
    Timers False
    For i = 0 To 32767
        Err.Clear
        On Local Error Resume Next
        Close #i
        If i Mod 10 = 0 Then
            tbCerrar.Text = i
            tbCerrar.Refresh
        End If
    Next
    Timers True
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
End Sub

Private Sub cmdControlFicherosTemporales_Click()
Dim sFichero As String
Dim sCad As String
Dim sPath As String

    If Not C_DEBUG Then On Local Error GoTo error
    'Localizamos los ficheros resultado de los PPC
    sCad = ""
    sFichero = Dir$(m_sFicheroEntrada & ".*.cTXT")
    While sFichero <> ""
        sCad = sCad & vbCrLf & sFichero
        sFichero = Dir
    Wend
    If sCad <> "" Then
        If MsgBox(mml_FRASE1141 & vbCrLf & sCad, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
            sFichero = Dir$(m_sFicheroEntrada & ".*.cTXT")
            While sFichero <> ""
                sCad = sCad & vbCrLf & sFichero
                sFichero = Mid$(sFichero, 1, Len(sFichero) - 5)
                sPath = sExtraerPath(m_sFicheroEntrada)
                Name sPath & "\" & sFichero & ".cTXT" As sPath & "\" & sFichero & ".TXT"
                sFichero = Dir
            Wend
            MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
        End If
    Else
        MsgBox mml_FRASE1142, vbOKOnly Or vbInformation, G_MSG_AVISO
    End If
    Exit Sub
error:

    PPCLog ProcesarError("cmdControlFicherosTemporales_Click", False)
End Sub

Private Sub cmdDatos_Click()
    MostrarDatosIntroducidos Val(tbCodCat.Text), Val(tbCodFase.Text), chkRep.Value

End Sub

Private Sub cmdDatosFaseAct_Click()
    MostrarDatosIntroducidos Val(tbCodCatAct.Text), Val(tbCodFaseAct.Text), 0
End Sub

Private Sub cmdDatosFaseSig_Click()
    MostrarDatosIntroducidos Val(tbCodCatSig.Text), Val(tbCodFaseSig.Text), 0
End Sub

Private Sub cmdDescalif_Click()
    
    If Not C_DEBUG Then On Local Error GoTo error
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        frmDescalificados.tbCodComp.Text = tbCodComp.Text
        frmDescalificados.tbDescComp.Text = tbDescComp.Text
        frmDescalificados.tbCodCateg.Text = tbCodCat.Text
        frmDescalificados.tbDescCateg.Text = tbDescCat.Text
        frmDescalificados.cbFase.ListIndex = Log(Val(tbCodFase.Text)) / Log(2)
        frmDescalificados.cmdActualizar_Click

        frmDescalificados.Show vbNomodal
    End If
    Exit Sub
error:
    PPCLog ProcesarError("cmdDescalif_Click", False)

End Sub

Private Sub cmdDorsales_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        frmADorsales.tbCodComp.Text = tbCodComp.Text
        frmADorsales.tbDescComp.Text = tbDescComp.Text
        frmADorsales.tbCodCateg.Text = tbCodCat.Text
        frmADorsales.tbDescCateg.Text = tbDescCat.Text
        frmADorsales.cbFase.ListIndex = Log(Val(tbCodFase.Text)) / Log(2) + 1
        
        frmADorsales.Show vbNomodal
        
    End If
    Exit Sub
error:
    PPCLog ProcesarError("cmdDorsales_Click", False)
End Sub

Private Sub cmdGenDatos_Click()
Dim i As Integer
    
    If Not C_DEBUG Then On Error GoTo error
        
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If MsgBox(mml_FRASE0588, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If Val(cbBailes.Text) = 0 Then
        CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
    End If
    If MsgBox(mml_FRASE1196, vbYesNo Or vbCritical, G_MSG_AVISO) = vbYes Then
        BorrarFicheros True
    Else
        BorrarFicheros False
    End If
    
    DoEvents: DoEvents
    If cbJuez.Text = "" Then
        For i = 0 To cbJuez.ListCount - 1
            GenerarFichero cbJuez.List(i), Val(tbCodCat.Text), tbDescCat.Text, Val(tbCodFase.Text), chkRep.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
        Next
    Else
        GenerarFichero cbJuez.Text, Val(tbCodCat.Text), tbDescCat.Text, Val(tbCodFase.Text), chkRep.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
    End If
    GenerarFicheroJuezPasosGenerico
    GenerarFicheroJueces Val(tbCodCat.Text), True
    
    If Val(cbBailes.Text) > 0 Then
        CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
    End If
    
    lblJuecesAct.Caption = ""
    tbCodCatAct.Text = tbCodCat.Text
    tbDescCatAct.Text = tbDescCat.Text
    tbCodFaseAct.Text = tbCodFase.Text
    tbDescFaseAct.Text = tbDescFase.Text
    chkRepAct.Value = chkRep.Value
    
    tbNumJueces.Text = RecuperarJueces(Val(tbCodCat.Text), cbJuezAct)
    BorrarDatosFaseSig
    
    Exit Sub
error:
    PPCLog ProcesarError("cmdGenDatos_Click", False)
    
End Sub
Sub GenerarFicheroJueces(iCodCat As Integer, Optional bActPanel As Boolean = False, Optional iActVersion As Integer = 1)
Dim rs As Recordset
Dim iFile As Integer
Dim i As Integer
Dim sFichero As String
Dim sNombreJuez As String

    If Not C_DEBUG Then On Local Error GoTo error
    iFile = FreeFile
    sFichero = m_sFicheroSalida & ".Jueces"
    Open sFichero & ".cTXT" For Output As #iFile
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM juez_categ jc, jueces j WHERE jc.cod_juez = j.codigo AND cod_categoria = " & iCodCat & " ORDER BY id_juez", dbOpenSnapshot)
    tbControlJueces.Text = ""
    
    If bActPanel Then
        i = 0
        For i = 0 To C_MAX_JUEZ_PANEL - 1
            tbJuezTx(i).Visible = False
            tbNumBAilesTx(i).Visible = False
        Next
    End If
    i = 0
    'Imprimimos la versión
    If iActVersion Then
        m_sVersiónFicheroJueces = Trim$(Str$(Int(Rnd * 100000)))
    End If
    Print #iFile, m_sVersiónFicheroJueces
    
    While Not rs.EOF
        Print #iFile, rs!id_juez & "-" & rs!Nombre
        sNombreJuez = Trim$(Left$(rs.Fields("nombre"), 5))
        sNombreJuez = CambiarCadena(" ", "_", sNombreJuez)
        tbControlJueces.Text = tbControlJueces.Text & rs!id_juez & "-" & sNombreJuez & " "
        
        If bActPanel Then
            If i < C_MAX_JUEZ_PANEL Then
                tbJuezTx(i).BackColor = vbYellow
                tbJuezTx(i).Caption = rs!id_juez
                tbNumBAilesTx(i).Caption = 0
                tbJuezTx(i).Visible = True
                tbNumBAilesTx(i).Visible = True
                i = i + 1
            End If
        End If
        
        rs.MoveNext
    Wend
    rs.Close
    Close #iFile
    On Local Error Resume Next
    Kill sFichero & ".TXT"
    Name sFichero & ".cTXT" As sFichero & ".TXT"
    Exit Sub
error:
    PPCLog ProcesarError("GenerarFicheroJueces", False)
End Sub
Sub GenerarAutorizacionRecarga(iAut As Integer)
Dim rs As Recordset
Dim iFile As Integer
Dim i As Integer
Dim sFile As String

    If Not C_DEBUG Then On Local Error GoTo error
    
    If iAut = 1 Then
        iFile = FreeFile
        sFile = m_sFicheroSalida & ".Recarga"
        Open sFile & ".cTXT" For Output As #iFile
        Print #iFile, "RECARGA/RELOAD"
        Close #iFile
        Name sFile & ".cTXT" As sFile & ".TXT"
        If chkAuto.Value = 1 Then
            'Generamos ficheros de recarga independientes para cada juez del panel
            For i = 0 To cbJuezAct.ListCount - 1
                sFile = m_sFicheroSalida & "." & cbJuezAct.List(i) & ".Recarga"
                iFile = FreeFile
                Open sFile & ".cTXT" For Output As #iFile
                Print #iFile, "RECARGA/RELOAD"
                Close #iFile
                Name sFile & ".cTXT" As sFile & ".TXT"
            Next
        End If
    Else
        chkAuto.Value = 0
        On Local Error Resume Next
        Kill m_sFicheroSalida & ".*.Recarga.TXT"
        Kill m_sFicheroSalida & ".Recarga.TXT"
    End If
    Exit Sub
error:
    PPCLog ProcesarError("GenerarAutorizacionRecarga", False)
End Sub
Sub GenerarFicheroJuezPasosGenerico()
Dim rs As Recordset, iFile As Long
Dim sFichero As String

    If Not C_DEBUG Then On Local Error GoTo error
    iFile = FreeFile
    sFichero = m_sFicheroSalida & ".JpGen"
    Open sFichero & ".cTXT" For Output As #iFile
    Set rs = db.OpenRecordset("SELECT codigo, descripcion FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY 2", dbOpenSnapshot)
    rs.MoveLast
    Print #iFile, Trim$(Str$(rs.RecordCount))
    rs.MoveFirst
    While Not rs.EOF
        Print #iFile, rs!codigo & "-" & rs!DESCRIPCION
        rs.MoveNext
    Wend
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT codigo, nombre FROM bailes ORDER BY 1", dbOpenSnapshot)
    rs.MoveLast
    Print #iFile, Trim$(Str$(rs.RecordCount))
    rs.MoveFirst
    While Not rs.EOF
        Print #iFile, rs!codigo & "-" & rs!Nombre
        rs.MoveNext
    Wend
    rs.Close
    Close #iFile
    
    'Renombramos el fichero
    On Local Error Resume Next
    Kill sFichero & ".TXT"
    Name sFichero & ".cTXT" As sFichero & ".TXT"
    Exit Sub
error:
    PPCLog ProcesarError("GenerarFicheroJuezPasosGenerico", False)
End Sub

Private Sub cmdGenDatosAct_Click()
Dim i As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    If tbCodCatAct.Text = "" Or tbCodFaseAct.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If MsgBox(mml_FRASE0588, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    CargarBailes Val(tbCodCatAct.Text), Val(tbCodFaseAct.Text), cbBailes
    If MsgBox(mml_FRASE1196, vbYesNo Or vbCritical, G_MSG_AVISO) = vbYes Then
        BorrarFicheros True
    Else
        BorrarFicheros False
    End If
    If cbJuezAct.Text = "" Then
        For i = 0 To cbJuezAct.ListCount - 1
            GenerarFichero cbJuezAct.List(i), Val(tbCodCatAct.Text), tbDescCatAct.Text, Val(tbCodFaseAct.Text), chkRepAct.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
        Next
    Else
        GenerarFichero cbJuezAct.Text, Val(tbCodCatAct.Text), tbDescCatAct.Text, Val(tbCodFaseAct.Text), chkRepAct.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
    End If
    GenerarFicheroJuezPasosGenerico
    GenerarFicheroJueces Val(tbCodCat.Text), False
    
    Exit Sub
error:
    PPCLog ProcesarError("cmdGenDatosAct_Click", False)

End Sub

Private Sub cmdGenDatosSig_Click()
Dim i As Integer
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    If tbCodCatSig.Text = "" Or tbCodFaseSig.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If MsgBox(mml_FRASE0588, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    CargarBailes Val(tbCodCatSig.Text), Val(tbCodFaseSig.Text), cbBailes
    If MsgBox(mml_FRASE1196, vbYesNo Or vbCritical, G_MSG_AVISO) = vbYes Then
        BorrarFicheros True
    Else
        BorrarFicheros False
    End If
    If cbJuezSig.Text = "" Then
        For i = 0 To cbJuezSig.ListCount - 1
            GenerarFichero cbJuezSig.List(i), Val(tbCodCatSig.Text), tbDescCatSig.Text, Val(tbCodFaseSig.Text), chkRepSig.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
        Next
    Else
        GenerarFichero cbJuezSig.List(i), Val(tbCodCatSig.Text), tbDescCatSig.Text, Val(tbCodFaseSig.Text), chkRepSig.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
    End If
    GenerarFicheroJuezPasosGenerico
    GenerarFicheroJueces Val(tbCodCat.Text), False
    Exit Sub
error:
    PPCLog ProcesarError("cmdGenDatosSig_Click", False)
End Sub
Private Function RecuperarPanelDeJuecesCompleto() As String
Dim rs As Recordset
Dim sJueces As String

    If Not C_DEBUG Then On Local Error GoTo error
    
    If Val(tbCodComp.Text) = 0 Then
        CamposSinCubrir
        Exit Function
    End If
    
    Set rs = db.OpenRecordset("SELECT DISTINCT id_juez FROM juez_categ jc, categorias c WHERE jc.cod_categoria = c.codigo AND c.codigo IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ") AND c.descripcion LIKE '*" & cbPista.List(cbPista.ListIndex) & "*' ORDER BY 1", dbOpenSnapshot)
    sJueces = ""
    While Not rs.EOF
        If sJueces <> "" Then sJueces = sJueces & ","
        sJueces = sJueces & rs.Fields("id_juez")
        rs.MoveNext
    Wend
    rs.Close
    
    RecuperarPanelDeJuecesCompleto = sJueces
    Exit Function
error:
    PPCLog ProcesarError("RecuperarPanelDeJuecesCompleto", False)
End Function

Private Sub cmdGenFichHora_Click()
Dim iFile As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    iFile = FreeFile
    Open m_sFicheroHora & ".cTXT" For Output As #iFile
    Print #iFile, Format$(Now, "dd/mm/yyyy")
    Print #iFile, Format$(Time, "hh:mm:ss")
    Close #iFile
    On Local Error Resume Next
    Kill m_sFicheroHora & ".TXT"
    If Not C_DEBUG Then On Local Error GoTo error
    Name m_sFicheroHora & ".cTXT" As m_sFicheroHora & ".TXT"
    
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    Exit Sub
error:
    ProcesarError "cmdGenFichHora_Click"
End Sub

Private Sub cmdMover_Click()
Dim i As Integer
Dim sFile As String
Dim sPathDestino As String
Dim sPath As String
    
    If Not PreguntaOperacion Then Exit Sub
    
    If Not C_DEBUG Then On Local Error GoTo error
    If lstFichRxOrigen.ListIndex >= 0 Then
        sFile = lstFichRxOrigen.List(lstFichRxOrigen.ListIndex)
        i = InStr(sFile, ".cTXT")
        If i > 0 Then
            sPathDestino = m_sPathOrigen & "\COPIAS_PDA\FORZADA\"
            sPath = m_sPathOrigen & "\"
            Name sPath & sFile As sPathDestino & sFile
        End If
    End If
    ActualizarListasFicheros
    Exit Sub
error:
    ProcesarError "cmdMover_Click"

End Sub

Private Sub cmdMoverFichero_Click()
Dim i As Integer
Dim sFile As String
Dim sPathDestino As String
Dim sPath As String
    
    If Not PreguntaOperacion Then Exit Sub
    
    If Not C_DEBUG Then On Local Error GoTo error
    If lstFichRx.ListIndex >= 0 Then
        sFile = lstFichRx.List(lstFichRxOrigen.ListIndex)
        i = InStr(sFile, ".cTXT")
        If i > 0 Then
            sPathDestino = m_sPathOrigen & "\COPIAS_PDA\FORZADA\"
            sPath = m_sPathOrigen & "\"
            Name sPath & sFile As sPathDestino & sFile
        End If
    End If
    ActualizarListasFicheros
    Exit Sub
error:
    ProcesarError "cmdMoverFichero_Click"

End Sub

Private Sub cmdPanelJueces_Click()
Dim rs As Recordset
Dim sJueces As String

    If Not C_DEBUG Then On Local Error GoTo error
    
    If Val(tbCodComp.Text) = 0 Or Val(tbCodCat.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM jueces j, juez_categ jc WHERE jc.cod_juez = j.codigo AND jc.cod_categoria = " & tbCodCat.Text & "  ORDER BY 1", dbOpenSnapshot)
    sJueces = ""
    While Not rs.EOF
        sJueces = sJueces & rs.Fields("id_juez") & " - " & rs.Fields("nombre") & vbCrLf
        rs.MoveNext
    Wend
    rs.Close
    
    MsgBox mml_FRASE1094 & vbCrLf & vbCrLf & sJueces, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    Exit Sub
error:
    PPCLog ProcesarError("cmdPanelJueces_Click", False)
    
End Sub

Private Sub cmdPuntuaciones_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        With frmAPuntuacionesBaile
            .tbCodComp.Text = tbCodComp.Text
            .tbDescComp.Text = tbDescComp.Text
            .tbCodCateg.Text = tbCodCat.Text
            .tbDescCateg.Text = tbDescCat.Text
            .cbFase.ListIndex = Log(Val(tbCodFase.Text)) / Log(2) + 1
            
            .Show vbNomodal
        End With
    End If
    Exit Sub
error:
    PPCLog ProcesarError("cmdPuntuaciones_Click", False)

End Sub

Private Sub cmdSalir_Click()
    On Local Error Resume Next
    Unload Me
End Sub

Private Sub cmdSelCat_Click()

    If Not C_DEBUG Then On Local Error GoTo error
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCat.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text, "descripcion", " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCat.Text = sResultado(2)
    tbCodFase.Text = ""
    tbDescFase.Text = ""
    
    RecuperarJueces Val(tbCodCat.Text), cbJuez
    Exit Sub
error:
    PPCLog ProcesarError("cmdSelCat_Click", False)
End Sub
Function RecuperarJueces(iCodCateg As Integer, cbJuez As ComboBox) As Integer
Dim rs As Recordset, i As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    ' Recuperamos los jueces
    cbJuez.Clear
    'Si el juez de pasos no examina todos los grupos examinables de la pista
    'obligatoriamente debe ser genérico
    'Si examina todos hasta cierta hora, en esa hora se activa laa opción
    If chkJPasosGen.Value = 1 Then
        Set rs = db.OpenRecordset("SELECT id_juez FROM juez_categ WHERE cod_categoria = " & iCodCateg & " ORDER BY 1", dbOpenSnapshot)
    Else
        Set rs = db.OpenRecordset("SELECT id_juez FROM juez_categ WHERE pasos = 0 and cod_categoria = " & iCodCateg & " ORDER BY 1", dbOpenSnapshot)
    End If
        cbJuez.Tag = 0
        i = 0
        While Not rs.EOF
            i = i + 1
            cbJuez.Tag = 1
            cbJuez.AddItem rs!id_juez
            rs.MoveNext
        Wend
    rs.Close
    cbJuez.Refresh
    RecuperarJueces = i
    cbJuez.ListIndex = -1
    Exit Function
error:
    PPCLog ProcesarError("RecuperarJueces", False)
End Function

Private Sub cmdSelComp_Click()
    
    If Not C_DEBUG Then On Local Error GoTo error
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCat.Text = ""
    tbDescCat.Text = ""
    tbCodFase.Text = ""
    tbDescFase.Text = ""

    Exit Sub
error:
    PPCLog ProcesarError("cmdSelComp_Click", False)
End Sub

Private Sub cmdSelFase_Click()

    If Not C_DEBUG Then On Local Error GoTo error
    If tbCodCat.Text = "" Then
        MsgBox mml_FRASE0504, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1")
    DescFase
    DoEvents
    CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
    Exit Sub
error:
    PPCLog ProcesarError("cmdSelFase_Click", False)
End Sub

Sub DescFase()
    tbDescFase.Text = sDescFase(Val(tbCodFase.Text))
End Sub

Sub GenerarFichero(sJuez As String, iCodCat As Integer, sDescCat As String, iCodFase As Integer, iCodRep As Integer, Optional iBaileIni As Integer = 0, Optional sBaile As String = "0")
Dim rs As Recordset
Dim iFile As Long
Dim bError As Boolean
Dim iCBailes As Integer
Dim bTeamMatch As Boolean
Dim iBaile As Integer


    If Not C_DEBUG Then On Local Error GoTo error
    
    'Comprobamos si hay que incrementar el nº de dorsales de cada tanda para dejar solo una en una semifinal
    If LimSemiUnaTanda > 0 Then
        If iCodFase = 2 Then
            If NumeroParejas(iCodCat, iCodFase, iCodRep) <= LimSemiUnaTanda Then
                AsignarDorsalesTanda iCodCat, LimSemiUnaTanda, True
            End If
        End If
    End If
    
    'Solo se genera el fichero si el juez pertenece al panel de la categoria
    If Not EsJuezDelPanel(iCodCat, sJuez) Then Exit Sub
    
    sJuez = Trim$(sJuez)
    
    iBaile = Val(sBaile)
    bTeamMatch = ComprobarSiTeamMatch(iCodCat)
    bError = False
    iFile = FreeFile
    Open m_sFicheroSalida & "." & sJuez & ".cTXT" For Output As #iFile
    Print #iFile, Trim$(Str$(iCodCat))
    Print #iFile, Trim$(Str$(iCodFase))
    Print #iFile, Trim$(Str$(IIf(iCodRep = 0, 0, 1)))
    Print #iFile, Trim$(Str$((iCodFase \ 2) * 6))
    If Len(sDescCat) > 5 Then
        Print #iFile, Left$(sDescCat, 1) & " " & Mid$(sDescCat, 5)
    Else
        Print #iFile, sDescCat
    End If
    'Informamos si es un TeamMatch
    If bTeamMatch Then
        Print #iFile, "TEAMMATCH"
    End If
    'Localizamos los dorsales
        'If G_DORSALES_COMBINADOS And CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
         '   Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & iRepesca & " AND d.fase =" & tbCodFase.Text & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
    
    If iCodFase > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(iCodCat) And G_PPC_GEN_DORSALES_COMBINADOS Then
    ' Si se generan dorsales combinados solo se envia un baile
        'Genera la información para la recombinación de dorsales
        
        Dim iMaxTandas As Integer, iTandasConMasDorsales As Integer, iTotalDorsales As Integer
        
        iMaxTandas = 0
        CalcularDorsalesPorTandaCatExt iCodCat, iCodFase, iCodRep, 1, iMaxTandas, iTandasConMasDorsales, iTotalDorsales
        If chkRecombinarAlgenerar.Value = 1 Then CombinarDorsales iCodCat, iCodFase, iCodRep, iMaxTandas, 0, True
        
        Print #iFile, "DORSALES_COMBINADOS"
        If iBaile = 0 Then ' Solo podemos generar un baile
            If cbBailes.ListCount > 1 Then
                If chkUltimos5Bailes.Value = 1 Then
                    If cbBailes.ListCount > 5 Then
                        iBaile = Val(cbBailes.List(6))
                    Else
                        MsgBox mml_FRASE1002, vbOKOnly Or vbCritical, mml_FRASE0096
                        Close #iFile
                        Exit Sub
                    End If
                Else
                    iBaile = Val(cbBailes.List(1))
                End If
            Else
                MsgBox mml_FRASE1002, vbOKOnly Or vbCritical, mml_FRASE0096
                Close #iFile
                Exit Sub
            End If
        End If
        Set rs = db.OpenRecordset("SELECT d.num_dorsal, dc.orden, no_presente, cod_baile FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & iCodCat & " AND d.repesca=" & iCodRep & " AND d.fase =" & iCodFase & " AND dc.cod_baile = " & iBaile & " ORDER BY orden, d.num_dorsal", dbOpenSnapshot)
        
        If Not rs.EOF Then
            rs.MoveLast
            If iBaile > 0 Then ' Solo generamos un baile
                'Imprimimos el número de dorsales
                Print #iFile, Trim$(Str$(rs.RecordCount))
            End If
            rs.MoveFirst
            While Not rs.EOF
                If iBaile > 0 Then ' Solo generamos un baile
                    If iBaile = rs!cod_Baile Then
                        Print #iFile, Trim$(Str$(rs!num_dorsal))
                    End If
                ElseIf iCBailes >= iBaileIni And iCBailes - iBaileIni < 5 Then
                    Print #iFile, Trim$(Str$(rs!num_dorsal))
                End If
                rs.MoveNext
            Wend
        Else
            MsgBox mml_FRASE0589 & sDescCat & mml_FRASE0590 & sDescFase(iCodFase), vbOKOnly Or vbCritical, mml_FRASE0096
            bError = True
        End If
    Else 'Imprimimos todos los dorsales ordenados por dorsal
        Set rs = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iCodRep & " ORDER BY num_dorsal", dbOpenSnapshot)
    
        If Not rs.EOF Then
            rs.MoveLast
            'Imprimimos el número de dorsales
            Print #iFile, Trim$(Str$(rs.RecordCount))
            rs.MoveFirst
            While Not rs.EOF
                Print #iFile, Trim$(Str$(rs!num_dorsal))
                rs.MoveNext
            Wend
        Else
            MsgBox mml_FRASE0589 & sDescCat & mml_FRASE0590 & sDescFase(iCodFase), vbOKOnly Or vbCritical, mml_FRASE0096
            bError = True
        End If
    End If
    rs.Close
    'Dorsales por tanda
    Print #iFile, iDorsalesPorTandaCateg(iCodCat)
    'Numero de bailes
    Set rs = db.OpenRecordset("SELECT b.codigo, b.nombre, bc.posicion FROM bailes_Categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & iCodCat & " AND fase = " & IIf(iCodFase = 1, 1, 2) & " ORDER BY posicion", dbOpenSnapshot)
    iCBailes = 0
    iBaileIni = IIf(iBaileIni > 0, 5, 0)
    If Not rs.EOF Then
        rs.MoveLast
        If iBaile > 0 Then
            Print #iFile, "1"
        Else
            Print #iFile, IIf(rs.RecordCount <= 5, Trim$(Str$(rs.RecordCount)), 5)
        End If
        rs.MoveFirst
        While Not rs.EOF
            If iBaile > 0 Then ' Solo generamos un baile
                If iBaile = rs!codigo Then
                    Print #iFile, sNombreBaileAbreviado(rs!Nombre)
                    Print #iFile, Trim$(Str$(rs!codigo))
                    iBaile = COD_MAX_BAILES
                End If
            ElseIf iCBailes >= iBaileIni And iCBailes - iBaileIni < 5 Then
                Print #iFile, sNombreBaileAbreviado(rs!Nombre)
                Print #iFile, Trim$(Str$(rs!codigo))
            End If
            Inc iCBailes
            rs.MoveNext
        Wend
        If iBaile <> COD_MAX_BAILES And iBaile > 0 Then
            MsgBox "Se ha intentado generar el único baile " & sBaile & " para la categoria " & sDescCategoria(iCodCat) & " pero la categoría no tiene ese baile.", vbCritical Or vbOKOnly, mml_FRASE0096
        End If
    Else
        MsgBox mml_FRASE0591 & sDescCat & mml_FRASE0590 & sDescFase(iCodFase), vbOKOnly Or vbCritical, mml_FRASE0096
        bError = True
    End If
    rs.Close
    'Imprimimos los dorsales no presentes
    Set rs = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iCodRep & " AND no_presente > 0  ORDER BY num_dorsal", dbOpenSnapshot)
    If Not rs.EOF Then
        Print #iFile, "NO_PRESENTES"
        rs.MoveLast
        Print #iFile, Trim$(Str$(rs.RecordCount))
        rs.MoveFirst
        While Not rs.EOF
            Print #iFile, Trim$(Str$(rs.Fields("num_dorsal")))
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    Close #iFile
    On Local Error Resume Next
    If bError Then
        Kill m_sFicheroSalida & "." & sJuez & ".cTXT"
    Else
        Kill m_sFicheroSalida & "." & sJuez & ".TXT"
        Name m_sFicheroSalida & "." & sJuez & ".cTXT" As m_sFicheroSalida & "." & sJuez & ".TXT"
    End If
    
    Exit Sub
error:
    PPCLog ProcesarError("GenerarFichero", False)
    On Local Error Resume Next
    Close #iFile
End Sub

Private Sub cmdSubirDatos_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    
    If Val(tbCodCatAct.Text) > 0 And Val(tbCodFaseAct.Text) > 0 Then
        tbCodComp.Text = VarCfg("horario_codcompeticion")
        tbCodCat.Text = tbCodCatAct.Text
        tbCodFase.Text = tbCodFaseAct.Text
        chkRep.Value = chkRepAct.Value
        tbDescCat.Text = tbDescCatAct.Text
        tbDescFase.Text = tbDescFaseAct.Text
        
        RecuperarJueces Val(tbCodCat.Text), cbJuez
        'CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
    End If
    Exit Sub
error:
    PPCLog ProcesarError("cmdSubirDatos_Click", False)
End Sub




Private Sub cmdSubirDatosSigFase_Click()

    If Not C_DEBUG Then On Local Error GoTo error
    
    If Val(tbCodCatSig.Text) > 0 And Val(tbCodFaseSig.Text) > 0 Then
        tbCodComp.Text = CodCompActiva
        tbCodCat.Text = tbCodCatSig.Text
        tbCodFase.Text = tbCodFaseSig.Text
        chkRep.Value = chkRepSig.Value
        tbDescCat.Text = tbDescCatSig.Text
        tbDescFase.Text = tbDescFaseSig.Text
        
        RecuperarJueces Val(tbCodCat.Text), cbJuez
    End If
    Exit Sub
error:
    PPCLog ProcesarError("cmdSubirDatosSigFase_Click", False)

End Sub

Private Sub Form_Load()

    If Not C_DEBUG Then On Local Error GoTo error
    
    m_sVersiónFicheroJueces = Trim$(Str$(Int(Rnd * 100000)))
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    IniciarDatosNoPresentes
    CargarPistas cbPista
    CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
    
    chkCalcAuto.Value = IIf(frmMenu.mnuGenAutoPPC.Checked, 1, 0)
    CargarPathFicheros
    
    If G_SOLO_UN_PC Then
        chkSoloUnPC.Value = 1
    Else
        chkSoloUnPC.Value = 0
    End If
    Exit Sub
error:
    ProcesarError "Form_Load"
End Sub
Sub CargarBailes(iCodCat As Integer, iCodFase As Integer, cbBailes As ComboBox)
Dim rs As Recordset
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    If iCodCat <> 0 And iCodFase <> 0 Then
        Set rs = db.OpenRecordset("SELECT codigo,nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE b.codigo = bc.cod_baile AND bc.cod_categoria = " & iCodCat & " AND bc.fase = " & IIf(iCodFase = 1, 1, 2) & " ORDER BY posicion", dbOpenSnapshot)
        cbBailes.Clear
        cbBailes.AddItem "0-Todos"
        While Not rs.EOF
            cbBailes.AddItem rs!codigo & "-" & rs!Nombre
            rs.MoveNext
        Wend
        rs.Close
        cbBailes.ListIndex = 0
    Else
        cbBailes.AddItem "0-Todos"
    End If
    cbBailes.Refresh
    Exit Sub
error:
    PPCLog ProcesarError("CargarBailes", False)
End Sub

Public Sub lblActTimer_Click()

    If Not C_DEBUG Then On Local Error GoTo error
    
    If lblActTimer.Tag = "0" Then
        If DistintaPista(cbPista.Text) Then
            lblActTimer.BackColor = vbGreen
            lblActTimer.Caption = mml_FRASE0592
            lblActTimer.Tag = "1"
            Timer1.Interval = G_INTERVALO_TIMER_PPC
            Timer1.Enabled = True
            tmrCalcular.Interval = G_TIEMPO_ESPERA_JPASOS_PPC
            tmrCalcular.Enabled = True
        Else
            MsgBox mml_FRASE1174, vbOKOnly Or vbCritical, G_MSG_ERROR
        End If
    Else
        lblActTimer.BackColor = vbRed
        lblActTimer.Caption = mml_FRASE0583
        lblActTimer.Tag = "0"
        Timer1.Enabled = False
        tmrCalcular.Enabled = False
    End If
    Exit Sub
error:
    PPCLog ProcesarError("lblActTimer_Click", False)
End Sub

Private Sub lblJuecesAct_DblClick()
    lblJuecesAct.Caption = ""
End Sub

Private Sub lblLock_Click()
    If lblLock.BackColor = vbRed Then
        frmPPC.Enabled = False
        lblLock.BackColor = vbGreen
    Else
        frmPPC.Enabled = True
        lblLock.BackColor = vbRed
    End If
End Sub

Private Sub tbCodCat_GotFocus()
    tbCodCat.SelStart = 0
    tbCodCat.SelLength = Len(tbCodCat.Text)
End Sub

Private Sub tbCodCat_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbCodCat_LostFocus()
Dim sCateg As String
    
    If Not C_DEBUG Then On Local Error GoTo error
    If Val(tbCodCat.Text) > 0 Then
        sCateg = sDescCategoria(tbCodCat.Text, tbCodComp.Text)
        If sCateg = "" Then
            tbCodCat.Text = ""
            tbDescCat.Text = ""
        Else
            tbDescCat.Text = sCateg
        End If
    Else
        tbCodCat.Text = ""
        tbDescCat.Text = ""
    End If
    Exit Sub
error:
    ProcesarError "tbCodCat_LostFocus"
End Sub


Private Sub tbCodCatAct_Change()
    If lblActTimer.Caption = mml_FRASE0592 Then tmrCalcular.Enabled = True

End Sub

Sub ActualizarListasFicheros()
Dim sFichero As String

    sFichero = Dir$(m_sPathOrigen & "\*.TXT")
    lstFichRxOrigen.Clear
    While sFichero <> ""
        lstFichRxOrigen.AddItem sFichero
        sFichero = Dir
    Wend
    sFichero = Dir$(m_sPathOrigen & ".*.cTXT")
    lstFichRxTmp.Clear
    While sFichero <> ""
        lstFichRxTmp.AddItem sFichero
        sFichero = Dir
    Wend
    
    
    'Localizamos los ficheros resultado de los PPC
    sFichero = Dir$(m_sFicheroEntrada & ".*.TXT")
    lstFichRx.Clear
    While sFichero <> ""
        lstFichRx.AddItem sFichero
        sFichero = Dir
    Wend
    sFichero = Dir$(m_sFicheroSalida & ".*.TXT")
    lstFichTx.Clear
    While sFichero <> ""
        lstFichTx.AddItem sFichero
        sFichero = Dir
    Wend
    sFichero = Dir$(m_sFicheroSalida & ".*.cTXT")
    lstFichTxTmp.Clear
    While sFichero <> ""
        lstFichTxTmp.AddItem sFichero
        sFichero = Dir
    Wend

End Sub

Private Sub Timer1_Timer()
Dim sFichero As String, sPath As String, rs As Recordset, sDir As String
Dim iCodCat As Integer, sDirFichas As String, sFich As String
Static iContadorBateria As Integer

    DoEvents: DoEvents: DoEvents
    
    If Not C_DEBUG Then On Error GoTo error
    Timers False
    
    On Local Error Resume Next
    'Localizamos los ficheros resultado de los PPC en los directorios originales
    ActualizarListasFicheros
    
    If G_MOVER_FICHEROS_PDA Then
        MoverFicherosPDAs VarCfg("dir_fichas"), G_RUTA_COPIA_FICH_PDA
    End If
    
    CopiarFicherosEntrada
    
    'Actualizamos los datos de control de bateria de los PDAs
    ControlBateria
    
    sDirFichas = VarCfg("dir_fichas") ' Base de directorio para la copia de seguridad
    
    'Debemos evitar los bloqueoa
    '1. Copiamos los ficheros a un directorio de seguridad antes de procesarlos
    '2. Movemos los ficheros al directorio de trabajo. Si no podemos mover el fichero no debemos procesarlo
    
    'Recuperamos los jueces actuales
    RecuperarJuecesActivos
    
    sFichero = Dir$(m_sFicheroEntrada & ".*.TXT")
    While sFichero <> ""
        'Procesamos el fichero
        iCodCat = RecuperarInfo(sExtraerPath(m_sFicheroEntrada) & "\" & sFichero)
        'Ahora debemos copiarlo a la carpeta correspondiente si lo he procesado
        If iCodCat > 0 Then
            Set rs = db.OpenRecordset("SELECT codigo, descripcion FROM categorias WHERE codigo = " & iCodCat, dbOpenSnapshot)
            If Not rs.EOF Then
                sDir = sDirFichas & "\TMP\" & rs!DESCRIPCION & "_" & rs!codigo
                On Local Error Resume Next
                MkDir sDir
                If Not C_DEBUG Then
                    On Local Error Resume Next
                Else
                    On Local Error GoTo 0
                End If
            End If
            rs.Close
            FileCopy sExtraerPath(m_sFicheroEntrada) & "\" & sFichero, sDir & "\" & sFichero
            If Err.Number <> 0 Then PPCLog ProcesarError("Timer1_FilesRead_FileCopy " & sFichero, False)
            Err.Clear
            Kill sExtraerPath(m_sFicheroEntrada) & "\" & sFichero
            If Err.Number <> 0 Then PPCLog ProcesarError("Timer1_FilesRead_Kill " & sFichero, False)
            Err.Clear
        End If
        sFichero = Dir
    Wend
    Timers True
    Exit Sub
error:
    PPCLog ProcesarError("Timer1_FilesRead", False)
    Timers True
End Sub
Sub RecuperarJuecesActivos()
Dim sFichero As String
Dim sFicheroJuez As String
Dim iFile As Integer
Dim sJuez As String
    On Local Error Resume Next
    sFichero = Dir$(m_sFicheroControlJueces & ".*.TXT")
    tbJuecesAct.Text = ""
    While sFichero <> ""
        iFile = FreeFile
        
        sFicheroJuez = sExtraerPath(m_sFicheroControlJueces) & "\" & sFichero
        EsperaGrabacionDeFichero sFicheroJuez, 100
        
        Open sFicheroJuez For Input As #iFile
        Line Input #iFile, sJuez
        sJuez = Left$(sJuez, 7)
        sJuez = CambiarCadena(" ", "_", sJuez)
        tbJuecesAct.Text = tbJuecesAct.Text & " " & sJuez
        Close #iFile
        sFichero = Dir
    Wend
    Exit Sub
error:
    PPCLog ProcesarError("RecuperarJuecesActivos", False)
    On Local Error Resume Next
    Close #iFile
End Sub
Function RecuperarInfo(sFichero As String) As Integer
Dim rs As Recordset
Dim iFile As Long
Dim sJuez As String, sCodCat As String, sCodFase As String
Dim sCodRep As String, sCodBaile As String, sNumDorsales As String
Dim i As Integer
Dim sEstado As String
Dim aDorsales() As String
Dim aPuestos() As String
Dim aDescalificados() As String
Dim sNumDescalificados As String
Dim sNumNoPresentes As String
Dim sCad As String
Dim bTeamMatch As Boolean
Dim rsBailes As Recordset
Dim lTam As Long
Dim iMaxBailes As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    RecuperarInfo = 0
    
    EsperaGrabacionDeFichero sFichero
    
    iFile = FreeFile
    Open sFichero For Input As #iFile
    Line Input #iFile, sJuez
    Line Input #iFile, sCodCat
    RecuperarInfo = Val(sCodCat)
    Line Input #iFile, sCodFase
    Line Input #iFile, sCodRep
    Line Input #iFile, sCodBaile
    Line Input #iFile, sNumDorsales
    ReDim aDorsales(Val(sNumDorsales)) As String
    ReDim aPuestos(Val(sNumDorsales)) As String
    
    
    'Si este fichero no corresponde a la categoria actual abandonamos su proceso, ya que puede corresponder a otra pista
    If G_NO_PROCESAR_FICH_DE_OTRA_CATEG Then
        If Val(sCodCat) <> Val(tbCodCatAct.Text) Or Val(sCodFase) <> Val(tbCodFaseAct.Text) Or Val(sCodRep) <> chkRepAct.Value Then
            RecuperarInfo = 0
            Close #iFile
            Exit Function
        End If
    End If
        
    'Actualizamos el panel de contadores de bailes
    iMaxBailes = 0
    For i = 0 To C_MAX_JUEZ_PANEL - 1
        If Trim$(tbJuezTx(i).Caption) = Trim$(sJuez) Then
            If Val(tbNumBAilesTx(i).Caption) < 90 Then
                tbNumBAilesTx(i).Caption = Val(tbNumBAilesTx(i).Caption) + 1
            End If
            
            If Val(tbNumBAilesTx(i).Caption) > iMaxBailes Then
                iMaxBailes = Val(tbNumBAilesTx(i).Caption)
            End If
        End If
    Next
    For i = 0 To C_MAX_JUEZ_PANEL - 1
        If Val(tbNumBAilesTx(i).Caption) = iMaxBailes Then
            tbJuezTx(i).BackColor = vbGreen
        Else
            tbJuezTx(i).BackColor = vbRed
        End If
        
    Next
    
    For i = 0 To Val(sNumDorsales) - 1
        Line Input #iFile, aDorsales(i)
        Line Input #iFile, aPuestos(i)
    Next
    Line Input #iFile, sEstado
    While sEstado = "DESC" Or sEstado = "NO_PRESENTES"
        Select Case sEstado
            Case "DESC"
                Line Input #iFile, sNumDescalificados
                ReDim aDescalificados(Val(sNumDescalificados)) As String
                For i = 0 To Val(sNumDescalificados) - 1
                    Line Input #iFile, aDescalificados(i)
                Next
            Case "NO_PRESENTES"
            Dim sTemp As String, bNoPresentesIguales As Boolean
            
                Line Input #iFile, sNumNoPresentes
                If g_iNoPresentesIniciado = C_NP_NO_INICIADO Or g_sNumNoPresentes <> sNumNoPresentes Then
                    ReDim aNoPresentes(Val(sNumNoPresentes)) As String
                    g_sNumNoPresentes = sNumNoPresentes
                    g_iNoPresentesIniciado = C_NP_INICIADO
                End If
                For i = 0 To Val(sNumNoPresentes) - 1
                    Line Input #iFile, sTemp
                    bNoPresentesIguales = True
                    If g_iNoPresentesIniciado = C_NP_PREPARADO Then
                        If aNoPresentes(i) <> sTemp Then
                            bNoPresentesIguales = False
                        End If
                    End If
                    aNoPresentes(i) = sTemp
                Next
                If Not bNoPresentesIguales Or Not g_iNoPresentesIniciado = C_NP_PREPARADO Then
                    If g_iCJueces < C_NUM_JUECES_ACEPTAR_NO_PRESENTES Then
                        g_iCJueces = 0
                    End If
                End If
                g_iNoPresentesIniciado = C_NP_PREPARADO
                If bNoPresentesIguales Then
                    Inc g_iCJueces
                End If
        End Select
        Line Input #iFile, sEstado
    Wend
    Close #iFile
    
    bTeamMatch = ComprobarSiTeamMatch(Val(sCodCat))
    
    'Comprobamos si los datos corresponden con la categ actual
    ' y el Juez no es de pasos genérico que puede enviar categoría en cualquier momento
    If sEstado <> "JPASOSGEN" And Not (sCodCat = tbCodCatAct.Text And sCodFase = tbCodFaseAct.Text And Val(sCodRep) = IIf(chkRepAct.Value = 1, 1, 0)) Then
    'Un juez ha transmitido una categoría que no es la actual
    'Comprobamos si el juez juzga la categoría actual
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE cod_categoria = " & sCodCat & " AND id_juez = '" & sJuez & "'", dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            rs.Close
            If G_CAMBIO_AUTO Then
                GoTo Cambiar
            End If
            If MsgBox("¿El juez " & sJuez & " ha transmitido una categoría que no es la actual, aceptarla como actual?", vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
Cambiar:
                tbCodCatAct.Text = sCodCat
                RecuperarInfo = Val(sCodCat)
                tbCodFaseAct.Text = sCodFase
                tbDescCatAct.Text = sDescCategoria(Val(sCodCat))
                tbDescFaseAct.Text = sDescFase(Val(sCodFase))
                chkRep.Value = Val(sCodRep)
                lblJuecesAct.Caption = ""
                tbNumJueces.Text = RecuperarJueces(Val(tbCodCatAct.Text), cbJuezAct)
                BorrarDatosFaseSig
                
                'CargarBailes Val(tbCodCatAct.Text), Val(tbCodFaseAct.Text), cbBailes
                IniciarDatosNoPresentes
            Else
                RecuperarInfo = 0
                Exit Function
            End If
        Else
            ' Si no juzga la categoría será de otra pista
            RecuperarInfo = 0
            rs.Close
            Exit Function
        End If
    End If
    
    'Si la categoria tiene dorsales combinados se estan enviando los bailes de uno en uno y
    ' debemos generar automaticamente el siguiente
    ' baile de la categoria y comprobar si el que nos envian es el ultimo para avanzar de categoria
    Dim sSQL As String
    Dim iPosicionBaile As Integer
    If sEstado = "DORSALES_COMBINADOS" Then
        'Localizamos la posición del baile actual
        sSQL = "SELECT bc.posicion FROM bailes_categ bc WHERE bc.cod_categoria = " & sCodCat & " AND bc.fase = " & IIf(sCodFase = "1", 1, 2) & " AND bc.cod_baile = " & sCodBaile & " ORDER BY bc.posicion"
        Set rsBailes = db.OpenRecordset(sSQL, dbOpenSnapshot)
        If Not rsBailes.EOF Then
            iPosicionBaile = rsBailes!posicion
        Else
            MsgBox mml_FRASE1005 & sDescCategoria(Val(sCodCat)), vbOKOnly Or vbCritical, G_MSG_ERROR
        End If
        sSQL = "SELECT TOP 1 bc.posicion, bc.cod_baile, b.nombre FROM bailes b, bailes_categ bc WHERE b.codigo = bc.cod_baile AND bc.cod_categoria = " & sCodCat & " AND bc.fase = " & IIf(sCodFase = "1", 1, 2) & " AND bc.posicion > " & iPosicionBaile & " ORDER BY bc.posicion"
        
        Set rsBailes = db.OpenRecordset(sSQL, dbOpenSnapshot)
        'Acabamos si no hay mas bailes o si terminamos con los 5 primeros bailes
        If rsBailes.EOF Then
            'Indicamos que el baile que acabamos de leer es el ultimo baile
            sEstado = "FIN"
        Else
            'Comprobamos si el siguiente es el sexto baile de la categoría, lo que implica que es el primero de la siguiente modalidad
            If NumeroBaileCateg(Val(sCodCat), Val(sCodFase), rsBailes.Fields("cod_baile")) = 6 Then
                sEstado = "FIN"
            Else
                'Generar Fichero actualizando versión
                GenerarFicheroJueces Val(sCodCat), False
                GenerarFichero sJuez, Val(sCodCat), sDescCategoria(Val(sCodCat)), Val(sCodFase), chkRep.Value, chkUltimos5Bailes.Value, rsBailes!cod_Baile
                'Informamos de la generación del siguiente baile
                lblJuecesAct.Caption = lblJuecesAct.Caption & Left$(rsBailes!Nombre, 1) & ">"
            End If
        End If
    End If
    
    'Comprobamos si es el ultimo baile del juez
    If sEstado = "FIN" Then ' FIN
        
        If tbCodCatSig.Text = "" Then
            'Calculamos la siguiente categoría sin recuperar GeneralLooks
            Dim iBaile As Integer
            iBaile = Val(cbBailes.List(cbBailes.ListIndex))
            Set rs = db.OpenRecordset("SELECT TOP 1 * FROM horario WHERE grupo LIKE '*" & cbPista.Text & "*' AND numfase <> 99 AND cod_competicion = " & VarCfg("horario_codcompeticion") & " AND orden > (SELECT TOP 1 orden FROM horario WHERE cod_baile " & IIf(chkUltimos5Bailes.Value, "<", "=") & iBaile & " AND cod_categoria = " & sCodCat & " and numfase = " & sCodFase & " AND repesca = " & sCodRep & " AND cod_competicion = " & VarCfg("horario_codcompeticion") & ") ORDER BY orden", dbOpenSnapshot)
            If Not rs.EOF Then
                tbDescCatSig.Text = sDescCategoria(rs!cod_categoria)
                tbCodCatSig.Text = rs!cod_categoria
                tbDescFaseSig.Text = sDescFase(rs!numfase)
                tbCodFaseSig.Text = rs!numfase
                chkRepSig.Value = rs!repesca
                
                lblJuecesSig.Text = ""
                'Generar Fichero actualizando versión
                GenerarFicheroJueces Val(tbCodCatSig.Text)
            
                CargarBailes Val(tbCodCatSig.Text), Val(tbCodFaseSig.Text), cbBailes
                If rs!cod_Baile > 0 Then
                    For i = 0 To cbBailes.ListCount - 1
                        If Val(cbBailes.List(i)) = rs!cod_Baile Then
                            cbBailes.ListIndex = i
                        End If
                    Next
                ElseIf rs!cod_Baile < 0 Then
                    chkUltimos5Bailes.Value = 1
                Else
                    If C_RESET_ULTIMOS_5_BAILES Then chkUltimos5Bailes.Value = 0
                End If
                
            Else
                'Finalizamos la competición
            
                'MsgBox mml_FRASE0595, vbOKOnly Or vbInformation, mml_FRASE0096
                'Paramos el timer
                'lblActTimer.Tag = "1"
                'lblActTimer_Click
            End If
            rs.Close
        End If
        
        'Si el juez no a transmitido antes las puntuaciones lo quitamos de los que
        'quedan por transmitir
        If InStr(lblJuecesSig.Text, sJuez) = 0 And Val(tbNumJueces.Text) > 0 Then
            lblJuecesSig.Text = lblJuecesSig.Text & sJuez & " "
            tbNumJueces.Text = Val(tbNumJueces.Text) - 1
        End If
        
        'Generamos el fichero
        If tbCodCatSig.Text <> "" Then
            Dim sDescCat As String
            sDescCat = sDescCategoria(Val(tbCodCatSig.Text))
            If chkGenSigCat.Value > 0 Then
                GenerarFichero sJuez, Val(tbCodCatSig.Text), sDescCat, Val(tbCodFaseSig.Text), chkRep.Value, chkUltimos5Bailes.Value, Val(cbBailes.List(cbBailes.ListIndex))
            End If
        End If
        
    End If
    'Grabamos el juez que ha transmitido el baile
    If sEstado <> "JPASOSGEN" Then
        lblJuecesAct.Caption = lblJuecesAct.Caption & sNombreBaileAbreviado(sNombreBaile(sCodBaile)) & "." & sJuez & " "
    End If
    
    'Grabar la información en la base de datos
    For i = 0 To Val(sNumDorsales) - 1
        db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal = " & aDorsales(i) & " AND cod_categoria = " & sCodCat & " AND cod_baile = " & sCodBaile & " AND cod_juez = '" & sJuez & "' AND fase = " & sCodFase & " AND repesca = " & sCodRep)
        If bTeamMatch Then
            db.Execute ("INSERT INTO puntuaciones VALUES (" & aDorsales(i) & "," & sCodCat & "," & sCodBaile & ",'" & sJuez & "','" & CDbl(aPuestos(i) + 1) / 2 & "'," & sCodFase & "," & sCodRep & ")")
        Else
            db.Execute ("INSERT INTO puntuaciones VALUES (" & aDorsales(i) & "," & sCodCat & "," & sCodBaile & ",'" & sJuez & "'," & aPuestos(i) & "," & sCodFase & "," & sCodRep & ")")
        End If
    Next
    
    'Grabamos la información de las descalificaciones
    For i = 0 To Val(sNumDescalificados) - 1
        'Si es un juez de pasos comprobamos que los datos están bien
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_Categoria = " & sCodCat & " AND fase = " & IIf(Val(sCodFase) > 1, 2, 1) & " AND cod_baile = " & sCodBaile, dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            rs.Close
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & sCodCat & " AND fase = " & sCodFase & " AND repesca = " & sCodRep & " AND num_dorsal = " & Mid$(aDescalificados(i), 1, 4) & " AND no_presente = 0", dbOpenSnapshot)
            If rs.Fields(0) > 0 Then
                db.Execute "DELETE FROM descalificaciones WHERE num_dorsal = " & Mid$(aDescalificados(i), 1, 4) & " AND cod_categoria = " & sCodCat & " AND cod_baile = " & sCodBaile & " AND id_juez = '" & sJuez & "' AND fase = " & sCodFase & " AND repesca = " & sCodRep
                sSQL = "INSERT INTO descalificaciones VALUES(" & MaxCod("descalificaciones") & "," & sCodCat & "," & sCodFase & ",'" & sJuez & "'," & sCodBaile & "," & Mid$(aDescalificados(i), 1, 4) & ",'" & Mid$(aDescalificados(i), 7) & "', NULL, " & sCodRep & ")"
                Debug.Print sSQL
                db.Execute sSQL
            End If
            rs.Close
        Else
            rs.Close
        End If
    Next
    
    'Grabamos la información de los no presentes
    If Not g_bNoPresentesProcesados And g_iCJueces >= C_NUM_JUECES_ACEPTAR_NO_PRESENTES Then
        If C_PREGUNTA_ACEPTAR_NO_PRESENTES Then
            sCad = ""
            For i = 0 To Val(sNumNoPresentes) - 1
                If sCad <> "" Then sCad = sCad & ", "
                sCad = sCad & aNoPresentes(i)
            Next
            If MsgBox("¿Acepta los dorsales: " & sCad & " como no presentes en pista?", vbYesNo Or vbQuestion, "") = vbNo Then
                GoTo continuar
            End If
        End If
        db.Execute ("UPDATE dorsales SET no_presente = 0 WHERE cod_categoria = " & sCodCat & " AND fase = " & sCodFase & " AND repesca = " & sCodRep)
        For i = 0 To Val(sNumNoPresentes) - 1
            db.Execute ("UPDATE dorsales SET no_presente = 1 WHERE num_dorsal = " & aNoPresentes(i) & " AND cod_categoria = " & sCodCat & " AND fase = " & sCodFase & " AND repesca = " & sCodRep)
        Next
        g_bNoPresentesProcesados = True
continuar:
    End If
    
    Dim bFinCateg As Boolean
    bFinCateg = False
    
    
    'Si todos los jueces han transmitido sus puntuacines avanzamos la categoría
    If sEstado <> "JPASOSGEN" And Val(tbNumJueces.Text) = 0 Then
    'Antes de avanzar de categoría comprobamos si hay que realizar un cálculo automático
        If frmMenu.mnuGenAutoPPC.Checked And chkCalcAuto.Value = 1 Then ComprobarPuntuaciones
        
        If Val(tbCodCatSig.Text) > 0 Then
            iCodCat = Val(tbCodCatAct.Text)
            iFase = Val(tbCodFaseAct.Text)
            iRepesca = chkRepAct.Value
            
            tbCodCatAct.Text = tbCodCatSig.Text
            tbDescCatAct.Text = tbDescCatSig.Text
            tbCodFaseAct.Text = tbCodFaseSig.Text
            tbDescFaseAct.Text = tbDescFaseSig.Text
            chkRepAct.Value = chkRepSig.Value
            
            cbJuezAct.Clear
            cbJuezAct.Text = ""
            tbNumJueces.Text = RecuperarJueces(Val(tbCodCatAct.Text), cbJuezAct)
            cbJuezAct.Refresh
            lblJuecesAct.Caption = ""
            'Con esto nos aseguramos que generamos todos los ficheros de todos los jueces aunque haya un cambio de panel
            For i = 0 To cbJuezAct.ListCount - 1
                GenerarFichero cbJuezAct.List(i), Val(tbCodCatAct.Text), tbDescCatAct.Text, Val(tbCodFaseAct.Text), chkRepAct.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
            Next
            For i = 0 To C_MAX_JUEZ_PANEL - 1
                If i <= cbJuezAct.ListCount - 1 Then
                    tbJuezTx(i).BackColor = vbYellow
                    tbJuezTx(i).Visible = True
                Else
                    tbJuezTx(i).Visible = False
                End If
            Next
            GenerarFicheroJueces Val(tbCodCatAct.Text), True, 0
            
            BorrarDatosFaseSig
            IniciarDatosNoPresentes
            
            tmrCalcular.Interval = G_TIEMPO_ESPERA_JPASOS_PPC
            tmrCalcular.Enabled = True
        Else
            'Finalizamos la competición
        
            MsgBox mml_FRASE0595, vbOKOnly Or vbInformation, mml_FRASE0096
            'Paramos el timer
            lblActTimer.Tag = "1"
            lblActTimer_Click
        End If
    End If
    
    Exit Function
error:
    PPCLog ProcesarError("RecuperarInfo", False)
    On Local Error Resume Next
    Close #iFile
End Function
Sub BorrarDatosFaseSig()
    cbJuezSig.Clear
    cbJuezSig.Text = ""
    tbCodCatSig.Text = ""
    tbCodFaseSig.Text = ""
    tbDescCatSig.Text = ""
    tbDescFaseSig.Text = ""
    chkRepSig.Value = 0
    lblJuecesSig.Text = ""
    cbJuezSig.Refresh
End Sub

Sub IniciarDatosNoPresentes()
    g_iNoPresentesIniciado = C_NP_NO_INICIADO
    g_iCJueces = 0
    g_bNoPresentesProcesados = False
End Sub

Private Sub ControlBateria()
Dim i As Integer
Dim sFichero As String
Dim sFicheroBat As String

    
    If Not C_DEBUG Then On Local Error GoTo error
    'Localizamos los ficheros resultado de los PPC
    sFichero = Dir$(m_sFicheroBateria & "*.TXT")
    i = 0
    While sFichero <> ""
        'Procesamos el fichero
        sFicheroBat = sExtraerPath(m_sFicheroBateria) & "\" & sFichero
        EsperaGrabacionDeFichero sFicheroBat, 100
        RecuperarInfoBateria sFicheroBat, i
        ' i = i + 1 se incremente dentro de RecuperarInfoBateria
        sFichero = Dir
    Wend
    For i = i To 12
        lblBat(i).Visible = False
    Next
    Exit Sub
error:
    PPCLog ProcesarError("tmrBateria_Timer", False)
    Timers True
End Sub

Sub RecuperarInfoBateria(sFich As String, iPos As Integer)
Dim iFile As Integer
Dim sCad As String
Dim sNombre As String
Dim sPorcentaje As String
Dim sHora As String
Dim sJuez As String
Dim i As Integer
Dim rs As Recordset
Dim iCargando As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    If iPos <= MAX_PDAS_BATERIAS Then
        iFile = FreeFile
        'Dejamos tiempo para que acabe de escribirse
        Open sFich For Input As #iFile
        Line Input #iFile, sNombre
        Line Input #iFile, sJuez
        Line Input #iFile, sPorcentaje
        Line Input #iFile, sHora
        Close #iFile
        
        lblBat(iPos).Caption = "(" & sJuez & ")" & Right$("   " & sPorcentaje, 3) & "%"
        If IsDate(sHora) Then
            lblBat(iPos).Caption = lblBat(iPos).Caption & vbCrLf & DateDiff("s", CDate(sHora), Time) & "s"
        Else
            sHora = ""
        End If
        
        If Abs(DateDiff("s", CDate(sHora), Time)) >= G_TIEMPO_PARA_PERDIDA_DE_CONEXION Then
            lblBat(iPos).BackColor = vbBlack
            lblBat(iPos).ForeColor = vbWhite
            'Si no hay conexión comprobamos si el juez pertenece a este panel
            If Val(tbCodCatAct.Text) > 0 Then
                ' Si no es juez del panel salimos corresponde a algún fichero no borrado
                If Not EsJuezDelPanel(tbCodCatAct.Text, sJuez) Then
                    Exit Sub
                End If
            End If
        Else
            lblBat(iPos).ForeColor = vbBlack
            If Val(sPorcentaje) > 30 Then
                lblBat(iPos).BackColor = vbGreen
            ElseIf Val(sPorcentaje) > 10 Then
                lblBat(iPos).BackColor = vbYellow
            Else
                lblBat(iPos).BackColor = vbRed
            End If
        End If
        lblBat(iPos).Visible = True
        
        If G_UNIDADES_NIVEL_MIN_CONTROL > 0 Then
            'Actualizamos la información de duración de bateria
            iCargando = 0
            If IsDate(sHora) And Val(sPorcentaje) > 0 And sJuez <> "" Then
                Set rs = db.OpenRecordset("SELECT * FROM bateria WHERE id_juez = '" & sJuez & "'", dbOpenSnapshot)
                If Not rs.EOF Then
                    'Si hay la marca de bateria comprobamos si es reciente
                    If CDate(rs.Fields("fecha")) = CDate(Format$(Now, "dd/mm/yyyy")) And DateDiff("s", CDate(rs.Fields("hora")), CDate(sHora)) < G_SEG_MAX_CONTROL_BATERIA Then
                        If rs.Fields("nivel") > Val(sPorcentaje) Then
                            If rs.Fields("cargando") = -1 Then
                                ' No se ha llegado a la marca de control, pero hemos pasado la marca de inicio de carga
                                'Iniciamos el sistema para comenzar a contar tiempos
                                db.Execute "UPDATE bateria SET nivel = " & Val(sPorcentaje) & ", fecha = '" & Format$(Now, "dd/mm/yyyy") & "', hora = '" & Format$(CDate(sHora), "hh:mm:ss") & "', tiempo_descarga = 0,cargando = 0 WHERE id_juez = '" & sJuez & "'"
                            ElseIf rs.Fields("nivel") = Val(sPorcentaje) + G_UNIDADES_NIVEL_MIN_CONTROL Then
                            'Si alcanzamos la marca de control y no tenemos grabado que todavía estamos en la marca inicial de control (cargando = -1)
                                Dim dUnidadesBateria As Double
                                Dim iSegUnidadBateria As Integer
                                
                                dUnidadesBateria = rs.Fields("nivel") - Val(sPorcentaje)
                                'Restamos el 5% por las imprecisiones
                                iSegUnidadBateria = (Val(sPorcentaje) / dUnidadesBateria * Abs(DateDiff("s", CDate(rs.Fields("hora")), CDate(sHora)))) * 0.95
                                db.Execute "UPDATE bateria SET nivel = " & Val(sPorcentaje) & ", fecha = '" & Format$(Now, "dd/mm/yyyy") & "', hora = '" & Format$(CDate(sHora), "hh:mm:ss") & "', tiempo_descarga = " & iSegUnidadBateria & ",cargando = 0 WHERE id_juez = '" & sJuez & "'"
                                iCargando = 0
                            ElseIf rs.Fields("nivel") > Val(sPorcentaje) + G_UNIDADES_NIVEL_MIN_CONTROL Then
                                ' Se ha superado la marca de control, colocamos el sistema en modo cálculo
                                db.Execute "UPDATE bateria SET nivel = " & Val(sPorcentaje) & ", fecha = '" & Format$(Now, "dd/mm/yyyy") & "', hora = '" & Format$(CDate(sHora), "hh:mm:ss") & "', tiempo_descarga = 0,cargando = -1 WHERE id_juez = '" & sJuez & "'"
                                iSegUnidadBateria = 0
                                iCargando = C_INI_CALCULO
                            Else
                                iSegUnidadBateria = rs.Fields("tiempo_descarga")
                                iCargando = rs.Fields("cargando")
                            End If
                        ElseIf rs.Fields("nivel") < Val(sPorcentaje) Then
                            'Cargando
                            db.Execute "UPDATE bateria SET nivel = " & Val(sPorcentaje) & ", fecha = '" & Format$(Now, "dd/mm/yyyy") & "', hora = '" & Format$(CDate(sHora), "hh:mm:ss") & "', cargando = 1 WHERE id_juez = '" & sJuez & "'"
                            iCargando = C_CARGANDO
                        Else
                            iSegUnidadBateria = rs.Fields("tiempo_descarga")
                            iCargando = rs.Fields("cargando")
                        End If
                        
                        'Actualizar información
                        If iSegUnidadBateria > 0 Then
                            lblBat(iPos).Caption = lblBat(iPos).Caption & vbCrLf & IIf(iCargando = C_CARGANDO, "(+) ", "(-) ") & iSegUnidadBateria \ 60 & "'"
                        ElseIf iCargando = C_INI_CALCULO Then
                            lblBat(iPos).Caption = lblBat(iPos).Caption & vbCrLf & mml_FRASE1215
                        Else
                            lblBat(iPos).Caption = lblBat(iPos).Caption & vbCrLf & mml_FRASE1214
                        End If
                    Else
                        db.Execute "UPDATE bateria SET nivel = " & Val(sPorcentaje) & ", fecha = '" & Format$(Now, "dd/mm/yyyy") & "', hora = '" & Format$(CDate(sHora), "hh:mm:ss") & "', tiempo_descarga = 0, cargando = -1 WHERE id_juez = '" & sJuez & "'"
                        lblBat(iPos).Caption = lblBat(iPos).Caption & vbCrLf & mml_FRASE1214 ' Calculando
                        PPCLog "Time of control exceeded " & G_SEG_MAX_CONTROL_BATERIA & """"
                    End If
                Else
                    'Si no hay marca de control de bateria para este juez, la creamos
                    db.Execute "INSERT INTO bateria VALUES ('" & sJuez & "'," & Val(sPorcentaje) & ",'" & Format$(Now, "dd/mm/yyyy") & "','" & Format$(CDate(sHora), "hh:mm:ss") & "',0,-1)"
                    lblBat(iPos).Caption = lblBat(iPos).Caption & vbCrLf & mml_FRASE1214 ' Calculando
                End If
                rs.Close
            End If
        End If
        
        iPos = iPos + 1
    End If
    Exit Sub
error:
    PPCLog ProcesarError("RecuperarInfoBateria", False)
    On Local Error Resume Next
    Close #iFile
End Sub

Private Sub tmrCalcular_Timer()
    'ComprobarPuntuaciones
End Sub

Sub ComprobarPuntuaciones()
    If Not C_DEBUG Then On Local Error GoTo error

    'Comprobamos si ya están todas las puntuaciones
    If G_CALCULO_AUTO_PPC Then
        If Val(tbCodCatAct.Text) = 0 Or Val(tbCodFaseAct.Text) = 0 Then Exit Sub
        
        If ComprobarSiEstanTodasPuntuaciones(Val(tbCodCatAct.Text), chkRepAct.Value, tbCodFaseAct.Text) Then
            tmrCalcular.Enabled = False
            If G_GEN_AUTO_RESULTADOS_PPC Then
                While frmCalcular.lblAutoPPC.BackColor = vbRed And lblActTimer.BackColor <> vbRed
                    Sleep 100
                    DoEvents
                Wend
                ' Mientras el control automático siga activo
                If lblActTimer.BackColor <> vbRed Then
                    frmCalcular.lblAutoPPC.BackColor = vbRed
                    frmCalcular.tbCodComp.Text = tbCodComp.Text
                    frmCalcular.tbDescComp.Text = tbDescComp.Text
                    frmCalcular.tbCodCat.Text = Val(tbCodCatAct.Text)
                    frmCalcular.tbDescCat.Text = sDescCategoria(Val(tbCodCatAct.Text))
                    frmCalcular.tbCodFase.Text = Val(tbCodFaseAct.Text)
                    frmCalcular.tbDescFase.Text = sDescFase(Val(tbCodFaseAct.Text))
                    frmCalcular.chkRep.Value = chkRepAct.Value
                    
                    frmCalcular.Visible = True
                    DoEvents
                    frmCalcular.cmdCalcular_Click
                    frmCalcular.Visible = False
                    frmCalcular.lblAutoPPC.BackColor = vbGreen
                    frmCalcular.lblAutoPPC.Refresh
                End If
            ElseIf MsgBox(mml_FRASE0445, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                frmCalcular.lblAutoPPC.BackColor = vbRed
                frmCalcular.tbCodComp.Text = tbCodComp.Text
                frmCalcular.tbDescComp.Text = tbDescComp.Text
                frmCalcular.tbCodCat.Text = Val(tbCodCatAct.Text)
                frmCalcular.tbDescCat.Text = sDescCategoria(Val(tbCodCatAct.Text))
                frmCalcular.tbCodFase.Text = Val(tbCodFaseAct.Text)
                frmCalcular.tbDescFase.Text = sDescFase(Val(tbCodFaseAct.Text))
                frmCalcular.chkRep.Value = chkRepAct.Value
                
                frmCalcular.Show vbModal
                frmCalcular.lblAutoPPC.BackColor = vbGreen
            End If
        End If
    End If
    Exit Sub

error:
    PPCLog ProcesarError("tmrCalcular_Timer", False)
End Sub



Sub PPCLog(sCad As String)
    If Len(tbLog.Text) + Len(sCad) > 30000 Then
        tbLog.Text = ""
    End If
    If sCad <> "" Then
        tbLog.Text = sCad & vbCrLf & tbLog.Text
    End If
End Sub

Sub CopiarFicherosEntrada()
Dim sFile As String
Dim sDirFichas As String
Dim sDirCopia As String

    lstError.Clear
    If Not C_DEBUG Then On Local Error GoTo error
    sDirFichas = sExtraerPath(m_sFicheroEntrada)
    sDirCopia = VarCfg("dir_fichas") & "\COPIA_PDA"
    
    sFile = Dir$(m_sFicheroEntrada & ".*.TXT")
    While sFile <> ""
        Err.Clear
        On Local Error Resume Next
        FileCopy sDirFichas & "\" & sFile, sDirCopia & "\" & sFile
        If Err.Number <> 0 Then
            PPCLog ProcesarError("CopiarFicherosEntrada: " & sFile, False)
            lstError.AddItem sFile
        End If
        sFile = Dir
    Wend
    Exit Sub
error:
    PPCLog ProcesarError("CopiarFicherosEntrada", False)
End Sub

Sub MoverFicherosPDAs(sOrigen As String, sDestino As String)
Dim sFile As String
Dim sFichHora As String
    If Not C_DEBUG Then On Local Error GoTo error
    
    sFichHora = sExtraerFichero(m_sFicheroHora)
    'Renombramos ficheros .TXT
    sFile = Dir$(sOrigen & "\*.TXT")
    While Not sFile = ""
        'No debemos transferir ficheros de sincronización de hora
        If sFile <> sFichHora & ".TXT" Then
            On Local Error Resume Next
            Name sOrigen & "\" & sFile As sOrigen & "\" & sFile & ".TMP"
        End If
        sFile = Dir
    Wend
    
    'Transferimos los ficheros renombrados
    sFile = Dir$(sOrigen & "\*.TMP")
    While Not sFile = ""
        On Local Error Resume Next
        Err.Clear
        FileCopy sOrigen & "\" & sFile, sDestino & "\" & Mid$(sFile, 1, Len(sFile) - 4)
        'Comprobamos si el archivo destino existe y está bloqueado por el PC
        If Err.Number = 0 Then
            Err.Clear
            Kill sOrigen & "\" & sFile
            If Err.Number <> 0 Then
                PPCLog ProcesarError("MoverFicherosPDAs_Kill " & sOrigen & "\" & sFile, False)
            End If
        End If
        sFile = Dir
    Wend
    Exit Sub
error:
        PPCLog ProcesarError("MoverFicherosPDAs", False)
End Sub
