VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmADorsales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0028"
   ClientHeight    =   9900
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   13905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdJuecesBailes 
      Caption         =   "mml_FRASE0051"
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
      Left            =   6090
      MaskColor       =   &H80000010&
      TabIndex        =   70
      Top             =   9390
      Width           =   4140
   End
   Begin VB.CommandButton cmdAusencia 
      Caption         =   "mml_FRASE1231"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9120
      TabIndex        =   69
      Top             =   6540
      Width           =   2655
   End
   Begin VB.CommandButton cmdPresente 
      Caption         =   "mml_FRASE0620"
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
      Left            =   12600
      TabIndex        =   68
      Top             =   7020
      Width           =   1275
   End
   Begin VB.CommandButton cmdCambiarPareja 
      Caption         =   "mml_FRASE1165"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1830
      TabIndex        =   67
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton cmdImpNumDorsal 
      Caption         =   "mml_FRASE1206"
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
      Left            =   6090
      MaskColor       =   &H80000010&
      TabIndex        =   66
      Top             =   7980
      Width           =   4140
   End
   Begin VB.CommandButton cmdBuscarParticipantes 
      Caption         =   "mml_FRASE1199"
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
      Left            =   10350
      MaskColor       =   &H80000010&
      TabIndex        =   65
      Top             =   7500
      Width           =   3495
   End
   Begin VB.CheckBox chkAgregarUltima 
      Caption         =   "mml_FRASE1189"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   570
      TabIndex        =   64
      Top             =   1500
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddDorsales 
      Caption         =   "GenerarDorsales"
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
      Height          =   405
      Left            =   180
      MaskColor       =   &H80000010&
      TabIndex        =   63
      Top             =   10080
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.CommandButton cmdTandas 
      Caption         =   "mml_FRASE0600"
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
      Left            =   6090
      MaskColor       =   &H80000010&
      TabIndex        =   62
      Top             =   8460
      Width           =   4140
   End
   Begin VB.CommandButton cmdAddCategAlHorario 
      Caption         =   "mml_FRASE1156"
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
      Left            =   6090
      MaskColor       =   &H80000010&
      TabIndex        =   61
      Top             =   8940
      Width           =   4140
   End
   Begin VB.CommandButton cmdImpHojasPuntuaciones 
      Caption         =   "mml_FRASE0463"
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
      Left            =   10350
      MaskColor       =   &H80000010&
      TabIndex        =   60
      Top             =   7980
      Width           =   3495
   End
   Begin VB.CommandButton cmdHayPuntuaciones 
      Caption         =   "mml_FRASE1128"
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
      Left            =   2880
      MaskColor       =   &H80000010&
      TabIndex        =   59
      Top             =   8940
      Width           =   3120
   End
   Begin VB.CommandButton cmdComprobarRecogida 
      Caption         =   "mml_FRASE1117"
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
      Left            =   10350
      MaskColor       =   &H80000010&
      TabIndex        =   57
      Top             =   8460
      Width           =   3495
   End
   Begin VB.CommandButton cmdImportarDorsalesProBaile 
      Caption         =   "mml_FRASE1053"
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
      Left            =   2880
      MaskColor       =   &H80000010&
      TabIndex        =   56
      Top             =   7980
      Width           =   3120
   End
   Begin VB.CommandButton cmdCombinarResultados 
      Caption         =   "mml_FRASE1038"
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
      Left            =   2880
      MaskColor       =   &H80000010&
      TabIndex        =   55
      Top             =   8460
      Width           =   3120
   End
   Begin VB.CommandButton spDorsal_SpinUp 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5430
      TabIndex        =   54
      Top             =   1350
      Width           =   225
   End
   Begin VB.CommandButton spDorsal_SpinDown 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5430
      TabIndex        =   53
      Top             =   1110
      Width           =   225
   End
   Begin VB.CommandButton cmdCateg 
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
      Left            =   2430
      Picture         =   "frmADorsales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   570
      Width           =   495
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
      Height          =   405
      Left            =   2430
      Picture         =   "frmADorsales.frx":046A
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   90
      Width           =   495
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
      Height          =   675
      Left            =   135
      TabIndex        =   51
      Top             =   8955
      Width           =   2655
   End
   Begin VB.CommandButton cmdNoPresente 
      Caption         =   "mml_FRASE0298"
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
      Left            =   11820
      TabIndex        =   50
      Top             =   6540
      Width           =   2055
   End
   Begin VB.CommandButton cmdCategAct 
      Height          =   405
      Left            =   12420
      Picture         =   "frmADorsales.frx":08D4
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "mml_FRASE0428"
      Top             =   90
      Width           =   525
   End
   Begin VB.Frame Frame3 
      Caption         =   "mml_FRASE0034"
      Height          =   915
      Left            =   0
      TabIndex        =   47
      Top             =   6570
      Width           =   735
      Begin VB.TextBox tbParejasReales 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Height          =   420
         Left            =   60
         TabIndex        =   48
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSigFase 
      Caption         =   "mml_FRASE0430"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9780
      TabIndex        =   46
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdCambiarDesCat 
      Caption         =   "mml_FRASE0252"
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
      Left            =   11400
      TabIndex        =   45
      Top             =   540
      Width           =   1245
   End
   Begin VB.CommandButton cmdNuevaDesc 
      Caption         =   "mml_FRASE0253"
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
      Left            =   10320
      TabIndex        =   44
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton cmdBorrarCateg 
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
      Left            =   12660
      Picture         =   "frmADorsales.frx":0BDE
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   540
      Width           =   330
   End
   Begin VB.ComboBox cbOrden 
      Height          =   315
      ItemData        =   "frmADorsales.frx":0F94
      Left            =   12360
      List            =   "frmADorsales.frx":0FAD
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   1320
      Width           =   1290
   End
   Begin VB.CommandButton cmdParejas 
      Caption         =   "mml_FRASE0004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      TabIndex        =   40
      Top             =   8415
      Width           =   2655
   End
   Begin VB.ComboBox cbAsigPista 
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
      ItemData        =   "frmADorsales.frx":0FD9
      Left            =   13020
      List            =   "frmADorsales.frx":0FFB
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   540
      Width           =   735
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "mml_FRASE0007"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1830
      TabIndex        =   37
      Top             =   1515
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenTeamMatch 
      Caption         =   "mml_FRASE0003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10350
      MaskColor       =   &H80000010&
      TabIndex        =   36
      Top             =   8940
      Width           =   3495
   End
   Begin VB.CommandButton cmdRegenHorario 
      Caption         =   "mml_FRASE0279"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1020
      TabIndex        =   35
      Top             =   7875
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   90
      Top             =   540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdComprobarFases 
      Caption         =   "mml_FRASE0280"
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
      Left            =   6090
      TabIndex        =   34
      Top             =   7500
      Width           =   4140
   End
   Begin VB.CommandButton cmdRegenHoras 
      Caption         =   "mml_FRASE0281"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   120
      TabIndex        =   33
      Top             =   7875
      Width           =   885
   End
   Begin VB.CommandButton cmdGenHorarioHTML 
      Caption         =   "mml_FRASE0282"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2025
      TabIndex        =   32
      Top             =   7875
      Width           =   750
   End
   Begin VB.CommandButton cmdCambiarDorsal 
      Caption         =   "mml_FRASE0283"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   6900
      TabIndex        =   30
      Top             =   6510
      Width           =   2085
   End
   Begin VB.CommandButton cmdHorario 
      Caption         =   "mml_FRASE0284"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   29
      Top             =   7515
      Width           =   2670
   End
   Begin VB.CommandButton cmdBorrarPuntuaciones 
      Caption         =   "mml_FRASE0285"
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
      Left            =   2880
      MaskColor       =   &H80000010&
      TabIndex        =   28
      Top             =   7500
      Width           =   3120
   End
   Begin VB.Frame Frame2 
      Caption         =   "mml_FRASE0008"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   780
      TabIndex        =   23
      Top             =   6510
      Width           =   2055
      Begin VB.TextBox tbParejasAdicionales 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Height          =   420
         Left            =   720
         TabIndex        =   31
         Top             =   450
         Width           =   615
      End
      Begin VB.TextBox tbParejas 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Height          =   420
         Left            =   60
         TabIndex        =   25
         Top             =   450
         Width           =   615
      End
      Begin VB.TextBox tbParejasEspeciales 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
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
         Height          =   420
         Left            =   1380
         TabIndex        =   24
         Top             =   450
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0009"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "mml_FRASE0010"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   26
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdGenGrupos 
      Caption         =   "mml_FRASE0288"
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
      Left            =   2880
      TabIndex        =   22
      Top             =   7050
      Width           =   3120
   End
   Begin VB.ComboBox cbOrd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmADorsales.frx":1037
      Left            =   3165
      List            =   "frmADorsales.frx":1041
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1260
      Width           =   1380
   End
   Begin VB.CheckBox cbRepesca 
      Caption         =   "mml_FRASE0289"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10860
      TabIndex        =   20
      Top             =   1020
      Width           =   1395
   End
   Begin VB.CommandButton cmdCambCat 
      Caption         =   "mml_FRASE0290"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   2880
      TabIndex        =   19
      Top             =   6510
      Width           =   2055
   End
   Begin VB.CommandButton cmdCambiarFase 
      Caption         =   "mml_FRASE0291"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   4980
      TabIndex        =   18
      Top             =   6510
      Width           =   1875
   End
   Begin VB.CommandButton cmdRegenerarDorsales 
      Caption         =   "mml_FRASE0292"
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
      Left            =   8595
      TabIndex        =   17
      Top             =   7050
      Width           =   2745
   End
   Begin VB.CommandButton cmdImprimirDorsales 
      Caption         =   "mml_FRASE0293"
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
      Left            =   6090
      TabIndex        =   16
      Top             =   7050
      Width           =   2445
   End
   Begin VB.ComboBox cbFase 
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
      ItemData        =   "frmADorsales.frx":105A
      Left            =   6900
      List            =   "frmADorsales.frx":107F
      TabIndex        =   3
      Text            =   "mml_FRASE0294"
      Top             =   1140
      Width           =   2835
   End
   Begin VB.TextBox tbDorsal 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4710
      TabIndex        =   13
      Text            =   "1"
      Top             =   1110
      Width           =   735
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "mml_FRASE0295"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10860
      TabIndex        =   4
      Top             =   1320
      Width           =   1485
   End
   Begin VB.TextBox tbDescCateg 
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
      Left            =   3750
      MaxLength       =   35
      TabIndex        =   11
      Top             =   570
      Width           =   6510
   End
   Begin VB.TextBox tbCodCateg 
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
      Left            =   2910
      MaxLength       =   5
      TabIndex        =   2
      Top             =   570
      Width           =   855
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
      Left            =   2910
      TabIndex        =   9
      Top             =   90
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
      Left            =   3750
      TabIndex        =   8
      Top             =   90
      Width           =   8505
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "mml_FRASE0296"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1830
      TabIndex        =   7
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton cmdPoner 
      Caption         =   "mml_FRASE0297"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   510
      TabIndex        =   6
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0028"
      Height          =   4785
      Left            =   45
      TabIndex        =   0
      Top             =   1710
      Width           =   13695
      Begin MSDataGridLib.DataGrid dgDorsales 
         Bindings        =   "frmADorsales.frx":1119
         Height          =   4545
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   13515
         _ExtentX        =   23839
         _ExtentY        =   8017
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
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
   Begin MSAdodcLib.Adodc adoDorsales 
      Height          =   495
      Left            =   390
      Top             =   -30
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      RecordSource    =   $"frmADorsales.frx":1133
      Caption         =   "mml_FRASE0028"
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
   Begin VB.CheckBox chkPresentes 
      Caption         =   "mml_FRASE1126"
      Height          =   345
      Left            =   11400
      TabIndex        =   58
      Top             =   7080
      Width           =   1185
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "mml_FRASE0006"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12360
      TabIndex        =   42
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "mml_FRASE0002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13020
      TabIndex        =   39
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label2 
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
      Left            =   5880
      TabIndex        =   15
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "mml_FRASE0300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3270
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label8 
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
      Left            =   750
      TabIndex        =   12
      Top             =   570
      Width           =   1575
   End
   Begin VB.Label Label5 
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
      Left            =   750
      TabIndex        =   10
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "frmADorsales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBorrar_Click()
End Sub

Private Sub cbFase_Click()
    Call cmdActualizar_Click
End Sub

Private Sub cbFase_KeyPress(KeyAscii As Integer)
    SelecCBFase cbFase, KeyAscii
End Sub

Private Sub cbOrd_Click()
    If cbOrd.ListIndex = 1 Then
        tbDorsal.Text = Int(Val(tbDorsal.Text)) Mod iMinDorsalOficial(tbCodComp.Text)
    ElseIf Val(tbDorsal.Text) < iMinDorsalOficial(tbCodComp.Text) Then
        tbDorsal.Text = Val(tbDorsal.Text) + iMinDorsalOficial(tbCodComp.Text)
    End If

End Sub

Private Sub cbRepesca_Click()
    Call cmdActualizar_Click

End Sub

Sub cmdActualizar_Click()
Dim rs As Recordset
    On Local Error GoTo error
    
    If tbCodCateg.Text <> "" Then
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        Sleep 500
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        adoDorsales.ConnectionString = "DSN=Escrutinio"
        adoDorsales.RecordSource = "SELECT d.codigo, num_dorsal, cod_categoria, fase, repesca, no_presente, grupoedad, combinar_edad as no_combinar, nombre_hombre, nombre_mujer, p.codigo, p.aebdc_codigo  FROM dorsales d, parejas p WHERE cod_pareja = p.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase LIKE '" & Trim$(Mid$(cbFase.Text, 1, 3)) & "' AND repesca=" & cbRepesca.Value & "  ORDER BY " & cbOrden.ItemData(cbOrden.ListIndex) \ 10
        adoDorsales.Refresh
        
        If Not adoDorsales.Recordset.EOF Then
            adoDorsales.Recordset.MoveLast
            tbParejas.Text = adoDorsales.Recordset.RecordCount
            adoDorsales.Recordset.MoveFirst
        Else
            tbParejas.Text = "0"
        End If
        'Parejas especiales
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE num_dorsal < " & iMinDorsalOficial(tbCodComp.Text) & " AND cod_categoria = " & tbCodCateg.Text & " AND fase Like '" & Trim$(Mid$(cbFase.Text, 1, 3)) & "' AND repesca=" & cbRepesca.Value, dbOpenSnapshot)
            tbParejasEspeciales.Text = rs.Fields(0)
        rs.Close
        'Parejas adicionales
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND p.pareja_adicional = 1 AND d.cod_categoria = " & tbCodCateg.Text & " AND fase Like '" & Trim$(Mid$(cbFase.Text, 1, 3)) & "' AND repesca=" & cbRepesca.Value, dbOpenSnapshot)
            tbParejasAdicionales.Text = rs.Fields(0)
        rs.Close
    ElseIf tbCodComp.Text <> "" Then
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        Sleep 500
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        adoDorsales.ConnectionString = "DSN=Escrutinio"
        adoDorsales.RecordSource = "SELECT d.codigo, num_dorsal, cod_categoria, fase, repesca, no_presente, grupoedad, combinar_edad as no_combinar, nombre_hombre, nombre_mujer, p.codigo  FROM dorsales d, parejas p WHERE cod_pareja = p.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " ) ORDER BY " & cbOrden.ItemData(cbOrden.ListIndex) \ 10
        adoDorsales.Refresh
        cbFase.ListIndex = 0
        
        If Not adoDorsales.Recordset.EOF Then
            adoDorsales.Recordset.MoveLast
            tbParejas.Text = adoDorsales.Recordset.RecordCount
            adoDorsales.Recordset.MoveFirst
        Else
            tbParejas.Text = "0"
        End If
        'Parejas especiales
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE num_dorsal < " & iMinDorsalOficial(tbCodComp.Text) & " AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " )", dbOpenSnapshot)
            tbParejasEspeciales.Text = rs.Fields(0)
        rs.Close
        'Parejas adicionales
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d, parejas p  WHERE d.cod_pareja = p.codigo AND p.pareja_adicional = 1 AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " )", dbOpenSnapshot)
            tbParejasAdicionales.Text = rs.Fields(0)
        rs.Close
        'Parejas distintas reales
        Set rs = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre FROM dorsales d, parejas p  WHERE d.cod_pareja = p.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " ) ", dbOpenSnapshot)
            If Not rs.EOF Then
                rs.MoveLast
                tbParejasReales.Text = rs.RecordCount
            Else
                tbParejasReales.Text = "0"
            End If
        rs.Close
    End If
    dgDorsales.Columns(0).Width = 800
    dgDorsales.Columns(1).Width = 800
    dgDorsales.Columns(2).Width = 800
    dgDorsales.Columns(3).Width = 800
    dgDorsales.Columns(4).Width = 800
    dgDorsales.Columns(5).Width = 800
    dgDorsales.Columns(6).Width = 1000
    dgDorsales.Columns(7).Width = 800
    dgDorsales.Columns(8).Width = 2000
    dgDorsales.Columns(9).Width = 2000
    Exit Sub
error:
    ProcesarError
End Sub

Private Sub cmdAddCategAlHorario_Click()
Dim rs As Recordset
Dim lOrden As Long
    If Not C_DEBUG Then On Local Error GoTo error
    If tbCodComp.Text = "" Or tbCodCateg.Text = "" Or Val(cbFase.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE1157, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
        Set rs = db.OpenRecordset("SELECT MAX(orden) FROM horario WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
        If rs.Fields(0) > 0 And rs.Fields(0) < MAX_NUMERO Then
            lOrden = rs.Fields(0) + 10
        Else
            lOrden = 10
        End If
        
        rs.Close
        db.Execute "INSERT INTO horario VALUES ('10:00:00','" & tbDescCateg.Text & "','" & sDescFase(Val(cbFase.Text)) & "'," & Val(cbFase.Text) & "," & tbCodCateg.Text & "," & cbRepesca.Value & "," & lOrden & ",0," & tbCodComp.Text & ",0," & Val(tbParejas.Text) & ",0,1)"
        
        MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    End If
    Exit Sub
error:
    ProcesarError "cmdAddCategAlHorario_Click"
    
End Sub

Private Sub cmdAddDorsales_Click()
Dim iNumDorsales As Integer
Dim iNumDorsal As Integer
Dim rs As Recordset

    If Val(tbCodCateg.Text) = 0 Or Val(cbFase.Text) = 0 Then
        CamposSinCubrir
    End If
    
    iNumDorsales = InputBox("Seleccione el número de parejas de la competición que quiere añadir", "", "10")
    
    If iNumDorsales > 0 Then
        Set rs = db.OpenRecordset("SELECT TOP " & iNumDorsales & " * FROM parejas WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
        iNumDorsal = 1
        While Not rs.EOF
            db.Execute "INSERT INTO dorsales VALUES(" & MaxCod("dorsales") + 1 & "," & iNumDorsal & "," & tbCodCateg.Text & "," & Val(cbFase.List(cbFase.ListIndex)) & "," & rs.Fields("codigo") & ",0,0)"
            Inc iNumDorsal
            rs.MoveNext
        Wend
    End If
    rs.Close
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
End Sub

Private Sub cmdBorrarCateg_Click()
Dim rs As Recordset
    MsgBox mml_FRASE0302, vbOKOnly Or vbCritical, mml_FRASE0084
    If tbCodCateg.Text = "" Then
        MsgBox mml_FRASE0303, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text)
    If rs.Fields(0) > 0 Then
        MsgBox mml_FRASE0304, vbOKOnly Or vbExclamation, mml_FRASE0096
        rs.Close
        Exit Sub
    Else
        MsgBox mml_FRASE0305, vbOKOnly Or vbCritical, mml_FRASE0084
        If MsgBox(mml_FRASE0306, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
            db.Execute ("DELETE FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text)
            db.Execute ("DELETE FROM juez_categ WHERE cod_categoria = " & tbCodCateg.Text)
            db.Execute ("DELETE FROM categorias WHERE codigo = " & tbCodCateg.Text)
        End If
    End If
    rs.Close
End Sub

Private Sub cmdBorrarPuntuaciones_Click()
Dim sFase As String

    If Not C_DEBUG Then On Local Error GoTo error

    sFase = Val(Mid$(cbFase.Text, 1, 3))
        
    If MsgBox(mml_FRASE0307, vbYesNo Or vbCritical, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    MsgBox mml_FRASE0308, vbOKOnly Or vbInformation, mml_FRASE0084
    
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0311, vbOKOnly Or vbExclamation, mml_FRASE0096
        Exit Sub
    End If
    
    If tbCodCateg.Text = "" Then
        If InputBox(mml_FRASE0309, vbYesNo Or vbCritical, "") = "AUTORIZO" Then
            BorrarPuntuaciones Val(tbCodComp.Text)
            
            MsgBox mml_FRASE0276, vbOKOnly Or vbExclamation, mml_FRASE0086
        End If
        Exit Sub
    End If
    
    If tbCodComp.Text = "" Or tbCodCateg.Text = "" Or cbFase.Text = mml_FRASE0310 Then
        MsgBox mml_FRASE0311, vbOKOnly Or vbExclamation, mml_FRASE0096
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0312, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If

    'Borrar todas las puntuaciones si es TeamMatch
    db.Execute "DELETE FROM ResultadosTeamMatch WHERE cod_categoria = " & tbCodCateg.Text
    'Borrar todas las puntuaciones, incluída la fase seleccionada
    db.Execute "DELETE FROM puntuaciones WHERE cod_categoria =" & tbCodCateg.Text & " AND fase <= " & sFase
    'Borrar todas las descalificaciones, incluída la fase seleccionada
    db.Execute "DELETE FROM descalificaciones WHERE cod_categoria =" & tbCodCateg.Text & " AND fase <= " & sFase
    'Borrar todas las hojas reconocidas, incluída la fase seleccionada
    db.Execute "DELETE FROM hojas_reconocidas WHERE cod_categoria =" & tbCodCateg.Text & " AND fase <= " & sFase
    'Borrar todas los dorsales de fases anteriores a la seleccionada
    db.Execute "DELETE FROM dorsales WHERE cod_categoria =" & tbCodCateg.Text & " AND ((fase < " & sFase & ") OR (fase = " & sFase & " AND repesca = 1))"
    'Borrar las combinaciones de dorsales
    db.Execute "DELETE FROM dorsalescombinados WHERE cod_categoria =" & tbCodCateg.Text & " AND ((fase < " & sFase & ") OR (fase = " & sFase & " AND repesca = 1))"
    'Borrar todas los datos de cal_conjunto (solo FINAL)
    db.Execute "DELETE FROM cal_conjunto WHERE cod_categoria =" & tbCodCateg.Text
    'Borrar todas los datos de cal_baile
    db.Execute "DELETE FROM cal_baile WHERE cod_categoria =" & tbCodCateg.Text & " AND fase <= " & sFase

    MsgBox mml_FRASE0276, vbOKOnly Or vbExclamation, mml_FRASE0086
    Exit Sub
error:
    ProcesarError "cmdBorrarPuntuaciones_Click"
End Sub

Private Sub cmdBuscarParticipantes_Click()
    On Local Error Resume Next
    frmBuscarPart.BuscarParticipantes Val(tbCodComp.Text)
End Sub

Private Sub cmdCambCat_Click()
    If tbCodComp.Text = "" Or tbCodCateg.Text = "" Or Val(cbFase.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    frmCambiarCateg.Cambiar tbCodComp.Text, tbCodCateg.Text
    
    Call cmdActualizar_Click
End Sub

Private Sub cmdCambiarDesCat_Click()
Dim iGrupo As Integer, iPos As Integer
    If tbCodCateg.Text = "" Then
        MsgBox mml_FRASE0303, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    iPos = InStr(tbDescCateg.Text, "(P")
    If cbAsigPista.ListIndex > 0 Then
        If iPos > 0 Then
            tbDescCateg.Text = Mid$(tbDescCateg.Text, 1, iPos - 1) & cbAsigPista.List(cbAsigPista.ListIndex)
        Else
            tbDescCateg.Text = tbDescCateg.Text & " " & cbAsigPista.List(cbAsigPista.ListIndex)
        End If
    Else
        If iPos > 0 Then
            tbDescCateg.Text = Trim$(Mid$(tbDescCateg.Text, 1, iPos - 1))
        End If
    End If
    
    iGrupo = IdentificarGrupoEdad(tbDescCateg.Text)
    If iGrupo > 0 Then
        db.Execute ("UPDATE categorias SET descripcion = '" & tbDescCateg.Text & "', cod_grupoedad = " & iGrupo & " WHERE codigo = " & tbCodCateg.Text)
    Else
        db.Execute ("UPDATE categorias SET descripcion = '" & tbDescCateg.Text & "' WHERE codigo = " & tbCodCateg.Text)
    End If
    MsgBox mml_FRASE0313, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Private Sub cmdCambiarFase_Click()
    If tbCodCateg.Text = "" Or Val(cbFase.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    MsgBox mml_FRASE1055, vbOKOnly Or vbCritical, G_MSG_AVISO
    frmCambiarFase.CambiarFase tbCodCateg.Text, Val(cbFase.List(cbFase.ListIndex)), cbRepesca.Value, Val(tbCodComp.Text)
End Sub

Private Sub cmdCambiarPareja_Click()
Dim iCodPareja As Long, iDorsal As Integer, iModalidad As Integer
Dim rs As Recordset
Dim sCod As String

    If Not C_DEBUG Then On Local Error GoTo error
    If MsgBox(G_PREGUNTA_OPERACION, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If
    If tbCodCateg.Text = "" Or Val(cbFase.Text) = 0 Then
        MsgBox mml_FRASE0349, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    iCodPareja = Val(sSeleccionar("SELECT p.codigo, nombre_hombre, nombre_mujer,m.nombre as modalidad,p.categoria,p.* FROM parejas p, modalidad m WHERE p.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text, "nombre_hombre", " ORDER BY " & C_ORDEN_PAREJAS))
    'Comprobamos el dorsal si ya está introducida la pareja
    If iCodPareja > 0 And dgDorsales.Row > 0 Then
        dgDorsales.Col = 0
        sCod = dgDorsales.Text
        If Val(sCod) > 0 Then
            db.Execute "UPDATE dorsales SET cod_pareja = " & iCodPareja & " WHERE codigo = " & sCod
            MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
        End If
    End If
    Exit Sub
error:
    ProcesarError "cmdCambiarPareja_Click"
End Sub

Private Sub cmdCateg_Click()
Dim rs As Recordset
   If Not C_DEBUG Then On Local Error GoTo error
   If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
   End If
    'JOIN para que aparezcan categ con 0 dorsales marcando que tienen 1 dorsal
    tbCodCateg.Text = sSeleccionar("SELECT c.codigo, c.descripcion, COUNT(*) AS NumDors, c.id_categoria, c.hora FROM (Categorias c LEFT JOIN Dorsales d ON d.cod_categoria = c.codigo) WHERE cod_competicion = " & tbCodComp.Text, "descripcion", " GROUP BY c.codigo, c.descripcion, c.id_categoria, c.hora  ORDER BY " & G_ORDEN_CATEGORIAS)
    ' No muestra las categorias vacias
    ' tbCodCateg.Text = sSeleccionar("SELECT c.codigo, c.descripcion, COUNT(*) AS NumDors, c.id_categoria, c.hora FROM Categorias c, Dorsales d WHERE d.cod_categoria = c.codigo AND cod_competicion = " & tbCodComp.Text & " GROUP BY c.codigo, c.descripcion, c.id_categoria, c.hora")
    tbDescCateg.Text = sResultado(2)
    
    cbFase.ListIndex = 0
    cbAsigPista.ListIndex = 0
    If Val(tbCodCateg.Text) > 0 Then
        'Calcular la menor fase
        EstablecerFase cbFase, MinFaseCateg(tbCodCateg.Text)
        
        Set rs = db.OpenRecordset("SELECT MAX(num_dorsal) FROM dorsales WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")", dbOpenSnapshot)
        If IsNull(rs.Fields(0)) Then
            tbDorsal.Text = iMinDorsalOficial(tbCodComp.Text)
        Else
            On Local Error Resume Next
            tbDorsal.Text = rs.Fields(0) + 1
        End If
        rs.Close
    Else
        cmdActualizar_Click
    End If

    Exit Sub
error:
    ProcesarError
End Sub


Private Sub cmdCategAct_Click()
    ActualizarCategActual
End Sub
Sub ActualizarCategActual()
Dim rs As Recordset, sMsj As String
    'h.cod_baile = 0 = Todos los bailes
    'h.cod_baile < 0  = indica el codigo de baile Std más bajo ya bailado
    Set rs = db.OpenRecordset("SELECT TOP 1 * FROM horario h WHERE grupo LIKE '*" & cbAsigPista.Text & "*' AND numfase <> " & C_FASE_GENERAL_LOOK & " AND cod_competicion = " & Val(VarCfg("horario_codcompeticion")) & " AND (SELECT COUNT(*) FROM puntuaciones WHERE ((h.cod_baile < 0 AND cod_baile " & G_ORDEN_10B_LAT_EST & " -h.cod_baile) OR h.cod_baile = 0 OR cod_baile = h.cod_baile) AND cod_categoria = h.cod_categoria AND fase = h.numfase AND repesca = h.repesca) = 0 ORDER BY orden", dbOpenSnapshot)
    If Not rs.EOF Then
        tbCodComp.Text = rs!cod_competicion
        tbDescComp.Text = sDescCompeticion(rs!cod_competicion)
        tbCodCateg.Text = rs!cod_categoria
        tbDescCateg.Text = sDescCategoria(rs!cod_categoria)
        cbFase.ListIndex = Log(rs!numfase) / Log(2) + 1
        cbFase_Click
    Else
        MsgBox mml_FRASE0442, vbOKOnly Or vbInformation, mml_FRASE0147
    End If
    rs.Close

End Sub

Private Sub cmdCombinarResultados_Click()
    On Local Error GoTo error
        If Val(tbCodCateg.Text) > 0 Then
            frmResultadosCombinados.CombinarResultados Val(tbCodCateg.Text), tbCodComp.Text
        Else
            MsgBox mml_FRASE0303, vbOKOnly Or vbInformation, G_MSG_ERROR
        End If
    Exit Sub
error:
    ProcesarError "cmdCombinarResultados_Click"
End Sub

Private Sub cmdComprobarFases_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    End If
    frmComprobarFases.Show vbNomodal
End Sub



Private Sub cmdComprobarRecogida_Click()
    If Val(tbCodComp.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
   
   frmControlPresencia.ControlPresencia Val(tbCodComp.Text)
End Sub

Private Sub cmdGenGrupos_Click()
    If MsgBox(mml_FRASE0307, vbYesNo Or vbCritical, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    If G_COUNTRY Then
        GenGruposCountry
    Else
        GenGruposAEBDC
    End If
End Sub

Sub GenGruposCountry()
Dim rsGrupos As Recordset, rsPar As Recordset, rsDorsal As Recordset
Dim sNombre As String
Dim iMod As Integer
Dim lCodGrupo As Long, lCodDorsal As Long
Dim iNumDorsal As Integer, iNumDorsales As Integer, iDorsal As Integer
Dim sSQL As String
Dim iFase As Integer

If Not C_DEBUG Then On Local Error GoTo error
    If Val(tbCodComp.Text) > 0 Then
        If MsgBox(mml_FRASE0322, vbYesNo Or vbCritical, mml_FRASE0084) = vbYes Then
            BorrarCompeticion Val(tbCodComp.Text), False, False
            
            iNumDorsal = iMinDorsalOficial(tbCodComp.Text)
            
            sSQL = "SELECT DISTINCT COUNT(*) as cont, grupoedad, cod_grupoedad, cod_modalidad, categoria, m.nombre, g.abreviatura FROM parejas p, modalidad m, gruposedad g WHERE g.codigo = p.cod_grupoedad AND p.cod_modalidad = m.codigo AND cod_competicion = " & Val(tbCodComp.Text) & " GROUP BY grupoedad, abreviatura, cod_grupoedad, cod_modalidad, categoria, m.nombre ORDER BY cod_modalidad, categoria, cod_grupoedad"
            Set rsGrupos = db.OpenRecordset("SELECT DISTINCT COUNT(*) as cont, grupoedad, cod_grupoedad, cod_modalidad, categoria, m.nombre, g.abreviatura, cod_modalidad FROM parejas p, modalidad m, gruposedad g WHERE g.codigo = p.cod_grupoedad AND p.cod_modalidad = m.codigo AND cod_competicion = " & Val(tbCodComp.Text) & " GROUP BY grupoedad, abreviatura, cod_grupoedad, cod_modalidad, categoria, m.nombre ORDER BY cod_modalidad, categoria, cod_grupoedad", dbOpenSnapshot)
            While Not rsGrupos.EOF
                sNombre = Left$(sDescModalidad(rsGrupos!cod_modalidad), 3) & " " & rsGrupos!categoria & " " & rsGrupos!grupoedad
                'Insertamos la nueva categoria
                lCodGrupo = MaxCod("categorias")
                db.Execute "INSERT INTO categorias VALUES (" & lCodGrupo & ", """ & sNombre & """, """ & rsGrupos!categoria & """," & Val(rsGrupos!cod_grupoedad) & "," & tbCodComp.Text & "," & rsGrupos!cod_modalidad & ",""10:00"",0,0," & VarCfg("max_dorsales_tanda") & ",0," & ImpHojaUnica & ")"
                'Ahora insertamos todas las parejas de la categoria
                Set rsPar = db.OpenRecordset("SELECT * FROM parejas WHERE cod_competicion = " & tbCodComp.Text & " AND cod_modalidad = " & rsGrupos!cod_modalidad & " AND cod_grupoedad = " & rsGrupos!cod_grupoedad & " AND categoria = '" & rsGrupos!categoria & "' ORDER BY cod_grupoedad, nombre_hombre", dbOpenSnapshot)
                    If Not rsPar.EOF Then
                        rsPar.MoveLast
                        iNumDorsales = rsPar.RecordCount
                        rsPar.MoveFirst
                        If iNumDorsales <= 7 Then
                            iFase = 1
                        ElseIf iNumDorsales <= 13 Then
                            iFase = 2
                        Else
                            iFase = 2 ^ (Int(Log((iNumDorsales - 1) / 6) / Log(2)) + 1)
                        End If
                        While Not rsPar.EOF
                            lCodDorsal = MaxCod("dorsales")
                            
                            ' Si esta pareja ya tiene dorsal en otra modalidad
                            sSQL = "SELECT num_dorsal FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND cod_categoria < " & lCodGrupo & " AND cod_modalidad <= " & rsGrupos!cod_modalidad & " AND cod_competicion = " & tbCodComp.Text & " AND ((nif_hombre <> '' AND nif_hombre = '" & rsPar!nif_hombre & "' AND nif_mujer = '" & rsPar!nif_mujer & "') OR (nombre_hombre <> '' AND nombre_hombre ='" & rsPar!nombre_hombre & "' AND nombre_mujer = '" & rsPar!nombre_mujer & "'))"
                            Debug.Print sSQL
                            Set rsDorsal = db.OpenRecordset(sSQL, dbOpenSnapshot)
                            If Not rsDorsal.EOF Then
                                iDorsal = rsDorsal!num_dorsal
                            Else
                                iDorsal = iNumDorsal
                                Inc iNumDorsal
                            End If
                                
                            rsDorsal.Close
                            
                            db.Execute "INSERT INTO dorsales VALUES (" & lCodDorsal & "," & iDorsal & "," & lCodGrupo & "," & iFase & "," & rsPar!codigo & ",0,0)"
                            rsPar.MoveNext
                        Wend
                    End If
                rsPar.Close
                rsGrupos.MoveNext
            Wend
            rsGrupos.Close
            MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
       End If
    End If
    
    Exit Sub
error:
    ProcesarError "cmdGenerar_Click"

End Sub
Private Sub GenGruposAEBDC()
Dim rs As Recordset
Dim rsParejas As Recordset
Dim rsDescEdad As Recordset
Dim iCont As Integer
Dim i As Integer, j As Integer
Dim sDesc As String
Dim aGruposAgrupados(50) As TGrupo
Dim iCodGrupoAnt As Integer
Dim iModalidad As Integer
Dim sCategoria As String
Dim iCodGrupo As Integer
Dim iNumDorsal As Integer
Dim iFase As Integer
Dim iPrimerGrupo As Integer
Dim sUltimaCategoria As String
Dim iUltimaModalidad As Integer
Dim rsDorsal As Recordset
Dim iDorsal As Integer
Dim iGrupoEdad As Integer
Dim bMinParejasAlcanzado As Boolean
Dim iGrupoActual As Integer
Dim sCatActual As String
Dim iNumParejas As Integer
Dim iPrimerGrupoAnt As Integer
Dim bNoAgruparHaciaAtras As Boolean
Dim bMinParejasAlcanzadoUlt As Boolean
Dim sCad As String

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")", dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            sCad = InputBox(mml_FRASE0321, mml_FRASE0084)
            sCad = UCase(sCad)
            If sCad <> "SI" And sCad <> "YES" Then
                Exit Sub
            End If
        End If
    rs.Close
    If MsgBox(mml_FRASE0322, vbYesNo Or vbCritical, mml_FRASE0084) = vbYes Then
        'Borramos todas las categorías existentes en esta competición y los dorsales asociados, junto con las puuntuaciones que pudieran tener
        BorrarCompeticion Val(tbCodComp.Text), False, False
        
        
        sExecSQL = "SELECT COUNT(*) as todas,cod_modalidad,categoria,cod_grupoedad, SUM(combinar_edad) as no_combinar FROM parejas WHERE cod_grupoedad <> 0 AND cod_competicion = " & tbCodComp.Text & " GROUP BY cod_modalidad,categoria,cod_grupoedad ORDER BY cod_modalidad,categoria,cod_grupoedad"
        Debug.Print sExecSQL
        Set rsParejas = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
        'Comenzamos
        iNumDorsal = iMinDorsalOficial(tbCodComp.Text)
        iPrimerGrupo = C_NO_HAY_GRUPO_ANTERIOR
        While Not rsParejas.EOF
            If bMinParejasAlcanzado Then
                rsParejas.MoveNext
            End If
            
            If rsParejas.EOF Then
                rsParejas.Close
                GoTo fin
            End If
            ' ********************************************************************************************************
            
            iModalidad = rsParejas!cod_modalidad
            sCategoria = rsParejas!categoria
            bMinParejasAlcanzado = False
            iCont = 0
            i = 0
            iPrimerGrupo = rsParejas!cod_grupoedad
            
            ' LINEAS EMPLEADAS PARA DEPURACIÓN -----------------------------
            If iModalidad = 3 And sCategoria = "H" And iPrimerGrupo = 1 Then
                DoEvents
            End If
            
            Do While Not rsParejas.EOF
                bNoAgruparHaciaAtras = False
                ' Si estamos en la misma mod y cat y podemos agrupar
                If (rsParejas!cod_modalidad = iModalidad And rsParejas!categoria = sCategoria) And _
                    rsParejas!cod_grupoedad - iPrimerGrupo < C_NUM_GRUPOS_AGRUPADOS Then
                    aGruposAgrupados(i).iGrupoEdad = rsParejas!cod_grupoedad
                    aGruposAgrupados(i).sCategoria = rsParejas!categoria
                    aGruposAgrupados(i).iModalidad = rsParejas!cod_modalidad
                    aGruposAgrupados(i).iTodasParejas = rsParejas!todas
                    
                    Inc i
                    
                    ' Si un grupo, el solo puede ser puntuable
                    If rsParejas!todas >= MinParejasGrupo(aGruposAgrupados(i - 1).iGrupoEdad, aGruposAgrupados(i - 1).sCategoria) Then
                        'Se cuentan todas las parejas
                        iCont = iCont + rsParejas!todas
                        If iPrimerGrupo >= G_ADULTO1 Then
                            ' Si es mayor o = que Ad1, como los anteriores no se pueden
                            ' agrupar a el se sale y se intenta agrupar solo los anteriores
                            iCont = iCont + rsParejas!todas - rsParejas!no_combinar
                            If iCont >= MinParejasGrupo(iPrimerGrupo, sCategoria) And bMinParejasAlcanzadoUlt Then
                                bNoAgruparHaciaAtras = True
                                iCont = 0
                            End If
                            rsParejas.MoveNext
                            Exit Do
                        End If
                    Else
                    ' Si el solo no puede ser puntuable se quitan las parejas que no quieren participar en una combinación de grupos
                        iCont = iCont + rsParejas!todas - rsParejas!no_combinar
                    End If
                    
                    'If iCont >= MinParejasGrupo(aGruposAgrupados(i - 1).iGrupoEdad, aGruposAgrupados(i - 1).sCategoria) Then
                    If iCont >= MinParejasGrupo(iPrimerGrupo, sCategoria) Then
                        bMinParejasAlcanzado = True
                        Exit Do
                    Else
                        rsParejas.MoveNext
                    End If
                Else
                    Exit Do
                End If
            Loop
            ' Cambio de categoría o se alcanzó el número de parejas o
            ' no se alcanzó pero pueden agrupar más de tres grupos
            
            ' Si salimos porque no se puede agrupar empezando por este grupo
            ' Avanzamos una posición y comenzamos de nuevo siempre que haya un intento de agrupación
            ' de más de un grupo
            If rsParejas.EOF Then ' Se acabaron los registros
                If i > 1 Then ' Pero quedaron más de un grupo que no se agrupan
                    For j = 1 To i - 1
                        rsParejas.MovePrevious
                    Next
                    'Creamos un grupo con el primer grupo solo
                    i = 1
                End If
                'Cambiamos de modalidad,cat o acabamos las posibilidades de agrupación o
                ' pasamos de ad1 y el último grupo debe ir solo por ser suficientemente grande
            ElseIf Not bMinParejasAlcanzado And i > 1 Then
                For j = 1 To i - 1
                    rsParejas.MovePrevious
                Next
                'Creamos un grupo con el primer grupo solo y comprobamos si se puede agrupar hacia atrás
                i = 1
            End If
            ' ********************************************************************************************************
            ' Solo se agrupa hacia atrás a partir de Adulto2
            ' Si salimos porque no podemos agrupar el primer grupo con los tres siguiente
            ' Comprobamos si hay posibilidad de agruparlo con el formado anteriormente
            If i = 1 And iPrimerGrupoAnt <> C_NO_HAY_GRUPO_ANTERIOR And _
                   iPrimerGrupo - iPrimerGrupoAnt < C_NUM_GRUPOS_AGRUPADOS And _
                   sUltimaCategoria = aGruposAgrupados(0).sCategoria And _
                   iUltimaModalidad = aGruposAgrupados(0).iModalidad And _
                   iPrimerGrupo >= G_ADULTO2 And Not bNoAgruparHaciaAtras Then
                Set rs = db.OpenRecordset("SELECT descripcion FROM categorias WHERE codigo = " & iCodGrupoAnt, dbOpenSnapshot)
                sDesc = rs!DESCRIPCION
                rs.Close
                ' Localizamos la descripción del grupo de edad
                Set rs = db.OpenRecordset("SELECT abreviatura, nombre FROM gruposedad WHERE codigo=" & aGruposAgrupados(0).iGrupoEdad, dbOpenSnapshot)
                sDesc = sDesc & "+" & rs!abreviatura
                rs.Close
                'Ahora cambiamos la descripción del grupo
                db.Execute ("UPDATE categorias SET descripcion = '" & sDesc & "' WHERE codigo = " & iCodGrupoAnt)
                
                'Ahora insertamos las parejas en este grupo
                Set rs = db.OpenRecordset("SELECT codigo, nif_hombre, nif_mujer, nombre_hombre, nombre_mujer, cod_modalidad, combinar_edad FROM parejas WHERE  cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad=" & aGruposAgrupados(0).iGrupoEdad & " AND cod_modalidad=" & aGruposAgrupados(0).iModalidad & " AND categoria='" & aGruposAgrupados(0).sCategoria & "' ORDER BY 4", dbOpenSnapshot)
                While Not rs.EOF
                    ' Como estamos agrupando hacia atrás, solo lo hacemos con las parejas que eligiron ser agrupadas
                    If rs!combinar_edad = C_COMBINAR_EDAD Then
                        ' Si esta pareja ya tiene dorsal en otra modalidad
                        Set rsDorsal = db.OpenRecordset("SELECT num_dorsal FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND cod_modalidad <> " & aGruposAgrupados(0).iModalidad & " AND cod_competicion = " & tbCodComp.Text & " AND ((nif_hombre = '" & rs!nif_hombre & "' AND nif_mujer = '" & rs!nif_mujer & "' AND '" & rs!nif_hombre & "' <>'') OR (nombre_hombre ='" & rs!nombre_hombre & "' AND nombre_mujer = '" & rs!nombre_mujer & "'))", dbOpenSnapshot)
                        If Not rsDorsal.EOF Then
                            iDorsal = rsDorsal!num_dorsal
                        Else
                            iDorsal = iNumDorsal
                            Inc iNumDorsal
                        End If
                        rsDorsal.Close
                        
                        db.Execute ("INSERT INTO dorsales VALUES(" & MaxCod("dorsales") & "," & iDorsal & "," & iCodGrupo & "," & iFase & "," & rs!codigo & ",0,0)")
                    End If
                    rs.MoveNext
                Wend
                rs.Close
            Else ' Hemos agrupado correctamente o tenemos solo un grupo para grabar **********************************
                sDesc = ""
                For j = 0 To i - 1
                    'Grupos de edad
                    Set rs = db.OpenRecordset("SELECT abreviatura, nombre FROM gruposedad WHERE codigo=" & aGruposAgrupados(j).iGrupoEdad, dbOpenSnapshot)
                    If sDesc <> "" Then sDesc = sDesc & "+"
                    ' Si solo hay una desc la ponemos completa
                    'If i = 1 Then
                    '    sDesc = sDesc & rs!Nombre
                    'Else
                        sDesc = sDesc & rs!abreviatura
                    'End If
                    rs.Close
                Next
                'Grupos de edad
                sDesc = aGruposAgrupados(j - 1).sCategoria & " " & sDesc
                'Modalidad
                Set rs = db.OpenRecordset("SELECT nombre FROM modalidad WHERE codigo=" & aGruposAgrupados(j - 1).iModalidad, dbOpenSnapshot)
                sDesc = Mid$(rs!Nombre, 1, C_CAR_DESC_MOD) & " " & sDesc
                rs.Close
                
                'Se agrupa hacia adelante en los menores de Adulto2
                If iPrimerGrupo <= G_YOUTH Then
                    iGrupoEdad = aGruposAgrupados(j - 1).iGrupoEdad
                    sCatActual = aGruposAgrupados(j - 1).sCategoria
                    iNumParejas = aGruposAgrupados(j - 1).iTodasParejas
                Else
                    iGrupoEdad = aGruposAgrupados(0).iGrupoEdad
                    sCatActual = aGruposAgrupados(0).sCategoria
                    iNumParejas = aGruposAgrupados(0).iTodasParejas
                End If
                'Ahora creamos el grupo
                iCodGrupo = MaxCod("categorias")
                db.Execute ("INSERT INTO categorias VALUES(" & iCodGrupo & ",'" & sDesc & "','" & aGruposAgrupados(j - 1).sCategoria & "'," & iGrupoEdad & "," & tbCodComp.Text & "," & aGruposAgrupados(j - 1).iModalidad & ",'12:00',0,0," & C_DORSALES_TANDA_DEFECTO & ",0," & ImpHojaUnica & ")")
                iCodGrupoAnt = iCodGrupo

                'Comprobamos la Fase de comienzo
                If iCont <= MinParejasGrupo(iGrupoEdad, sCatActual) Then
                    iFase = 1
                Else
                    If iCont <= 7 Then
                        iFase = 1
                    ElseIf iCont <= 13 Then
                        iFase = 2
                    Else
                        iFase = 2 ^ (Int(Log((iCont - 1) / 6) / Log(2)) + 1)
                    End If
                End If
                'Ahora insertamos las parejas en este grupo
                For j = 0 To i - 1
                    'Grupos de edad
                    Debug.Print "SELECT codigo FROM parejas WHERE  cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad=" & aGruposAgrupados(j).iGrupoEdad & " AND cod_modalidad=" & aGruposAgrupados(j).iModalidad & " AND categoria=" & aGruposAgrupados(j).sCategoria
                    Set rs = db.OpenRecordset("SELECT codigo, nif_hombre, nif_mujer, nombre_hombre, nombre_mujer, cod_modalidad, combinar_edad, cod_grupoedad FROM parejas WHERE  cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad=" & aGruposAgrupados(j).iGrupoEdad & " AND cod_modalidad=" & aGruposAgrupados(j).iModalidad & " AND categoria='" & aGruposAgrupados(j).sCategoria & "'", dbOpenSnapshot)
                    While Not rs.EOF
                    
                        '(Si queremos combinar) o (Si no queremos combinar pero pertenecemos al grupo de edad
                        ' en donde se realiza la combinación y el número de parejas es suficiente) => combinamos
                        If rs!combinar_edad = C_COMBINAR_EDAD Or _
                            (rs!cod_grupoedad = iGrupoEdad And rs!combinar_edad = C_NO_COMBINAR_EDAD And _
                            iNumParejas >= MinParejasGrupo(iGrupoEdad, sCatActual)) Then
                            ' Si esta pareja ya tiene dorsal en otra modalidad
                            Set rsDorsal = db.OpenRecordset("SELECT num_dorsal FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND cod_modalidad <> " & aGruposAgrupados(0).iModalidad & " AND cod_competicion = " & tbCodComp.Text & " AND ((nif_hombre = '" & rs!nif_hombre & "' AND nif_mujer = '" & rs!nif_mujer & "' AND '" & rs!nif_hombre & "' <>'') OR (nombre_hombre ='" & rs!nombre_hombre & "' AND nombre_mujer = '" & rs!nombre_mujer & "'))", dbOpenSnapshot)
                            If Not rsDorsal.EOF Then
                                iDorsal = rsDorsal!num_dorsal
                            Else
                                iDorsal = iNumDorsal
                                Inc iNumDorsal
                            End If
                            rsDorsal.Close
                        
                            db.Execute ("INSERT INTO dorsales VALUES(" & MaxCod("dorsales") & "," & iDorsal & "," & iCodGrupo & "," & iFase & "," & rs!codigo & ",0,0)")
                        End If
                        
                        rs.MoveNext
                    Wend
                    rs.Close
                Next
                'If iCont < MinParejasGrupo(iGrupoEdad, sCatActual) Then
                If iCont < MinParejasGrupo(iPrimerGrupo, sCatActual) Then
                    MsgBox mml_FRASE0323 & sDesc & mml_FRASE0324, vbOKOnly Or vbInformation, mml_FRASE0084
                End If
                
                'Hemos creado un nuevo grupo y cambiamos los valores
                iPrimerGrupoAnt = iPrimerGrupo
                sUltimaCategoria = aGruposAgrupados(0).sCategoria
                iUltimaModalidad = aGruposAgrupados(0).iModalidad
                bMinParejasAlcanzadoUlt = bMinParejasAlcanzado
            
            End If
            
        Wend
        rsParejas.Close
    Else
        Exit Sub
    End If
fin:
    MsgBox mml_FRASE0325, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub
Private Function MinParejasGrupo(iGrupoEdad As Integer, sCateg As String) As Integer
    Select Case UCase(sCateg)
        Case "A"
            MinParejasGrupo = 4
        Case "B"
            MinParejasGrupo = 4
        Case "C"
            MinParejasGrupo = 4
        Case Else
            If iGrupoEdad = G_JUVENIL Then
                MinParejasGrupo = 4
            Else
                MinParejasGrupo = 6
            End If
    End Select
End Function


Private Sub cmdGenHorarioHTML_Click()
Dim rs As Recordset
Dim iFase As Integer
Dim sFase As String
Dim iMinutos As Integer

   If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
   End If
   
   If Not C_DEBUG Then On Local Error GoTo error
    Open G_ARCH_HORARIO For Output As #100
    Print #100, mml_FRASE0326
    
    iMinutos = 0
    Set rs = db.OpenRecordset("SELECT * FROM horario WHERE cod_competicion =" & tbCodComp.Text & " ORDER BY hora, orden", dbOpenSnapshot)
    While Not rs.EOF
        Print #100, "<TR><TD>"
        Print #100, Format$(rs!hora, "hh:mm")
        Print #100, "</TD><TD>"
        Print #100, rs!grupo
        Print #100, "</TD><TD>"
        Print #100, rs!fase
        Print #100, "</TD></TR>"
        rs.MoveNext
    Wend
    rs.Close
    
    Print #100, "</TABLE></CENTER></BODY></HTML>"
    Close #100

    MsgBox mml_FRASE0327 & G_ARCH_HORARIO, vbOKOnly Or vbInformation, mml_FRASE0086
error:
    ProcesarError
End Sub

Private Sub cmdGenTeamMatch_Click()
Dim rs As Recordset
    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If MsgBox("¿Seguro que desea generar las categorias del TeamMatch?", vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    frmTeamMatch.GenerarCategorias Val(tbCodComp.Text)
End Sub

Private Sub cmdHayPuntuaciones_Click()
Dim rs As Recordset
Dim sCad As String

    If Not C_DEBUG Then On Local Error GoTo error
    Set rs = db.OpenRecordset("SELECT DISTINCT c.descripcion FROM puntuaciones p, categorias c WHERE c.codigo = p.cod_categoria AND c.cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    sCad = ""
    While Not rs.EOF
        If Len(sCad) < C_LIM_MSGBOX Then
            sCad = sCad & vbCrLf & rs.Fields("descripcion")
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    If sCad <> "" Then
        MsgBox mml_FRASE1129 & vbCrLf & sCad, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    Else
        MsgBox mml_FRASE1130, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    End If
    Exit Sub
error:
    ProcesarError "cmdHayPuntuaciones_Click"
End Sub

Private Sub cmdHorario_Click()
Dim rs As Recordset, rs1 As Recordset
Dim iFase As Integer
Dim sFase As String
Dim iMinutos As Integer
Dim iCOrden As Integer
Dim iNumDorsales As Integer

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0328, vbYesNo Or vbInformation, mml_FRASE0084) = vbNo Then Exit Sub
    
    Open G_ARCH_HORARIO For Output As #100
    Print #100, mml_FRASE0326
    
    db.Execute ("DELETE FROM horario WHERE cod_competicion = " & tbCodComp.Text)
    iMinutos = 0
    iCOrden = 11
    Set rs = db.OpenRecordset("SELECT [cod_categoria], [descripcion], [hora], MAX([fase]) as f FROM dorsales AS d, categorias AS c Where d.cod_categoria = C.codigo And cod_competicion = " & tbCodComp.Text & " GROUP BY cod_categoria,descripcion, hora", dbOpenSnapshot)
    While Not rs.EOF
        Set rs1 = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & rs!cod_categoria)
        iNumDorsales = rs1.Fields(0)
        If iNumDorsales = 7 And Not G_COUNTRY Then
            'Insertamos un GeneralLook
            iCOrden = iCOrden + 10
            iFase = C_FASE_GENERAL_LOOK
            db.Execute ("INSERT INTO horario VALUES ('" & DateAdd("n", iMinutos, rs!hora) & "','" & rs!DESCRIPCION & "','GeneralLook', " & iFase & "," & rs!cod_categoria & ",0," & iCOrden & ",0," & tbCodComp.Text & ",0," & iNumDorsales & ",0,1)")
        End If
        rs1.Close
        iFase = rs!f
        Do
            Select Case iFase
                Case 1:
                    sFase = mml_FRASE0329
                Case 2:
                    sFase = mml_FRASE0330
                Case Else
                    sFase = "1/" & Trim$(Str$(iFase)) & mml_FRASE0331
            End Select
            
            Print #100, "<TR><TD>"
            iCOrden = iCOrden + 10
            db.Execute ("INSERT INTO horario VALUES ('" & DateAdd("n", iMinutos, rs!hora) & "','" & rs!DESCRIPCION & "','" & sFase & "', " & iFase & "," & rs!cod_categoria & ",0," & iCOrden & ",0," & tbCodComp.Text & ",0," & iNumDorsales & ",0,1)")
            Print #100, "</TD><TD>"
            Print #100, rs!hora
            Print #100, "</TD><TD>"
            Print #100, rs!DESCRIPCION
            Print #100, "</TD><TD>"
            Print #100, sFase
            Print #100, "</TD></TR>"
            iFase = iFase / 2
        Loop While iFase >= 1
        iMinutos = iMinutos + G_MINUTOS_POR_CATEG
        rs.MoveNext
    Wend
    rs.Close
    
    Print #100, "</TABLE></CENTER></BODY></HTML>"
    Close #100

    MsgBox mml_FRASE0332, vbOKOnly Or vbInformation, mml_FRASE0086
    
End Sub

Private Sub cmdImpHojasPuntuaciones_Click()
    If Val(tbCodComp.Text) = 0 Or Val(tbCodCateg.Text) = 0 Or Val(cbFase.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If

    With frmImpHojasPuntuaciones
        .tbCodComp.Text = tbCodComp.Text
        .tbDescComp.Text = tbDescComp.Text
        .tbCodCat.Text = tbCodCateg.Text
        .tbDescCat.Text = tbDescCateg.Text
        .tbCodFase.Text = Val(cbFase.Text)
        .tbDescFase.Text = sDescFase(Val(cbFase.Text))
        
        .Show vbModal
    End With
End Sub
Sub RenumerarDorsales()
Dim sSQL As String
Dim iNumDorsal As Integer
Dim rs As Recordset
    'Seleccionamos todas las parejas convencionales de la competición y asignamos dorsales muy altos
    ' Para evitar problemas de colisión renumeramos los dorsales con números muy altos
    sSQL = "SELECT codigo FROM dorsales d WHERE cod_categoria in (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")"
    Debug.Print sSQL
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
    iNumDorsal = C_DORSAL_INI_RENUMERACION
    While Not rs.EOF
        db.Execute ("UPDATE dorsales SET num_dorsal = " & iNumDorsal & " WHERE num_dorsal <= " & 2 * C_DORSAL_INI_RENUMERACION & " AND codigo = " & rs.Fields("codigo"))
        Inc iNumDorsal
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub cmdImpNumDorsal_Click()
Dim sTitulo As String
Dim rs As Recordset
Dim Ancho As Integer
Dim sHombre As String
Dim sMujer As String
Dim sDorsal As String
Dim iTamFuente As Integer
Dim iCont As Integer
Dim bUnDorsalPorHoja As Boolean
Dim bApaisado As Boolean
Dim sCad As String

    If MsgBox(mml_FRASE0333, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If

    sTitulo = tbDescComp.Text
    iTamFuente = Val(VarCfg("fuente_dorsal"))
    sCad = ""
    
    Ancho = Printer.Height * 60 / 100
            
    sCad = InputBox(mml_FRASE1192, "", sCad)
    sDorsal = InputBox(mml_FRASE0336)
    If Val(sDorsal) > 0 Then
        Set rs = db.OpenRecordset("SELECT DISTINCT nombre_hombre, nombre_mujer, num_dorsal FROM dorsales d, parejas p WHERE cod_competicion = " & tbCodComp.Text & " AND d.num_dorsal = " & sDorsal & " AND d.cod_pareja = p.codigo ORDER BY 3", dbOpenSnapshot)
    Else
        MsgBox mml_FRASE0337, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    
    bApaisado = False
    bUnDorsalPorHoja = True
    
    Printer.Orientation = vbPRORPortrait
    
    Printer.FontName = G_NOMBRE_FUENTE_DORSAL
    Printer.Print ""
    If G_LOGO_DORSAL_IZQ <> "" Then
        If Dir$(G_LOGO_DORSAL_IZQ) = "" Then
            MsgBox mml_FRASE0340 & G_LOGO_DORSAL_IZQ & mml_FRASE0341, vbOKOnly Or vbInformation, mml_FRASE0096
        End If
    End If
    If G_LOGO_DORSAL_DER <> "" Then
        If Dir$(G_LOGO_DORSAL_DER) = "" Then
            MsgBox mml_FRASE0340 & G_LOGO_DORSAL_DER & mml_FRASE0342, vbOKOnly Or vbInformation, mml_FRASE0096
        End If
    End If
        
    Printer.CurrentY = G_DORSAL_POR_Y
    While Not rs.EOF
        Inc iCont
        sDorsal = IIf(IsNull(rs!num_dorsal), "", rs!num_dorsal)
        sHombre = IIf(IsNull(rs!nombre_hombre), "", rs!nombre_hombre)
        sMujer = IIf(IsNull(rs!nombre_mujer), "", rs!nombre_mujer)
        'Printer.PaintPicture LoadPicture("fondo_dorsal.bmp"), 1200, 0
        Printer.FontSize = G_FUENTE_GRANDE_DORSAL
        Printer.FontBold = False
        Printer.CurrentX = (Ancho - Printer.TextWidth(sTitulo)) / 2 + G_MARGEN_DORSAL_X
        Printer.Print Trim$(sTitulo);
        Printer.FontSize = G_FUENTE_PEQUE_DORSAL
        Printer.Print
        Printer.FontSize = iTamFuente
        Printer.FontBold = True
        Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sDorsal))) / 2 + G_MARGEN_DORSAL_X
        Printer.Print Trim$(sDorsal)
        Printer.FontSize = G_FUENTE_PEQUE_DORSAL
        Printer.FontBold = False
        Printer.Print
        If G_LOGO_DORSAL_IZQ <> "" And Dir$(G_LOGO_DORSAL_IZQ) <> "" Then
            Printer.PaintPicture LoadPicture(G_LOGO_DORSAL_IZQ), G_MARGEN_DORSAL_X + 1200, Printer.CurrentY - 800
        End If
        If G_LOGO_DORSAL_DER <> "" And Dir$(G_LOGO_DORSAL_DER) <> "" Then
            Printer.PaintPicture LoadPicture(G_LOGO_DORSAL_DER), G_MARGEN_DORSAL_X + 8000, Printer.CurrentY - 800
        End If
        Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sHombre))) / 2 + 600 + G_MARGEN_DORSAL_X
        Printer.Print Trim$(sHombre)
        Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sMujer))) / 2 + 600 + G_MARGEN_DORSAL_X
        Printer.Print Trim$(sMujer)
        Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sCad))) / 2 + 600 + G_MARGEN_DORSAL_X
        Printer.Print Trim$(sCad)
        
        If bUnDorsalPorHoja Then
            Printer.NewPage
            Printer.CurrentY = G_DORSAL_POR_Y
        Else
            If iCont Mod 2 = 0 Then
                Printer.NewPage
                Printer.CurrentY = G_DORSAL_POR_Y
            Else
                Printer.Print
                Printer.Print
                Printer.Print
                Printer.Print
            End If
        End If
        rs.MoveNext
        
    Wend
    
    Printer.EndDoc
    rs.Close
    
    MsgBox mml_FRASE0343, vbOKOnly Or vbInformation, mml_FRASE0086


End Sub

Private Sub cmdImportarDorsalesProBaile_Click()
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim rsPar As Recordset
    Dim lCod As Long
    Dim iDorsal As Integer
    Dim sSQL As String
    Dim iNumDorsal As Long
    Dim iOpc As Integer
    
        If Not C_DEBUG Then On Local Error GoTo error
        If tbCodComp.Text = "" Then
            CamposSinCubrir
            Exit Sub
        End If
    
        If MsgBox(mml_FRASE1054, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbNo Then
            Exit Sub
        End If
        
        iOpc = Val(InputBox(mml_FRASE1139, "Select"))
        If iOpc = 2 Then
            RenumerarDorsales
            'Recuperamos los dorsales de [Lista de participantes]
            Set rs = db.OpenRecordset("SELECT c.cod_modalidad, d.codigo, p.nombre_hombre, p.nombre_mujer FROM dorsales d, parejas p, categorias c WHERE c.codigo = d.cod_categoria AND d.cod_pareja = p.codigo AND c.cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
            While Not rs.EOF
                If rs.Fields("cod_modalidad") = COD_MODALIDAD_STD Then
                    Set rs1 = db.OpenRecordset("SELECT campo42 AS dorsal FROM [Lista de participantes] WHERE campo22 <> '' AND NOT campo22 IS NULL AND campo1&campo2&campo3 = '" & rs.Fields("nombre_hombre") & "' AND campo4&campo5&campo6 = '" & rs.Fields("nombre_mujer") & "'", dbOpenSnapshot)
                    If Not rs1.EOF Then
                        db.Execute "UPDATE dorsales SET num_dorsal = " & rs1.Fields("dorsal") & " WHERE codigo = " & rs.Fields("codigo")
                    End If
                    rs.MoveNext
                Else
                    Set rs1 = db.OpenRecordset("SELECT campo42 AS dorsal FROM [Lista de participantes] WHERE campo23 <> '' AND NOT campo23 IS NULL AND campo1&campo2&campo3 = '" & rs.Fields("nombre_hombre") & "' AND campo4&campo5&campo6 = '" & rs.Fields("nombre_mujer") & "'", dbOpenSnapshot)
                    If Not rs1.EOF Then
                        db.Execute "UPDATE dorsales SET num_dorsal = " & rs1.Fields("dorsal") & " WHERE codigo = " & rs.Fields("codigo")
                    End If
                    rs.MoveNext
                End If
            Wend
            rs.Close
        ElseIf iOpc = 1 Then
            RenumerarDorsales
                
            If Not C_DEBUG Then On Local Error GoTo error
            sSQL = "SELECT codigo, dorsal_probaile FROM dorsales d, enlaceprobaile e " & _
                    " WHERE d.cod_pareja = e.cod_pareja " & _
                    " AND d.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")" & _
                    " AND e.cod_competicion = " & tbCodComp.Text
    
            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
            While Not rs.EOF
                If rs!dorsal_probaile = 0 Then
                    'Probaile exporta un dorsal 0
                    Set rsPar = db.OpenRecordset("SELECT nombre_hombre, nombre_mujer FROM parejas p, dorsales d WHERE p.codigo = d.cod_pareja AND d.codigo = " & rs.Fields("codigo"), dbOpenSnapshot)
                    MsgBox mml_FRASE1095 & rsPar.Fields("nombre_hombre") & ", " & rsPar.Fields("nombre_mujer"), vbOKOnly Or vbCritical, G_MSG_ERROR
                    rsPar.Close
                Else
                    db.Execute "UPDATE dorsales SET num_dorsal = " & rs!dorsal_probaile & " WHERE codigo = " & rs!codigo
                End If
                rs.MoveNext
            Wend
            rs.Close
            
            'Buscamos todas las parejas que no se importaron de probaile y les asignamos dorsales diferentes
            Set rs = db.OpenRecordset("SELECT codigo FROM dorsales d WHERE NOT cod_pareja IN " & _
                                      "(SELECT cod_pareja FROM enlaceprobaile as pb WHERE cod_competicion = " & tbCodComp.Text & " ) " & _
                                      " AND d.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")", dbOpenSnapshot)
            While Not rs.EOF
                Set rs1 = db.OpenRecordset("SELECT MAX(num_dorsal) FROM dorsales WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")", dbOpenSnapshot)
                If IsNull(rs1.Fields(0)) Then
                    iDorsal = iMinDorsalOficial(tbCodComp.Text)
                Else
                    On Local Error Resume Next
                    iDorsal = rs1.Fields(0) + 1
                End If
                
                db.Execute "UPDATE dorsales SET dorsal = " & iDorsal & " WHERE codigo = " & rs!codigo
                rs1.Close
                rs.MoveNext
            Wend
                    
            rs.Close
        End If
        MsgBox mml_FRASE0134, vbOKOnly Or vbInformation, mml_FRASE0086
        Exit Sub
error:
    ProcesarError "cmdImportarDorsalesProBaile_Click"
End Sub

Private Sub cmdImprimirDorsales_Click()
Dim sTitulo As String
Dim rs As Recordset
Dim Ancho As Integer
Dim sHombre As String
Dim sMujer As String
Dim sDorsal As String
Dim iTamFuente As Integer
Dim iCont As Integer
Dim bUnDorsalPorHoja As Boolean
Dim bApaisado As Boolean
Dim sCad As String

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE0333, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
   
    sTitulo = tbDescComp.Text
    iTamFuente = Val(VarCfg("fuente_dorsal"))
    sCad = ""
    
    Ancho = Printer.Height * 60 / 100
    If MsgBox(mml_FRASE0334, vbYesNo Or vbQuestion, mml_FRASE0099) = vbNo Then
        If MsgBox(mml_FRASE0335, vbYesNo Or vbQuestion, mml_FRASE0099) = vbNo Then
            sCad = InputBox(mml_FRASE1192, "", sCad)
            sDorsal = InputBox(mml_FRASE0336)
            If Val(sDorsal) > 0 Then
                Set rs = db.OpenRecordset("SELECT DISTINCT nombre_hombre, nombre_mujer, num_dorsal FROM dorsales d, parejas p WHERE cod_competicion = " & tbCodComp.Text & " AND d.num_dorsal = " & sDorsal & " AND d.cod_pareja = p.codigo ORDER BY 3", dbOpenSnapshot)
            Else
                MsgBox mml_FRASE0337, vbOKOnly Or vbInformation, mml_FRASE0096
                Exit Sub
            End If
        Else
            If tbCodCateg.Text = "" Then
                CamposSinCubrir
                Exit Sub
            End If
            sCad = tbDescCateg.Text
            sCad = InputBox(mml_FRASE1192, "", sCad)
            'Dorsales del grupo
            Set rs = db.OpenRecordset("SELECT DISTINCT nombre_hombre, nombre_mujer, num_dorsal FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND d.cod_categoria = " & tbCodCateg.Text & " AND repesca=" & cbRepesca.Value & " ORDER BY 3", dbOpenSnapshot)
        End If
    Else
        'Dorsales de la competición
        Set rs = db.OpenRecordset("SELECT DISTINCT nombre_hombre, nombre_mujer, num_dorsal FROM dorsales d, parejas p, categorias c WHERE d.cod_pareja = p.codigo AND d.cod_categoria = c.codigo AND c.cod_competicion = " & tbCodComp.Text & " ORDER BY 3", dbOpenSnapshot)
    End If
    
    If MsgBox(mml_FRASE0338, vbYesNo Or vbQuestion, mml_FRASE0099) = vbNo Then
        bApaisado = False
        If MsgBox(mml_FRASE0339, vbYesNo Or vbQuestion, mml_FRASE0099) = vbNo Then
            bUnDorsalPorHoja = True
        Else
            bUnDorsalPorHoja = False
        End If
    Else
        bApaisado = True
    End If
    
    If bApaisado Then
        Printer.Orientation = vbPRORLandscape
    Else
        Printer.Orientation = vbPRORPortrait
    End If
    Printer.FontName = G_NOMBRE_FUENTE_DORSAL
    Printer.Print ""
    If G_LOGO_DORSAL_IZQ <> "" Then
        If Dir$(G_LOGO_DORSAL_IZQ) = "" Then
            MsgBox mml_FRASE0340 & G_LOGO_DORSAL_IZQ & mml_FRASE0341, vbOKOnly Or vbInformation, mml_FRASE0096
        End If
    End If
    If G_LOGO_DORSAL_DER <> "" Then
        If Dir$(G_LOGO_DORSAL_DER) = "" Then
            MsgBox mml_FRASE0340 & G_LOGO_DORSAL_DER & mml_FRASE0342, vbOKOnly Or vbInformation, mml_FRASE0096
        End If
    End If
    If bApaisado Then
        While Not rs.EOF
            sDorsal = IIf(IsNull(rs!num_dorsal), "", rs!num_dorsal)
            sHombre = IIf(IsNull(rs!nombre_hombre), "", rs!nombre_hombre)
            sMujer = IIf(IsNull(rs!nombre_mujer), "", rs!nombre_mujer)
            'Printer.PaintPicture LoadPicture("fondo_dorsal.bmp"), 1200, 0
            Printer.CurrentY = G_DORSAL_POR_Y
            Printer.FontSize = G_FUENTE_GRANDE_DORSAL
            Printer.FontBold = False
            Printer.CurrentX = (Ancho - Printer.TextWidth(sTitulo)) / 2 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sTitulo)
            Printer.FontSize = G_FUENTE_PEQUE_DORSAL
            Printer.Print
            Printer.FontSize = iTamFuente
            Printer.FontBold = True
            Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sDorsal))) / 2 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sDorsal)
            Printer.FontSize = G_FUENTE_PEQUE_DORSAL
            Printer.FontBold = False
            Printer.Print
            If G_LOGO_DORSAL_IZQ <> "" And Dir$(G_LOGO_DORSAL_IZQ) <> "" Then
                Printer.PaintPicture LoadPicture(G_LOGO_DORSAL_IZQ), G_MARGEN_DORSAL_X + 1200, Printer.CurrentY - 800
            End If
            If G_LOGO_DORSAL_DER <> "" And Dir$(G_LOGO_DORSAL_DER) <> "" Then
                Printer.PaintPicture LoadPicture(G_LOGO_DORSAL_DER), G_MARGEN_DORSAL_X + 8000, Printer.CurrentY - 800
            End If
            Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sHombre))) / 2 + 600 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sHombre)
            Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sMujer))) / 2 + 600 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sMujer)
            Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sCad))) / 2 + 600 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sCad)
            
            Printer.NewPage
            
            rs.MoveNext
            
        Wend
     Else
        Printer.CurrentY = G_DORSAL_POR_Y
        While Not rs.EOF
            Inc iCont
            sDorsal = IIf(IsNull(rs!num_dorsal), "", rs!num_dorsal)
            sHombre = IIf(IsNull(rs!nombre_hombre), "", rs!nombre_hombre)
            sMujer = IIf(IsNull(rs!nombre_mujer), "", rs!nombre_mujer)
            'Printer.PaintPicture LoadPicture("fondo_dorsal.bmp"), 1200, 0
            Printer.FontSize = G_FUENTE_GRANDE_DORSAL
            Printer.FontBold = False
            Printer.CurrentX = (Ancho - Printer.TextWidth(sTitulo)) / 2 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sTitulo);
            Printer.FontSize = G_FUENTE_PEQUE_DORSAL
            Printer.Print
            Printer.FontSize = iTamFuente
            Printer.FontBold = True
            Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sDorsal))) / 2 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sDorsal)
            Printer.FontSize = G_FUENTE_PEQUE_DORSAL
            Printer.FontBold = False
            Printer.Print
            If G_LOGO_DORSAL_IZQ <> "" And Dir$(G_LOGO_DORSAL_IZQ) <> "" Then
                Printer.PaintPicture LoadPicture(G_LOGO_DORSAL_IZQ), G_MARGEN_DORSAL_X + 1200, Printer.CurrentY - 800
            End If
            If G_LOGO_DORSAL_DER <> "" And Dir$(G_LOGO_DORSAL_DER) <> "" Then
                Printer.PaintPicture LoadPicture(G_LOGO_DORSAL_DER), G_MARGEN_DORSAL_X + 8000, Printer.CurrentY - 800
            End If
            Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sHombre))) / 2 + 600 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sHombre)
            Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sMujer))) / 2 + 600 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sMujer)
            Printer.CurrentX = (Ancho - Printer.TextWidth(Trim$(sCad))) / 2 + 600 + G_MARGEN_DORSAL_X
            Printer.Print Trim$(sCad)
            
            If bUnDorsalPorHoja Then
                Printer.NewPage
                Printer.CurrentY = G_DORSAL_POR_Y
            Else
                If iCont Mod 2 = 0 Then
                    Printer.NewPage
                    Printer.CurrentY = G_DORSAL_POR_Y
                Else
                    Printer.Print
                    Printer.Print
                    Printer.Print
                    Printer.Print
                End If
            End If
            rs.MoveNext
            
        Wend
    End If
    Printer.EndDoc
    rs.Close
    
    MsgBox mml_FRASE0343, vbOKOnly Or vbInformation, mml_FRASE0086

End Sub

Private Sub cmdJuecesBailes_Click()
On Local Error GoTo error
    frmAJuezBaile.cbFase.ListIndex = IIf(cbFase.ListIndex = 1, 0, 1)
    
    frmAJuezBaile.tbCodComp.Text = tbCodComp.Text
    frmAJuezBaile.tbDescComp.Text = tbDescComp.Text
    frmAJuezBaile.tbCodCateg.Text = tbCodCateg.Text
    frmAJuezBaile.tbDescCateg.Text = tbDescCateg.Text
    
    frmAJuezBaile.cmdActualizar_Click
    frmAJuezBaile.Show vbModal
    Exit Sub
error:
    ProcesarError "cmdJuecesBailes_Click"
End Sub

Private Sub cmdNoPresente_Click()
Dim iCodDorsal As Long
    
    dgDorsales.Col = 0
    If dgDorsales.Text = "" Or tbCodCateg.Text = "" Or cbFase.Text = "" Then
        MsgBox mml_FRASE0344, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0345, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    dgDorsales.Col = 0
    iCodDorsal = dgDorsales.Text
    db.Execute ("UPDATE dorsales SET no_presente = 2 WHERE codigo =" & iCodDorsal)
    Call cmdActualizar_Click

End Sub

Private Sub cmdNuevaDesc_Click()
Dim lCodigo As Long
Dim rs As Recordset
Dim iGrupo As Integer
    If Trim$(tbDescCateg.Text) = "" Then
        MsgBox mml_FRASE0346, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If Val(tbCodCateg.Text) = 0 Then
        MsgBox mml_FRASE0347, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    iGrupo = IdentificarGrupoEdad(tbDescCateg.Text)
    lCodigo = MaxCod("categorias")
    Set rs = db.OpenRecordset("SELECT * FROM categorias WHERE codigo = " & tbCodCateg.Text, dbOpenSnapshot)
    If Not rs.EOF Then
        If iGrupo = 0 Then
            db.Execute ("INSERT INTO categorias VALUES (" & lCodigo & ",'" & tbDescCateg.Text & "','" & rs!id_categoria & "'," & rs!cod_grupoedad & "," & rs!cod_competicion & "," & rs!cod_modalidad & ",'" & rs!hora & "',0,0," & C_DORSALES_TANDA_DEFECTO & ",0," & ImpHojaUnica & ")")
        Else
            db.Execute ("INSERT INTO categorias VALUES (" & lCodigo & ",'" & tbDescCateg.Text & "','" & rs!id_categoria & "'," & iGrupo & "," & rs!cod_competicion & "," & rs!cod_modalidad & ",'" & rs!hora & "',0,0," & C_DORSALES_TANDA_DEFECTO & ",0," & ImpHojaUnica & ")")
        End If
        tbCodCateg.Text = lCodigo
        MsgBox mml_FRASE0348, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
    rs.Close
End Sub

Private Sub cmdParejas_Click()
    If tbCodComp.Text <> "" Then
        frmAParejas.tbCodComp.Text = tbCodComp.Text
        frmAParejas.tbDescComp.Text = tbDescComp.Text
    End If
    frmAParejas.Show vbNomodal
End Sub

Private Sub cmdPoner_Click()
Dim iCodPareja As Long, iDorsal As Integer, iModalidad As Integer
Dim rs As Recordset

    If MsgBox(G_PREGUNTA_OPERACION, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If

    If tbCodCateg.Text = "" Or Val(cbFase.Text) = 0 Then
        MsgBox mml_FRASE0349, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If g_lCodUltimaPareja > 0 And chkAgregarUltima.Value = 1 Then
        iCodPareja = g_lCodUltimaPareja
    Else
        iCodPareja = Val(sSeleccionar("SELECT p.codigo, nombre_hombre, nombre_mujer,m.nombre as modalidad,p.categoria,p.* FROM parejas p, modalidad m WHERE p.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text, "nombre_hombre", " ORDER BY " & C_ORDEN_PAREJAS))
    End If
    'Comprobamos el dorsal si ya está introducida la pareja
    If iCodPareja > 0 Then
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM parejas p WHERE p.codigo = " & iCodPareja & " AND cod_modalidad = (SELECT cod_modalidad FROM categorias WHERE codigo = " & tbCodCateg.Text & ")", dbOpenSnapshot)
        If rs.Fields(0) = 0 Then
            If MsgBox(mml_FRASE1109, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbNo Then
                Exit Sub
            End If
        End If
        rs.Close
        iDorsal = Val(tbDorsal.Text)
        Set rs = db.OpenRecordset("SELECT num_dorsal FROM dorsales d, parejas p WHERE cod_competicion = " & tbCodComp.Text & " AND p.codigo = d.cod_pareja AND nombre_hombre = '" & sResultado(2) & "' AND nombre_mujer = '" & sResultado(3) & "'", dbOpenSnapshot)
        If Not rs.EOF Then
            If MsgBox(mml_FRASE0350 & rs!num_dorsal & mml_FRASE0351, vbYesNo Or vbInformation, mml_FRASE0084) = vbYes Then
                iDorsal = rs!num_dorsal
            End If
        End If
        rs.Close
        If iCodPareja > 0 Then
            sSQL = "INSERT INTO dorsales VALUES(" & MaxCod("dorsales") & "," & iDorsal & "," & tbCodCateg.Text & "," & Mid$(cbFase.Text, 1, 3) & "," & iCodPareja & ",0," & cbRepesca.Value & ")"
            Debug.Print sSQL
            db.Execute (sSQL)
            tbDorsal.Text = tbDorsal.Text + 1
            Call cmdActualizar_Click
        End If
        chkAgregarUltima.Value = 0
    End If
End Sub


Private Sub cmdPresente_Click()
Dim iCodDorsal As Long
    
    dgDorsales.Col = 0
    If dgDorsales.Text = "" Or tbCodCateg.Text = "" Or cbFase.Text = "" Then
        MsgBox mml_FRASE0344, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0345, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    dgDorsales.Col = 0
    iCodDorsal = dgDorsales.Text
    If chkPresentes.Value = 1 Then
        If Val(cbFase.Text) = 0 Then
            CamposSinCubrir
            Exit Sub
        End If
        db.Execute ("UPDATE dorsales SET no_presente = 0 WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & Val(cbFase.Text))
    Else
        db.Execute ("UPDATE dorsales SET no_presente = 0 WHERE codigo =" & iCodDorsal)
    End If
    Call cmdActualizar_Click


End Sub

Private Sub cmdQuitar_Click()
Dim iCodDorsal As Long
Dim aDorsales(200) As Long
Dim iDorsal, i As Integer
Dim iCDorsal As Integer
Dim rs As Recordset
Dim sCausaEliminacion As String
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    If MsgBox(G_PREGUNTA_OPERACION, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If

    dgDorsales.Col = 0
    If dgDorsales.Text = "" Or tbCodCateg.Text = "" Or Val(cbFase.Text) = 0 Then
        MsgBox mml_FRASE0344, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    sCausaEliminacion = InputBox(mml_FRASE1190, "", mml_FRASE0298)
    If sCausaEliminacion = "" Then
        Exit Sub
    End If
    
    If chkTodos.Value = 1 Then
        If MsgBox(mml_FRASE1101, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            Exit Sub
        End If
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones p WHERE  p.cod_categoria = " & tbCodCateg.Text & _
            " AND p.fase = " & Val(cbFase.Text) & " AND p.repesca = " & cbRepesca.Value & _
            " AND p.num_dorsal IN (SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text & _
            " AND fase = " & Val(cbFase.List(cbFase.ListIndex)) & " AND repesca = " & cbRepesca.Value & ")", dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            rs.Close
            MsgBox mml_FRASE1100, vbOKOnly Or vbCritical, G_MSG_ERROR
            Exit Sub
        End If
        rs.Close
        db.Execute ("DELETE FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & Val(cbFase.List(cbFase.ListIndex)) & " AND repesca = " & cbRepesca.Value)
    Else
        Dim sDorsal As String
        Dim sNombreHombre As String
        Dim sNombreMujer As String
    
        dgDorsales.Col = 0
        iCodDorsal = dgDorsales.Text
        
        ' Recuperamos los dorsales
        dgDorsales.Col = 0
        iCDorsal = 0
        For Each iDorsal In dgDorsales.SelBookmarks
            dgDorsales.Col = 0
            dgDorsales.Bookmark = iDorsal
            aDorsales(iCDorsal) = Val(dgDorsales.Text)
            Inc iCDorsal
            
            dgDorsales.Col = 1
            sDorsal = dgDorsales.Text
            dgDorsales.Col = 8
            sNombreHombre = dgDorsales.Text
            dgDorsales.Col = 9
            sNombreMujer = dgDorsales.Text
            
            InsertarEliminado tbCodCateg.Text, tbDescCateg.Text, tbCodComp.Text, tbDescComp.Text, Val(cbFase.List(cbFase.ListIndex)), cbRepesca.Value, sDorsal, sNombreHombre, sNombreMujer, sCausaEliminacion
        Next
        
        For i = 0 To iCDorsal - 1
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones p WHERE  p.cod_categoria = " & tbCodCateg.Text & _
                " AND p.fase = " & Val(cbFase.Text) & " AND p.repesca = " & cbRepesca.Value & _
                " AND p.num_dorsal = (SELECT num_dorsal FROM dorsales WHERE codigo = " & aDorsales(i) & ")", dbOpenSnapshot)
            If rs.Fields(0) > 0 Then
                rs.Close
                MsgBox mml_FRASE1100 & " Cod. " & aDorsales(i), vbOKOnly Or vbCritical, G_MSG_ERROR
                Exit Sub
            End If
            rs.Close
            db.Execute ("DELETE FROM dorsales WHERE codigo =" & aDorsales(i))
        Next
    End If
    Call cmdActualizar_Click
    
    Exit Sub

error:
    ProcesarError "cmdQuitar_Click"
End Sub

Private Sub CommandButton1_Click()

End Sub


Private Sub cmdRegenerarDorsales_Click()
Dim iDorsal As Integer
Dim iNumDorsal As Integer
Dim rsDorsal As Recordset
Dim rs As Recordset
Dim sBusqueda As String
Dim bRecuperarDorsal As Boolean

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE0307, vbYesNo Or vbCritical, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    iDorsal = Val(InputBox(mml_FRASE1181, "", iMinDorsalOficial(tbCodComp.Text)))
    If iDorsal = 0 Then Exit Sub
    
    If MsgBox(mml_FRASE0352, vbYesNo Or vbCritical, mml_FRASE0084) = vbYes Then
        If MsgBox(mml_FRASE1179, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
            sBusqueda = " AND d.cod_categoria = " & tbCodCateg.Text & " "
        Else
            sBusqueda = ""
        End If
        If MsgBox(mml_FRASE0961, vbYesNo Or vbCritical, mml_FRASE0084) = vbYes Then
            Exit Sub
        End If
        If MsgBox(mml_FRASE1180, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
            bRecuperarDorsal = True
        Else
            bRecuperarDorsal = False
        End If
        'Seleccionamos todas las parejas convencionales de la competición y asignamos dorsales muy altos
        ' Para evitar problemas de colisión
        sSQL = "SELECT d.*, c.codigo, c.cod_modalidad as mod, p.cod_modalidad, p.nif_hombre, p.nif_mujer, p.nombre_hombre, p.nombre_mujer FROM dorsales d, categorias c, parejas p WHERE p.codigo = d.cod_pareja AND num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND c.codigo = d.cod_categoria AND d.cod_categoria in (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ") " & sBusqueda & " ORDER BY p.cod_modalidad, c.codigo"
        Debug.Print sSQL
        Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
        iNumDorsal = C_DORSAL_INI_RENUMERACION
        While Not rs.EOF
            db.Execute ("UPDATE dorsales SET num_dorsal = " & iNumDorsal & " WHERE codigo = " & rs![d.codigo])
            Inc iNumDorsal
            rs.MoveNext
        Wend
        
        'Seleccionamos todas las parejas convencionales de la competición y asignamos dorsales
        sSQL = "SELECT d.*, c.codigo, c.cod_modalidad as mod, p.cod_modalidad, p.nif_hombre, p.nif_mujer, p.nombre_hombre, p.nombre_mujer FROM dorsales d, categorias c, parejas p WHERE p.codigo = d.cod_pareja AND num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND c.codigo = d.cod_categoria AND d.cod_categoria in (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")  " & sBusqueda & " ORDER BY p.cod_modalidad, c.codigo"
        Debug.Print sSQL
        Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
        While Not rs.EOF
            If bRecuperarDorsal Then
                ' Si esta pareja ya tiene dorsal en otra modalidad
                sSQL = "SELECT num_dorsal FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND cod_categoria < " & rs![c.codigo] & " AND cod_modalidad <= " & rs!Mod & " AND cod_competicion = " & tbCodComp.Text & " AND ((nif_hombre <> '' AND nif_hombre = '" & rs!nif_hombre & "' AND nif_mujer = '" & rs!nif_mujer & "') OR (nombre_hombre <> '' AND nombre_hombre ='" & rs!nombre_hombre & "' AND nombre_mujer = '" & rs!nombre_mujer & "'))"
                Debug.Print sSQL
                Set rsDorsal = db.OpenRecordset(sSQL, dbOpenSnapshot)
                If Not rsDorsal.EOF Then
                    iNumDorsal = rsDorsal!num_dorsal
                Else
                    iNumDorsal = iDorsal
                    Inc iDorsal
                End If
                rsDorsal.Close
            Else
                iNumDorsal = iDorsal
                Inc iDorsal
            End If
            db.Execute ("UPDATE dorsales SET num_dorsal = " & iNumDorsal & " WHERE codigo = " & rs![d.codigo])
            rs.MoveNext
        Wend
        rs.Close
        cmdActualizar_Click
        MsgBox mml_FRASE0353, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
End Sub

Public Sub cmdRegenHorario_Click()
    If G_COUNTRY Then
        RegenHorarioCountry
    Else
        RegenHorario
    End If
End Sub
Function RecuperarOrdenCfg(iMod As Integer, aCateg() As String)
    Select Case iMod
        Case 1
                RecuperarOrdenCfg = DividirCampo(G_ORDEN_CATEG_EST, aCateg, ":")
        Case 2
                RecuperarOrdenCfg = DividirCampo(G_ORDEN_CATEG_LAT, aCateg, ":")
        Case 3
                RecuperarOrdenCfg = DividirCampo(G_ORDEN_CATEG_COM, aCateg, ":")
    End Select
End Function

Private Sub RegenHorario()
Dim rs As Recordset
Dim iOrden As Integer
Dim sHora As String
Dim i As Integer
Dim aCateg(20) As String, iContCat As Integer
Dim iModalidad As Integer

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    
    iOrden = 12
    'Reordenamos las categorias para poder trabajar con ellas
    ReordenarHorario 10000, 1
    
    iModalidad = Val(Mid$(C_ORDEN_HORARIO_MODALIDAD, 1, 1))
    iContCat = RecuperarOrdenCfg(iModalidad, aCateg)
    For i = 0 To iContCat - 1
        'Primero ordenamos las categorías que hemos configurado en parámetros
        OrdenarModalidadHorario Str$(iModalidad), iOrden, aCateg(i)
        'Si nos hemos olvidado de alguna se reordenan el resto
        OrdenarModalidadHorario Str$(iModalidad), iOrden, aCateg(i), True
    Next
    
    iModalidad = Val(Mid$(C_ORDEN_HORARIO_MODALIDAD, 2, 1))
    iContCat = RecuperarOrdenCfg(iModalidad, aCateg)
    For i = 0 To iContCat - 1
        OrdenarModalidadHorario Str$(iModalidad), iOrden, aCateg(i)
        OrdenarModalidadHorario Str$(iModalidad), iOrden, aCateg(i), True
    Next
    
    iModalidad = Val(Mid$(C_ORDEN_HORARIO_MODALIDAD, 3, 1))
    iContCat = RecuperarOrdenCfg(iModalidad, aCateg)
    For i = 0 To iContCat - 1
        OrdenarModalidadHorario Str$(iModalidad), iOrden, aCateg(i)
        OrdenarModalidadHorario Str$(iModalidad), iOrden, aCateg(i), True
    Next
    
    MsgBox mml_FRASE0354 & Chr$(13) & Chr$(10) & mml_FRASE0355, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub
Function RecuperarOrdenCfgCountry(iMod As Integer, aCateg() As String)
    If Not C_DEBUG Then On Local Error GoTo error
    RecuperarOrdenCfgCountry = DividirCampo(VarCfg(C_BASE_ORDEN_COUNTRY & Trim$(Str$(iMod))), aCateg, ":")
    Exit Function
error:
    ProcesarError "RecuperarOrdenCfgCountry"
End Function

Private Sub RegenHorarioCountry()
Dim rs As Recordset
Dim iOrden As Integer
Dim sHora As String
Dim i As Integer, j As Integer
Dim aCateg(20) As String, iContCat As Integer
Dim iModalidad As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    
    iOrden = 12
    'Reordenamos las categorias para poder trabajar con ellas
    ReordenarHorario 10000, 1
    
    For j = 0 To Len(C_ORDEN_HORARIO_MODALIDAD) / 2
        iModalidad = Val(Mid$(C_ORDEN_HORARIO_MODALIDAD, j * 2 + 1, 2))
        iContCat = RecuperarOrdenCfg(iModalidad, aCateg)
        For i = 0 To iContCat - 1
            OrdenarModalidadHorario Str$(iModalidad), iOrden, aCateg(i)
        Next
    Next
    
    MsgBox mml_FRASE0354 & Chr$(13) & Chr$(10) & mml_FRASE0355, vbOKOnly Or vbInformation, mml_FRASE0086
    Exit Sub
error:
    ProcesarError "RegenHorarioCountry"
End Sub
Sub OrdenarModalidadHorario(sMod As String, iOrden As Integer, sCatExcluidas As String, Optional bResto As Boolean = False)
Dim rs As Recordset
Dim iInicio As Integer
Dim iCateg As Integer
Dim iCateg1 As Integer
Dim iOrdenCat As Integer
Dim sResto As String

    iInicio = 1
    iOrden = iOrden + 30
    If bResto Then
        sResto = " NOT "
    End If
    sSQL = "SELECT numfase - INT(numfase/" & C_FASE_GENERAL_LOOK & " )*(" & C_FASE_GENERAL_LOOK & " -1) as ordenacion,  h.* FROM horario h,categorias c WHERE c.id_categoria " & sResto & " IN " & sCatExcluidas & " AND c.cod_modalidad = " & sMod & " AND h.cod_categoria = c.codigo AND h.cod_competicion = " & tbCodComp.Text & " ORDER BY 1 DESC, h.numfase DESC, h.grupo DESC"
    Debug.Print sSQL
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
    While Not rs.EOF
        If iCateg1 = rs!cod_categoria Then
            iOrdenCat = iOrden + G_SALTO_CATEG / 2 + 1
        ElseIf iCateg = rs!cod_categoria Then
            iOrdenCat = iOrden + G_SALTO_CATEG + 1
        Else
            iOrdenCat = iOrden
        End If
        iCateg1 = iCateg
        db.Execute ("UPDATE horario SET orden = " & iOrdenCat & ", inicio_sesion = " & iInicio & " WHERE orden = " & rs!orden & " AND cod_competicion = " & rs!cod_competicion)
        iCateg = rs!cod_categoria
        iOrden = iOrden + 10
        iInicio = 0
        rs.MoveNext
    Wend
    rs.Close

End Sub
Sub ReordenarHorario(iInicial As Integer, iIncremento As Integer)
Dim rs As Recordset
Dim iOrden As Integer
    
    iOrden = iInicial
    Set rs = db.OpenRecordset("SELECT h.* FROM horario h WHERE h.cod_competicion = " & tbCodComp.Text & " ORDER BY h.numfase DESC, h.grupo", dbOpenSnapshot)
    While Not rs.EOF
        db.Execute ("UPDATE horario SET orden = " & iOrden & ",inicio_sesion = 0 WHERE orden = " & rs!orden & " AND cod_competicion = " & rs!cod_competicion)
        iOrden = iOrden + iIncremento
        rs.MoveNext
    Wend
    rs.Close

End Sub
Public Sub cmdRegenHoras_Click()
    RegenerarHoras Val(tbCodComp.Text)
End Sub

Public Sub RegenerarHoras(lCodComp As Long)
Dim rsHorario As Recordset
Dim sHora As String, dHora As Date, iOrden As Integer, iMaxMin As Integer, iMin As Integer
Dim sPista As String
Dim rs As Recordset

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE1184, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If

    sPista = InputBox(mml_FRASE1177, "", "")
    
    If Val(sPista) > 0 Then
        sPista = " AND grupo LIKE '*(P" & Val(sPista) & ")*' "
    End If

    iOrden = G_SALTO_ORDEN
    sHora = InputBox(mml_FRASE0356)
    
    If Not IsDate(sHora) Then
        MsgBox mml_FRASE0357, vbOKOnly Or vbExclamation, mml_FRASE0096
        Exit Sub
    End If
    
    dHora = CDate(sHora)

    'Regeneramos las horas del horario según el órden actual, solo para los grupos iniciales de cada conjunto de salida a pista
    If C_COUNTRY And G_COUNTRY_COMB_GRUP_HOJAS Then
        Set rsHorario = db.OpenRecordset("SELECT h.*, c.id_categoria FROM horario h, categorias c WHERE h.cod_categoria = c.codigo AND h.cod_competicion = " & lCodComp & sPista & " AND h.inicio_grupo = 1 " & _
                                         " UNION " & _
                                         " SELECT h.*, 0 FROM horario AS h WHERE h.cod_categoria = 0 AND h.cod_competicion = " & lCodComp & sPista & " ORDER BY orden,h.hora")
    Else
        'En caso de baile de salón ignoramos inicio_grupo ya que nunca bailan varios grupos simultáneamente
        Set rsHorario = db.OpenRecordset("SELECT h.*, c.id_categoria FROM horario h, categorias c WHERE h.cod_categoria = c.codigo AND h.cod_competicion = " & lCodComp & sPista & _
                                         " UNION " & _
                                         " SELECT h.*, 0 FROM horario AS h WHERE h.cod_categoria = 0 AND h.cod_competicion = " & lCodComp & sPista & " ORDER BY orden,h.hora")
    End If
    
    
    While Not rsHorario.EOF
        'Los inicios de sesión tienen la primera hora válida
        If (rsHorario!inicio_sesion And C_INICIO_SESION) = 1 Then
            sHora = InputBox(mml_FRASE0358 & Format$(dHora, "hh:mm") & mml_FRASE0359 & rsHorario!grupo, mml_FRASE0360, rsHorario!hora)
    
            If Not IsDate(sHora) Then
                MsgBox mml_FRASE0357, vbOKOnly Or vbExclamation, mml_FRASE0096
                Exit Sub
            End If

            dHora = CDate(sHora)
        End If
        db.Execute "UPDATE horario SET hora = #" & Format$(dHora, "hh:mm:ss") & "# WHERE cod_competicion = " & rsHorario!cod_competicion & " AND orden = " & rsHorario.Fields("orden")
        Dim sCad As String
        Dim iPos As Integer
        
        sCad = rsHorario.Fields("grupo")
        iPos = InStr(sCad, " (")
        If iPos > 0 Then
            sCad = Left(sCad, iPos - 1)
        End If
        'Eliminamos la pista
        If rsHorario!id_categoria = 0 Then
            'Localizamos la duración de la fase en la tabla de separadores de fases
            Set rs = db.OpenRecordset("SELECT duracion FROM faseshorario WHERE des_fase = '" & sCad & "'", dbOpenSnapshot)
            If Not rs.EOF Then
                iMin = rs.Fields("duracion")
            Else
                iMin = G_DURACION_DEFECTO_FASES_SEPARACION
            End If
            rs.Close
        Else
            iMin = MinPorBaile(rsHorario!numfase, rsHorario!id_categoria, rsHorario!cod_competicion, rsHorario!repesca, rsHorario!cod_categoria)
        End If
        If (rsHorario!inicio_sesion And C_NO_ACT_HORA) = 1 Then
            If iMin > iMaxMin Then
                iMaxMin = iMin
            End If
        Else ' Tenemos que actualizar con la mayor duración
            If iMaxMin > iMin Then
                iMin = iMaxMin
            End If
            dHora = DateAdd("n", iMin, dHora)
            iMaxMin = 0
        End If
        
        rsHorario.MoveNext
    Wend
    
    If C_COUNTRY Then
        'Propagamos la hora de los grupos iniciales al resto de los grupos que bailan juntos
        Set rsHorario = db.OpenRecordset("SELECT orden, hora, inicio_grupo, cod_competicion FROM horario h WHERE cod_competicion = " & lCodComp & sPista & " ORDER BY orden,h.hora")
        While Not rsHorario.EOF
            With rsHorario
                If .Fields("inicio_grupo") = 1 Then
                    dHora = .Fields("hora")
                Else
                    db.Execute ("UPDATE horario SET hora = #" & Format$(dHora, "hh:mm:ss") & "# WHERE cod_competicion = " & .Fields("cod_competicion") & " AND orden = " & .Fields("orden"))
                End If
                .MoveNext
            End With
        Wend
    End If
    rsHorario.Close
    
    MsgBox mml_FRASE0361, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Private Sub cmdSalir_Click()
    On Local Error Resume Next
    'MsgBox mml_FRASE1055, vbOKOnly Or vbCritical, G_MSG_AVISO
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCateg.Text = ""
    tbDescCateg.Text = ""
    If tbDescComp.Text <> "" Then Call cmdActualizar_Click
End Sub


Private Sub cmdSigFase_Click()
    If cbFase.ListIndex < cbFase.ListCount - 1 Then
        cbFase.ListIndex = cbFase.ListIndex + 1
    End If
End Sub

Private Sub cmdCambiarDorsal_Click()
    If tbDorsal.Text = "" Or tbCodComp.Text = "" Or tbCodCateg.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0362 & tbDorsal.Text & mml_FRASE0363, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then Exit Sub
    dgDorsales.Col = 0
    If Val(tbDorsal.Text) > 0 Then
        db.Execute ("UPDATE dorsales SET num_dorsal = '" & tbDorsal.Text & "' WHERE codigo = " & dgDorsales.Text)
    End If

End Sub

Private Sub cmdTandas_Click()
    On Local Error Resume Next
    With frmCambiarRonda
        .tbCodComp.Text = tbCodComp.Text
        .tbDescComp.Text = tbDescComp.Text
        .tbCodCat.Text = tbCodCateg.Text
        .tbDescCat.Text = tbDescCateg.Text
        .tbCodFase.Text = Val(cbFase.Text)
        .tbDescFase.Text = sDescFase(.tbCodFase.Text)
        .chkRep.Value = cbRepesca.Value
        .RecargarBailes
        .Show vbModal
    End With
End Sub

Private Sub cmdAusencia_Click()
Dim iCodDorsal As Long
    
    dgDorsales.Col = 0
    If dgDorsales.Text = "" Or tbCodCateg.Text = "" Or cbFase.Text = "" Then
        MsgBox mml_FRASE0344, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0345, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    dgDorsales.Col = 0
    iCodDorsal = dgDorsales.Text
    db.Execute ("UPDATE dorsales SET no_presente = 1 WHERE codigo =" & iCodDorsal)
    Call cmdActualizar_Click


End Sub

Private Sub dgDorsales_DblClick()
    If tbCodComp.Text <> "" Then
        dgDorsales.Col = 2
        tbCodCateg.Text = dgDorsales.Text
        cbFase.ListIndex = 0
        tbDescCateg.Text = sDescCategoria(Val(tbCodCateg.Text))
        cmdActualizar_Click
    End If
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    
    cbOrd.ListIndex = 0
    cbOrden.ListIndex = 0
    dgDorsales.RecordSelectors = True
    
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    If Val(tbCodComp.Text) > 0 Then
        Call cmdActualizar_Click
    End If
    
    If C_DEBUG Then cmdAddDorsales.Visible = True
End Sub

Private Sub spDorsal_SpinDown_Click()
    If Val(tbDorsal.Text) > 0 Then
        tbDorsal.Text = Val(tbDorsal.Text) - 1
    End If
End Sub

Private Sub spDorsal_SpinUp_Click()
    If Val(tbDorsal.Text) < 999 Then
        tbDorsal.Text = Val(tbDorsal.Text) + 1
    End If

End Sub

Private Sub tbCodCateg_GotFocus()
    tbCodCateg.SelStart = 0
    tbCodCateg.SelLength = Len(tbCodCateg.Text)
End Sub

Private Sub tbCodCateg_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbCodCateg_LostFocus()
Dim sCateg As String
    
    If Not C_DEBUG Then On Local Error GoTo error
    If Val(tbCodCateg.Text) > 0 Then
        sCateg = sDescCategoria(tbCodCateg.Text, tbCodComp.Text)
        If sCateg = "" Then
            tbCodCateg.Text = ""
            tbDescCateg.Text = ""
        Else
            tbDescCateg.Text = sCateg
        End If
    Else
        tbCodCateg.Text = ""
        tbDescCateg.Text = ""
    End If
    cbFase.ListIndex = 0
    Exit Sub
error:
    ProcesarError "tbCodCateg_LostFocus"
End Sub


Private Sub tbDescCateg_KeyDown(KeyCode As Integer, Shift As Integer)
    'If (KeyCode = 46 Or KeyCode = 8) And tbDescCateg.SelStart < 6 Then KeyCode = 0
End Sub

Private Sub tbDescCateg_KeyPress(KeyAscii As Integer)
    LimitarDescCateg KeyAscii
    'If tbDescCateg.SelStart < 5 Then KeyAscii = 0
End Sub

Private Sub tbDorsal_Change()
    If Val(tbDorsal.Text) < iMinDorsalOficial(tbCodComp.Text) Then
        tbDorsal.BackColor = &HC0C0FF
    Else
        tbDorsal.BackColor = &HC0E0FF
    End If
End Sub

Private Sub tbDorsal_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub


Sub InsertarEliminado(ByVal lCodCateg As Long, ByVal sDescCateg As String, ByVal lCodComp As Long, ByVal sDescComp As String, ByVal iFase As Integer, ByVal iRepesca As Integer, ByVal iDorsal As Integer, ByVal sNombreHombre As String, ByVal sNombreMujer As String, ByVal sCausa As String)
Dim iFile As Integer
    
    On Local Error Resume Next
    iFile = FreeFile
    Open G_FICHERO_ELIMINADOS For Append As #iFile
    Print #iFile, "Cod.Comp " & lCodComp & " - " & sDescComp & ". Cod.Categ " & lCodCateg & " - " & sDescCateg & ", Fase " & iFase & " Rep " & iRepesca & " Dorsal Nº" & iDorsal & ", " & sNombreHombre & " - " & sNombreMujer & " (" & sCausa & ")"
    Close #iFile
End Sub

