VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0036"
   ClientHeight    =   6255
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9930
   ControlBox      =   0   'False
   Icon            =   "GenerarForms.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   9120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame mrcPDAs 
      Caption         =   "PDAs "
      Height          =   5985
      Left            =   0
      TabIndex        =   54
      Top             =   90
      Visible         =   0   'False
      Width           =   2475
      Begin VB.CommandButton cmdAdmEquipos 
         Height          =   495
         Left            =   810
         Picture         =   "GenerarForms.frx":0BC2
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "mml_FRASE1210"
         Top             =   3690
         Width           =   615
      End
      Begin VB.CommandButton cmdPDA4 
         Height          =   495
         Left            =   810
         Picture         =   "GenerarForms.frx":17BC
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "mml_FRASE0039"
         Top             =   2886
         Width           =   615
      End
      Begin VB.CommandButton cmdPDA3 
         Height          =   495
         Left            =   810
         Picture         =   "GenerarForms.frx":23B6
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "mml_FRASE0039"
         Top             =   2084
         Width           =   615
      End
      Begin VB.CommandButton cmdPDA2 
         Height          =   495
         Left            =   810
         Picture         =   "GenerarForms.frx":2FB0
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "mml_FRASE0039"
         Top             =   1282
         Width           =   615
      End
      Begin VB.CommandButton cmdPDA1 
         Height          =   495
         Left            =   810
         Picture         =   "GenerarForms.frx":3BAA
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "mml_FRASE0039"
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCorregirBD 
      Caption         =   "CorregirBD"
      Height          =   615
      Left            =   3390
      TabIndex        =   53
      Top             =   6990
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Height          =   795
      Left            =   660
      TabIndex        =   47
      Top             =   5340
      Width           =   8445
      Begin VB.CommandButton cmdCopiaDB1 
         Caption         =   "mml_FRASE1143"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2190
         TabIndex        =   51
         Top             =   180
         Width           =   1995
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "mml_FRASE0037"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   90
         TabIndex        =   50
         Top             =   180
         Width           =   1995
      End
      Begin VB.CommandButton cmdCopiaCompleta 
         Caption         =   "mml_FRASE1146"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4260
         TabIndex        =   49
         Top             =   180
         Width           =   1995
      End
      Begin VB.CommandButton cmdBorrarTodo 
         Caption         =   "mml_FRASE1147"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   6360
         TabIndex        =   48
         Top             =   180
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdGenFormularios 
      Caption         =   "Gen. Formularios"
      Height          =   525
      Left            =   5280
      TabIndex        =   45
      Top             =   780
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   12000
      Left            =   8850
      Top             =   2700
   End
   Begin VB.Timer Timer8 
      Interval        =   61450
      Left            =   8160
      Top             =   2070
   End
   Begin VB.Frame frmPeque 
      Height          =   6075
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   2520
      Begin VB.CommandButton cmdCopiaBD 
         Caption         =   "mml_FRASE1143"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   46
         Top             =   3930
         Width           =   1995
      End
      Begin VB.CommandButton cmdHora1 
         Caption         =   "12:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3420
         Width           =   945
      End
      Begin VB.PictureBox picElim 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1440
         Picture         =   "GenerarForms.frx":47A4
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   41
         ToolTipText     =   "Grupos que tiene eliminatorias pendientes"
         Top             =   2940
         Width           =   315
      End
      Begin VB.CommandButton Command13 
         Height          =   495
         Left            =   1200
         Picture         =   "GenerarForms.frx":4C96
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "mml_FRASE0039"
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton Command12 
         Height          =   495
         Left            =   480
         Picture         =   "GenerarForms.frx":5890
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "mml_FRASE0028"
         Top             =   225
         Width           =   615
      End
      Begin VB.CommandButton Command11 
         Height          =   495
         Left            =   480
         Picture         =   "GenerarForms.frx":63B2
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "mml_FRASE0046"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         Height          =   495
         Left            =   1200
         Picture         =   "GenerarForms.frx":6F94
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "mml_FRASE0045"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command9 
         Height          =   495
         Left            =   1200
         Picture         =   "GenerarForms.frx":7D3E
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "mml_FRASE0044"
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Height          =   495
         Left            =   480
         Picture         =   "GenerarForms.frx":9000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "mml_FRASE0043"
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Height          =   495
         Left            =   1200
         Picture         =   "GenerarForms.frx":9BE2
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "mml_FRASE0042"
         Top             =   780
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         Height          =   495
         Left            =   1200
         Picture         =   "GenerarForms.frx":AB8C
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "mml_FRASE0041"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Height          =   495
         Left            =   480
         Picture         =   "GenerarForms.frx":BB36
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "mml_FRASE0040"
         Top             =   780
         Width           =   615
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   480
         Picture         =   "GenerarForms.frx":C718
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   30
         Top             =   2940
         Width           =   315
      End
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   480
         Picture         =   "GenerarForms.frx":CB82
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "mml_FRASE0039"
         Top             =   2400
         Width           =   615
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   960
         Picture         =   "GenerarForms.frx":D77C
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   28
         ToolTipText     =   "mml_FRASE0038"
         Top             =   2940
         Width           =   315
      End
   End
   Begin VB.TextBox tbCompActiva 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1650
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   3915
      Width           =   6615
   End
   Begin VB.PictureBox picEPADance 
      AutoSize        =   -1  'True
      Height          =   2070
      Left            =   4500
      Picture         =   "GenerarForms.frx":DABE
      ScaleHeight     =   2010
      ScaleWidth      =   3630
      TabIndex        =   25
      Top             =   5490
      Visible         =   0   'False
      Width           =   3690
   End
   Begin MSComDlg.CommonDialog DlgFicheros 
      Left            =   270
      Top             =   1710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   40000
      Left            =   6600
      Top             =   840
   End
   Begin VB.Timer Timer6 
      Interval        =   59000
      Left            =   8160
      Top             =   1680
   End
   Begin VB.PictureBox picEPADance2 
      AutoSize        =   -1  'True
      Height          =   2070
      Left            =   765
      Picture         =   "GenerarForms.frx":25810
      ScaleHeight     =   2010
      ScaleWidth      =   3630
      TabIndex        =   23
      Top             =   5490
      Visible         =   0   'False
      Width           =   3690
   End
   Begin VB.Timer tmrCalcularRetraso 
      Left            =   9240
      Top             =   3300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   555
      Left            =   495
      TabIndex        =   21
      Top             =   3870
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox picEPA 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   7290
      Picture         =   "GenerarForms.frx":277B3
      ScaleHeight     =   765
      ScaleWidth      =   1620
      TabIndex        =   20
      Top             =   4590
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   3360
      Picture         =   "GenerarForms.frx":2B881
      ScaleHeight     =   570
      ScaleWidth      =   2595
      TabIndex        =   19
      Top             =   4815
      Width           =   2595
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   40000
      Left            =   7440
      Top             =   840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   58000
      Left            =   7020
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   8160
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   8160
      Top             =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CN"
      Height          =   495
      Left            =   615
      TabIndex        =   18
      Top             =   4905
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   8820
      Top             =   3300
   End
   Begin VB.Frame frmGrande 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9930
      Begin VB.CommandButton cmdActivarPDA 
         Height          =   495
         Left            =   6180
         Picture         =   "GenerarForms.frx":2EF73
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "mml_FRASE0039"
         Top             =   165
         Width           =   615
      End
      Begin VB.PictureBox picElimPorSalir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8340
         Picture         =   "GenerarForms.frx":2F3B5
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   40
         ToolTipText     =   "Grupos que tiene eliminatorias pendientes"
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox picUltimoGrupo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7980
         Picture         =   "GenerarForms.frx":2F8A7
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   24
         ToolTipText     =   "mml_FRASE0038"
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdRecOptico 
         Height          =   495
         Left            =   5490
         Picture         =   "GenerarForms.frx":2FBE9
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "mml_FRASE0039"
         Top             =   165
         Width           =   615
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7605
         Picture         =   "GenerarForms.frx":307E3
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   17
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdPuntBaile 
         Height          =   495
         Left            =   765
         Picture         =   "GenerarForms.frx":30C4D
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "mml_FRASE0040"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdHora 
         Caption         =   "12:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton cmdImpHojasPuntuac 
         Height          =   495
         Left            =   4110
         Picture         =   "GenerarForms.frx":3182F
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "mml_FRASE0041"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdImpResultados 
         Height          =   495
         Left            =   3450
         Picture         =   "GenerarForms.frx":327D9
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "mml_FRASE0042"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdDescalif 
         Height          =   495
         Left            =   2055
         Picture         =   "GenerarForms.frx":33783
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "mml_FRASE0043"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdPublicar 
         Height          =   495
         Left            =   4815
         Picture         =   "GenerarForms.frx":34365
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "mml_FRASE0044"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdCalcular 
         Height          =   495
         Left            =   2760
         Picture         =   "GenerarForms.frx":35627
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "mml_FRASE0045"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdHorario 
         Height          =   495
         Left            =   1410
         Picture         =   "GenerarForms.frx":363D1
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "mml_FRASE0067"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdDorsales 
         Height          =   495
         Left            =   120
         Picture         =   "GenerarForms.frx":36FB3
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "mml_FRASE0028"
         Top             =   165
         Width           =   615
      End
      Begin VB.CommandButton cmdEscanear 
         Height          =   495
         Left            =   6870
         Picture         =   "GenerarForms.frx":37AD5
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "mml_FRASE0039"
         Top             =   165
         Width           =   615
      End
   End
   Begin VB.PictureBox picEPA1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   4200
      Picture         =   "GenerarForms.frx":386CF
      ScaleHeight     =   570
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   780
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   420
      Left            =   1650
      TabIndex        =   43
      Top             =   3480
      Width           =   6720
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "mml_FRASE1084"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   0
         TabIndex        =   44
         Top             =   30
         Width           =   6675
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2130
      Picture         =   "GenerarForms.frx":394A1
      ScaleHeight     =   2175
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   1290
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   3810
      Picture         =   "GenerarForms.frx":459C3
      ScaleHeight     =   2415
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   615
         Left            =   2040
         TabIndex        =   52
         Top             =   120
         Width           =   75
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   5970
      Picture         =   "GenerarForms.frx":58B45
      ScaleHeight     =   2415
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   1350
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "mml_FRASE0047"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   2700
      TabIndex        =   4
      Top             =   4290
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "mml_FRASE0047"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2670
      TabIndex        =   3
      Top             =   4320
      Width           =   4575
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "mml_FRASE0029"
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "mml_FRASE0048"
      Begin VB.Menu mnuACompeticiones 
         Caption         =   "mml_FRASE0049"
      End
      Begin VB.Menu mnuGrupos 
         Caption         =   "mml_FRASE0032"
      End
      Begin VB.Menu mnuAJueces 
         Caption         =   "mml_FRASE0050"
      End
      Begin VB.Menu mnuAParejas 
         Caption         =   "mml_FRASE0034"
      End
      Begin VB.Menu mnuADorsales 
         Caption         =   "mml_FRASE0028"
      End
      Begin VB.Menu mnuJuecesBailes 
         Caption         =   "mml_FRASE0051"
      End
      Begin VB.Menu mnuInsertarSociosAnulados 
         Caption         =   "mml_FRASE0052"
      End
      Begin VB.Menu mnuImportacion 
         Caption         =   "mml_FRASE0053"
      End
      Begin VB.Menu mnuImpComp 
         Caption         =   "mml_FRASE0054"
      End
      Begin VB.Menu mnuCopiarPart 
         Caption         =   "mml_FRASE0055"
      End
      Begin VB.Menu mnuControlPagos 
         Caption         =   "mml_FRASE1023"
      End
      Begin VB.Menu mnuImportarProBaile 
         Caption         =   "mml_FRASE1049"
      End
      Begin VB.Menu mnuAgrupacionManual 
         Caption         =   "mml_FRASE1010"
      End
      Begin VB.Menu mnuComprobarComp 
         Caption         =   "mml_FRASE1068"
      End
      Begin VB.Menu mnuRecogidaDorsales 
         Caption         =   "mml_FRASE1110"
      End
      Begin VB.Menu mnuVerFicheroEliminados 
         Caption         =   "mml_FRASE1191"
      End
      Begin VB.Menu mnuCargarParejasAEBDC 
         Caption         =   "mml_FRASE1232"
      End
   End
   Begin VB.Menu mnuCompeticion 
      Caption         =   "mml_FRASE0056"
      Begin VB.Menu mnuPuntBaile 
         Caption         =   "mml_FRASE0040"
      End
      Begin VB.Menu mnuPuntBaile2 
         Caption         =   "mml_FRASE0005"
      End
      Begin VB.Menu mnuPuntBaile3 
         Caption         =   "mml_FRASE0014"
      End
      Begin VB.Menu mnuPuntBaile4 
         Caption         =   "mml_FRASE0015"
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuElacePPC 
         Caption         =   "mml_FRASE0063"
      End
      Begin VB.Menu mnuEnlacePPC1 
         Caption         =   "mml_FRASE0016"
      End
      Begin VB.Menu mnuEnlacePPC2 
         Caption         =   "mml_FRASE0182"
      End
      Begin VB.Menu mnuEnlacePPC3 
         Caption         =   "mml_FRASE0869"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnlacePPC_HTML 
         Caption         =   "mml_FRASE1217"
      End
      Begin VB.Menu mnuEnlacePPC_HTML1 
         Caption         =   "mml_FRASE1218"
      End
      Begin VB.Menu mnuEnlacePPC_HTML2 
         Caption         =   "mml_FRASE1219"
      End
      Begin VB.Menu mnuEnlacePPC_HTML3 
         Caption         =   "mml_FRASE1220"
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAPuntuaciones 
         Caption         =   "mml_FRASE0057"
      End
      Begin VB.Menu mnuCalcular 
         Caption         =   "mml_FRASE0058"
      End
      Begin VB.Menu mnuPublicar 
         Caption         =   "mml_FRASE0186"
      End
      Begin VB.Menu mnuDescalif 
         Caption         =   "mml_FRASE0043"
      End
      Begin VB.Menu mnuHora 
         Caption         =   "mml_FRASE0060"
      End
      Begin VB.Menu mnuEnviarCorreo 
         Caption         =   "mml_FRASE0061"
      End
      Begin VB.Menu mnuAsocCompHorario 
         Caption         =   "mml_FRASE0062"
      End
      Begin VB.Menu mnuProDance 
         Caption         =   "mml_FRASE0064"
      End
      Begin VB.Menu mnuSorteoRondas 
         Caption         =   "mml_FRASE1103"
      End
      Begin VB.Menu mnuPaneles 
         Caption         =   "mml_FRASE1094"
      End
   End
   Begin VB.Menu mnuImprimir 
      Caption         =   "mml_FRASE0065"
      Begin VB.Menu mnuImprimirCateg 
         Caption         =   "mml_FRASE0066"
      End
      Begin VB.Menu mnuImpHorario 
         Caption         =   "mml_FRASE0067"
      End
      Begin VB.Menu mnuImprimirParticipantes 
         Caption         =   "mml_FRASE0030"
      End
      Begin VB.Menu mnuImprimirFinal 
         Caption         =   "mml_FRASE0068"
      End
      Begin VB.Menu mnuImpHojasPuntuaciones 
         Caption         =   "mml_FRASE0069"
      End
      Begin VB.Menu mnuImpOrdenCombinado 
         Caption         =   "mml_FRASE0070"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuImpInternet 
         Caption         =   "mml_FRASE0071"
      End
      Begin VB.Menu mnuImpDiplomas 
         Caption         =   "mml_FRASE1076"
      End
   End
   Begin VB.Menu mnuLecVis 
      Caption         =   "mml_FRASE0072"
      Begin VB.Menu mnuResultados 
         Caption         =   "mml_FRASE0073"
      End
      Begin VB.Menu mnuLecOptica 
         Caption         =   "mml_FRASE0074"
      End
   End
   Begin VB.Menu mnuGenerador 
      Caption         =   "mml_FRASE0027"
      Begin VB.Menu mnuGenNumBase 
         Caption         =   "mml_FRASE0076"
      End
      Begin VB.Menu mnuIntroNumBase 
         Caption         =   "mml_FRASE0183"
      End
   End
   Begin VB.Menu mnuArchivoAvanzado 
      Caption         =   "mml_FRASE0078"
      Begin VB.Menu mnuTablasDatos 
         Caption         =   "mml_FRASE0991"
      End
      Begin VB.Menu mnuBar91 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnaPagPorJuez 
         Caption         =   "mml_FRASE1003"
      End
      Begin VB.Menu mnuHojasOpticas 
         Caption         =   "mml_FRASE0985"
      End
      Begin VB.Menu mnuBar85 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPuntuacionesPorJuez 
         Caption         =   "mml_FRASE1004"
      End
      Begin VB.Menu mnuBar86 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenAutoPPC 
         Caption         =   "mml_FRASE1020"
      End
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "mml_FRASE0080"
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' NOTAS *******************************************************************************
'En la tabla horario:
    'h.cod_baile = 0 = Todos los bailes
    'h.cod_baile < 0  = indica el codigo de baile Std más bajo ya bailado
    'h.cod_baile = n = baile

'**************************************************************************************

Option Explicit
Const C_NUM_CAMPOS_FICHERO_PARTICIPANTES = 43
Dim Msj As String
Dim iContTimer As Integer
     
Sub ActualizacionBaseDatos()
    On Local Error Resume Next
    'Versión 2.6.1 Mejoras en importación de participantes desde fichero de texto
    db.Execute "ALTER TABLE [lista de participantes] ADD campo43 varchar(30)"
    'Versión 2.6 Mejoras para Country
    db.Execute "ALTER TABLE Modalidad ADD orden integer"
    db.Execute "ALTER TABLE DescCategoria ADD orden integer"
    db.Execute "ALTER TABLE GruposEdad ADD orden integer"
    db.Execute "ALTER TABLE Horario ADD inicio_grupo integer"
    'Versión 2.5
    db.Execute "ALTER TABLE Competiciones ADD aebdc_codigo varchar(10)"
    db.Execute "ALTER TABLE Parejas ADD aebdc_codigo integer"
    
    db.Execute "CREATE TABLE [cfg_html] (fase varchar(15),cod_comp integer,cod_categoria long,cod_fase integer,cod_rep integer,cod_baile integer)"
    db.Execute "CREATE TABLE [jueces_html] (cod_juez varchar(4), cod_categoria long, id_pda integer)"
    db.Execute "CREATE TABLE [bateria_html] (id_pda integer, bateria integer, hora_op varchar(15))"
    db.Execute "DELETE FROM cfg_html"
    db.Execute "INSERT INTO cfg_html VALUES ('ACTUAL',0,0,0,0,0)"
    db.Execute "INSERT INTO cfg_html VALUES ('SIGUIENTE',0,0,0,0,0)"
    db.Execute "DELETE FROM jueces_html"
    db.Execute "DELETE FROM bateria_html"
    db.Execute "ALTER TABLE Bailes ADD abrev varchar(5)"
    'Version 2.4.2
    db.Execute "CREATE TABLE Bateria (id_juez varchar(5), nivel integer, fecha varchar(15), hora varchar(15), tiempo_descarga integer, cargando integer)"
    'Version 2.4.1
    db.Execute "CREATE TABLE Errores (version varchar(50), modulo varchar(50), descripcion varchar(150), hora varchar(50))"
    'Version 2.4
    db.Execute "CREATE TABLE Paneles (cod_panel varchar(2), cod_categoria long)"
    db.Execute "CREATE TABLE Juez_Panel (cod_competicion integer, cod_juez long, cod_panel varchar(2), id_juez varchar(2), pasos integer)"
    'Version 2.2
    db.Execute "CREATE TABLE DescCategoria (cod_desccategoria integer, descripcion varchar(50), modalidades varchar(50))"
    db.Execute "CREATE TABLE EnlaceProBaile (cod_competicion integer, cod_pareja_probaile integer, dorsal_probaile integer, cod_pareja integer)"
    db.Execute "CREATE TABLE Agrupaciones (cod_competicion integer, modalidad varchar(50), categoria varchar(50), grupos varchar(50), cont integer, cod_grupos varchar(50), cod_modalidad integer)"
    db.Execute "ALTER TABLE EnlaceProBaile ADD cod_competicion_probaile integer"
    db.Execute "CREATE TABLE ResultadosFinales (cod_competicion integer, cod_categoria integer, dorsal integer, cod_baile integer, posicion char(5))"
    db.Execute "ALTER TABLE Juez_categ ALTER id_juez varchar(2)"
    db.Execute "ALTER TABLE categorias ADD imprimir_una_hoja_puntuaciones integer"
    db.Execute "CREATE TABLE RecogidaDorsales (num_dorsal integer, cod_categoria integer, cod_pareja integer)"
    db.Execute "ALTER TABLE RecogidaDorsales ADD PRIMARY KEY (num_dorsal, cod_categoria)"
    db.Execute "ALTER TABLE [Lista de participantes] ADD campo42 varchar(10)"
    db.Execute "CREATE TABLE [Lista de participantes2] (campo1 varchar(80),campo2 varchar(80),campo3 varchar(80),campo4 varchar(80),campo5 varchar(80),campo6 varchar(80),campo7 varchar(80),campo8 varchar(80),campo9 varchar(80),campo10 varchar(80),campo11 varchar(80),campo12 varchar(80),campo13 varchar(80),campo14 varchar(80),campo15 varchar(80),campo16 varchar(80),campo17 varchar(80),campo18 varchar(80),campo19 varchar(80),campo20 varchar(80),campo21 varchar(80),campo22 varchar(80),campo23 varchar(80),campo24 varchar(80),campo25 varchar(80),campo26 varchar(80),campo27 varchar(80),campo28 varchar(80),campo29 varchar(80),campo30 varchar(80),campo31 varchar(80),campo32 varchar(80),campo33 varchar(80),campo34 varchar(80),campo35 varchar(80),campo36 varchar(80),campo37 varchar(80),campo38 varchar(80),campo39 varchar(80),campo40 varchar(80),campo41 varchar(80),campo42 varchar(80))"

    'Versión 2.5 -> Incorporar generación de fichero exportación AEBDC
    If VersionBD() = 0 Then
        db.Execute "CREATE TABLE [version] (version_bd integer)"
        db.Execute "INSERT INTO version VALUES (1)"
    End If
    If VersionBD() < 3 Then
        db.Execute "UPDATE version SET version_bd = 3"
        db.Execute "ALTER TABLE [Modalidad] ADD aebdc_codigo varchar(2)"
        db.Execute "INSERT INTO Modalidad VALUES ('ACTUAL','10 Bailes','10')"
        DoEvents
        Sleep 2000
        db.Execute "UPDATE Modalidad SET aebdc_codigo = 'ST' WHERE codigo = 1"
        db.Execute "UPDATE Modalidad SET aebdc_codigo = 'LA' WHERE codigo = 2"
        db.Execute "UPDATE Modalidad SET aebdc_codigo = 'NA' WHERE codigo = 3,4" ' No aplicable
        
        db.Execute "ALTER TABLE [gruposedad] ADD aebdc_codigo varchar(2)"
        DoEvents
        Sleep 2000
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '10' WHERE codigo = 1" 'Juvenil
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '20' WHERE codigo = 2" 'Junior I
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '30' WHERE codigo = 3" 'Junior II
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '40' WHERE codigo = 4" 'Youth
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '50' WHERE codigo = 5" 'Adulto I
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '60' WHERE codigo = 6" 'Adulto II
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '70' WHERE codigo = 7" 'Senior I
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '80' WHERE codigo = 8" 'Senior II
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '90' WHERE codigo = 9" 'Senior III
        db.Execute "UPDATE gruposedad SET aebdc_codigo = '99' WHERE codigo = 10" 'TOP Senior
        db.Execute "INSERT INTO gruposedad VALUES (11,'Top Senior','TSen','99')"
        
        db.Execute "ALTER TABLE [competiciones] ADD aebdc_codigo varchar(7)"
        Sleep 2000
        db.Execute "UPDATE competiciones SET aebdc_codigo = '0'"
        
        db.Execute "ALTER TABLE [parejas] ADD aebdc_codigo integer"
        Sleep 2000
        db.Execute "UPDATE parejas SET aebdc_codigo = 0"
    ElseIf VersionBD() < 4 Then
        db.Execute "UPDATE version SET version_bd = 4"
        db.Execute "CREATE TABLE [parejas_aebdc] (aebdc_codigo varchar(8),num_socio_hombre varchar(10),nombre_hombre varchar(100),num_socio_mujer varchar(60),nombre_mujer varchar(100), escuelas varchar(100))"
    End If
End Sub
Function VersionBD() As Integer
Dim rs As Recordset
    If Not C_DEBUG Then On Local Error GoTo error
        
    Set rs = db.OpenRecordset("SELECT version_bd FROM version", dbOpenSnapshot)
    If Not rs.EOF Then
        VersionBD = rs.Fields("version_bd")
    Else
        VersionBD = 0
    End If
    rs.Close
        
    Exit Function
error:
    ProcesarError "VersionBD"
End Function

Private Sub cmdAdmEquipos_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    Shell G_ADM_EQUIPOS, vbNormalFocus
    Exit Sub
error:
    ProcesarError "cmdAdmEquipos_Click"
End Sub

Private Sub cmdCorregirBD_Click()
    If InputBox("Introduzca código") = "1060" Then
        'Corrección del error por el que se creaba la tabla con campos char y no varchar en una versión antigüa
        db.Execute "ALTER TABLE Agrupaciones ALTER modalidad varchar(50)"
        db.Execute "ALTER TABLE Agrupaciones ALTER categoria varchar(50)"
        db.Execute "ALTER TABLE Agrupaciones ALTER grupos varchar(50)"
        db.Execute "ALTER TABLE Agrupaciones ALTER cod_grupos varchar(50)"
        'Corrección del error de modificación a char(2) en vez de varchar(2)
        db.Execute "UPDATE Juez_Categ SET id_juez = Trim(id_juez)"
        db.Execute "UPDATE cfg SET valor = 'nombre_hombre' WHERE variable = 'orden_parejas' AND valor = 'codigo'"
        'Actualización de versiones viejas
        db.Execute "ALTER TABLE horario ADD num_dorsales integer"
        db.Execute "ALTER TABLE horario ADD num_grupo integer"
        
        MsgBox "Operación realizada", vbOKOnly Or vbInformation, "Aviso"
    End If

End Sub
Private Sub cmdAct_Click()
    Form_Load
    Form_Paint
End Sub

Private Sub cmdBorrarTodo_Click()
Dim sDirFichas As String
Dim sDirFichasCopia As String
Dim sPath As String
Dim sCad As String
    If Not ExisteCompeticion(CodCompActiva) Then
        MsgBox mml_FRASE1151, vbOKOnly Or vbCritical, G_MSG_ERROR
        Exit Sub
    End If
    If MsgBox(mml_FRASE1150, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
        If InputBox(mml_FRASE1149, G_MSG_AVISO, "") = "AUTORIZO" Then
            If MsgBox(mml_FRASE1148, vbYesNo, G_MSG_PREGUNTA) = vbYes Then
                CopiarBD True
                Sleep 5000
                MsgBox mml_FRASE1152, vbOKOnly Or vbInformation, G_MSG_MENSAJE
            End If
        
            BorrarPuntuaciones CodCompActiva
        
            sDirFichas = VarCfg("dir_fichas")
            sDirFichasCopia = VarCfg("pda_copia_fich")
            sPath = sExtraerPath(VarCfg("dir_fichas"))
            sCad = "cmd /C " & G_PATH_ESCRUTINIO & "COPIA_BD\BORRAR.BAT """ & sDirFichas & """"
            sCad = "cmd /C " & G_PATH_ESCRUTINIO & "COPIA_BD\BORRAR.BAT """ & sDirFichasCopia & """"
            Shell sCad, vbNormalFocus
            
            Sleep 2000
            MsgBox mml_FRASE1152, vbOKOnly Or vbInformation, G_MSG_MENSAJE
            On Local Error Resume Next
            
            CrearDirectorios sDirFichas
            CrearDirectorios sDirFichasCopia
            MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
        End If
    End If
End Sub

Private Sub cmdCalcular_Click()
    frmCalcular.Show 1
End Sub

Private Sub CopiarBD(Optional bCopiarTodo As Boolean = False)
Dim sCad As String
Dim sCopiaFich As String
Dim sNombreCopia As String
    If Not C_DEBUG Then On Local Error GoTo error
    sNombreCopia = sDescCompeticion(CodCompActiva)
    If sNombreCopia <> "" Then
        CambiarCadena " ", "_", sNombreCopia
    Else
        sNombreCopia = "COPIA"
    End If
    If bCopiarTodo Then
        sNombreCopia = sNombreCopia & "_COPIA_COMPLETA"
        sNombreCopia = InputBox(mml_FRASE1145, G_MSG_PREGUNTA, sNombreCopia)
        sCopiaFich = "COPIA_FICH"
    Else
        sNombreCopia = sNombreCopia & Format$(Now, "yy-mm-dd") & "." & Format(Time, "hh_mm_ss")
        sNombreCopia = InputBox(mml_FRASE1145, G_MSG_PREGUNTA, sNombreCopia)
        If sNombreCopia <> "" Then
            If MsgBox(mml_FRASE1144, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
                sCopiaFich = "COPIA_FICH"
            Else
                sCopiaFich = "NO_COPIA_FICH"
            End If
        End If
    End If
    If sNombreCopia <> "" Then
        sCad = "cmd /C " & G_PATH_ESCRUTINIO & "COPIA_BD\COPIAR.BAT """ & G_PATH_ESCRUTINIO & """ """ & G_DIR_COPIA_BD & "\" & sNombreCopia & "\"" " & sCopiaFich
        Shell sCad, vbNormalFocus
    End If
    Exit Sub
error:
    ProcesarError "cmdCopiaBD_Click"
End Sub

Private Sub cmdCopiaBD_Click()
    CopiarBD
End Sub

Private Sub cmdCopiaCompleta_Click()
    CopiarBD True
End Sub

Private Sub cmdCopiaDB1_Click()
    CopiarBD
End Sub



Private Sub cmdGenFormularios_Click()
    If InputBox("Introduzca código") = "1060" Then
        frmGenerarFormularios.Show vbModal
    End If
End Sub

Private Sub cmdHorario_Click()
    On Local Error Resume Next
    frmTablasDatos.Show vbModal

End Sub

Private Sub cmdPDA1_Click()
    On Local Error Resume Next
    frmEnlacePPC_HTML.Show
End Sub

Private Sub cmdPDA2_Click()
    On Local Error Resume Next
    frmEnlacePPC_HTML1.Show

End Sub

Private Sub cmdPDA3_Click()
    On Local Error Resume Next
    frmEnlacePPC_HTML2.Show

End Sub

Private Sub cmdPDA4_Click()
    On Local Error Resume Next
    frmEnlacePPC_HTML3.Show

End Sub

Private Sub cmdRecOptico_Click()
    frmRecOptico.Show vbNomodal, Me
End Sub

Private Sub Command10_Click()
cmdCalcular_Click
End Sub

Private Sub Command11_Click()
    cmdHorario_Click
End Sub

Private Sub Command12_Click()
cmdDorsales_Click
End Sub

Private Sub Command13_Click()
cmdEscanear_Click
End Sub


Private Sub Command3_Click()
cmdRecOptico_Click
End Sub

Private Sub Command4_Click()
cmdPuntBaile_Click
End Sub



Private Sub Command6_Click()
cmdImpHojasPuntuac_Click
End Sub

Private Sub Command7_Click()
cmdImpResultados_Click
End Sub

Private Sub Command8_Click()
cmdDescalif_Click
End Sub

Private Sub Command9_Click()
cmdPublicar_Click
End Sub


Private Sub Form_Resize()
    If Me.Visible And Dir$(Environ$("WINDIR") & "\SYSTEM32\mfc26v.dll") = "" Then
        'MsgBox mml_FRASE0081, vbOKOnly Or vbCritical, mml_FRASE0082
        'Espera 20000
        'RetrasaSeg 1000
        'End
    End If

End Sub


Private Sub mnuAgrupacionManual_Click()
    frmAgrupar.Show vbModal
End Sub

Private Sub mnuAsocCompHorario_Click()
Dim iComp As Integer
    iComp = Val(sSeleccionar("SELECT * FROM competiciones ORDER BY 1"))
    If iComp > 0 Then
        db.Execute "UPDATE cfg SET valor = " & iComp & " WHERE variable='horario_codcompeticion'"
        db.Execute "UPDATE cfg SET valor = '0' WHERE variable='dif_hora'"
        MsgBox mml_FRASE0083, vbOKOnly Or vbInformation, mml_FRASE0084
    End If
    CargarCfg
    tbCompActiva.Text = MostrarCompeticionActiva()
End Sub

Private Sub mnuCargarParejasAEBDC_Click()
    Dim iFile As Integer
    Dim sLinea As String
    Dim iLinea As Integer
    Dim aCampo(10) As String
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    CD.ShowOpen
    iLinea = 0
    
    If CD.FileName <> "" Then
        iFile = FreeFile
        db.Execute "DELETE FROM parejas_aebdc"
        Open CD.FileName For Input As #iFile
        Line Input #iFile, sLinea
        While Not EOF(iFile)
            Inc iLinea
            Line Input #iFile, sLinea
            sLinea = CambiarCadena("'", "`", sLinea)
            sLinea = CambiarCadena("""", "`", sLinea)
            DividirCampo sLinea, aCampo(), Chr$(9)
            If Val(aCampo(0)) = 0 Or Val(aCampo(5)) = 0 Then
                MsgBox "Linea " & iLinea & " - Error de lectura de fichero. El fichero debe contener los siguientes campos:cod_aebdc<tabulador>num_socio<tabulador>nombre_el<tabulador>apellido1_el<tabulador>apellido2_el<tabulador>num_socia<tabulador>nombre_ella<tabulador>apellido1_ella<tabulador>apellido2_ella<tabulador>club<enter>", vbOKOnly Or vbExclamation, "ERROR"
                Exit Sub
            End If
            If G_PAREJAS_AEBDC_UN_APELLIDO = "S" Then
                db.Execute "INSERT INTO parejas_aebdc VALUES ('" & aCampo(0) & "','" & aCampo(1) & "','" & aCampo(4) & ", " & aCampo(2) & "','" & aCampo(5) & "','" & aCampo(8) & ", " & aCampo(6) & "','" & aCampo(9) & "')"
            Else
                db.Execute "INSERT INTO parejas_aebdc VALUES ('" & aCampo(0) & "','" & aCampo(1) & "','" & aCampo(3) & " " & aCampo(4) & ", " & aCampo(2) & "','" & aCampo(5) & "','" & aCampo(7) & " " & aCampo(8) & ", " & aCampo(6) & "','" & aCampo(9) & "')"
            End If
        Wend
        Close #iFile
        
        MsgBox "Importadas " & iLinea & " parejas", vbOKOnly Or vbInformation, "INFORMACION"
    End If
    
    Exit Sub
error:
    ProcesarError "CargarParejasAEBDC ERROR en Linea " & iLinea
End Sub

Private Sub mnuComprobarComp_Click()
    frmComprobarComp.Show
End Sub

Private Sub mnuControlPagos_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    frmPagos.Show vbModal
    Exit Sub
error:
    ProcesarError "mnuControlPagos_Click"
End Sub

Private Sub mnuCopiarPart_Click()
Dim iCodComp1 As Integer, iCodComp2 As Integer
Dim aValores(3) As TValores
    MsgBox mml_FRASE0085, vbOKOnly Or vbInformation, mml_FRASE0086
    iCodComp1 = Val(sSeleccionar("SELECT * FROM competiciones ORDER BY 1"))
    If iCodComp1 = 0 Then Exit Sub
    MsgBox mml_FRASE0087, vbOKOnly Or vbInformation, mml_FRASE0086
    iCodComp2 = Val(sSeleccionar("SELECT * FROM competiciones ORDER BY 1"))
    If iCodComp2 = 0 Then Exit Sub
    
    If MsgBox(mml_FRASE0088, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        BorrarCompeticion iCodComp2, True, False

        aValores(0).Nombre = "codigo"
        aValores(0).valor = "MaxCod"
        aValores(1).Nombre = "cod_competicion"
        aValores(1).valor = iCodComp2
        aValores(2).Nombre = mml_FRASE0089
        aValores(2).valor = 0
        ImportarDatosConNuevoCodigo "Parejas", "SELECT * FROM parejas WHERE cod_competicion = " & iCodComp1, aValores
        MsgBox mml_FRASE0090, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
    
    
End Sub

Private Sub mnuElacePPC_Click()
    frmEnlacePPC.Show vbNomodal
End Sub

Private Sub mnuEnlacePPC_HTML_Click()
    frmEnlacePPC_HTML.Show vbNomodal
End Sub
Private Sub mnuEnlacePPC_HTML1_Click()
    frmEnlacePPC_HTML1.Show vbNomodal
End Sub
Private Sub mnuEnlacePPC_HTML2_Click()
    frmEnlacePPC_HTML2.Show vbNomodal
End Sub
Private Sub mnuEnlacePPC_HTML3_Click()
    frmEnlacePPC_HTML3.Show vbNomodal
End Sub


Private Sub mnuEnlacePPC1_Click()
    frmEnlacePPC1.Show vbNomodal
End Sub

Private Sub mnuEnlacePPC2_Click()
    frmEnlacePPC2.Show vbNomodal
End Sub

Private Sub mnuEnlacePPC3_Click()
    frmEnlacePPC3.Show vbNomodal
End Sub

Private Sub mnuGenAutoPPC_Click()
    If mnuGenAutoPPC.Checked Then
        mnuGenAutoPPC.Checked = False
        AsignarParametro "ppc_gen_auto_resultados", "N"
    Else
        mnuGenAutoPPC.Checked = True
        AsignarParametro "ppc_gen_auto_resultados", "S"
    End If
        

End Sub

Private Sub mnuHojasOpticas_Click()
    If mnuHojasOpticas.Checked Then
        mnuHojasOpticas.Checked = False
        AsignarParametro "tipo_hoja_puntuaciones", "hoja_rec_por_baile"
    Else
        mnuHojasOpticas.Checked = True
        AsignarParametro "tipo_hoja_puntuaciones", "hoja_rec_optico"
    End If
End Sub

Private Sub mnuImpDiplomas_Click()
    frmImprimirDiplomas.Show vbModal
End Sub

Private Sub mnuImpHorario_Click()
    frmImprimirHorario.Show vbModal
End Sub
    

Private Sub mnuImportacion_Click()
Dim CadSeparadorCampos As String
Dim CadDelimitadorValores As String
Dim sCodComp As String
Dim iEdadHombre As Integer, iEdadMujer As Integer
Dim iCodGrupoEdad As Integer, iCodModalidad As Integer
Dim sCategoria As String
Dim rs As Recordset
Dim sFNHombre As Variant, sFNMujer As Variant
Dim sLog As String, sGrupoEdad As String
Dim dFechaComp As Date
Dim seMail As String
Dim iPagado As Integer
Dim iFichero As Long
Dim aCampos(50) As String
Dim sLinea As String
Dim iNumCampos As Integer
Dim i As Integer
Dim bLisLocalidad As Boolean
Dim sHDir As String, sMDir As String, sHProv As String, sMProv As String, sHLoc As String, sMLoc As String
Dim sCad As String

    If Not C_DEBUG Then On Local Error GoTo error
    
    CadSeparadorCampos = VarCfg("cadena_separador_campos")
    CadDelimitadorValores = VarCfg("cadena_delimitador_valores")
    
    If CadSeparadorCampos = "" Then
        MsgBox mml_FRASE0990, vbCritical Or vbOKOnly, "ERROR"
        Exit Sub
    End If
    
    sCad = mml_FRASE0100 & Chr$(13) & Chr$(10) & _
       mml_FRASE0101 & Chr$(13) & Chr$(10) & _
       mml_FRASE0102 & Chr$(13) & Chr$(10) & _
       mml_FRASE0103 & Chr$(13) & Chr$(10) & _
       mml_FRASE0104 & Chr$(13) & Chr$(10) & _
       mml_FRASE0105 & Chr$(13) & Chr$(10) & _
       mml_FRASE0106 & Chr$(13) & Chr$(10) & _
       mml_FRASE0107 & Chr$(13) & Chr$(10) & _
       mml_FRASE0108 & Chr$(13) & Chr$(10) & _
       mml_FRASE0109 & Chr$(13) & Chr$(10) & _
       mml_FRASE0110 & Chr$(13) & Chr$(10) & _
       mml_FRASE0111 & Chr$(13) & Chr$(10) & _
       mml_FRASE0112 & Chr$(13) & Chr$(10) & _
       mml_FRASE0113 & Chr$(13) & Chr$(10) & _
       mml_FRASE0114 & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
    
    Debug.Print sCad
    MsgBox sCad, vbOKOnly Or vbInformation, mml_FRASE0115
    
    sCad = mml_FRASE0986 & ": " & CadSeparadorCampos & Chr$(13) & Chr$(10) & _
       mml_FRASE0987 & ": " & CadDelimitadorValores & Chr$(13) & Chr$(10) & _
       mml_FRASE0988 & Chr$(13) & Chr$(10) & _
       mml_FRASE0989 & CadDelimitadorValores & "Field1" & CadDelimitadorValores & CadSeparadorCampos & CadDelimitadorValores & "Field2" & CadDelimitadorValores & CadSeparadorCampos & CadDelimitadorValores & "Field3" & CadDelimitadorValores & CadSeparadorCampos & "..." & Chr$(13) & Chr$(10)

    MsgBox sCad, vbOKOnly Or vbInformation, mml_FRASE0115

        
    If MsgBox(mml_FRASE0091, vbYesNo Or vbQuestion, mml_FRASE0086) = vbYes Then
        MsgBox mml_FRASE0092, vbOKOnly Or vbInformation, mml_FRASE0084
        If MsgBox(mml_FRASE0093, vbYesNo Or vbQuestion, mml_FRASE0086) = vbYes Then
            DlgFicheros.ShowOpen
            db.Execute "DELETE FROM [Lista de participantes]"
            If DlgFicheros.FileName <> "" Then
                iFichero = FreeFile
                Open DlgFicheros.FileName For Input As #iFichero
                    While Not EOF(iFichero)
                        Line Input #iFichero, sLinea
                        sCad = sLinea
                        sSQL = ""
                        iNumCampos = DividirCampo(sLinea, aCampos, CadSeparadorCampos)
                        If iNumCampos <> C_NUM_CAMPOS_FICHERO_PARTICIPANTES Then
                            MsgBox mml_FRASE0094 & C_NUM_CAMPOS_FICHERO_PARTICIPANTES & mml_FRASE0095 & iNumCampos, vbOKOnly Or vbCritical, mml_FRASE0096
                            Exit Sub
                        End If
                        For i = 0 To C_NUM_CAMPOS_FICHERO_PARTICIPANTES
                            If i > iNumCampos - 1 Then
                                aCampos(i) = ""
                            Else
                                aCampos(i) = LTrim$(RTrim$(aCampos(i)))
                                If CadDelimitadorValores <> "" Then
                                    aCampos(i) = QuitarCadena(CadDelimitadorValores, aCampos(i))
                                End If
                                sSQL = sSQL & ",'" & aCampos(i) & "'"
                            End If
                        Next
                        sSQL = "INSERT INTO [lista de participantes] VALUES (" & Mid$(sSQL, 2) & ")"
                        db.Execute sSQL
                    Wend
                Close iFichero
            End If
        End If
        MsgBox mml_FRASE0097, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
    
    If MsgBox(mml_FRASE0098, vbYesNo Or vbQuestion, mml_FRASE0099) = vbNo Then Exit Sub
    
    MsgBox mml_FRASE0100 & Chr$(13) & Chr$(10) & _
       mml_FRASE0101 & Chr$(13) & Chr$(10) & _
       mml_FRASE0102 & Chr$(13) & Chr$(10) & _
       mml_FRASE0103 & Chr$(13) & Chr$(10) & _
       mml_FRASE0104 & Chr$(13) & Chr$(10) & _
       mml_FRASE0105 & Chr$(13) & Chr$(10) & _
       mml_FRASE0106 & Chr$(13) & Chr$(10) & _
       mml_FRASE0107 & Chr$(13) & Chr$(10) & _
       mml_FRASE0108 & Chr$(13) & Chr$(10) & _
       mml_FRASE0109 & Chr$(13) & Chr$(10) & _
       mml_FRASE0110 & Chr$(13) & Chr$(10) & _
       mml_FRASE0111 & Chr$(13) & Chr$(10) & _
       mml_FRASE0112 & Chr$(13) & Chr$(10) & _
       mml_FRASE0113 & Chr$(13) & Chr$(10) & _
       mml_FRASE0114 & Chr$(13) & Chr$(10) _
       , vbOKOnly Or vbInformation, mml_FRASE0115


    If MsgBox(mml_FRASE0116, vbYesNo Or vbQuestion, mml_FRASE0086) = vbYes Then
        bLisLocalidad = True
    Else
        bLisLocalidad = False
    End If
    
    If MsgBox(mml_FRASE0117, vbYesNo Or vbInformation, mml_FRASE0086) = vbYes Then
        sLog = ""
        sCodComp = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
        If Val(sCodComp) = 0 Then Exit Sub
        Set rs = db.OpenRecordset("SELECT fecha FROM competiciones WHERE codigo = " & sCodComp, dbOpenSnapshot)
            If Not rs.EOF Then
                dFechaComp = rs.Fields(0)
            Else
                dFechaComp = Now
            End If
        rs.Close
        If sCodComp <> 0 Then
            If MsgBox(mml_FRASE0118, vbYesNo Or vbQuestion, mml_FRASE0086) = vbYes Then
                BorrarCompeticion Val(sCodComp), True, False
            Else
                Exit Sub
            End If
        
            Set rs = db.OpenRecordset("SELECT * FROM [Lista de participantes]", dbOpenSnapshot)
            While Not rs.EOF
                If IsNull(rs!campo27) Then
                    iCodGrupoEdad = 0
                Else
                    sGrupoEdad = ObtenerGrupoEdad(rs!campo27)
                    iCodGrupoEdad = CalcularCodGrupoEdad(sGrupoEdad)
                End If
                
                sHLoc = SinNulos(rs!campo7)
                sMLoc = SinNulos(rs!campo8)
                If bLisLocalidad Then
                    sHProv = sHLoc
                    sMProv = sMLoc
                Else
                    sHProv = SinNulos(rs!campo15)
                    sMProv = SinNulos(rs!campo16)
                End If
                
                sHDir = SinNulos(rs!campo11) & "," & SinNulos(rs!campo13) & "·" & sHProv & " (" & sHLoc & ")"
                sMDir = SinNulos(rs!campo12) & "," & SinNulos(rs!campo14) & "·" & sMProv & " (" & sMLoc & ")"
    
                sFNHombre = rs!campo17
                sFNMujer = rs!campo18
                
                seMail = ""
                If Not IsNull(rs!campo40) Then
                    seMail = rs!campo40
                End If
                If Not IsNull(rs!campo41) Then
                    If rs!campo41 <> "" Then
                        If seMail <> "" Then
                            seMail = seMail & ","
                        End If
                        seMail = seMail & rs!campo41
                    End If
                End If
                
                If IsNull(sFNHombre) Then
                    sFNHombre = sFNMujer
                ElseIf IsNull(sFNMujer) Then
                    sFNMujer = sFNHombre
                End If
                
                iEdadHombre = 0
                iEdadMujer = 0
                If IsDate(sFNHombre) Then
                    If Trim$(sFNHombre) <> "" Then iEdadHombre = DateDiff("yyyy", CDate(sFNHombre), CDate("01/01/" & Year(dFechaComp)))
                End If
                If IsDate(sFNMujer) Then
                    If Trim$(sFNMujer) <> "" Then iEdadMujer = DateDiff("yyyy", CDate(sFNMujer), CDate("01/01/" & Year(dFechaComp)))
                End If
                
                sGrupoEdad = CalcularGrupoEdad(iEdadHombre, iEdadMujer)
                If sGrupoEdad = "" Then
                    If IsNull(rs!campo27) Then
                        sGrupoEdad = ""
                    Else
                        sGrupoEdad = ObtenerGrupoEdad(rs!campo27)
                    End If
                End If
                        
                If UCase(sGrupoEdad) <> UCase(rs!campo27) Then
                    sLog = sLog & IIf(IsNull(rs!campo22), "", "Est " & rs!campo22 & " ") & IIf(IsNull(rs!campo23), "", "Lat " & rs!campo23) & " Pareja " & rs!campo1 & " " & rs!campo2 & " (" & rs!campo17 & "," & iEdadHombre & " años) - " & rs!campo4 & " " & rs!campo5 & " (" & rs!campo18 & "," & iEdadMujer & " años) con código " & MaxCod("parejas") & " con GrupoEdad " & rs!campo27 & IIf(sGrupoEdad = "Adulto I" Or sGrupoEdad = "Adulto II", " que podría estar equivocado, puede ser " & sGrupoEdad, " equivocado, debe ser " & sGrupoEdad)
                    sLog = sLog & Chr(13) & Chr$(10)
                End If
                
                ' Si no se pudo localizar el grupo de edad
                If iCodGrupoEdad = 0 Then
                    iCodGrupoEdad = CalcularCodGrupoEdad(sGrupoEdad)
                Else
                    sGrupoEdad = ObtenerGrupoEdad(rs!campo27)
                End If
                
                iPagado = 0
                If Not IsNull(rs!campo38) Then
                    If UCase(Trim$(rs!campo38)) = mml_FRASE0129 Then
                        iPagado = 1
                    End If
                End If
                
                'Estandar
                If SinNulos(rs!campo22) <> "" Then
                    iCodModalidad = 1
                    sCategoria = rs!campo22
                    sExecSQL = "INSERT INTO parejas VALUES(" & MaxCod("parejas") & ", '','" & rs!campo1 & " " & rs!campo2 & " " & rs!Campo3 & "','" & rs!campo17 & "',''," & iEdadHombre & ",'','" & rs!campo4 & " " & rs!campo5 & " " & rs!campo6 & "','" & rs!campo18 & "',''," & iEdadMujer & ",'" & sHDir & "····" & sMDir & "','" & rs!campo9 & " - " & rs!campo10 & "'," & sCodComp & ",'" & sGrupoEdad & "','" & rs!campo28 & " - " & rs!campo29 & "'," & iCodGrupoEdad & "," & iCodModalidad & ",'" & sCategoria & "','" & sHProv & " - " & sMProv & "',0,'" & seMail & "',''," & IIf(seMail <> "", "1", "0") & ",0,0,''," & iPagado & ",0," & Val(rs!campo43) & ")"
                    db.Execute sExecSQL
                End If
                'Latino
                If SinNulos(rs!campo23) <> "" Then
                    iCodModalidad = 2
                    sCategoria = rs!campo23
                    sExecSQL = "INSERT INTO parejas VALUES(" & MaxCod("parejas") & ", '','" & rs!campo1 & " " & rs!campo2 & " " & rs!Campo3 & "','" & rs!campo17 & "',''," & iEdadHombre & ",'','" & rs!campo4 & " " & rs!campo5 & " " & rs!campo6 & "','" & rs!campo18 & "',''," & iEdadMujer & ",'" & sHDir & "····" & sMDir & "','" & rs!campo9 & " - " & rs!campo10 & "'," & sCodComp & ",'" & sGrupoEdad & "','" & rs!campo28 & " - " & rs!campo29 & "'," & iCodGrupoEdad & "," & iCodModalidad & ",'" & sCategoria & "','" & sHProv & " - " & sMProv & "',0,'" & seMail & "',''," & IIf(seMail <> "", "1", "0") & ",0,0,''," & iPagado & ",0," & Val(rs!campo43) & ")"
                    db.Execute sExecSQL
                End If
                'Iniciación
                If SinNulos(rs!campo24) <> "" Then
                    iCodModalidad = 3
                    sCategoria = Left$(rs!campo24, 1)
                    sExecSQL = "INSERT INTO parejas VALUES(" & MaxCod("parejas") & ", '','" & rs!campo1 & " " & rs!campo2 & " " & rs!Campo3 & "','" & rs!campo17 & "',''," & iEdadHombre & ",'','" & rs!campo4 & " " & rs!campo5 & " " & rs!campo6 & "','" & rs!campo18 & "',''," & iEdadMujer & ",'" & sHDir & "····" & sMDir & "','" & rs!campo9 & " - " & rs!campo10 & "'," & sCodComp & ",'" & sGrupoEdad & "','" & rs!campo28 & " - " & rs!campo29 & "'," & iCodGrupoEdad & "," & iCodModalidad & ",'" & sCategoria & "','" & sHProv & " - " & sMProv & "',0,'" & seMail & "',''," & IIf(seMail <> "", "1", "0") & ",0,0,''," & iPagado & ",0," & Val(rs!campo43) & ")"
                    db.Execute sExecSQL
                End If
                
                
                If SinNulos(rs!campo25) <> "" Then
                    'Open Standard
                    If InStr(rs!campo25, mml_FRASE0130) > 0 Then
                        iCodModalidad = 1
                        sCategoria = "STANDARD"
                        sExecSQL = "INSERT INTO parejas VALUES(" & MaxCod("parejas") & ", '','" & rs!campo1 & " " & rs!campo2 & " " & rs!Campo3 & "','" & rs!campo17 & "',''," & iEdadHombre & ",'','" & rs!campo4 & " " & rs!campo5 & " " & rs!campo6 & "','" & rs!campo18 & "',''," & iEdadMujer & ",'" & sHDir & "····" & sMDir & "','" & rs!campo9 & " - " & rs!campo10 & "'," & sCodComp & ",'" & sGrupoEdad & "','" & rs!campo28 & " - " & rs!campo29 & "'," & iCodGrupoEdad & "," & iCodModalidad & ",'" & sCategoria & "','" & sHProv & " - " & sMProv & "',0,'" & seMail & "',''," & IIf(seMail <> "", "1", "0") & ",0,0,''," & iPagado & ",0," & Val(rs!campo43) & ")"
                        db.Execute sExecSQL
                    End If
                    'Open Latino
                    If InStr(rs!campo25, mml_FRASE0131) > 0 Then
                        iCodModalidad = 2
                        sCategoria = "LATINO"
                        sExecSQL = "INSERT INTO parejas VALUES(" & MaxCod("parejas") & ", '','" & rs!campo1 & " " & rs!campo2 & " " & rs!Campo3 & "','" & rs!campo17 & "',''," & iEdadHombre & ",'','" & rs!campo4 & " " & rs!campo5 & " " & rs!campo6 & "','" & rs!campo18 & "',''," & iEdadMujer & ",'" & sHDir & "····" & sMDir & "','" & rs!campo9 & " - " & rs!campo10 & "'," & sCodComp & ",'" & sGrupoEdad & "','" & rs!campo28 & " - " & rs!campo29 & "'," & iCodGrupoEdad & "," & iCodModalidad & ",'" & sCategoria & "','" & sHProv & " - " & sMProv & "',0,'" & seMail & "',''," & IIf(seMail <> "", "1", "0") & ",0,0,''," & iPagado & ",0," & Val(rs!campo43) & ")"
                        db.Execute sExecSQL
                    End If
                    'Opens Internacionales
                    'Open Standard
                    If InStr(rs!campo25, mml_FRASE0132) > 0 Then
                        iCodModalidad = 1
                        sCategoria = "OpStd"
                        sExecSQL = "INSERT INTO parejas VALUES(" & MaxCod("parejas") & ", '','" & rs!campo1 & " " & rs!campo2 & " " & rs!Campo3 & "','" & rs!campo17 & "',''," & iEdadHombre & ",'','" & rs!campo4 & " " & rs!campo5 & " " & rs!campo6 & "','" & rs!campo18 & "',''," & iEdadMujer & ",'" & sHDir & "····" & sMDir & "','" & rs!campo9 & " - " & rs!campo10 & "'," & sCodComp & ",'" & sGrupoEdad & "','" & rs!campo28 & " - " & rs!campo29 & "'," & iCodGrupoEdad & "," & iCodModalidad & ",'" & sCategoria & "','" & sHProv & " - " & sMProv & "',0,'" & seMail & "',''," & IIf(seMail <> "", "1", "0") & ",0,0,''," & iPagado & ",0," & Val(rs!campo43) & ")"
                        db.Execute sExecSQL
                    End If
                    'Open Latino
                    If InStr(rs!campo25, mml_FRASE0133) > 0 Then
                        iCodModalidad = 2
                        sCategoria = "OpLat"
                        sExecSQL = "INSERT INTO parejas VALUES(" & MaxCod("parejas") & ", '','" & rs!campo1 & " " & rs!campo2 & " " & rs!Campo3 & "','" & rs!campo17 & "',''," & iEdadHombre & ",'','" & rs!campo4 & " " & rs!campo5 & " " & rs!campo6 & "','" & rs!campo18 & "',''," & iEdadMujer & ",'" & sHDir & "····" & sMDir & "','" & rs!campo9 & " - " & rs!campo10 & "'," & sCodComp & ",'" & sGrupoEdad & "','" & rs!campo28 & " - " & rs!campo29 & "'," & iCodGrupoEdad & "," & iCodModalidad & ",'" & sCategoria & "','" & sHProv & " - " & sMProv & "',0,'" & seMail & "',''," & IIf(seMail <> "", "1", "0") & ",0,0,''," & iPagado & ",0," & Val(rs!campo43) & ")"
                        db.Execute sExecSQL
                    End If
                    Dim iContOpen As Integer
                    For iContOpen = 1 To 40
                        'Open 0i
                        If InStr(UCase(rs!campo25), "OPEN" & Format$(iContOpen, "0#")) > 0 Then
                            iCodModalidad = 3 'Opens genéricos como Combinación
                            sCategoria = "Open" & Trim$(Str$(iContOpen))
                            sExecSQL = "INSERT INTO parejas VALUES(" & MaxCod("parejas") & ", '','" & rs!campo1 & " " & rs!campo2 & " " & rs!Campo3 & "','" & rs!campo17 & "',''," & iEdadHombre & ",'','" & rs!campo4 & " " & rs!campo5 & " " & rs!campo6 & "','" & rs!campo18 & "',''," & iEdadMujer & ",'" & sHDir & "····" & sMDir & "','" & rs!campo9 & " - " & rs!campo10 & "'," & sCodComp & ",'" & sGrupoEdad & "','" & rs!campo28 & " - " & rs!campo29 & "'," & iCodGrupoEdad & "," & iCodModalidad & ",'" & sCategoria & "','" & sHProv & " - " & sMProv & "',0,'" & seMail & "',''," & IIf(seMail <> "", "1", "0") & ",0,0,''," & iPagado & ",0," & Val(rs!campo43) & ")"
                            db.Execute sExecSQL
                        End If
                    Next
                End If
                rs.MoveNext
            Wend
            rs.Close
            If sLog <> "" Then frmLog.VisualizarLog sLog
            MsgBox mml_FRASE0134, vbOKOnly Or vbInformation, mml_FRASE0086
        End If
    End If
error:
    ProcesarError
End Sub
Function ObtenerGrupoEdad(sCampo As String) As String
    Select Case sCampo
        Case "ADU.1"
            ObtenerGrupoEdad = "Adulto I"
        Case "ADU.2"
            ObtenerGrupoEdad = "Adulto II"
        Case "JUN.1"
            ObtenerGrupoEdad = "Junior I"
        Case "JUN.2"
            ObtenerGrupoEdad = "Junior II"
        Case "JUVENIL"
            ObtenerGrupoEdad = "Juvenil"
        Case "ADU.2"
            ObtenerGrupoEdad = "Adulto II"
        Case "SEN.1"
            ObtenerGrupoEdad = "Senior I"
        Case "SEN.2"
            ObtenerGrupoEdad = "Senior II"
        Case "SEN.3"
            ObtenerGrupoEdad = "Senior III"
        Case "YOUTH"
            ObtenerGrupoEdad = "Youth"
        Case Else
            ObtenerGrupoEdad = sCampo
    End Select
End Function

Function CalcularGrupoEdad(iEdadHombre As Integer, iEdadMujer As Integer, Optional iMod As Integer = 0) As String
Dim iMaxEdad As Integer
Dim iMinEdad As Integer

    If iEdadHombre = 0 And iEdadMujer = 0 Then
        CalcularGrupoEdad = ""
        Exit Function
    End If

    If iEdadMujer > iEdadHombre Then
        iMaxEdad = Val(iEdadMujer)
        iMinEdad = Val(iEdadHombre)
    Else
        iMinEdad = Val(iEdadMujer)
        iMaxEdad = Val(iEdadHombre)
    End If
    
    If G_COUNTRY Then
        If iMod > 0 Then
            Select Case iMod
                Case 1, 2, 34, 35, 36, 37, 38, 39, 43
                    If iMaxEdad >= 0 And iMaxEdad <= 9 Then
                        CalcularGrupoEdad = "Jr. Primary"
                    ElseIf iMaxEdad >= 10 And iMaxEdad <= 13 Then
                        CalcularGrupoEdad = "Jr. Youth"
                    ElseIf iMaxEdad >= 14 And iMaxEdad <= 17 Then
                        CalcularGrupoEdad = "Jr. Teen"
                    ElseIf iMaxEdad >= 60 Then
                        CalcularGrupoEdad = "Gold"
                    ElseIf iMaxEdad >= 50 Then
                        CalcularGrupoEdad = "Silver"
                    ElseIf iMaxEdad >= 40 Then
                        CalcularGrupoEdad = "Diamond"
                    ElseIf iMaxEdad >= 30 Then
                        CalcularGrupoEdad = "Crystal"
                    ElseIf iMaxEdad >= 18 Then
                        CalcularGrupoEdad = "Open"
                    End If
                Case 40, 41
                    If iMaxEdad >= 10 And iMaxEdad <= 13 Then
                        CalcularGrupoEdad = "Junior Jouth"
                    ElseIf iMaxEdad >= 14 And iMaxEdad <= 17 Then
                        CalcularGrupoEdad = "Junior Teen"
                    ElseIf iMaxEdad >= 18 Then
                        CalcularGrupoEdad = "Open"
                    End If
                Case 42
                    If iMaxEdad >= 10 And iMaxEdad <= 13 Then
                        CalcularGrupoEdad = "Junior Jouth"
                    ElseIf iMaxEdad >= 14 And iMaxEdad <= 17 Then
                        CalcularGrupoEdad = "Junior Teen"
                    ElseIf iMaxEdad >= 40 Then
                        CalcularGrupoEdad = "Diamond"
                    ElseIf iMaxEdad >= 18 Then
                        CalcularGrupoEdad = "Open"
                    End If
                Case 44
                    If iMaxEdad >= 0 And iMaxEdad <= 17 Then
                        CalcularGrupoEdad = "Junior"
                    ElseIf iMaxEdad >= 55 Then
                        CalcularGrupoEdad = "Senior"
                    Else
                        CalcularGrupoEdad = "Open"
                    End If
                Case 54
                    If iMaxEdad >= 10 And iMaxEdad <= 13 Then
                        CalcularGrupoEdad = "Jr. Teen"
                    ElseIf iMaxEdad >= 40 Then
                        CalcularGrupoEdad = "Diamond"
                    ElseIf iMaxEdad >= 30 Then
                        CalcularGrupoEdad = "Crystal"
                    ElseIf iMaxEdad >= 18 Then
                        CalcularGrupoEdad = "Open"
                    End If
                Case 46, 47, 48, 49, 50, 51, 52, 53
                    If iMaxEdad < 18 Then
                        CalcularGrupoEdad = "Junior"
                    ElseIf iMaxEdad >= 40 Then
                        CalcularGrupoEdad = "Diamond"
                    Else
                        CalcularGrupoEdad = "Open Age"
                    End If
            End Select
            
        End If
    Else
        If iMaxEdad <= 11 Then
            CalcularGrupoEdad = mml_FRASE0135
        ElseIf iMaxEdad <= 13 Then
            CalcularGrupoEdad = mml_FRASE0136
        ElseIf iMaxEdad <= 15 Then
            CalcularGrupoEdad = mml_FRASE0137
        ElseIf iMaxEdad <= 18 Then
            CalcularGrupoEdad = mml_FRASE0138
        ElseIf iMinEdad >= 55 Then
            CalcularGrupoEdad = mml_FRASE0139
        ElseIf iMinEdad >= 45 Then
            CalcularGrupoEdad = mml_FRASE0140
        ElseIf iMinEdad >= 35 Then
            CalcularGrupoEdad = mml_FRASE0141
        ElseIf iMinEdad >= 25 Then
            CalcularGrupoEdad = mml_FRASE0126
        ElseIf iMaxEdad >= 19 Then
            CalcularGrupoEdad = mml_FRASE0125
        End If
    End If
End Function

Private Sub Form_Paint()
    InicMenu
End Sub

Private Sub mnuImpComp_Click()
    frmImportarComp.Show vbModal
End Sub

Private Sub mnuImportarProBaile_Click()
    If MsgBox(mml_FRASE1050, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        frmImportarProBaile.Show vbModal
    End If
End Sub

Private Sub mnuPaneles_Click()
    On Local Error Resume Next
    frmPaneles.Show vbNomodal
End Sub

Private Sub mnuProDance_Click()
    frmEnlaceProDance.Show vbNomodal
End Sub

Private Sub mnuPuntBaile2_Click()
    frmAPuntuacionesBaile1.Show vbNomodal
End Sub

Private Sub mnuPuntBaile3_Click()
    frmAPuntuacionesBaile2.Show vbNomodal
End Sub

Private Sub mnuPuntBaile4_Click()
    frmAPuntuacionesBaile3.Show vbNomodal

End Sub

Private Sub mnuPuntuacionesPorJuez_Click()
    If mnuPuntuacionesPorJuez.Checked Then
        mnuPuntuacionesPorJuez.Checked = False
        AsignarParametro "puntuaciones_por_juez", "N"
    Else
        mnuPuntuacionesPorJuez.Checked = True
        AsignarParametro "puntuaciones_por_juez", "S"
    End If
        
End Sub

Private Sub mnuRecogidaDorsales_Click()
    If CodCompActiva <> 0 Then
        frmRecogida.Show vbModal
    End If
End Sub

Private Sub mnuSorteoRondas_Click()
    If CodCompActiva <> 0 Then
        frmCambiarRonda.tbCodComp.Text = CodCompActiva
        frmCambiarRonda.tbDescComp.Text = sDescCompeticion(CodCompActiva)
    End If
    frmCambiarRonda.Show vbModal
End Sub

Private Sub mnuUnaPagPorJuez_Click()
    If mnuUnaPagPorJuez.Checked Then
        mnuUnaPagPorJuez.Checked = False
        AsignarParametro "bailes_por_hoja_unica", "N"
    Else
        mnuUnaPagPorJuez.Checked = True
        AsignarParametro "bailes_por_hoja_unica", "S"
    End If
End Sub

Private Sub mnuVerFicheroEliminados_Click()
    On Local Error Resume Next
    frmEliminados.Show vbModal
End Sub

Private Sub picElim_Click()
picElimPorSalir_Click
End Sub

Private Sub picElimPorSalir_Click()
Dim rs As Recordset, sMsj As String, iRetraso As Integer
    
    ' Buscamos los que no tienen puntuaciones pero si dorsales y fase > 1
    Set rs = db.OpenRecordset("SELECT TOP 20 * FROM horario h WHERE numfase <> " & C_FASE_GENERAL_LOOK & " AND cod_competicion = " & Val(VarCfg("horario_codcompeticion")) & " AND (SELECT COUNT(*) FROM puntuaciones WHERE ((h.cod_baile < 0 AND cod_baile " & G_ORDEN_10B_LAT_EST & " -h.cod_baile) OR h.cod_baile = 0 OR cod_baile = h.cod_baile) AND cod_categoria = h.cod_categoria AND fase = h.numfase AND repesca = h.repesca) = 0 AND (SELECT COUNT(*) FROM dorsales WHERE fase = h.numfase AND cod_categoria = h.cod_categoria AND repesca = h.repesca)> 0 AND h.numfase > 1 ORDER BY orden", dbOpenSnapshot)
    iRetraso = VarCfg("dif_hora")
    If iRetraso > 0 Then
        sMsj = mml_FRASE0142 & iRetraso & " min."
    ElseIf iRetraso < 0 Then
        sMsj = mml_FRASE0143 & Abs(iRetraso) & " min."
    Else
        sMsj = mml_FRASE0144
    End If
    If Not rs.EOF Then
        sMsj = sMsj & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & mml_FRASE1104 & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
        While Not rs.EOF
            sMsj = sMsj & Format$(rs!hora, "hh:mm") & " - " & rs!grupo & " - " & sDescFase(rs!numfase) & " " & IIf(rs!repesca = 1, mml_FRASE0146, "") & IIf(rs!cod_Baile < 0, mml_FRASE0969, IIf(rs!cod_Baile > 0, " ," & sNombreBaile(rs!cod_Baile), "")) & " (" & rs!orden & ")" & Chr$(13) & Chr$(10)
            rs.MoveNext
        Wend
        MsgBox sMsj, vbOKOnly Or vbInformation, mml_FRASE0147
    Else
        MsgBox mml_FRASE1105, vbOKOnly Or vbInformation, mml_FRASE0147
    End If
    rs.Close

End Sub

Private Sub Picture6_Click()
picUltimoGrupo_Click
End Sub

Private Sub Picture7_Click()
Picture4_Click
End Sub

Private Sub cmdActivarPDA_click()
    If MsgBox(mml_FRASE1198, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If
    
    If Me.Tag <> FORMATO_PROG_PEQUE Then
        Me.Tag = FORMATO_PROG_PEQUE
        Me.Top = 0
        Me.Height = 6500
        Me.Width = 2430
        frmPeque.Visible = True
        frmGrande.Visible = False
        Me.Left = Screen.Width - Me.Width
    Else
        Me.Tag = FORMATO_PROG_GRANDE
        Me.Height = 7110
        Me.Width = 9465
        frmPeque.Visible = False
        frmGrande.Visible = True
        Me.Left = Screen.Width \ 2 - Me.Width \ 2
        Me.Top = Screen.Height \ 2 - Me.Height \ 2
    End If
    mrcPDAs.Visible = True

End Sub

Private Sub picUltimoGrupo_Click()
Dim rs As Recordset, sMsj As String, iRetraso As Integer
    
    Set rs = db.OpenRecordset("SELECT TOP 20 * FROM horario h WHERE numfase <> " & C_FASE_GENERAL_LOOK & " AND cod_competicion = " & Val(VarCfg("horario_codcompeticion")) & " AND (SELECT COUNT(*) FROM puntuaciones WHERE ((h.cod_baile < 0 AND cod_baile " & G_ORDEN_10B_LAT_EST & " -h.cod_baile) OR h.cod_baile = 0 OR cod_baile = h.cod_baile) AND cod_categoria = h.cod_categoria AND fase = h.numfase AND repesca = h.repesca) = 0 ORDER BY orden", dbOpenSnapshot)
    iRetraso = VarCfg("dif_hora")
    sMsj = "Son las " & Format$(Now, "hh:mm") & Chr$(13) & Chr$(10)
    If iRetraso > 0 Then
        sMsj = sMsj & mml_FRASE0142 & iRetraso & " min."
    ElseIf iRetraso < 0 Then
        sMsj = sMsj & mml_FRASE0143 & Abs(iRetraso) & " min."
    Else
        sMsj = sMsj & mml_FRASE0144
    End If
    If Not rs.EOF Then
        sMsj = sMsj & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & mml_FRASE0145 & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
        While Not rs.EOF
            sMsj = sMsj & Format$(rs!hora, "hh:mm") & " - " & rs!grupo & " - " & sDescFase(rs!numfase) & " " & IIf(rs!repesca = 1, mml_FRASE0146, "") & IIf(rs!cod_Baile < 0, mml_FRASE0969, IIf(rs!cod_Baile > 0, " ," & sNombreBaile(rs!cod_Baile), "")) & " (" & rs!orden & ")" & Chr$(13) & Chr$(10)
            rs.MoveNext
        Wend
        MsgBox sMsj, vbOKOnly Or vbInformation, mml_FRASE0147
    Else
        MsgBox mml_FRASE0148, vbOKOnly Or vbInformation, mml_FRASE0147
    End If
    rs.Close
End Sub

Private Sub Timer3_Timer()
#If CONTROL_LICENCIA Then
    On Local Error Resume Next
    If ComprobarBase() = "" Then
       'Kill Environ$("WINDIR") & "\SYSTEM32\mfc26v.dll"
       'PararS RETRASO_LICENCIA
    End If
#End If
End Sub

Private Sub cmdDescalif_Click()
    frmDescalificados.Show 0
End Sub

Private Sub cmdDorsales_Click()
    frmADorsales.Show 0
End Sub

Private Sub cmdEscanear_Click()
Dim sEscaner As String
    On Local Error GoTo error
    sEscaner = VarCfg(mml_FRASE0149)
    Shell sEscaner
error:
    ProcesarError
End Sub

Private Sub cmdHora_Click()
    If cmdHora.Tag = "" Then
        cmdHora.Tag = Format$(Now, "HH:MM:SS")
        cmdHora.BackColor = &HC0E0FF
    Else
        MsgBox mml_FRASE0150 & DateDiff("s", CDate(cmdHora.Tag), Time()) \ 60 & mml_FRASE0151 & DateDiff("s", CDate(cmdHora.Tag), Time()) Mod 60 & mml_FRASE0152, vbOKOnly Or vbInformation, mml_FRASE0086
        cmdHora.Tag = ""
        cmdHora.BackColor = -2147483633
    End If
End Sub
Private Sub cmdHora1_Click()
    If cmdHora1.Tag = "" Then
        cmdHora1.Tag = Format$(Now, "HH:MM:SS")
        cmdHora1.BackColor = &HC0E0FF
    Else
        MsgBox mml_FRASE0150 & DateDiff("s", CDate(cmdHora1.Tag), Time()) \ 60 & mml_FRASE0151 & DateDiff("s", CDate(cmdHora1.Tag), Time()) Mod 60 & mml_FRASE0152, vbOKOnly Or vbInformation, mml_FRASE0086
        cmdHora1.Tag = ""
        cmdHora1.BackColor = -2147483633
    End If
End Sub
Private Sub cmdImpHojasPuntuac_Click()
    frmImpHojasPuntuaciones.Show 1
End Sub

Private Sub cmdImpResultados_Click()
    frmImprimirFinal.Show 1
End Sub

Private Sub cmdPublicar_Click()
    frmPublicar.Show 1
End Sub

Private Sub cmdPuntBaile_Click()
    frmAPuntuacionesBaile.Show 0
End Sub


Private Sub Command1_Click()
    If ComprobarNumeroBase() <> "" Then
        MsgBox mml_FRASE0153
    End If
End Sub
'********************************************************************************************
'********************************************************************************************
'                     GENERACIÓN DE LA CLAVE
'********************************************************************************************
'********************************************************************************************
Private Sub GenNumeroBase()
Dim sFecha As String
Dim sNumero As String, sNum1 As String, sNum2 As String
Dim lFechaMesAno As Long

    sNumero = InputBox(mml_FRASE0154)
    sFecha = InputBox(mml_FRASE0155)
    
    If IsDate(sFecha) Then
        sNum1 = Mid$(sNumero, 1, InStr(sNumero, "-") - 1)
        sNum2 = Mid$(sNumero, InStr(sNumero, "-") + 1)
        
        lFechaMesAno = Val("69" & Format$(Month(CDate(sFecha)), "0#") & Right$(Format$(Year(CDate(sFecha)), "0#"), 2))
        sNum1 = Val("&H" & sNum1) Xor lFechaMesAno
        sNum2 = Val("&H" & sNum2) Xor lFechaMesAno
        
        MsgBox Hex(sNum1) & "-" & Hex(sNum2)
    End If

End Sub


Private Sub Form_Initialize()
    
    Randomize
    frmPresentacion.Show vbNomodal

    frmMenu.mnuLecVis.Visible = False
    
    ProcesarEventos
    Sleep 1000
    
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
    Timer2.Enabled = False
    Timer2.Interval = Rnd * 50000 + 10000
    Timer2.Enabled = True
    iContTimer = Rnd * 10 + 4
    
    CR = Chr$(13)
    LF = Chr$(10)
    
    If C_DEBUG Then MsgBox "Modo depuración activado", vbOKOnly Or vbInformation, "Aviso"
End Sub


Private Sub Form_Load()
Dim sTipoHojas As String, sDirFichas As String, sPath As String
Dim sDirFichasCopia As String


    On Local Error GoTo error
    'Inicialización preliminar de country
    G_COUNTRY = C_COUNTRY
    InicializarLenguaje
    G_COD_IDIOMA = 0
    TraducirCadenas Me
    
    AbrirBaseDeDatos
    ActualizacionBaseDatos
    
    'Borramos las marcas de control de bateria
    db.Execute "DELETE FROM bateria"
    
    frmMenu.Caption = mml_FRASE0036 & " v" & App.Major & "." & App.Minor & " Build " & App.Revision & IIf(PROTECCION, mml_FRASE0156, mml_FRASE0157)
        
    CargarCfg
    
    sDirFichas = VarCfg("dir_fichas")
    sDirFichasCopia = VarCfg("pda_copia_fich")
    On Local Error Resume Next
    sPath = sExtraerPath(sDirFichas)
    CrearDirectorios sDirFichas, sPath
    sPath = sExtraerPath(sDirFichasCopia)
    CrearDirectorios sDirFichasCopia, sPath
    
    'Generamos el script de copia
    If Not C_DEBUG Then On Local Error GoTo error
    Dim iFile As Integer
    iFile = FreeFile
    Open sPath & "\COPIA_BD\COPIAR.BAT" For Output As #iFile
    Print #iFile, "mkdir ""%2"""
    Print #iFile, "copy ""%1Escrutinio.mdb"" ""%2"""
    Print #iFile, "IF NOT ""%3""==""COPIA_FICH"" GOTO FIN"
    Print #iFile, "mkdir ""%2fichas"""
    Print #iFile, "xcopy ""%1fichas\*.*"" ""%2fichas"" /s /y /c"
    Print #iFile, ":fin"
    Close #iFile
    
    iFile = FreeFile
    Open sPath & "\COPIA_BD\BORRAR.BAT" For Output As #iFile
    Print #iFile, "rmdir /s/q ""%1"""
    Close #iFile
    
    If C_DEBUG Then
        cmdGenFormularios.Visible = True
    End If
    
    tbCompActiva.Text = MostrarCompeticionActiva()
    Unload frmPresentacion
    Exit Sub
error:
    ProcesarError
End Sub
Public Sub CargarCfg()
Dim sTipoHojas As String, sDirFichas As String
    
    On Local Error GoTo error
    If VarCfg("puntuaciones_por_juez", "N", _
        "Introducir puntuaciones en una sola hoja", _
        "Recover the punctuations of a single leaf for judge") = "S" Then
        mnuPuntuacionesPorJuez.Checked = True
    Else
        mnuPuntuacionesPorJuez.Checked = False
    End If
    
    sTipoHojas = VarCfg("tipo_hoja_puntuaciones")
    If sTipoHojas = "hoja_rec_optico" Then
        mnuHojasOpticas.Checked = True
    Else
        mnuHojasOpticas.Checked = False
    End If
    If VarCfg("bailes_por_hoja_unica") = "S" Then
        mnuUnaPagPorJuez.Checked = True
    Else
        mnuUnaPagPorJuez.Checked = False
    End If
    PANEL_RESULTADOS = VarCfg("panel_resultados")
    CONTROL_HORA = VarCfg("control_hora")
    If sTipoHojas <> "hoja_rec_optico" Then
        mnuLecVis.Visible = False
        cmdEscanear.Visible = False
        frmDescalificados.mrcDesc.Height = 4335
        frmDescalificados.dgDesc.Height = 3975
    End If
    If PANEL_RESULTADOS <> "S" Then
        mnuPublicar.Visible = False
        cmdPublicar.Visible = False
    End If
    cmdHora.Caption = Format$(Time(), "HH:MM")
    cmdHora1.Caption = Format$(Time(), "HH:MM")
    
    CargarVariablesConfiguracion
    
    tmrCalcularRetraso.Interval = Val(VarCfg("timer_calcular_retraso"))
    
    picEPADance.Picture = LoadPicture(G_PATH_GRAFICO_HOJAS)

    mnuGenAutoPPC.Checked = G_GEN_AUTO_RESULTADOS_PPC

    If G_COUNTRY <> C_COUNTRY Then
        MsgBox "El parámetro COUNTRY no coincide con la constante de lenguaje C_COUNTRY", vbOKOnly Or vbExclamation, "ERROR"
    End If

error:
    ProcesarError
End Sub
Private Sub Form_Terminate()
    db.Close
    End
End Sub


Private Sub mnuACompeticiones_Click()
    frmACompeticiones.Show 1
End Sub


Private Sub mnuADorsales_Click()
    frmADorsales.Show 0
End Sub

Private Sub mnuAJueces_Click()
    frmAJueces.Show 1
End Sub

Private Sub mnuAParejas_Click()
    frmAParejas.Show
    
End Sub

Private Sub mnuAPuntuaciones_Click()
    On Local Error Resume Next
    frmAPuntuaciones.Show vbModal
End Sub

Private Sub mnuBailes_Click()
    frmBailes.Show 1
End Sub

Private Sub mnuCalcular_Click()
    frmCalcular.Show 1
End Sub

Private Sub mnuCompeticiones_Click()
    frmCompeticiones.Show 1
End Sub

Private Sub mnuDescalif_Click()
    frmDescalificados.Show 0
End Sub


Private Sub mnuEnviarCorreo_Click()
    correo.Show 1
End Sub

Private Sub mnuGenNumBase_Click()
  Dim ComputerInfo As cComputerInfo
  
  Set ComputerInfo = New cComputerInfo
  
  ' Print the value of each property
  Debug.Print "ActiveProcessorMask = " & ComputerInfo.ActiveProcessorMask
  Debug.Print "AllocationGranularity = " & ComputerInfo.AllocationGranularity
  Debug.Print "CompareExchangeDouble = " & ComputerInfo.CompareExchangeDouble
  Debug.Print "FloatingPointEmulated = " & ComputerInfo.FloatingPointEmulated
  Debug.Print "FloatingPointError = " & ComputerInfo.FloatingPointError
  Debug.Print "LowMemory = " & ComputerInfo.LowMemory
  Debug.Print "MaxAppAddress = " & ComputerInfo.MaxAppAddress
  Debug.Print "MinAppAddress = " & ComputerInfo.MinAppAddress
  Debug.Print "MMXAvailable = " & ComputerInfo.MMXAvailable
  Debug.Print "NumberOfProcessors = " & ComputerInfo.NumberOfProcessors
  Debug.Print "PageSize = " & ComputerInfo.PageSize
  
  Select Case ComputerInfo.ProcessorArchitecture
    Case cmiIntel
      Debug.Print "Intel processor"
    Case cmiMIPS
      Debug.Print "MIPS processor"
    Case cmiALPHA
      Debug.Print "Alpha processor"
    Case cmiPPC
      Debug.Print "Power PC processor"
    Case cmiUnknown
      Debug.Print "Unknown processor"
  End Select
  
  Debug.Print "ProcessorLevel = " & ComputerInfo.ProcessorLevel
  Debug.Print "ProcessorRevision = " & ComputerInfo.ProcessorRevision
  
  
  'MsgBox Format$(Hex(ComputerInfo.ProcessorRevision Xor 10600604) & "-" & Hex(GetHDDSerialNumber("C:\") Xor 10600604))
  
  MsgBox "Nº base: " & Format$(Hex(ComputerInfo.ProcessorRevision) & "-" & Hex(GetHDDSerialNumber("C:\")))
  
  Select Case ComputerInfo.ProcessorType
    Case cmiIntel386
      Debug.Print "Intel 386"
    Case cmiIntel486
      Debug.Print "Intel 486"
    Case cmiIntelPENTIUM
      Debug.Print "Intel Pentium"
    Case cmiMIPSR4000
      Debug.Print "MIPS R4000"
    Case cmiALPHA21064
      Debug.Print "Alpha 21064"
  End Select
  
  Debug.Print "SlowGraphics = " & ComputerInfo.SlowGraphics
  Debug.Print "SlowMachine = " & ComputerInfo.SlowMachine
  
  Set ComputerInfo = Nothing

End Sub


Private Sub mnuGrupos_Click()
    frmACategorias.Show 0
End Sub

Private Sub mnuHora_Click()
    frmCambiarHora.Show 1
End Sub

Private Sub mnuImpHojasPuntuaciones_Click()
    frmImpHojasPuntuaciones.Show 1
End Sub

Private Sub mnuImpInternet_Click()
    frmImprimirInternet.Show 1
End Sub

Private Sub mnuImpOrdenCombinado_Click()
    frmImprimirOrdenCombinado.Show 1
End Sub

Private Sub mnuImprimirCateg_Click()
    frmImprimirCateg.Show 1
End Sub

Private Sub mnuImprimirFinal_Click()
    frmImprimirFinal.Show 1
End Sub

Private Sub mnuImprimirParticipantes_Click()
    frmImprimirParticipantes.Show 1
End Sub

Private Sub mnuInsertarSociosAnulados_Click()
Dim sSocio As String

    If MsgBox(mml_FRASE0207, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then Exit Sub
    MsgBox mml_FRASE0192, vbOKOnly Or vbInformation, mml_FRASE0115
    On Local Error GoTo error
    db.Execute "DELETE FROM sociosanulados"
    Open G_ARCH_SOCIOS_ANULADOS For Input As #100
    While Not EOF(100)
        Line Input #100, sSocio
        db.Execute ("INSERT INTO sociosanulados VALUES (" & sSocio & ");")
    Wend
    Close #100
    Exit Sub
error:
If Err.Number <> 0 Then
   Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
         & Err.Source & Chr(13) & Err.Description & mml_FRASE0209 & G_ARCH_SOCIOS_ANULADOS
   MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
   Close #100
End If
End Sub

Private Function ComprobarNumeroBase() As String
Dim lValor As Long, lLong As Long, sValor As String * 100
Dim sNumero As String, sNum1 As String, sNum2 As String
Dim ComputerInfo As cComputerInfo
Dim sCodigo As String
  
  On Local Error GoTo error
  Set ComputerInfo = New cComputerInfo

    ComprobarNumeroBase = ""
    RegOpenKey HKEY_CURRENT_USER, "Applications\GenNumber" & Chr$(0), lValor
    If lValor = 0 Then Exit Function
    lLong = 80
    RegQueryValue lValor, "" & Chr$(0), sValor, lLong
    RegCloseKey lValor
    
    sNumero = Mid$(sValor, 1, InStr(sValor, Chr$(0)) - 1)
    
    sNum1 = Mid$(sNumero, 1, InStr(sNumero, "-") - 1)
    sNum2 = Mid$(sNumero, InStr(sNumero, "-") + 1)
    
    sNum1 = Val("&H" & sNum1) Xor 69
    sNum2 = Val("&H" & sNum2) Xor 69
    
    sCodigo = Trim$(Str$(Val(sNum1) Xor ComputerInfo.ProcessorRevision))

    If (Left$(sCodigo, 2) = "69" Or Left$(sCodigo, 2) = "70" Or Left$(sCodigo, 2) = "71" Or Left$(sCodigo, 2) = "72") And _
        sCodigo = Trim$(Str$(Val(sNum2) Xor GetHDDSerialNumber("C:\"))) And _
        CDate("01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2)) > Now Then
        ComprobarNumeroBase = "01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2)
    End If
error:
End Function

Private Sub mnuIntroNumBase_Click()
Dim sNumero As String, sNum1 As String, sNum2 As String
Dim lValor As Long, iFile As Integer, i As Long
    sNumero = InputBox(mml_FRASE0210)
    
    If InStr(sNumero, "-") = 0 Then Exit Sub
    If sNumero <> "" Then
        sNum1 = Mid$(sNumero, 1, InStr(sNumero, "-") - 1)
        sNum2 = Mid$(sNumero, InStr(sNumero, "-") + 1)
        
        sNum1 = Val("&H" & sNum1) Xor 69
        sNum2 = Val("&H" & sNum2) Xor 69
        
        sNumero = Hex$(sNum1) & "-" & Hex$(sNum2)
        
        RegCreateKey HKEY_CURRENT_USER, "Applications\GenNumber" & Chr$(0), lValor
        RegSetValue lValor, "" & Chr$(0), REG_SZ, sNumero & Chr$(0), 0
        RegCloseKey lValor
        
        On Local Error GoTo error
        'iFile = FreeFile
        'Open Environ$("WINDIR") & "\SYSTEM32\mfc26v.dll" For Output As #iFile
        'Randomize 200
        'For i = 0 To 45435
        '    Print #iFile, Chr$(Rnd() * 255);
        'Next
        'Close iFile
    End If
    Exit Sub
error:
End Sub


Private Sub mnuJuecesBailes_Click()
    frmAJuezBaile.Show 1
End Sub

Private Sub mnuJuecesCateg_Click()
    frmJuecesCompet.Show 1
    
End Sub

Private Sub mnuLecOptica_Click()
On Error GoTo error
    frmRecOptico.Show vbNomodal, Me
    Exit Sub
error:
    MsgBox mml_FRASE0211, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub


Private Sub mnuPublicar_Click()
    frmPublicar.Show 1
End Sub

Private Sub mnuPuntBaile_Click()
    frmAPuntuacionesBaile.Show 0
End Sub

Private Sub mnuPuntuaciones_Click()
    frmPuntuaciones.Show 0
End Sub

Private Sub mnuResultados_Click()
On Error GoTo error
    frmResultados.Show noModal
    Exit Sub
error:
    MsgBox mml_FRASE0211, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Private Sub mnuSalir_Click()
    On Local Error Resume Next
    If MsgBox(mml_FRASE0212, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        db.Close
        End
    End If
End Sub

Private Sub Timer5_Timer()
#If CONTROL_LICENCIA Then
    RetrasaSeg RETRASO_LICENCIA
#End If
End Sub

Private Sub mnuSobre_Click()
    frmSobre.Show 1
End Sub

Private Sub mnuTablasDatos_Click()
    frmTablasDatos.Show vbNomodal
End Sub

Private Sub Picture4_Click()
    If Me.Tag <> FORMATO_PROG_PEQUE Then
        Me.Tag = FORMATO_PROG_PEQUE
        Me.Top = 0
        Me.Height = 6500
        Me.Width = 2430
        frmPeque.Visible = True
        frmGrande.Visible = False
        Me.Left = 0
    Else
        Me.Tag = FORMATO_PROG_GRANDE
        Me.Height = 7110
        Me.Width = 10020
        frmPeque.Visible = False
        frmGrande.Visible = True
        Me.Left = Screen.Width \ 2 - Me.Width \ 2
        Me.Top = Screen.Height \ 2 - Me.Height \ 2
    End If
End Sub

Private Sub Timer1_Timer()
    cmdHora.Caption = Format$(Time(), "HH:MM")
    cmdHora1.Caption = Format$(Time(), "HH:MM")
End Sub

Private Sub Timer2_Timer()
#If CONTROL_LICENCIA Then
    If iContTimer = 0 Then
        If ComprobarNumeroBase() = "" Then
            frmMenu.Timer4.Enabled = True
        End If
        Timer2.Enabled = False
    Else
        iContTimer = iContTimer - 1
    End If
#End If
End Sub

Private Sub Timer4_Timer()
#If CONTROL_LICENCIA Then
    On Local Error Resume Next
    RetrasaSeg RETRASO_LICENCIA
#End If
End Sub
'*****************************************************************************************************************
Public Function ComprobarBase() As String
Dim lValor As Long, lLong As Long, sValor As String * 100
Dim sNumero As String, sNum1 As String, sNum2 As String
Dim ComputerInfo As cComputerInfo
Dim sCodigo As String
  
  On Local Error GoTo error
  Set ComputerInfo = New cComputerInfo

    ComprobarBase = ""
    RegOpenKey HKEY_CURRENT_USER, "Applications\GenNumber" & Chr$(0), lValor
    If lValor = 0 Then Exit Function
    lLong = 80
    RegQueryValue lValor, "" & Chr$(0), sValor, lLong
    RegCloseKey lValor
    
    sNumero = Mid$(sValor, 1, InStr(sValor, Chr$(0)) - 1)
    
    sNum1 = Mid$(sNumero, 1, InStr(sNumero, "-") - 1)
    sNum2 = Mid$(sNumero, InStr(sNumero, "-") + 1)
    
    sNum1 = Val("&H" & sNum1) Xor 69
    sNum2 = Val("&H" & sNum2) Xor 69
    
    sCodigo = Trim$(Str$(Val(sNum1) Xor ComputerInfo.ProcessorRevision))

    dBase = CDate("01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2))
    
    If (Left$(sCodigo, 2) = "69" Or Left$(sCodigo, 2) = "70" Or Left$(sCodigo, 2) = "71" Or Left$(sCodigo, 2) = "72" Or Left$(sCodigo, 2) = "73") And _
        sCodigo = Trim$(Str$(Val(sNum2) Xor GetHDDSerialNumber("C:\"))) And _
        CDate("01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2)) > Now Then
        ComprobarBase = "01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2)
    End If
error:
End Function
Public Function IniciarBD() As String
Dim lValor As Long, lLong As Long, sValor As String * 100
Dim sNumero As String, sNum1 As String, sNum2 As String
Dim ComputerInfo As cComputerInfo
Dim sCodigo As String
  
On Local Error GoTo error
  Set ComputerInfo = New cComputerInfo

    IniciarBD = ""
    RegOpenKey HKEY_CURRENT_USER, "Applications\GenNumber" & Chr$(0), lValor
    If lValor = 0 Then Exit Function
    lLong = 80
    RegQueryValue lValor, "" & Chr$(0), sValor, lLong
    RegCloseKey lValor
    
    sNumero = Mid$(sValor, 1, InStr(sValor, Chr$(0)) - 1)
    
    sNum1 = Mid$(sNumero, 1, InStr(sNumero, "-") - 1)
    sNum2 = Mid$(sNumero, InStr(sNumero, "-") + 1)
    
    sNum1 = Val("&H" & sNum1) Xor 69
    sNum2 = Val("&H" & sNum2) Xor 69
    
    sCodigo = Trim$(Str$(Val(sNum1) Xor ComputerInfo.ProcessorRevision))

    If (Left$(sCodigo, 2) = "69" Or Left$(sCodigo, 2) = "70" Or Left$(sCodigo, 2) = "71" Or Left$(sCodigo, 2) = "72" Or Left$(sCodigo, 2) = "73") And _
        sCodigo = Trim$(Str$(Val(sNum2) Xor GetHDDSerialNumber("C:\"))) And _
        CDate("01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2)) > Now Then
        IniciarBD = "01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2)
    End If
error:
End Function

'****************************************************************************************************************
Public Sub InicMenu()
Dim lValor As Long, lLong As Long, sValor As String * 100
Dim sNumero As String, sNum1 As String, sNum2 As String
Dim ComputerInfo As cComputerInfo
Dim sCodigo As String
  
  On Local Error GoTo error
  Set ComputerInfo = New cComputerInfo

    RegOpenKey HKEY_CURRENT_USER, "Applications\GenNumber" & Chr$(0), lValor
    If lValor = 0 Then Exit Sub
    lLong = 80
    RegQueryValue lValor, "" & Chr$(0), sValor, lLong
    RegCloseKey lValor
    
    sNumero = Mid$(sValor, 1, InStr(sValor, Chr$(0)) - 1)
    
    sNum1 = Mid$(sNumero, 1, InStr(sNumero, "-") - 1)
    sNum1 = Val("&H" & sNum1) Xor 69
    sCodigo = Trim$(Str$(Val(sNum1) Xor ComputerInfo.ProcessorRevision))

    On Local Error Resume Next
    
    If Left$(sCodigo, 2) = "69" Then
        frmMenu.mnuLecVis.Visible = True
        frmMenu.mnuLecOptica.Visible = True
        frmMenu.mnuResultados.Visible = True
    ElseIf Left$(sCodigo, 2) = "70" Then
        frmMenu.mnuLecVis.Visible = True
        frmMenu.mnuLecOptica.Visible = True
        frmMenu.mnuResultados.Visible = False
        cmdPublicar.Visible = False
    ElseIf Left$(sCodigo, 2) = "71" Then
        frmMenu.mnuLecVis.Visible = True
        frmMenu.mnuResultados.Visible = True
        frmMenu.mnuLecOptica.Visible = False
        cmdEscanear.Visible = False
        cmdRecOptico.Visible = False
    Else
        frmMenu.mnuLecOptica.Visible = False
        cmdEscanear.Visible = False
        cmdRecOptico.Visible = False
        frmMenu.mnuResultados.Visible = False
        cmdPublicar.Visible = False
        frmMenu.mnuLecVis.Visible = False
    End If
error:
End Sub


Private Sub Timer6_Timer()
Dim rs As Recordset
#If CONTROL_LICENCIA Then
    Set rs = db.OpenRecordset("SELECT MAX(fecha) FROM competiciones", dbOpenSnapshot)
    If IsDate(frmMenu.IniciarBD()) Then
        If DateDiff("m", frmMenu.IniciarBD(), rs.Fields(0)) > 2 Then
            frmMenu.Timer7.Enabled = True
        End If
    End If
    rs.Close
#End If
End Sub

Private Sub Timer7_Timer()
#If CONTROL_LICENCIA Then
    RetrasaSeg RETRASO_LICENCIA
#End If
End Sub

Private Sub Timer8_Timer()
Dim sBase As String * 255
Dim iBDLon As Integer
Dim sFecha1 As String

#If CONTROL_LICENCIA Then
    iBDLon = GetPrivateProfileString(mml_FRASE0033, "CodUltReg", "", sBase, 255, "Escrutinio.ini")
    If iBDLon > 0 Then
        sBase = Left$(sBase, iBDLon) & Chr$(0)
        sBase = Format$(CDbl(sBase) - 567890#, "000000")
        sFecha1 = Mid$(sBase, 5, 2) & "/" & Mid$(sBase, 3, 2) & "/" & Mid$(sBase, 1, 2)
        'Comprobamos si caduco la licencia o si el usuario retrasó la fecha
        If CDate(sFecha1) > Now Then
            frmMenu.Timer9.Enabled = True
            Exit Sub
        End If
    End If
    sBase = Str$(CDbl(Format$(Now, "yymmdd")) + 567890#)
    sBase = Left$(sBase, 8) & Chr$(0)
    WritePrivateProfileString mml_FRASE0033, "CodUltReg", sBase, "Escrutinio.ini"
#End If
End Sub

Private Sub Timer9_Timer()
#If CONTROL_LICENCIA Then
    MensajeSeg 10
    MsgBox mml_FRASE0967, vbOKOnly Or vbInformation, mml_FRASE0084
    End
#End If
End Sub

Private Sub tmrCalcularRetraso_Timer()
#If CONTROL_LICENCIA Then
#End If
End Sub
