VERSION 5.00
Begin VB.Form frmEnlacePPC_HTML 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0577"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPPC 
      Caption         =   "mml_FRASE0577"
      Height          =   9465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.Timer tmrEspera 
         Enabled         =   0   'False
         Interval        =   40000
         Left            =   0
         Top             =   6120
      End
      Begin VB.CommandButton cmdGenDatosFaseSig 
         Caption         =   "mml_FRASE0580"
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
         Left            =   8820
         TabIndex        =   99
         Top             =   1260
         Width           =   1710
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
         TabIndex        =   80
         Top             =   1305
         Width           =   1365
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
         TabIndex        =   79
         Top             =   1320
         Width           =   3870
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
         TabIndex        =   78
         Top             =   1320
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
         TabIndex        =   77
         Top             =   840
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
         TabIndex        =   76
         Top             =   840
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
         TabIndex        =   75
         Top             =   360
         Width           =   5895
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
         TabIndex        =   74
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdCategAct 
         Height          =   375
         Left            =   9060
         Picture         =   "EnlacePPC_HTML.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "mml_FRASE0428"
         Top             =   360
         Width           =   675
      End
      Begin VB.CommandButton cmdDorsales 
         Height          =   375
         Left            =   9840
         Picture         =   "EnlacePPC_HTML.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "mml_FRASE0028"
         Top             =   360
         Width           =   645
      End
      Begin VB.CommandButton cmdDescalif 
         Height          =   435
         Left            =   10020
         Picture         =   "EnlacePPC_HTML.frx":0E2C
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "mml_FRASE0043"
         Top             =   780
         Width           =   465
      End
      Begin VB.Frame Frame2 
         Caption         =   "mml_FRASE0582"
         Height          =   1530
         Left            =   450
         TabIndex        =   28
         Top             =   2640
         Width           =   9915
         Begin VB.TextBox tbCodBaileAct 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   7440
            TabIndex        =   102
            Top             =   600
            Width           =   2355
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
            Left            =   945
            TabIndex        =   34
            Top             =   630
            Width           =   4590
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
            Left            =   105
            TabIndex        =   33
            Top             =   630
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
            Left            =   1365
            TabIndex        =   32
            Top             =   210
            Width           =   6720
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
            Left            =   525
            TabIndex        =   31
            Top             =   210
            Width           =   855
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
            Left            =   8190
            TabIndex        =   30
            Top             =   180
            Width           =   1320
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
            Picture         =   "EnlacePPC_HTML.frx":1A0E
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   225
            Width           =   405
         End
         Begin VB.Label Label4 
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
            Left            =   5580
            TabIndex        =   100
            Top             =   690
            Width           =   1770
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
            Left            =   90
            TabIndex        =   70
            Top             =   1050
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
            Left            =   420
            TabIndex        =   69
            Top             =   1050
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
            Left            =   630
            TabIndex        =   68
            Top             =   1050
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
            Left            =   960
            TabIndex        =   67
            Top             =   1050
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
            Left            =   1710
            TabIndex        =   66
            Top             =   1050
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
            Left            =   2040
            TabIndex        =   65
            Top             =   1050
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
            Left            =   2250
            TabIndex        =   64
            Top             =   1050
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
            Left            =   2580
            TabIndex        =   63
            Top             =   1050
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
            Left            =   2790
            TabIndex        =   62
            Top             =   1050
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
            Left            =   3120
            TabIndex        =   61
            Top             =   1050
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
            Left            =   3330
            TabIndex        =   60
            Top             =   1050
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
            Left            =   3660
            TabIndex        =   59
            Top             =   1050
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
            Left            =   3870
            TabIndex        =   58
            Top             =   1050
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
            Left            =   4200
            TabIndex        =   57
            Top             =   1050
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
            Left            =   4410
            TabIndex        =   56
            Top             =   1050
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
            Left            =   4740
            TabIndex        =   55
            Top             =   1050
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
            Left            =   4950
            TabIndex        =   54
            Top             =   1050
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
            Left            =   5280
            TabIndex        =   53
            Top             =   1050
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
            Left            =   5490
            TabIndex        =   52
            Top             =   1050
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
            Left            =   5820
            TabIndex        =   51
            Top             =   1050
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
            Left            =   6030
            TabIndex        =   50
            Top             =   1050
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
            Left            =   6360
            TabIndex        =   49
            Top             =   1050
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
            Left            =   6570
            TabIndex        =   48
            Top             =   1050
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
            Left            =   6900
            TabIndex        =   47
            Top             =   1050
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
            Left            =   7110
            TabIndex        =   46
            Top             =   1050
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
            Left            =   7440
            TabIndex        =   45
            Top             =   1050
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
            Left            =   7650
            TabIndex        =   44
            Top             =   1050
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
            Left            =   7980
            TabIndex        =   43
            Top             =   1050
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
            Left            =   8190
            TabIndex        =   42
            Top             =   1050
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
            Index           =   15
            Left            =   8520
            TabIndex        =   41
            Top             =   1050
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
            Index           =   2
            Left            =   1500
            TabIndex        =   40
            Top             =   1050
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
            Left            =   1170
            TabIndex        =   39
            Top             =   1050
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
            Left            =   9060
            TabIndex        =   38
            Top             =   1050
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
            Left            =   8730
            TabIndex        =   37
            Top             =   1050
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
            Left            =   9600
            TabIndex        =   36
            Top             =   1050
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
            Index           =   17
            Left            =   9270
            TabIndex        =   35
            Top             =   1050
            Visible         =   0   'False
            Width           =   345
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "mml_FRASE0580"
         Height          =   1605
         Left            =   450
         TabIndex        =   20
         Top             =   4200
         Width           =   9915
         Begin VB.CommandButton cmdBorrarSig 
            Caption         =   "mml_FRASE0251"
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
            Left            =   7950
            TabIndex        =   104
            Top             =   210
            Width           =   1860
         End
         Begin VB.TextBox tbCodBaileSig 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   7380
            TabIndex        =   103
            Top             =   660
            Width           =   2415
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
            Left            =   6120
            TabIndex        =   27
            Top             =   270
            Width           =   1770
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
            TabIndex        =   26
            Top             =   270
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
            TabIndex        =   25
            Top             =   270
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
            TabIndex        =   23
            Top             =   660
            Width           =   4140
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
            TabIndex        =   22
            Top             =   1080
            Width           =   9765
         End
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
            Picture         =   "EnlacePPC_HTML.frx":1E78
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   270
            Width           =   405
         End
         Begin VB.Label Label5 
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
            Left            =   5430
            TabIndex        =   101
            Top             =   720
            Width           =   1770
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   0
         Top             =   2730
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
         TabIndex        =   19
         Top             =   1740
         Value           =   1  'Checked
         Width           =   5025
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
         Left            =   1890
         TabIndex        =   18
         Top             =   2070
         Width           =   2565
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
         ItemData        =   "EnlacePPC_HTML.frx":22E2
         Left            =   765
         List            =   "EnlacePPC_HTML.frx":2304
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1950
         Width           =   1065
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
         Picture         =   "EnlacePPC_HTML.frx":2337
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "mml_FRASE0425"
         Top             =   780
         Width           =   465
      End
      Begin VB.ComboBox cbBailes 
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
         Left            =   6975
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2070
         Width           =   1800
      End
      Begin VB.CommandButton cmdSelComp 
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
         Height          =   360
         Left            =   1845
         Picture         =   "EnlacePPC_HTML.frx":2619
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
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
         Picture         =   "EnlacePPC_HTML.frx":2A83
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   855
         Width           =   450
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
         Picture         =   "EnlacePPC_HTML.frx":2EED
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1305
         Width           =   450
      End
      Begin VB.CommandButton cmdGenDatosFaseAct 
         Caption         =   "mml_FRASE0582"
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
         TabIndex        =   11
         Top             =   1650
         Width           =   1710
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
         Left            =   450
         TabIndex        =   10
         Top             =   6990
         Width           =   2865
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
         Left            =   6420
         TabIndex        =   9
         Top             =   6990
         Width           =   2370
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
         TabIndex        =   8
         Top             =   6825
         Width           =   1515
      End
      Begin VB.Frame Frame4 
         Caption         =   "mml_FRASE1041"
         Height          =   930
         Left            =   420
         TabIndex        =   5
         Top             =   5850
         Width           =   9945
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
            TabIndex        =   7
            Top             =   240
            Width           =   9015
         End
         Begin VB.CommandButton cmdBorrarControlJueces 
            Caption         =   "mml_FRASE1221"
            Height          =   585
            Left            =   9090
            TabIndex        =   6
            Top             =   240
            Width           =   705
         End
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
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   8370
         Width           =   10275
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
         Left            =   3360
         TabIndex        =   3
         Top             =   6990
         Width           =   2970
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
         Left            =   1890
         TabIndex        =   2
         Top             =   2370
         Width           =   2385
      End
      Begin VB.CommandButton cmdPuntuaciones 
         Height          =   435
         Left            =   9540
         Picture         =   "EnlacePPC_HTML.frx":3357
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "mml_FRASE0040"
         Top             =   780
         Width           =   465
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
         TabIndex        =   98
         Top             =   1320
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
         TabIndex        =   97
         Top             =   840
         Width           =   1575
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
         TabIndex        =   96
         Top             =   360
         Width           =   1575
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
         TabIndex        =   95
         Tag             =   "0"
         Top             =   2070
         Width           =   1740
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
         TabIndex        =   94
         Top             =   1995
         Width           =   630
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
         TabIndex        =   93
         Top             =   1785
         Width           =   1770
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
         Left            =   120
         TabIndex        =   92
         Top             =   7590
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
         Left            =   930
         TabIndex        =   91
         Top             =   7590
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
         Left            =   1740
         TabIndex        =   90
         Top             =   7590
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
         Left            =   2550
         TabIndex        =   89
         Top             =   7590
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
         Left            =   3360
         TabIndex        =   88
         Top             =   7590
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
         Left            =   4170
         TabIndex        =   87
         Top             =   7590
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
         Left            =   4980
         TabIndex        =   86
         Top             =   7590
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
         Left            =   5760
         TabIndex        =   85
         Top             =   7590
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
         Left            =   6570
         TabIndex        =   84
         Top             =   7590
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
         Left            =   7380
         TabIndex        =   83
         Top             =   7590
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
         Left            =   8190
         TabIndex        =   82
         Top             =   7590
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
         Left            =   9000
         TabIndex        =   81
         Top             =   7590
         Visible         =   0   'False
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmEnlacePPC_HTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const C_MAX_JUEZ_PANEL = 18
Dim TiempoEspera As Boolean

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
    tbJuecesAct.Text = ""

    db.Execute "DELETE FROM jueces_html"
End Sub

Private Sub cmdBorrarSig_Click()
    tbCodCatSig.Text = ""
    tbDescCatSig.Text = ""
    tbCodFaseSig.Text = ""
    tbDescFaseSig.Text = ""
    chkRepSig.Value = 0
    tbCodBaileSig.Text = ""
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
        
        If rs!cod_baile > 0 Then
            Dim i As Integer
            
            For i = 0 To cbBailes.ListCount - 1
                If Val(cbBailes.List(i)) = rs!cod_baile Then
                    cbBailes.ListIndex = i
                End If
            Next
        ElseIf rs!cod_baile < 0 Then
            chkUltimos5Bailes.Value = 1
        Else
            If C_RESET_ULTIMOS_5_BAILES Then chkUltimos5Bailes.Value = 0
        End If
        
    Else
        MsgBox mml_FRASE0442, vbOKOnly Or vbInformation, mml_FRASE0147
    End If
    rs.Close

    Exit Sub
error:
    PPCLog ProcesarError("cmdCategAct_Click", False)
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


Private Sub cmdGenDatosFaseAct_Click()
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    If tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    
    tbCodCatAct.Text = tbCodCat.Text
    tbDescCatAct.Text = tbDescCat.Text
    tbCodFaseAct.Text = tbCodFase.Text
    tbDescFaseAct.Text = tbDescFase.Text
    chkRepAct.Value = chkRep.Value
    tbCodBaileAct.Text = cbBailes.Text

    ActualizarCateg tbCodComp.Text, tbCodCatAct.Text, Val(tbCodFaseAct.Text), chkRepAct.Value, Val(tbCodBaileAct.Text), cbPista.Text, "ACTUAL"

error:
    PPCLog ProcesarError("cmdGenDatosFaseAct_Click", False)
End Sub

Private Sub cmdGenDatosFaseSig_Click()
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    If tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    tbCodCatSig.Text = tbCodCat.Text
    tbCodFaseSig.Text = tbCodFase.Text
    chkRepSig.Value = chkRep.Value
    tbCodBaileSig.Text = cbBailes.Text
    tbDescCatSig.Text = tbDescCat.Text
    tbDescFaseSig.Text = tbDescFase.Text
    
    ActualizarCateg tbCodComp.Text, tbCodCatSig.Text, tbCodFaseSig.Text, chkRepSig.Value, Val(tbCodBaileSig.Text), cbPista.Text, "SIGUIENTE"

error:
    PPCLog ProcesarError("cmdGenDatosFaseSig_Click", False)
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
        frmAPuntuacionesBaile.tbCodComp.Text = tbCodComp.Text
        frmAPuntuacionesBaile.tbDescComp.Text = tbDescComp.Text
        frmAPuntuacionesBaile.tbCodCateg.Text = tbCodCat.Text
        frmAPuntuacionesBaile.tbDescCateg.Text = tbDescCat.Text
        frmAPuntuacionesBaile.cbFase.ListIndex = Log(Val(tbCodFase.Text)) / Log(2) + 1
        
        frmAPuntuacionesBaile.Show vbNomodal
        
    End If
    Exit Sub
error:
    PPCLog ProcesarError("cmdPuntuaciones_Click", False)
End Sub

Private Sub cmdSalir_Click()
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
    
    Exit Sub
error:
    PPCLog ProcesarError("cmdSelCat_Click", False)
End Sub

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

Private Sub cmdSubirDatos_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    
    If Val(tbCodCatAct.Text) > 0 And Val(tbCodFaseAct.Text) > 0 Then
        tbCodComp.Text = VarCfg("horario_codcompeticion")
        tbCodCat.Text = tbCodCatAct.Text
        tbCodFase.Text = tbCodFaseAct.Text
        chkRep.Value = chkRepAct.Value
        tbDescCat.Text = tbDescCatAct.Text
        tbDescFase.Text = tbDescFaseAct.Text
        
        CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
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
        
    End If
    Exit Sub
error:
    PPCLog ProcesarError("cmdSubirDatosSigFase_Click", False)

End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    G_INTERVALO_CONTROL_CALCULO_PDA = Val(VarCfg("pda_intervalo_min_calculo", "60", "Intervalo m�nimo entre c�lculos autom�ticos de tablets y PDAs en segundos", "Minimum interval between automatic calculations of tablets and PDAs in seconds"))
    
    CargarPistas cbPista
    CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
    
    chkCalcAuto.Value = IIf(frmMenu.mnuGenAutoPPC.Checked, 1, 0)
    
    IniciarControlTiempo
    
    Exit Sub
error:
    ProcesarError "Form_Load"
End Sub

Sub PPCLog(sCad As String)
    If Len(tbLog.Text) + Len(sCad) > 30000 Then
        tbLog.Text = ""
    End If
    If sCad <> "" Then
        tbLog.Text = sCad & vbCrLf & tbLog.Text
    End If
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
        If tbCodCatAct.Text = "" Or tbCodFaseAct.Text = "" Then
            CamposSinCubrir
            Exit Sub
        End If
        
        If DistintaPista(cbPista.Text) Then
            lblActTimer.BackColor = vbGreen
            lblActTimer.Caption = mml_FRASE0592
            lblActTimer.Tag = "1"
            Timer1.Interval = G_INTERVALO_TIMER_PPC
            Timer1.Enabled = True
            IniciarControlTiempo
            
            'Si no est� calculada la fase siguiente, la calculamos
            If chkGenSigCat.Value = 1 And tbCodCatSig.Text = "" Then
                Set rs = db.OpenRecordset(SQL_SIG_CATEG, dbOpenSnapshot)
                If Not rs.EOF Then
                    tbDescCatSig.Text = sDescCategoria(rs!cod_categoria)
                    tbCodCatSig.Text = rs!cod_categoria
                    tbDescFaseSig.Text = sDescFase(rs!numfase)
                    tbCodFaseSig.Text = rs!numfase
                    chkRepSig.Value = rs!repesca
                    tbCodBaileSig.Text = rs!cod_baile & " - " & sNombreBaile(rs!cod_baile)
                Else
                    tbDescCatSig.Text = ""
                    tbCodCatSig.Text = ""
                    tbDescFaseSig.Text = ""
                    tbCodFaseSig.Text = ""
                    chkRepSig.Value = 0
                    tbCodBaileSig.Text = ""
                End If
                rs.Close
                    
                'Actualizamos la informaci�n de la siguiente categoria
                ActualizarCateg tbCodComp.Text, Val(tbCodCatSig.Text), Val(tbCodFaseSig.Text), chkRepSig.Value, Val(tbCodBaileSig.Text), cbPista.Text, "SIGUIENTE"
                
            End If
        Else
            MsgBox mml_FRASE1174, vbOKOnly Or vbCritical, G_MSG_ERROR
        End If
    Else
        DesactivarControlAutomatico
    End If
    Exit Sub
error:
    PPCLog ProcesarError("lblActTimer_Click", False)
End Sub
Private Sub DesactivarControlAutomatico()
    lblActTimer.BackColor = vbRed
    lblActTimer.Caption = mml_FRASE0583
    lblActTimer.Tag = "0"
    Timer1.Enabled = False
    tmrEspera.Enabled = False
End Sub

Private Sub tbCodCat_Change()
    tbCodFase.Text = ""
    tbDescFase.Text = ""
End Sub

Private Sub tbCodCat_LostFocus()
    tbDescCat.Text = sDescCategoria(Val(tbCodCat.Text), Val(tbCodComp.Text))
    If tbDescCat.Text = "" Then tbCodCat.Text = ""
End Sub

Private Sub Timer1_Timer()
    Dim rs As Recordset
    Dim i As Integer
    Dim iNumJueces As Integer
    Dim iMaxBailes As Integer
    Dim iNumBailes As Integer
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    'Comprobamos jueces introducidos
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM jueces_html WHERE cod_categoria = " & tbCodCatAct.Text & " ORDER BY cod_juez", dbOpenSnapshot)
    tbJuecesAct.Text = ""
    While Not rs.EOF
        tbJuecesAct.Text = tbJuecesAct.Text & " " & rs.Fields("cod_juez")
        rs.MoveNext
    Wend
    rs.Close
    
    'Comprobamos el estado de puntuaciones por juez
    For i = 0 To C_MAX_JUEZ_PANEL - 1
        tbJuezTx(i).Tag = ""
    Next
    
    i = 0
    Set rs = db.OpenRecordset("SELECT DISTINCT id_juez FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCatAct.Text & " ORDER BY id_juez", dbOpenSnapshot)
    While Not rs.EOF
        tbJuezTx(i).Caption = rs.Fields("id_juez")
        tbJuezTx(i).Tag = rs.Fields("id_juez")
        tbNumBAilesTx(i).Caption = "0"
        tbJuezTx(i).Visible = True
        tbNumBAilesTx(i).Visible = True
        rs.MoveNext
        Inc i
    Wend
    iNumJueces = i
    For i = iNumJueces To C_MAX_JUEZ_PANEL - 1
        tbJuezTx(i).Visible = False
        tbNumBAilesTx(i).Visible = False
    Next
    iMaxBailes = 0
    rs.Close
    Set rs = db.OpenRecordset("SELECT cod_juez, cod_baile, COUNT(*) FROM puntuaciones p WHERE p.cod_categoria = " & tbCodCatAct.Text & " AND fase = " & tbCodFaseAct.Text & " AND repesca = " & chkRepAct.Value & " GROUP BY cod_juez, cod_baile ORDER BY cod_juez", dbOpenSnapshot)
    i = 0
    While Not rs.EOF And i < iNumJueces
        If tbJuezTx(i).Caption = rs.Fields("cod_juez") Then
            tbNumBAilesTx(i).Caption = Val(tbNumBAilesTx(i).Caption) + 1
            Inc iNumBailes
            rs.MoveNext
        Else
            iNumBailes = 0
            Inc i
        End If
        
        If iNumBailes > iMaxBailes Then
            iMaxBailes = iNumBailes
        End If
    Wend
    
    For i = 0 To iNumJueces - 1
        If tbNumBAilesTx(i).Caption = iMaxBailes Then
            tbJuezTx(i).BackColor = C_COLOR_VERDE
        Else
            tbJuezTx(i).BackColor = C_COLOR_ROJO
        End If
    Next
    
    
    If chkCalcAuto.Value = 1 Then
        ComprobarPuntuaciones
    End If
    
    Exit Sub
error:
    PPCLog ProcesarError("Timer1_Click", False)
End Sub
Private Sub IniciarControlTiempo()
    tmrEspera.Enabled = False
    TiempoEspera = False
    tmrEspera.Interval = G_INTERVALO_CONTROL_CALCULO_PDA * 1000
    tmrEspera.Enabled = True
End Sub

Sub ComprobarPuntuaciones()
Dim lCateg As Long
Dim iFase As Integer
Dim iRep As Integer
Dim iTodasPuntuaciones As Integer
Dim sSQL As String

    If Not C_DEBUG Then On Local Error GoTo error
    lCateg = Val(tbCodCatAct.Text)
    iFase = Val(tbCodFaseAct.Text)
    iRep = chkRepAct.Value
    'Comprobamos si ya est�n todas las puntuaciones
    'Solo calculamos automaticamente si podemos avanzar de categoria
    If G_CALCULO_AUTO_PPC And chkGenSigCat.Value = 1 Then
        If Val(tbCodCatAct.Text) = 0 Or Val(tbCodFaseAct.Text) = 0 Then Exit Sub
        
        iTodasPuntuaciones = ComprobarSiEstanTodasPuntuacionesDorsalesPresentes(Val(tbCodCatAct.Text), chkRepAct.Value, tbCodFaseAct.Text)
        
        If iTodasPuntuaciones = C_TODAS_PUNTUACIONES Or _
           (iTodasPuntuaciones = C_PUNTUACIONES_CINCO_BAILES And Val(tbCodBaileAct.Text) = 0) Then
            'Est�n todas las puntuaciones por lo que avanzamos de categor�a
            tbCodCatAct.Text = tbCodCatSig.Text
            tbDescCatAct.Text = tbDescCatSig.Text
            tbCodFaseAct.Text = tbCodFaseSig.Text
            tbDescFaseAct.Text = tbDescFaseSig.Text
            chkRepAct.Value = chkRepSig.Value
            tbCodBaileAct.Text = tbCodBaileSig.Text
            
            If tbCodCatSig.Text <> "" Then
                'Si no se selecciona el modo de 10 bailes, independientemente del baile de partida configurado, siempre encontrar� la fase
                'Si est� activo el modo de 10 bailes, solo localiza cada fase si el baile seleccionado coincide con el del horario
                'Calculamos la siguiente categoria si antes de entrar exist�a
                sSQL = SQL_SIG_CATEG
                Debug.Print sSQL
                Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
                If Not rs.EOF Then
                    tbDescCatSig.Text = sDescCategoria(rs!cod_categoria)
                    tbCodCatSig.Text = rs!cod_categoria
                    tbDescFaseSig.Text = sDescFase(rs!numfase)
                    tbCodFaseSig.Text = rs!numfase
                    chkRepSig.Value = rs!repesca
                    tbCodBaileSig.Text = rs!cod_baile & " - " & sNombreBaile(rs!cod_baile)
                Else
                    tbDescCatSig.Text = ""
                    tbCodCatSig.Text = ""
                    tbDescFaseSig.Text = ""
                    tbCodFaseSig.Text = ""
                    chkRepSig.Value = 0
                    tbCodBaileSig.Text = ""
                    
                End If
                rs.Close
                
                'Asignamos las nuevas categor�as en la base de datos
                ActualizarCateg tbCodComp.Text, Val(tbCodCatAct.Text), Val(tbCodFaseAct.Text), chkRepAct.Value, Val(tbCodBaileAct.Text), cbPista.Text, "ACTUAL"
                ActualizarCateg tbCodComp.Text, Val(tbCodCatSig.Text), Val(tbCodFaseSig.Text), chkRepSig.Value, Val(tbCodBaileSig.Text), cbPista.Text, "SIGUIENTE"
                
            End If
            If G_GEN_AUTO_RESULTADOS_PPC Then
                While frmCalcular.lblAutoPPC.BackColor = vbRed And lblActTimer.BackColor <> vbRed
                    Sleep 100
                    DoEvents
                Wend
                ' Mientras el control autom�tico siga activo
                If lblActTimer.BackColor <> vbRed Then
                    'Si estamos calculando de nuevo antes de que pasen 40 segundos desactivamos el sitema autom�tico
                    If Not TiempoEspera Then
                        DesactivarControlAutomatico
                        If MsgBox(mml_FRASE1285 & G_INTERVALO_CONTROL_CALCULO_PDA, vbYesNo Or vbQuestion, mml_FRASE0096) = vbYes Then
                            GoTo continuar
                        End If
                    Else
continuar:
                        'Solo calculamos si tenemos todas las puntuaciones, no cinco bailes
                        If iTodasPuntuaciones = C_TODAS_PUNTUACIONES Then
                            frmCalcular.lblAutoPPC.BackColor = vbRed
                            frmCalcular.tbCodComp.Text = tbCodComp.Text
                            frmCalcular.tbDescComp.Text = tbDescComp.Text
                            frmCalcular.tbCodCat.Text = lCateg
                            frmCalcular.tbDescCat.Text = sDescCategoria(lCateg)
                            frmCalcular.tbCodFase.Text = iFase
                            frmCalcular.tbDescFase.Text = sDescFase(iFase)
                            frmCalcular.chkRep.Value = iRep
                            
                            frmCalcular.Visible = True
                            DoEvents
                            frmCalcular.cmdCalcular_Click
                            frmCalcular.Visible = False
                            frmCalcular.lblAutoPPC.BackColor = vbGreen
                            frmCalcular.lblAutoPPC.Refresh
                        End If
                        
                        IniciarControlTiempo
                        
                        If tbCodCatAct.Text = "" Then
                            'Entramos a calcular la �ltima categor�a
                            If lblActTimer.Tag = "1" Then
                                lblActTimer_Click
                            End If
                            MsgBox mml_FRASE1284, vbOKOnly Or vbInformation, mml_FRASE0096
                        End If
                    End If
                End If
            ElseIf MsgBox(mml_FRASE0445, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                frmCalcular.lblAutoPPC.BackColor = vbRed
                frmCalcular.tbCodComp.Text = tbCodComp.Text
                frmCalcular.tbDescComp.Text = tbDescComp.Text
                frmCalcular.tbCodCat.Text = lCateg
                frmCalcular.tbDescCat.Text = sDescCategoria(lCateg)
                frmCalcular.tbCodFase.Text = iFase
                frmCalcular.tbDescFase.Text = sDescFase(iFase)
                frmCalcular.chkRep.Value = iRep
                
                frmCalcular.Show vbModal
                frmCalcular.lblAutoPPC.BackColor = vbGreen
            End If
        End If
    End If
    Exit Sub

error:
    PPCLog ProcesarError("tmrCalcular_Timer", False)
End Sub

Sub ActualizarCateg(CodComp As String, CodCatSig As String, CodFaseSig As String, RepSig As String, CodBaile As Integer, Pista As String, fase As String)
Dim sSQL As String
    db.Execute "DELETE FROM cfg_html WHERE fase = '" & fase & "' AND pista = '" & Pista & "'"
    If Pista = "" Then Pista = ""
    sSQL = "INSERT INTO cfg_html VALUES ('" & fase & "'," & CodComp & "," & CodCatSig & "," & CodFaseSig & "," & RepSig & "," & CodBaile & ",'" & Pista & "')"
    If C_DEBUG Then Debug.Print sSQL
    db.Execute sSQL
End Sub

Private Sub tmrEspera_Timer()
    TiempoEspera = True
End Sub

Function SQL_SIG_CATEG() As String
    SQL_SIG_CATEG = "SELECT TOP 1 * FROM horario WHERE grupo LIKE '*" & cbPista.Text & "*' AND numfase <> 99 AND cod_competicion = " & VarCfg("horario_codcompeticion") & " AND orden > (SELECT MAX(orden) FROM horario WHERE cod_baile " & IIf(chkUltimos5Bailes.Value, "=", "<=") & Val(tbCodBaileAct.Text) & " AND cod_categoria = " & tbCodCatAct.Text & " and numfase = " & tbCodFaseAct.Text & " AND repesca = " & chkRepAct.Value & " AND cod_competicion = " & VarCfg("horario_codcompeticion") & ") ORDER BY orden"
End Function
