VERSION 5.00
Begin VB.Form frmMatchAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MatchAnalysis"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10695
   Icon            =   "frmMatchAnalisys.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   10695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmPPC 
      Caption         =   "mml_FRASE0577"
      Height          =   9465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin VB.Frame Frame2 
         Caption         =   "mml_FRASE0582"
         Height          =   2070
         Left            =   450
         TabIndex        =   11
         Top             =   2640
         Width           =   9915
         Begin VB.TextBox tbPareja 
            Alignment       =   2  'Center
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
            Left            =   120
            TabIndex        =   103
            Top             =   1500
            Width           =   9675
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
            Picture         =   "frmMatchAnalisys.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   225
            Width           =   405
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
            TabIndex        =   17
            Top             =   180
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
            Left            =   525
            TabIndex        =   16
            Top             =   210
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
            TabIndex        =   15
            Top             =   210
            Width           =   6720
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
            TabIndex        =   14
            Top             =   630
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
            Left            =   945
            TabIndex        =   13
            Top             =   630
            Width           =   4590
         End
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
            TabIndex        =   12
            Top             =   600
            Width           =   2355
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            Index           =   2
            Left            =   1500
            TabIndex        =   50
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
            Index           =   15
            Left            =   8520
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
            Index           =   15
            Left            =   8190
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
            Index           =   14
            Left            =   7980
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
            Index           =   14
            Left            =   7650
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
            Index           =   13
            Left            =   7110
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
            Index           =   12
            Left            =   6900
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
            Index           =   12
            Left            =   6570
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
            Index           =   11
            Left            =   6360
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            Index           =   0
            Left            =   90
            TabIndex        =   20
            Top             =   1050
            Visible         =   0   'False
            Width           =   345
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
            TabIndex        =   19
            Top             =   690
            Width           =   1770
         End
      End
      Begin VB.CommandButton cmdPuntuaciones 
         Height          =   435
         Left            =   9720
         Picture         =   "frmMatchAnalisys.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "mml_FRASE0040"
         Top             =   780
         Width           =   705
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
         TabIndex        =   83
         Top             =   2370
         Width           =   2385
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
         Left            =   3330
         TabIndex        =   82
         Top             =   6990
         Width           =   2970
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
         TabIndex        =   81
         Top             =   8370
         Width           =   10275
      End
      Begin VB.Frame Frame4 
         Caption         =   "mml_FRASE1041"
         Height          =   930
         Left            =   420
         TabIndex        =   78
         Top             =   5850
         Width           =   9945
         Begin VB.CommandButton cmdBorrarControlJueces 
            Caption         =   "mml_FRASE1221"
            Height          =   585
            Left            =   9090
            TabIndex        =   80
            Top             =   240
            Width           =   705
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
            TabIndex        =   79
            Top             =   240
            Width           =   9015
         End
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
         TabIndex        =   77
         Top             =   6825
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
         Left            =   6390
         TabIndex        =   76
         Top             =   6990
         Width           =   2370
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
         TabIndex        =   75
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
         Picture         =   "frmMatchAnalisys.frx":148E
         Style           =   1  'Graphical
         TabIndex        =   74
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
         Picture         =   "frmMatchAnalisys.frx":18F8
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   855
         Width           =   450
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
         Picture         =   "frmMatchAnalisys.frx":1D62
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   450
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
         TabIndex        =   71
         Top             =   2070
         Width           =   1800
      End
      Begin VB.Timer tmrCalcular 
         Enabled         =   0   'False
         Left            =   0
         Top             =   6420
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
         Picture         =   "frmMatchAnalisys.frx":21CC
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "mml_FRASE0425"
         Top             =   780
         Width           =   585
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
         ItemData        =   "frmMatchAnalisys.frx":24AE
         Left            =   765
         List            =   "frmMatchAnalisys.frx":24D0
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   1950
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
         Left            =   1890
         TabIndex        =   68
         Top             =   2070
         Visible         =   0   'False
         Width           =   2565
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
         TabIndex        =   67
         Top             =   1740
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   5025
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   0
         Top             =   2730
      End
      Begin VB.Frame Frame3 
         Caption         =   "mml_FRASE0580"
         Height          =   1605
         Left            =   450
         TabIndex        =   56
         Top             =   4200
         Visible         =   0   'False
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
            Picture         =   "frmMatchAnalisys.frx":2503
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   270
            Width           =   405
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
            TabIndex        =   64
            Top             =   1080
            Width           =   9765
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
            TabIndex        =   63
            Top             =   660
            Width           =   4140
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
            TabIndex        =   62
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
            TabIndex        =   61
            Top             =   270
            Width           =   4740
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
            TabIndex        =   60
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
            Left            =   6120
            TabIndex        =   59
            Top             =   270
            Width           =   1770
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
            TabIndex        =   58
            Top             =   660
            Width           =   2415
         End
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
            TabIndex        =   57
            Top             =   210
            Width           =   1860
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
            TabIndex        =   66
            Top             =   720
            Width           =   1770
         End
      End
      Begin VB.CommandButton cmdDorsales 
         Height          =   375
         Left            =   9750
         Picture         =   "frmMatchAnalisys.frx":296D
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "mml_FRASE0028"
         Top             =   360
         Width           =   645
      End
      Begin VB.CommandButton cmdCategAct 
         Height          =   375
         Left            =   9060
         Picture         =   "frmMatchAnalisys.frx":348F
         Style           =   1  'Graphical
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   1320
         Width           =   3870
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
         TabIndex        =   2
         Top             =   1305
         Width           =   1365
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
         TabIndex        =   1
         Top             =   1260
         Visible         =   0   'False
         Width           =   1710
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
         TabIndex        =   102
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
         TabIndex        =   101
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
         TabIndex        =   100
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
         TabIndex        =   99
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
         TabIndex        =   98
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
         TabIndex        =   97
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   94
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
         TabIndex        =   93
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
         TabIndex        =   91
         Top             =   7590
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
         TabIndex        =   90
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
         TabIndex        =   89
         Top             =   1995
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
         TabIndex        =   88
         Tag             =   "0"
         Top             =   2070
         Width           =   1740
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
         Top             =   1320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmMatchAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const C_MAX_JUEZ_PANEL = 18
Const C_MAX_BATERIA_PANEL = 12


Private Sub cbPista_Click()
    Me.Caption = cbPista.Text & " " & mml_FRASE0577
End Sub

Private Sub cmdAvanzar_Click()
    If tbCodCatSig.Text = "" Then
        'No hay siguiente categoría
        MsgBox mml_FRASE0595, vbOKOnly Or vbInformation, mml_FRASE0096
    Else
        tbCodCatAct.Text = tbCodCatSig.Text
        tbDescCatAct.Text = tbDescCatSig.Text
        tbCodFaseAct.Text = tbCodFaseSig.Text
        tbDescFaseAct.Text = tbDescFaseSig.Text
        chkRepAct.Value = chkRepSig.Value
        tbCodBaileAct.Text = tbCodBaileSig.Text
        
        'Calculamos la siguiente categoria
        Set rs = db.OpenRecordset("SELECT TOP 1 * FROM horario WHERE grupo LIKE '" & scarLike & "" & cbPista.Text & "" & scarLike & "' AND numfase <> 99 AND cod_competicion = " & VarCfg("horario_codcompeticion") & " AND orden > (SELECT TOP 1 orden FROM horario WHERE cod_baile " & IIf(chkUltimos5Bailes.Value, "<", "=") & Val(tbCodBaileAct.Text) & " AND cod_categoria = " & tbCodCatAct.Text & " and numfase = " & tbCodFaseAct.Text & " AND repesca = " & chkRepAct.Value & " AND cod_competicion = " & VarCfg("horario_codcompeticion") & ") ORDER BY orden", dbOpenSnapshot)
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
        
        'Asignamos las nuevas categorías en la base de datos
        db.Execute "UPDATE cfg_html SET cod_comp = " & Val(tbCodComp.Text) & ",cod_categoria = " & Val(tbCodCatAct.Text) & ",cod_fase = " & Val(tbCodFaseAct.Text) & ",cod_rep = " & chkRepAct.Value & ", cod_baile = " & Val(tbCodBaileAct.Text) & " WHERE fase = 'ACTUAL'"
        db.Execute "UPDATE cfg_html SET cod_comp = " & Val(tbCodComp.Text) & ",cod_categoria = " & Val(tbCodCatSig.Text) & ",cod_fase = " & Val(tbCodFaseSig.Text) & ",cod_rep = " & chkRepSig.Value & ", cod_baile = " & Val(tbCodBaileSig.Text) & " WHERE fase = 'SIGUIENTE'"
        
    End If

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
    tbJuecesAct.Text = ""

    db.Execute "DELETE FROM jueces_html"
    db.Execute "DELETE FROM bateria_html"
End Sub

Private Sub cmdBorrarSig_Click()
    tbCodCatSig.Text = ""
    tbDescCatSig.Text = ""
    tbCodFaseSig.Text = ""
    tbDescFaseSig.Text = ""
    chkRepSig.Value = 0
    tbCodBaileSig.Text = ""
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

    db.Execute "UPDATE MatchAnalysis SET cod_comp = " & tbCodComp.Text & ",cod_categoria = " & tbCodCatAct.Text & ",cod_fase = " & tbCodFaseAct.Text & ",cod_rep = " & chkRepAct.Value & ", cod_baile = " & Val(tbCodBaileAct.Text) & " WHERE fase = 'ACTUAL'"

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
    
    db.Execute "UPDATE MatchAnalysis SET cod_comp = " & tbCodComp.Text & ",cod_categoria = " & tbCodCatSig.Text & ",cod_fase = " & tbCodFaseSig.Text & ",cod_rep = " & chkRepSig.Value & ", cod_baile = " & Val(tbCodBaileSig.Text) & " WHERE fase = 'SIGUIENTE'"

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
        With frmMAPuntuaciones
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
        
    End If
    Exit Sub
error:
    PPCLog ProcesarError("cmdSubirDatosSigFase_Click", False)

End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    CargarPistas cbPista
    CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
    
    chkCalcAuto.Value = IIf(frmMenu.mnuGenAutoPPC.Checked, 1, 0)
    
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
    
    If tbCodCatAct.Text = "" Or tbCodFaseAct.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If lblActTimer.Tag = "0" Then
        If DistintaPista(cbPista.Text, "frmMatchAnalysis") Then
            lblActTimer.BackColor = vbGreen
            lblActTimer.Caption = mml_FRASE0592
            lblActTimer.Tag = "1"
            Timer1.Interval = G_INTERVALO_TIMER_PPC
            Timer1.Enabled = True
            tmrCalcular.Interval = G_TIEMPO_ESPERA_JPASOS_PPC
            tmrCalcular.Enabled = True
            
            'Si no está calculada la fase siguiente, la calculamos
            If chkGenSigCat.Value = 1 And tbCodCatSig.Text = "" Then
                Set rs = db.OpenRecordset("SELECT TOP 1 * FROM horario WHERE grupo LIKE '" & scarLike & "" & cbPista.Text & "" & scarLike & "' AND numfase <> 99 AND cod_competicion = " & VarCfg("horario_codcompeticion") & " AND orden > (SELECT TOP 1 orden FROM horario WHERE cod_baile " & IIf(chkUltimos5Bailes.Value, "<", "=") & Val(tbCodBaileAct.Text) & " AND cod_categoria = " & tbCodCatAct.Text & " and numfase = " & tbCodFaseAct.Text & " AND repesca = " & chkRepAct.Value & " AND cod_competicion = " & VarCfg("horario_codcompeticion") & ") ORDER BY orden", dbOpenSnapshot)
                If Not rs.EOF Then
                    tbDescCatSig.Text = sDescCategoria(rs!cod_categoria)
                    tbCodCatSig.Text = rs!cod_categoria
                    tbDescFaseSig.Text = sDescFase(rs!numfase)
                    tbCodFaseSig.Text = rs!numfase
                    chkRepSig.Value = rs!repesca
                    tbCodBaileSig.Text = rs!cod_baile & " - " & sNombreBaile(rs!cod_baile)
                    
                    'Actualizamos la información de la siguiente categoria
                    db.Execute "UPDATE MatchAnalysis SET cod_comp = " & tbCodComp.Text & ",cod_categoria = " & tbCodCatSig.Text & ",cod_fase = " & tbCodFaseSig.Text & ",cod_rep = " & chkRepSig.Value & ", cod_baile = " & Val(tbCodBaileSig.Text) & " WHERE fase = 'SIGUIENTE'"
                    
                Else
                    MsgBox mml_FRASE0595, vbOKOnly Or vbInformation, mml_FRASE0096
                End If
                rs.Close
            End If
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

Private Sub tbCodCat_LostFocus()
    tbDescCat.Text = sDescCategoria(Val(tbCodCat.Text), Val(tbCodComp.Text))
    If tbDescCat.Text = "" Then tbCodCat.Text = ""
End Sub

Private Sub Timer1_Timer()
    Dim rs As Recordset
    Dim rsJuez As Recordset
    Dim i As Integer
    Dim sIdJuez As String
    Dim sHora As String
    Dim lDif As Long
    Dim iNumJueces As Integer
    Dim iMaxBailes As Integer
    Dim iNumBailes As Integer
    Dim iNumParametros As Integer
    Dim iNumPuntuaciones As Integer
    Dim iDorsal As Integer
    
    
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
    Set rs = db.OpenRecordset("SELECT DISTINCT id_juez FROM juez_categ WHERE cod_categoria = " & tbCodCatAct.Text & " ORDER BY id_juez", dbOpenSnapshot)
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
    'Recuperamos la información de la primera pareja que no tiene todos los bailes
    'Recuperamos la información del número de bailes
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCat.Text & " AND fase = " & IIf(Val(tbCodFase.Text) = 1, 1, 2), dbOpenSnapshot)
    iNumBailes = rs.Fields(0)
    rs.Close
    'Recuperamos la información del número de parámetros
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM MA_parametros WHERE activo = 'S'", dbOpenSnapshot)
    iNumParametros = rs.Fields(0)
    rs.Close
    
    'Recuperar el número total de parametros que deben tener puntuación en la categoría
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM Juez_parametro WHERE cod_categoria = " & tbCodCat.Text & " AND fase = " & IIf(Val(tbCodFase.Text) = 1, 1, 2), dbOpenSnapshot)
    iNumPuntuaciones = rs.Fields(0)
    rs.Close
    
    If iNumBailes = 0 Then
        lblActTimer_Click
        MsgBox mml_FRASE0591, vbOKOnly Or vbCritical, "ERROR"
        Exit Sub
    End If
    
    sSQL = "SELECT num_dorsal, nombre_hombre, nombre_mujer FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND d.cod_categoria = " & tbCodCat.Text & _
            " AND d.fase = " & tbCodFase.Text & " AND d.repesca = " & chkRep.Value & _
            " AND (SELECT COUNT(*) FROM MA_puntuaciones pu WHERE pu.num_dorsal = d.num_dorsal AND pu.cod_categoria = d.cod_categoria " & _
            " AND pu.fase = d.fase AND pu.repesca = d.repesca) < " & iNumBailes * iNumPuntuaciones & " ORDER BY num_dorsal"
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
    If Not rs.EOF Then
    Dim sPareja As String
        sPareja = tbPareja.Text
        tbPareja.Text = "(" & rs.Fields("num_dorsal") & ") - " & rs.Fields("nombre_hombre") & " & " & rs.Fields("nombre_mujer")
        iDorsal = rs.Fields("num_dorsal")
        'En el momento de cambiar de pareja calculamos la anterior
        If chkCalcAuto.Value = 1 And sPareja <> tbPareja.Text Then
            If sPareja <> "" Then
                'Realizamos una impresión automática con el cambiamos de pareja
                With frmMAPuntuaciones
                    .tbCodComp.Text = tbCodComp.Text
                    .tbCodCateg.Text = tbCodCatAct.Text
                    .chkRep.Value = chkRepAct.Value
                    .cbFase.ListIndex = Log(Val(tbCodFaseAct.Text)) / Log(2) + 1
                    .cbDorsal.Text = Val(Mid(sPareja, 2))
                    .cmdImprimir_Click
                End With
                Unload frmAPuntuaciones
            End If
        End If
    Else
        MsgBox "Están introducidas todas las puntuaciones", vbOKOnly Or vbInformation, ""
        Timer1.Enabled = False
        Exit Sub
    End If
    iMaxBailes = 0
    iNumBailes = 0
    rs.Close
    Set rs = db.OpenRecordset("SELECT cod_juez, cod_baile, COUNT(*) FROM MA_puntuaciones p WHERE p.num_dorsal = " & iDorsal & " AND p.cod_categoria = " & tbCodCatAct.Text & " AND fase = " & tbCodFaseAct.Text & " AND repesca = " & chkRepAct.Value & " GROUP BY cod_juez, cod_baile ORDER BY cod_juez", dbOpenSnapshot)
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
    
    Exit Sub
error:
    PPCLog ProcesarError("Timer1_Click", False)
End Sub

