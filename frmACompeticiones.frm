VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmACompeticiones 
   Caption         =   "mml_FRASE0049"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "mml_FRASE0029"
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
      Left            =   8400
      TabIndex        =   63
      Top             =   7890
      Width           =   2055
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "mml_FRASE0250"
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
      Left            =   5640
      TabIndex        =   38
      Top             =   5085
      Width           =   2835
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
      Height          =   495
      Left            =   4335
      TabIndex        =   37
      Top             =   5085
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
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
      Height          =   495
      Left            =   3015
      TabIndex        =   36
      Top             =   5085
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
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
      Height          =   495
      Left            =   1710
      TabIndex        =   35
      Top             =   5085
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   5025
      Left            =   30
      TabIndex        =   41
      Top             =   0
      Width           =   10455
      Begin VB.TextBox tbCodigoAEBDC 
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
         Left            =   6900
         MaxLength       =   7
         TabIndex        =   65
         Top             =   390
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "mml_FRASE1048"
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
         Left            =   2760
         TabIndex        =   64
         Top             =   360
         Width           =   1575
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
         Height          =   405
         Left            =   1920
         Picture         =   "frmACompeticiones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   2220
         Width           =   495
      End
      Begin VB.CheckBox chkControl 
         Caption         =   "mml_FRASE0268"
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
         Left            =   1950
         TabIndex        =   39
         Top             =   2745
         Width           =   6345
      End
      Begin VB.Frame Horario 
         Caption         =   "mml_FRASE0269"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   90
         TabIndex        =   49
         Top             =   3135
         Width           =   10275
         Begin VB.TextBox tbGeneralLooKA 
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
            Left            =   2790
            TabIndex        =   26
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbGeneralLooKB 
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
            Left            =   3540
            TabIndex        =   27
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbGeneralLooKC 
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
            Left            =   4290
            TabIndex        =   28
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbGeneralLooKD 
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
            Left            =   5040
            TabIndex        =   29
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbGeneralLooKE 
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
            Left            =   5790
            TabIndex        =   30
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbGeneralLooKF 
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
            Left            =   6540
            TabIndex        =   31
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbGeneralLooKG 
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
            Left            =   7290
            TabIndex        =   32
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbGeneralLooKH 
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
            Left            =   8040
            TabIndex        =   33
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbGeneralLooKI 
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
            Left            =   8790
            TabIndex        =   34
            Top             =   1275
            Width           =   735
         End
         Begin VB.TextBox tbFinalI 
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
            Left            =   8790
            TabIndex        =   25
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbFinalH 
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
            Left            =   8040
            TabIndex        =   24
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbFinalG 
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
            Left            =   7290
            TabIndex        =   23
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbFinalF 
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
            Left            =   6540
            TabIndex        =   22
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbFinalE 
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
            Left            =   5790
            TabIndex        =   21
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbFinalD 
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
            Left            =   5040
            TabIndex        =   20
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbFinalC 
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
            Left            =   4290
            TabIndex        =   19
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbFinalB 
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
            Left            =   3540
            TabIndex        =   18
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbFinalA 
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
            Left            =   2790
            TabIndex        =   17
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox tbElimI 
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
            Left            =   8790
            TabIndex        =   16
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox tbElimH 
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
            Left            =   8040
            TabIndex        =   15
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox tbElimG 
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
            Left            =   7290
            TabIndex        =   14
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox tbElimF 
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
            Left            =   6540
            TabIndex        =   13
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox tbElimE 
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
            Left            =   5790
            TabIndex        =   12
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox tbElimD 
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
            Left            =   5040
            TabIndex        =   11
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox tbElimC 
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
            Left            =   4290
            TabIndex        =   10
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox tbElimB 
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
            Left            =   3540
            TabIndex        =   9
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox tbElimA 
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
            Left            =   2790
            TabIndex        =   8
            Top             =   540
            Width           =   735
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "mml_FRASE0270"
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
            Left            =   330
            TabIndex        =   61
            Top             =   1305
            Width           =   2385
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "mml_FRASE0271"
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
            Left            =   330
            TabIndex        =   60
            Top             =   930
            Width           =   2385
         End
         Begin VB.Label Label18 
            Caption         =   "I"
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
            Left            =   9030
            TabIndex        =   59
            Top             =   180
            Width           =   225
         End
         Begin VB.Label Label17 
            Caption         =   "H"
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
            TabIndex        =   58
            Top             =   180
            Width           =   225
         End
         Begin VB.Label Label16 
            Caption         =   "G"
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
            Left            =   7530
            TabIndex        =   57
            Top             =   180
            Width           =   225
         End
         Begin VB.Label Label15 
            Caption         =   "F"
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
            Left            =   6780
            TabIndex        =   56
            Top             =   180
            Width           =   225
         End
         Begin VB.Label Label14 
            Caption         =   "E"
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
            Left            =   6030
            TabIndex        =   55
            Top             =   180
            Width           =   225
         End
         Begin VB.Label Label13 
            Caption         =   "D"
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
            Left            =   5310
            TabIndex        =   54
            Top             =   180
            Width           =   225
         End
         Begin VB.Label Label12 
            Caption         =   "A"
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
            Left            =   3060
            TabIndex        =   53
            Top             =   150
            Width           =   225
         End
         Begin VB.Label Label10 
            Caption         =   "C"
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
            Left            =   4560
            TabIndex        =   52
            Top             =   180
            Width           =   225
         End
         Begin VB.Label Label9 
            Caption         =   "B"
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
            Left            =   3810
            TabIndex        =   51
            Top             =   180
            Width           =   225
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "mml_FRASE0272"
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
            Left            =   330
            TabIndex        =   50
            Top             =   540
            Width           =   2385
         End
      End
      Begin VB.TextBox tbMinDorsalOficial 
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
         Left            =   8985
         MaxLength       =   4
         TabIndex        =   7
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox tbDorsalesTanda 
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
         Left            =   6105
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox tbEscuela 
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
         Left            =   1920
         MaxLength       =   150
         TabIndex        =   4
         Top             =   1800
         Width           =   7800
      End
      Begin VB.TextBox tbFechaComp 
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
         Left            =   2400
         TabIndex        =   5
         Top             =   2220
         Width           =   1455
      End
      Begin VB.TextBox tbDirComp 
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
         Left            =   1920
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1320
         Width           =   7800
      End
      Begin VB.TextBox tbDescComp 
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
         Left            =   1920
         MaxLength       =   35
         TabIndex        =   2
         Top             =   840
         Width           =   5310
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
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "mml_FRASE1227"
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
         Left            =   4680
         TabIndex        =   66
         Top             =   390
         Width           =   2115
      End
      Begin VB.Label Label7 
         Caption         =   "mml_FRASE0273"
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
         Left            =   6900
         TabIndex        =   48
         Top             =   2220
         Width           =   2085
      End
      Begin VB.Label Label6 
         Caption         =   "mml_FRASE0256"
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
         Left            =   4110
         TabIndex        =   47
         Top             =   2220
         Width           =   1920
      End
      Begin VB.Label Label5 
         Caption         =   "mml_FRASE0959"
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
         Left            =   240
         TabIndex        =   46
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "mml_FRASE0155"
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
         Left            =   240
         TabIndex        =   45
         Top             =   2220
         Width           =   1575
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
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "mml_FRASE0260"
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
         Left            =   240
         TabIndex        =   43
         Top             =   840
         Width           =   1575
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
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0049"
      Height          =   2175
      Left            =   15
      TabIndex        =   0
      Top             =   5655
      Width           =   10455
      Begin MSDataGridLib.DataGrid dgComp 
         Bindings        =   "frmACompeticiones.frx":0342
         Height          =   1845
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   3254
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
   Begin MSAdodcLib.Adodc adoComp 
      Height          =   495
      Left            =   225
      Top             =   7905
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
      RecordSource    =   "SELECT * FROM competiciones ORDER BY 1"
      Caption         =   "mml_FRASE0049"
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
End
Attribute VB_Name = "frmACompeticiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAct_Click()
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    Sleep 500
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    adoComp.Refresh
    dgComp.Refresh

    If adoComp.Recordset.EOF Then
        dgComp.Enabled = False
    Else
        dgComp.Enabled = True
    End If
End Sub

Private Sub cmdBorrar_Click()
    'On Local Error GoTo Error
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    If MsgBox(mml_FRASE1096, vbYesNo Or vbCritical, mml_FRASE0084) = vbYes Then
        If Val(VarCfg("horario_codcompeticion")) = Val(tbCodComp.Text) Then
            MsgBox mml_FRASE1097, vbOKOnly Or vbCritical, G_MSG_ERROR
            Exit Sub
        End If
        BorrarDatosComp Val(tbCodComp.Text), True
    End If
    cmdNuevo_Click
    Call cmdAct_Click
    MsgBox mml_FRASE0276, vbOKOnly Or vbExclamation, mml_FRASE0086
    Exit Sub
error:
    ProcesarError
End Sub

Private Sub cmdGrabar_Click()
Dim rs As Recordset

    If Not C_DEBUG Then On Error GoTo error
    
    If tbDirComp.Text = "" Or tbDescComp.Text = "" Or tbFechaComp.Text = "" Or tbMinDorsalOficial.Text = "" Or tbDorsalesTanda.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    ' Comprobamos que la fecha de la competición no difiere en más de un dia de la fecha de
    ' fin de licencia
    'If IsDate(frmMenu.ComprobarBase()) Then
    '    If DateDiff("d", frmMenu.ComprobarBase(), Now) > 1 Then
    '        frmMenu.Timer4.Enabled = True
    '    End If
    'Else
    '    frmMenu.Timer4.Enabled = True
    'End If
    'Control simple de fecha límite
    If C_ERROR_FECHA And Date > CDate(C_FECHA_ERROR) Then
        MsgBox "File comctl32.dll not found or it is corrupt", vbOKOnly Or vbCritical, "ERROR"
        End
    End If
    
    If tbCodComp.Text = "" Then
        tbCodComp.Text = MaxCod("competiciones")
        db.Execute ("INSERT INTO competiciones VALUES(" & MaxCod("competiciones") & ", '" & tbDescComp.Text & "','" & tbDirComp.Text & "','" & tbFechaComp.Text & "','" & tbEscuela.Text & "'," & Val(tbMinDorsalOficial.Text) & "," & Val(tbDorsalesTanda.Text) & "," & Val(tbFinalA.Text) & "," & Val(tbFinalB.Text) & "," & Val(tbFinalC.Text) & "," & Val(tbFinalD.Text) & "," & Val(tbFinalE.Text) & "," & Val(tbFinalF.Text) & "," & Val(tbFinalG.Text) & "," & Val(tbFinalH.Text) & "," & Val(tbFinalI.Text) & "," & _
                    Val(tbElimA.Text) & "," & Val(tbElimB.Text) & "," & Val(tbElimC.Text) & "," & Val(tbElimD.Text) & "," & Val(tbElimE.Text) & "," & Val(tbElimF.Text) & "," & Val(tbElimG.Text) & "," & Val(tbElimH.Text) & "," & Val(tbElimI.Text) & "," & Val(tbGeneralLooKA.Text) & "," & Val(tbGeneralLooKB.Text) & "," & Val(tbGeneralLooKC.Text) & "," & Val(tbGeneralLooKD.Text) & "," & Val(tbGeneralLooKE.Text) & "," & Val(tbGeneralLooKF.Text) & "," & Val(tbGeneralLooKG.Text) & "," _
                    & Val(tbGeneralLooKH.Text) & "," & Val(tbGeneralLooKI.Text) & "," & chkControl.Value & ",'" & tbCodigoAEBDC.Text & "')")
    Else
        db.Execute ("UPDATE competiciones SET descripcion = '" & tbDescComp.Text & _
                    "',direccion ='" & tbDirComp.Text & _
                    "',fecha='" & tbFechaComp.Text & _
                    "',escuela='" & tbEscuela.Text & _
                    "',min_dorsal_oficial=" & Val(tbMinDorsalOficial.Text) & _
                    ",dorsales_tanda=" & Val(tbDorsalesTanda.Text) & _
                    ",FinalA=" & Val(tbFinalA.Text) & ",FinalB=" & Val(tbFinalB.Text) & ",FinalC=" & Val(tbFinalC.Text) & ",FinalD=" & Val(tbFinalD.Text) & ",FinalE=" & Val(tbFinalE.Text) & ",FinalF=" & Val(tbFinalF.Text) & ",FinalG=" & Val(tbFinalG.Text) & ",FinalH=" & Val(tbFinalH.Text) & ",FinalI=" & Val(tbFinalI.Text) & _
                    ",ElimA=" & Val(tbElimA.Text) & ",ElimB=" & Val(tbElimB.Text) & ",ElimC=" & Val(tbElimC.Text) & ",ElimD=" & Val(tbElimD.Text) & ",ElimE=" & Val(tbElimE.Text) & ",ElimF=" & Val(tbElimF.Text) & ",ElimG=" & Val(tbElimG.Text) & ",ElimH=" & Val(tbElimH.Text) & ",ElimI=" & Val(tbElimI.Text) & _
                    ",GeneralLooKA=" & Val(tbGeneralLooKA.Text) & ",GeneralLooKB=" & Val(tbGeneralLooKB.Text) & ",GeneralLooKC=" & Val(tbGeneralLooKC.Text) & ",GeneralLooKD=" & Val(tbGeneralLooKD.Text) & ",GeneralLooKE=" & Val(tbGeneralLooKE.Text) & ",GeneralLooKF=" & Val(tbGeneralLooKF.Text) & ",GeneralLooKG=" & Val(tbGeneralLooKG.Text) & ",GeneralLooKH=" & Val(tbGeneralLooKH.Text) & ",GeneralLooKI=" & Val(tbGeneralLooKI.Text) & _
                    ",control=" & chkControl.Value & _
                    ",aebdc_codigo='" & tbCodigoAEBDC.Text & "'" & _
                    " WHERE codigo = " & tbCodComp.Text)
    End If
    MsgBox mml_FRASE0278, vbOKOnly Or vbInformation, mml_FRASE0086
    Call cmdAct_Click
    
error:
    ProcesarError
End Sub

Private Sub cmdNuevo_Click()
    tbCodComp.Text = ""
    tbDescComp.Text = ""
    tbDirComp.Text = ""
    tbFechaComp.Text = ""
    tbEscuela.Text = ""
    tbDorsalesTanda.Text = ""
    tbMinDorsalOficial.Text = ""
    
    tbElimA.Text = ""
    tbFinalA.Text = ""
    tbGeneralLooKA.Text = ""
    tbElimB.Text = ""
    tbFinalB.Text = ""
    tbGeneralLooKB.Text = ""
    tbElimC.Text = ""
    tbFinalC.Text = ""
    tbGeneralLooKC.Text = ""
    tbElimD.Text = ""
    tbFinalD.Text = ""
    tbGeneralLooKD.Text = ""
    tbElimE.Text = ""
    tbFinalE.Text = ""
    tbGeneralLooKE.Text = ""
    tbElimF.Text = ""
    tbFinalF.Text = ""
    tbGeneralLooKF.Text = ""
    tbElimG.Text = ""
    tbFinalG.Text = ""
    tbGeneralLooKG.Text = ""
    tbElimH.Text = ""
    tbFinalH.Text = ""
    tbGeneralLooKH.Text = ""
    tbElimI.Text = ""
    tbFinalI.Text = ""
    tbGeneralLooKI.Text = ""
    chkControl.Value = 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelComp_Click()
    tbFechaComp.Text = frmCalendario.Mostrar
End Sub

Private Sub Command1_Click()
    tbCodComp.Text = ""
    tbDescComp.Text = ""
    tbDirComp.Text = ""
    tbFechaComp.Text = ""
End Sub

Private Sub dgComp_Click()
    dgComp.Col = 0
    tbCodComp.Text = dgComp.Text
    dgComp.Col = 1
    tbDescComp.Text = dgComp.Text
    dgComp.Col = 2
    tbDirComp.Text = dgComp.Text
    dgComp.Col = 3
    tbFechaComp.Text = dgComp.Text
    dgComp.Col = 4
    tbEscuela.Text = dgComp.Text
    dgComp.Col = 5
    tbMinDorsalOficial.Text = dgComp.Text
    dgComp.Col = 6
    tbDorsalesTanda.Text = dgComp.Text
    dgComp.Col = 7
    tbFinalA.Text = dgComp.Text
    dgComp.Col = 8
    tbFinalB.Text = dgComp.Text
    dgComp.Col = 9
    tbFinalC.Text = dgComp.Text
    dgComp.Col = 10
    tbFinalD.Text = dgComp.Text
    dgComp.Col = 11
    tbFinalE.Text = dgComp.Text
    dgComp.Col = 12
    tbFinalF.Text = dgComp.Text
    dgComp.Col = 13
    tbFinalG.Text = dgComp.Text
    dgComp.Col = 14
    tbFinalH.Text = dgComp.Text
    dgComp.Col = 15
    tbFinalI.Text = dgComp.Text
    dgComp.Col = 16
    tbElimA.Text = dgComp.Text
    dgComp.Col = 17
    tbElimB.Text = dgComp.Text
    dgComp.Col = 18
    tbElimC.Text = dgComp.Text
    dgComp.Col = 19
    tbElimD.Text = dgComp.Text
    dgComp.Col = 20
    tbElimE.Text = dgComp.Text
    dgComp.Col = 21
    tbElimF.Text = dgComp.Text
    dgComp.Col = 22
    tbElimG.Text = dgComp.Text
    dgComp.Col = 23
    tbElimH.Text = dgComp.Text
    dgComp.Col = 24
    tbElimI.Text = dgComp.Text
    dgComp.Col = 25
    tbGeneralLooKA.Text = dgComp.Text
    dgComp.Col = 26
    tbGeneralLooKB.Text = dgComp.Text
    dgComp.Col = 27
    tbGeneralLooKC.Text = dgComp.Text
    dgComp.Col = 28
    tbGeneralLooKD.Text = dgComp.Text
    dgComp.Col = 29
    tbGeneralLooKE.Text = dgComp.Text
    dgComp.Col = 30
    tbGeneralLooKF.Text = dgComp.Text
    dgComp.Col = 31
    tbGeneralLooKG.Text = dgComp.Text
    dgComp.Col = 32
    tbGeneralLooKH.Text = dgComp.Text
    dgComp.Col = 33
    tbGeneralLooKI.Text = dgComp.Text
    dgComp.Col = 34
    chkControl.Value = Val(dgComp.Text)
    dgComp.Col = 35
    tbCodigoAEBDC.Text = dgComp.Text
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    If VarCfg("tipo_hoja_puntuaciones") = "hora_rec_por_baile" Then
        chkControl.Value = 0
        chkControl.Visible = False
    End If
End Sub


Private Sub tbCodigoAEBDC_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub
