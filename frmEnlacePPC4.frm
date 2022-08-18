VERSION 5.00
Begin VB.Form frmEnlacePPC4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "(4) mml_FRASE0577"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0577"
      Height          =   6105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10635
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
         Height          =   480
         Left            =   8460
         TabIndex        =   57
         Top             =   5535
         Width           =   1875
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
         Height          =   480
         Left            =   5400
         TabIndex        =   56
         Top             =   5520
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
         Height          =   480
         Left            =   2385
         TabIndex        =   55
         Top             =   5535
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
         Height          =   480
         Left            =   405
         TabIndex        =   54
         Top             =   5535
         Width           =   1875
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
         Top             =   1305
         Width           =   1650
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
         Picture         =   "frmEnlacePPC4.frx":0000
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
         Picture         =   "frmEnlacePPC4.frx":046A
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
         Picture         =   "frmEnlacePPC4.frx":08D4
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
         Left            =   630
         Top             =   2160
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
         Top             =   2040
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
         Height          =   405
         Left            =   9150
         Picture         =   "frmEnlacePPC4.frx":0D3E
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "mml_FRASE0425"
         Top             =   840
         Width           =   600
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
         ItemData        =   "frmEnlacePPC4.frx":1020
         Left            =   765
         List            =   "frmEnlacePPC4.frx":1042
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
         Height          =   375
         Left            =   4950
         TabIndex        =   36
         Top             =   2010
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
         Top             =   2160
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
         Height          =   1170
         Left            =   435
         TabIndex        =   20
         Top             =   4245
         Width           =   9915
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
            Height          =   480
            Left            =   7920
            TabIndex        =   53
            Top             =   135
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
            Picture         =   "frmEnlacePPC4.frx":1075
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
            Left            =   5925
            TabIndex        =   33
            Top             =   630
            Width           =   3885
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
            Left            =   4905
            TabIndex        =   28
            Top             =   675
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
            Width           =   3150
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
            Left            =   915
            TabIndex        =   23
            Top             =   270
            Width           =   4290
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
            Left            =   75
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
            Left            =   4095
            TabIndex        =   29
            Top             =   765
            Width           =   765
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "mml_FRASE0582"
         Height          =   1590
         Left            =   420
         TabIndex        =   14
         Top             =   2610
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
            Top             =   135
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
            Picture         =   "frmEnlacePPC4.frx":1357
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
            Picture         =   "frmEnlacePPC4.frx":17C1
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
               Size            =   12
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
            Top             =   660
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
            Top             =   660
            Width           =   2640
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
            Left            =   45
            TabIndex        =   16
            Top             =   1050
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
            Left            =   885
            TabIndex        =   15
            Top             =   1050
            Width           =   2640
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
            Height          =   825
            Left            =   3540
            TabIndex        =   34
            Top             =   660
            Width           =   6285
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
         Height          =   495
         Left            =   9900
         Picture         =   "frmEnlacePPC4.frx":1AA3
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "mml_FRASE0043"
         Top             =   765
         Width           =   615
      End
      Begin VB.CommandButton cmdDorsales 
         Height          =   495
         Left            =   9900
         Picture         =   "frmEnlacePPC4.frx":2685
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "mml_FRASE0028"
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdCategAct 
         Height          =   495
         Left            =   9135
         Picture         =   "frmEnlacePPC4.frx":31A7
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "mml_FRASE0428"
         Top             =   300
         Width           =   615
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
         Height          =   375
         Left            =   8820
         TabIndex        =   32
         Tag             =   "0"
         Top             =   1665
         Width           =   1680
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
Attribute VB_Name = "frmEnlacePPC4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim g_sNoPresentes As String

Const C_NP_NO_INICIADO = 0
Const C_NP_INICIADO = 1
Const C_NP_PREPARADO = 2

Dim aNoPresentes() As String

Dim g_iNoPresentesIniciado As Integer
Dim g_iCJueces As Integer
Dim g_bNoPresentesProcesados As Boolean
Dim g_sNumNoPresentes As Integer

Dim iCodCat As Integer
Dim iFase As Integer
Dim iRepesca As Integer




Private Sub cmdBailes_Click()
Dim rs As Recordset, sMsj As String
    Set rs = db.OpenRecordset("SELECT * FROM bailes")
    While Not rs.EOF
        sMsj = sMsj & rs!codigo & " - " & rs!Nombre & Chr$(13) & Chr$(10)
        rs.MoveNext
    Wend
    rs.Close
    
    MsgBox sMsj, vbOKOnly Or vbInformation, mml_FRASE0185
End Sub

Private Sub cmdBorrarFicheros_Click()
    If MsgBox(mml_FRASE0585, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        On Local Error Resume Next
        Kill G_FICHERO_SALIDA & ".?.TXT"
        If MsgBox(mml_FRASE0586, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
            Kill G_FICHERO_ENTRADA & "*"
        End If
        MsgBox mml_FRASE0587, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
    
End Sub

Private Sub cmdCalcular_Click()
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        frmCalcular.tbCodComp.Text = tbCodComp.Text
        frmCalcular.tbDescComp.Text = tbDescComp.Text
        frmCalcular.tbCodCat.Text = tbCodCatAct.Text
        frmCalcular.tbDescCat.Text = tbDescCatAct.Text
        frmCalcular.tbCodFase.Text = Val(tbCodFaseAct.Text)
        frmCalcular.tbDescFase.Text = sDescFase(Val(tbCodFaseAct.Text))
        frmCalcular.chkRep.Value = chkRepAct.Value
    
        frmCalcular.Show vbModal
    End If

End Sub

Private Sub cmdCategAct_Click()
Dim rs As Recordset, sMsj As String
    Set rs = db.OpenRecordset("SELECT TOP 1 * FROM horario h WHERE grupo LIKE '*" & cbPista.Text & "*' AND numfase <> 99 AND cod_competicion = " & Val(VarCfg("horario_codcompeticion")) & " AND (SELECT COUNT(*) FROM puntuaciones WHERE ((h.cod_baile < 0 AND cod_baile " & G_ORDEN_10B_LAT_EST & " -h.cod_baile) OR h.cod_baile = 0 OR cod_baile = h.cod_baile) AND cod_categoria = h.cod_categoria AND fase = h.numfase AND repesca = h.repesca) = 0 ORDER BY orden", dbOpenSnapshot)
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

    RecuperarJueces Val(tbCodCat.Text), cbJuez
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
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        frmDescalificados.tbCodComp.Text = tbCodComp.Text
        frmDescalificados.tbDescComp.Text = tbDescComp.Text
        frmDescalificados.tbCodCateg.Text = tbCodCat.Text
        frmDescalificados.tbDescCateg.Text = tbDescCat.Text
        frmDescalificados.cbFase.ListIndex = Log(Val(tbCodFase.Text)) / Log(2)
        frmDescalificados.cmdActualizar_Click

        frmDescalificados.Show vbNomodal
    End If

End Sub

Private Sub cmdDorsales_Click()
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        frmADorsales.tbCodComp.Text = tbCodComp.Text
        frmADorsales.tbDescComp.Text = tbDescComp.Text
        frmADorsales.tbCodCateg.Text = tbCodCat.Text
        frmADorsales.tbDescCateg.Text = tbDescCat.Text
        frmADorsales.cbFase.ListIndex = Log(Val(tbCodFase.Text)) / Log(2) + 1
        
        frmADorsales.Show vbNomodal
        
    End If

End Sub

Private Sub cmdGenDatos_Click()
Dim i As Integer
    
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If MsgBox(mml_FRASE0588, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
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
    lblJuecesAct.Caption = ""
    tbCodCatAct.Text = tbCodCat.Text
    tbDescCatAct.Text = tbDescCat.Text
    tbCodFaseAct.Text = tbCodFase.Text
    tbDescFaseAct.Text = tbDescFase.Text
    chkRepAct.Value = chkRep.Value
    
    tbNumJueces.Text = RecuperarJueces(Val(tbCodCat.Text), cbJuezAct)
    BorrarDatosFaseSig
End Sub

Sub GenerarFicheroJuezPasosGenerico()
Dim rs As Recordset, iFile As Long
    iFile = FreeFile
    Open G_FICHERO_SALIDA & ".JpGen.TXT" For Output As #iFile
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
End Sub

Private Sub cmdGenDatosAct_Click()
Dim i As Integer
    If tbCodCatAct.Text = "" Or tbCodFaseAct.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If MsgBox(mml_FRASE0588, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    If cbJuezAct.Text = "" Then
        For i = 0 To cbJuezAct.ListCount - 1
            GenerarFichero cbJuezAct.List(i), Val(tbCodCatAct.Text), tbDescCatAct.Text, Val(tbCodFaseAct.Text), chkRepAct.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
        Next
    Else
        GenerarFichero cbJuezAct.Text, Val(tbCodCatAct.Text), tbDescCatAct.Text, Val(tbCodFaseAct.Text), chkRepAct.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
    End If

End Sub

Private Sub cmdGenDatosSig_Click()
Dim i As Integer
    If tbCodCatSig.Text = "" Or tbCodFaseSig.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If MsgBox(mml_FRASE0588, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    If cbJuezSig.Text = "" Then
        For i = 0 To cbJuezSig.ListCount - 1
            GenerarFichero cbJuezSig.List(i), Val(tbCodCatSig.Text), tbDescCatSig.Text, Val(tbCodFaseSig.Text), chkRepSig.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
        Next
    Else
        GenerarFichero cbJuezSig.List(i), Val(tbCodCatSig.Text), tbDescCatSig.Text, Val(tbCodFaseSig.Text), chkRepSig.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCat_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCat.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCat.Text = sResultado(2)
    tbCodFase.Text = ""
    tbDescFase.Text = ""
    
    RecuperarJueces Val(tbCodCat.Text), cbJuez
End Sub
Function RecuperarJueces(iCodCateg As Integer, cbJuez As ComboBox) As Integer
Dim rs As Recordset, i As Integer
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
End Function

Private Sub cmdSelComp_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCat.Text = ""
    tbDescCat.Text = ""
    tbCodFase.Text = ""
    tbDescFase.Text = ""

End Sub

Private Sub cmdSelFase_Click()
    If tbCodCat.Text = "" Then
        MsgBox mml_FRASE0504, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1")
    DescFase
    DoEvents
    CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes

End Sub

Sub DescFase()
    Select Case tbCodFase.Text
        Case 1:
            tbDescFase.Text = mml_FRASE0329
        Case 2:
            tbDescFase.Text = "SEMI-FINAL"
        Case 4:
            tbDescFase.Text = "CUARTOS DE FINAL"
        Case 8:
            tbDescFase.Text = "OCTAVOS DE FINAL"
        Case "":
            tbDescFase.Text = ""
        Case Else
            tbDescFase.Text = tbCodFase.Text & "OS DE FINAL"
    End Select
End Sub

Sub GenerarFichero(sJuez As String, iCodCat As Integer, sDescCat As String, iCodFase As Integer, iCodRep As Integer, Optional iBaileIni As Integer = 0, Optional sBaile As String = "0")
Dim rs As Recordset
Dim iFile As Long
Dim bError As Boolean
Dim iCBailes As Integer
Dim bTeamMatch As Boolean
Dim iBaile As Integer

    iBaile = Val(sBaile)
    bTeamMatch = ComprobarSiTeamMatch(iCodCat)
    bError = False
    iFile = FreeFile
    Open G_FICHERO_SALIDA & "." & sJuez & ".cTXT" For Output As #iFile
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
    
    If G_DORSALES_COMBINADOS And CombinarDorsalesCateg(iCodCat) And G_PPC_GEN_DORSALES_COMBINADOS Then
    ' Si se generan dorsales combinados solo se envia un baile
        'Genera la información para la recombinación de dorsales
        
        Dim iMaxTandas As Integer, iTandasConMasDorsales As Integer, iTotalDorsales As Integer
        
        iMaxTandas = 0
        CalcularDorsalesPorTandaCatExt iCodCat, iCodFase, iCodRep, 1, iMaxTandas, iTandasConMasDorsales, iTotalDorsales
        CombinarDorsales iCodCat, iCodFase, iCodRep, iMaxTandas, 1, True
        
        Print #iFile, "DORSALES_COMBINADOS"
        If iBaile = 0 Then ' Solo podemos generar un baile
            If cbBailes.ListCount > 1 Then
                iBaile = Val(cbBailes.List(1))
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
                    If iBaile = rs!cod_baile Then
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
    Else 'Imprimimos todos los dorsales ordenador por dorsal
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
    Close #iFile
    On Local Error Resume Next
    If bError Then
        Kill G_FICHERO_SALIDA & "." & sJuez & ".cTXT"
    Else
        Kill G_FICHERO_SALIDA & "." & sJuez & ".TXT"
        Name G_FICHERO_SALIDA & "." & sJuez & ".cTXT" As G_FICHERO_SALIDA & "." & sJuez & ".TXT"
    End If
End Sub

Private Sub cmdSubirDatos_Click()
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbCodCat.Text = tbCodCatAct.Text
    tbCodFase.Text = tbCodFaseAct.Text
    chkRep.Value = chkRepAct.Value
    tbDescCat.Text = tbDescCatAct.Text
    tbDescFase.Text = tbDescFaseAct.Text
    
    RecuperarJueces Val(tbCodCat.Text), cbJuez
    'CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    IniciarDatosNoPresentes
    CargarPistas cbPista
    CargarBailes Val(tbCodCat.Text), Val(tbCodFase.Text), cbBailes
End Sub
Sub CargarBailes(iCodCat As Integer, iCodFase As Integer, cbBailes As ComboBox)
Dim rs As Recordset
    
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
End Sub

Private Sub lblActTimer_Click()
    If lblActTimer.Tag = "0" Then
        lblActTimer.BackColor = vbGreen
        lblActTimer.Caption = mml_FRASE0592
        lblActTimer.Tag = "1"
        Timer1.Interval = G_INTERVALO_TIMER_PPC
        Timer1.Enabled = True
        tmrCalcular.Interval = G_TIEMPO_ESPERA_JPASOS_PPC
        tmrCalcular.Enabled = True
    Else
        lblActTimer.BackColor = vbRed
        lblActTimer.Caption = mml_FRASE0583
        lblActTimer.Tag = "0"
        Timer1.Enabled = False
        tmrCalcular.Enabled = False
    End If
End Sub

Private Sub lblJuecesAct_DblClick()
    lblJuecesAct.Caption = ""
End Sub

Private Sub tbCodCatAct_Change()
    If lblActTimer.Caption = mml_FRASE0592 Then tmrCalcular.Enabled = True

End Sub

Private Sub Timer1_Timer()
Dim sFichero As String, sPath As String, rs As Recordset, sDir As String
Dim iCodCat As Integer, sDirFichas As String

    If Not C_DEBUG Then On Local Error GoTo error
    Timer1.Enabled = False
    lblActTimer.Enabled = False
    sDirFichas = VarCfg("dir_fichas")
    'Localizamos los ficheros resultado de los PPC
    sFichero = Dir$(G_FICHERO_ENTRADA & "*.TXT")
    While sFichero <> ""
        'Procesamos el fichero
        iCodCat = RecuperarInfo(sExtraerPath(G_FICHERO_ENTRADA) & "\" & sFichero)
        'Ahora debemos copiarlo a la carpeta correspondiente
        If iCodCat > 0 Then
            Set rs = db.OpenRecordset("SELECT codigo, descripcion FROM categorias WHERE codigo = " & iCodCat, dbOpenSnapshot)
            If Not rs.EOF Then
                On Local Error Resume Next
                sDir = sDirFichas & "\TMP\" & rs!DESCRIPCION & "_" & rs!codigo
                MkDir sDir
                If Not C_DEBUG Then
                    On Local Error GoTo error
                Else
                    On Local Error GoTo 0
                End If
            End If
            rs.Close
            FileCopy sExtraerPath(G_FICHERO_ENTRADA) & "\" & sFichero, sDir & "\" & sFichero
            Kill sExtraerPath(G_FICHERO_ENTRADA) & "\" & sFichero
        End If
        sFichero = Dir
    Wend
    Timer1.Enabled = True
    lblActTimer.Enabled = True
error:
    ProcesarError
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

    RecuperarInfo = 0
    iFile = FreeFile
    
    'Comprobamos si todavía se está grabando el fichero
    Sleep 100
    lTam = 0
    While lTam <> FileLen(sFichero)
        lTam = FileLen(sFichero)
        DoEvents
        Sleep 200
        DoEvents
    Wend
    
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
    
    bTeamMatch = ComprobarSiTeamMatch(Val(sCodCat))
    
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
            'Es el ultimo baile
            sEstado = "FIN"
        Else
            GenerarFichero sJuez, Val(sCodCat), sDescCategoria(Val(sCodCat)), Val(sCodFase), chkRep.Value, chkUltimos5Bailes.Value, rsBailes!cod_baile
            'Informamos de la generación del siguiente baile
            lblJuecesAct.Caption = lblJuecesAct.Caption & Left$(rsBailes!Nombre, 1) & ">"
        End If
    End If
    
    'Comprobamos si es el ultimo baile del juez
    If sEstado = mml_FRASE0594 Then ' FIN
        
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
                
                lblJuecesSig.Text = ""
            
                CargarBailes Val(tbCodCatSig.Text), Val(tbCodFaseSig.Text), cbBailes
                If rs!cod_baile > 0 Then
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
    If sEstado <> "JPASOSGEN" And (Val(tbCodCatSig.Text) > 0 And Val(tbNumJueces.Text) = 0) Then
    'Antes de avanzar de categoría comprobamos si hay que realizar un cálculo automático
        If frmMenu.mnuGenAutoPPC.Checked Then ComprobarPuntuaciones
        
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
        For i = 0 To cbJuezAct.ListCount - 1
            GenerarFichero cbJuezAct.List(i), Val(tbCodCatAct.Text), tbDescCatAct.Text, Val(tbCodFaseAct.Text), chkRepAct.Value, chkUltimos5Bailes.Value, cbBailes.List(cbBailes.ListIndex)
        Next
        
        BorrarDatosFaseSig
        IniciarDatosNoPresentes
        
        tmrCalcular.Interval = G_TIEMPO_ESPERA_JPASOS_PPC
        tmrCalcular.Enabled = True
        
    End If
    
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
            ElseIf MsgBox(mml_FRASE0445, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                frmCalcular.tbCodComp.Text = tbCodComp.Text
                frmCalcular.tbDescComp.Text = tbDescComp.Text
                frmCalcular.tbCodCat.Text = Val(tbCodCatAct.Text)
                frmCalcular.tbDescCat.Text = sDescCategoria(Val(tbCodCatAct.Text))
                frmCalcular.tbCodFase.Text = Val(tbCodFaseAct.Text)
                frmCalcular.tbDescFase.Text = sDescFase(Val(tbCodFaseAct.Text))
                frmCalcular.chkRep.Value = chkRepAct.Value
                
                frmCalcular.Show vbModal
            End If
        End If
    End If
    Exit Sub

error:
    ProcesarError "tmrCalcular_Timer"
End Sub
