VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAPuntuaciones 
   Caption         =   "mml_FRASE0057"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   585
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
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   90
      Width           =   495
   End
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
      Height          =   480
      Left            =   6795
      TabIndex        =   41
      Top             =   1080
      Width           =   1470
   End
   Begin VB.CheckBox chkRep 
      Caption         =   "mml_FRASE0418"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   34
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "mml_FRASE0419"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   8295
      Begin VB.CommandButton cmdBorrarTodas 
         Caption         =   "mml_FRASE0383"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6960
         TabIndex        =   40
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   5
         Left            =   2160
         TabIndex        =   19
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   6
         Left            =   3120
         TabIndex        =   20
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   7
         Left            =   4080
         TabIndex        =   21
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   8
         Left            =   5040
         TabIndex        =   22
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   9
         Left            =   6000
         TabIndex        =   23
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "mml_FRASE0420"
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
         Left            =   6960
         TabIndex        =   25
         Top             =   240
         Width           =   1215
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
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Left            =   6960
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   6000
         TabIndex        =   18
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   3
         Left            =   5040
         TabIndex        =   17
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   2
         Left            =   4080
         TabIndex        =   16
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   1
         Left            =   3120
         TabIndex        =   15
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbBaile 
         Alignment       =   2  'Center
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
         Index           =   0
         Left            =   2160
         TabIndex        =   14
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbDorsal 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   2160
         TabIndex        =   39
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   3120
         TabIndex        =   38
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   4080
         TabIndex        =   37
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   5040
         TabIndex        =   36
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   6000
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0421"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   6000
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   5040
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   4080
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3120
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2160
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "mml_FRASE0300"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.ComboBox cbFase 
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
      ItemData        =   "frmMAPuntuaciones.frx":0000
      Left            =   4440
      List            =   "frmMAPuntuaciones.frx":0002
      TabIndex        =   11
      Text            =   "2 ,SEMIFINAL"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "mml_FRASE0295"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox tbDescCateg 
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
      Left            =   3360
      TabIndex        =   7
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox tbCodCateg 
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
      Left            =   2520
      TabIndex        =   6
      Top             =   600
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
      Left            =   2520
      TabIndex        =   4
      Top             =   120
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
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   4935
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
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0422"
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgPuntuaciones 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
   Begin MSAdodcLib.Adodc adoPuntuaciones 
      Height          =   495
      Left            =   2040
      Top             =   6720
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
      RecordSource    =   "SELECT * FROM puntuaciones"
      Caption         =   "mml_FRASE0422"
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
      Left            =   3720
      TabIndex        =   10
      Top             =   1200
      Width           =   735
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
      Left            =   360
      TabIndex        =   8
      Top             =   600
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
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmAPuntuaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBorrar_Click()
End Sub

Private Sub cbFase_Click()
    Call ActualizarTodo

End Sub

Private Sub cbFase_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbJuez_Click()
Dim i As Integer
    tbDorsal.Text = ""
    For i = 0 To 9
        tbBaile(i).Text = "0"
    Next
End Sub

Private Sub cbJuez_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub ActualizarTodo()
Dim rs As Recordset
Dim sFase As String
Dim i As Integer
    If tbCodCateg.Text = "" Or tbCodComp.Text = "" Or Trim$(Mid$(cbFase.Text, 1, 2)) = "" Then
        Exit Sub
    End If
    ' Recuperamos los jueces y los bailes
    cbJuez.Clear
    Set rs = db.OpenRecordset("SELECT id_juez FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rs.EOF
            cbJuez.AddItem rs!id_juez
            rs.MoveNext
        Wend
    rs.Close
    ' Final = 1 o No final = 2
    sFase = IIf(Val(Mid$(cbFase.Text, 1, 2)) = 1, "1", "2")
    Set rs = db.OpenRecordset("SELECT cod_baile, nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE cod_categoria = " & tbCodCateg.Text & " AND fase =" & sFase & " AND bc.cod_baile = b.codigo ORDER BY posicion", dbOpenSnapshot)
        For i = 0 To 9
            lblBaile(i).Visible = False
            tbBaile(i).Visible = False
        Next
        i = 0
        While Not rs.EOF And i <= 9
            lblBaile(i).Tag = rs!cod_baile
            lblBaile(i).Caption = "Cod. " & rs!cod_baile & Chr$(13) & Chr$(10) & rs!Nombre
            lblBaile(i).Visible = True
            tbBaile(i).Visible = True
            i = i + 1
            rs.MoveNext
        Wend
    rs.Close
    Call cmdActualizar_Click
    cbJuez.Refresh
End Sub


Private Sub cmdActualizar_Click()

    If tbCodCateg.Text <> "" Then
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        Sleep 100
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        adoPuntuaciones.ConnectionString = "DSN=Escrutinio"
        Debug.Print "SELECT * FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase LIKE '" & Trim$(Mid$(cbFase.Text, 1, 2)) & "%' AND cod_juez Like '" & cbJuez.Text & "%' AND num_dorsal Like '" & tbDorsal.Text & "%'"
        adoPuntuaciones.RecordSource = "SELECT * FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase LIKE '" & Trim$(Mid$(cbFase.Text, 1, 2)) & "%' AND cod_juez Like '" & cbJuez.Text & "%' AND num_dorsal Like '" & tbDorsal.Text & "%' ORDER BY 2,6,1,3,4"
        adoPuntuaciones.Refresh
    End If

End Sub



Private Sub cmdBorrarTodas_Click()
    If tbCodCateg.Text = "" Or tbCodComp.Text = "" Or Trim$(Mid$(cbFase.Text, 1, 2)) = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If MsgBox(mml_FRASE0423, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        Debug.Print "DELETE FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = '" & Trim$(Mid$(cbFase.Text, 1, 2)) & "'"
        db.Execute "DELETE FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Trim$(Mid$(cbFase.Text, 1, 2))
    End If
End Sub

Private Sub cmdCateg_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCateg.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCateg.Text = sResultado(2)
    Call ActualizarTodo

End Sub





Private Sub cmdQuitar_Click()
Dim sNumDorsal As String
Dim sCodCategoria As String
Dim sCodCategorias As String
Dim sCodBaile As String
Dim sCodJuez As String
Dim sFase As String
Dim sRep As String

    If tbCodCateg.Text = "" Or cbFase.Text = "" Then
        MsgBox mml_FRASE0424, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    dgPuntuaciones.Col = 0
    sNumDorsal = dgPuntuaciones.Text
    dgPuntuaciones.Col = 1
    sCodCategoria = dgPuntuaciones.Text
    dgPuntuaciones.Col = 2
    sCodBaile = dgPuntuaciones.Text
    dgPuntuaciones.Col = 3
    sCodJuez = dgPuntuaciones.Text
    dgPuntuaciones.Col = 5
    sFase = dgPuntuaciones.Text
    dgPuntuaciones.Col = 6
    sRep = dgPuntuaciones.Text
    
    Debug.Print "DELETE FROM puntuaciones WHERE num_dorsal = " & sNumDorsal & " AND cod_categoria = " & sCodCategoria & " AND cod_baile = " & sCodBaile & " AND cod_juez = '" & sCodJuez & "' AND fase = " & sFase & " AND repesca=" & sRep
    db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal = " & sNumDorsal & " AND cod_categoria = " & sCodCategoria & " AND cod_baile = " & sCodBaile & " AND cod_juez = '" & sCodJuez & "' AND fase = " & sFase & " AND repesca=" & sRep)
    Call cmdActualizar_Click
End Sub

Private Sub CommandButton1_Click()

End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    Call cmdActualizar_Click
End Sub

Private Sub dgJueces_Click()

End Sub

Private Sub Label7_Click()
End Sub

Private Sub addPuntuaciones(bModificar As Boolean)
Dim i As Integer
    If cbJuez.Text = "" Or tbDorsal.Text = "" Or tbCodCateg.Text = "" Or Trim$(Mid$(cbFase.Text, 1, 2)) = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    i = 0
    For i = 0 To 9
        If lblBaile(i).Visible = True Then
            If tbBaile(i).Text <> "" Then
                Debug.Print "INSERT INTO puntuaciones VALUES (" & tbDorsal.Text & "," & tbCodCateg.Text & "," & lblBaile(i).Tag & "," & cbJuez.Text & "," & tbBaile(i).Text & "," & Val(Trim$(Mid$(cbFase.Text, 1, 2))) & "," & chkRep.Value & ")"
                ' Borramos cualquier puntuación que pudiera tener
                If bModificar Then
                    Debug.Print "DELETE FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal=" & tbDorsal.Text & " AND cod_juez = '" & cbJuez.Text & "' AND repesca=" & chkRep.Value & " AND fase=" & Val(Trim$(Mid$(cbFase.Text, 1, 2))) & " AND cod_baile=" & lblBaile(i).Tag
                    db.Execute ("DELETE FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal=" & tbDorsal.Text & " AND cod_juez = '" & cbJuez.Text & "' AND repesca=" & chkRep.Value & " AND fase=" & Val(Trim$(Mid$(cbFase.Text, 1, 2))) & " AND cod_baile=" & lblBaile(i).Tag)
                End If
                db.Execute ("INSERT INTO puntuaciones VALUES (" & tbDorsal.Text & "," & tbCodCateg.Text & "," & lblBaile(i).Tag & ",'" & cbJuez.Text & "'," & tbBaile(i).Text & "," & Val(Trim$(Mid$(cbFase.Text, 1, 2))) & "," & chkRep.Value & ")")
            End If
            tbBaile(i).Text = "0"
        Else
            Exit For
        End If
    Next
    Call cmdActualizar_Click
    tbDorsal.SetFocus
End Sub



Private Sub Command1_Click()
    addPuntuaciones False
    tbDorsal_DblClick
    tbBaile(0).SetFocus
End Sub

Private Sub Command2_Click()
    addPuntuaciones True
End Sub

Private Sub dgPuntuaciones_AfterColUpdate(ByVal ColIndex As Integer)
Dim sNumDorsal As String
Dim sCodCategoria As String
Dim sCodCategorias As String
Dim sCodBaile As String
Dim sCodJuez As String
Dim sFase As String
Dim sRep As String
Dim sPuesto As String

    If tbCodCateg.Text = "" Or cbFase.Text = "" Then
        MsgBox mml_FRASE0424, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    dgPuntuaciones.Col = 0
    sNumDorsal = dgPuntuaciones.Text
    dgPuntuaciones.Col = 1
    sCodCategoria = dgPuntuaciones.Text
    dgPuntuaciones.Col = 2
    sCodBaile = dgPuntuaciones.Text
    dgPuntuaciones.Col = 3
    sCodJuez = dgPuntuaciones.Text
    dgPuntuaciones.Col = 4
    sPuesto = dgPuntuaciones.Text
    dgPuntuaciones.Col = 5
    sFase = dgPuntuaciones.Text
    dgPuntuaciones.Col = 6
    sRep = dgPuntuaciones.Text
    
    Debug.Print "DELETE FROM puntuaciones WHERE num_dorsal = " & sNumDorsal & " AND cod_categoria = " & sCodCategoria & " AND cod_baile = " & sCodBaile & " AND cod_juez = '" & sCodJuez & "' AND fase = " & sFase & " AND repesca=" & sRep
    db.Execute ("UPDATE puntuaciones SET puesto = " & sPuesto & " WHERE num_dorsal = " & sNumDorsal & " AND cod_categoria = " & sCodCategoria & " AND cod_baile = " & sCodBaile & " AND cod_juez = '" & sCodJuez & "' AND fase = " & sFase & " AND repesca=" & sRep)
    
    Call cmdActualizar_Click

End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub

Private Sub tbBaile_DblClick(Index As Integer)
    If Val(Mid$(cbFase.Text, 1, 2)) = 1 Then
        tbBaile(Index).Text = Val(tbBaile(Index).Text) + 1
    Else
        If tbBaile(Index).Text = "0" Then
            tbBaile(Index).Text = "1"
        Else
            tbBaile(Index).Text = "0"
        End If
    End If
End Sub

Private Sub tbBaile_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
        tbBaile(Index).Text = Chr$(KeyAscii)
        KeyAscii = 0
    End If
End Sub

Private Sub tbDorsal_DblClick()
Dim rs As Recordset
    If tbCodCateg.Text <> "" And Trim$(Mid$(cbFase.Text, 1, 2)) <> "" Then
        Set rs = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca = " & chkRep.Value & " AND repesca = " & chkRep.Value & " AND fase = " & Mid$(cbFase.Text, 1, 2) & " ORDER BY 1", dbOpenSnapshot)
        If Not rs.EOF Then
            Do While Not rs.EOF
                If rs!num_dorsal > Val(tbDorsal.Text) Then
                    tbDorsal.Text = rs!num_dorsal
                    Exit Do
                Else
                    rs.MoveNext
                End If
            Loop
            If rs.EOF Then
                If cbJuez.ListIndex < cbJuez.ListCount - 1 Then
                    cbJuez.ListIndex = cbJuez.ListIndex + 1
                    rs.MoveFirst
                    tbDorsal.Text = rs!num_dorsal
                End If
            End If
        End If
        rs.Close
    End If
End Sub
