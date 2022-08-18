VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAJuezBaile 
   Caption         =   "mml_FRASE0051"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   8580
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCambiarJuez 
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
      Left            =   5430
      TabIndex        =   42
      Top             =   2640
      Width           =   1245
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
      Height          =   420
      Left            =   7530
      Picture         =   "frmAJuezBaile.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton cmdCambiar 
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
      Left            =   1350
      TabIndex        =   37
      Top             =   2700
      Width           =   1305
   End
   Begin VB.CheckBox chkSoloJueces 
      BackColor       =   &H0080C0FF&
      Caption         =   "mml_FRASE1121"
      Height          =   375
      Left            =   2880
      TabIndex        =   36
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox chkAutorizar 
      Height          =   375
      Left            =   3900
      TabIndex        =   35
      Top             =   7530
      Width           =   225
   End
   Begin VB.CommandButton cmdImportarBailes 
      Caption         =   "mml_FRASE1029"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   90
      TabIndex        =   34
      Top             =   7470
      Width           =   3795
   End
   Begin VB.TextBox tbCodCategCopia 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4440
      MaxLength       =   5
      TabIndex        =   33
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox tbDescCategCopia 
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
      Left            =   5280
      TabIndex        =   32
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmdCategCopia 
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
      Left            =   3990
      Picture         =   "frmAJuezBaile.frx":046A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1110
      Width           =   465
   End
   Begin VB.CommandButton cmdCopiarDatos 
      Caption         =   "mml_FRASE1021"
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
      TabIndex        =   30
      Top             =   1080
      Width           =   2535
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
      Height          =   375
      Left            =   2070
      Picture         =   "frmAJuezBaile.frx":08D4
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   600
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
      Left            =   2070
      Picture         =   "frmAJuezBaile.frx":0D3E
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   120
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
      Height          =   915
      Left            =   6660
      TabIndex        =   27
      Top             =   7230
      Width           =   1875
   End
   Begin VB.ComboBox cbPista 
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
      ItemData        =   "frmAJuezBaile.frx":11A8
      Left            =   7590
      List            =   "frmAJuezBaile.frx":11CA
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2010
      Width           =   915
   End
   Begin VB.TextBox tbPos 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   23
      Text            =   "1"
      Top             =   2400
      Width           =   420
   End
   Begin VB.CheckBox chkJuezPorCat 
      Caption         =   "mml_FRASE0369"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6090
      TabIndex        =   22
      Top             =   2040
      Width           =   1485
   End
   Begin VB.CheckBox chkJuezPorComp 
      Caption         =   "mml_FRASE0370"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4335
      TabIndex        =   21
      Top             =   2025
      Width           =   1755
   End
   Begin VB.CheckBox chkBailesPorCateg 
      Caption         =   "mml_FRASE0371"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   165
      TabIndex        =   20
      Top             =   2070
      Width           =   4065
   End
   Begin VB.CheckBox chkPasos 
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Top             =   2640
      Width           =   255
   End
   Begin VB.ComboBox cbIdJuez 
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
      ItemData        =   "frmAJuezBaile.frx":11FD
      Left            =   6720
      List            =   "frmAJuezBaile.frx":1387
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2400
      Width           =   795
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
      Height          =   405
      Left            =   6060
      TabIndex        =   16
      Top             =   1560
      Width           =   1485
   End
   Begin VB.ComboBox cbFase 
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
      ItemData        =   "frmAJuezBaile.frx":1579
      Left            =   2040
      List            =   "frmAJuezBaile.frx":1583
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      Caption         =   "mml_FRASE0050"
      Height          =   3975
      Left            =   4320
      TabIndex        =   12
      Top             =   3000
      Width           =   4095
      Begin MSDataGridLib.DataGrid dgJueces 
         Bindings        =   "frmAJuezBaile.frx":15A3
         Height          =   3645
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6429
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
   Begin VB.CommandButton cmdPonerJuez 
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
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuitarJuez 
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
      Left            =   5430
      TabIndex        =   10
      Top             =   2400
      Width           =   1245
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
      TabIndex        =   8
      Top             =   600
      Width           =   4935
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
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   7
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      Height          =   315
      Left            =   1350
      TabIndex        =   3
      Top             =   2400
      Width           =   1305
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
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0185"
      Height          =   4125
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   4095
      Begin MSDataGridLib.DataGrid dgBailes 
         Bindings        =   "frmAJuezBaile.frx":15BB
         Height          =   3795
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6694
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
   Begin MSAdodcLib.Adodc adoBailes 
      Height          =   495
      Left            =   90
      Top             =   7290
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
      RecordSource    =   "SELECT codigo,nombre,fase FROM bailes_categ, bailes WHERE cod_baile = codigo"
      Caption         =   "mml_FRASE0185"
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
   Begin MSAdodcLib.Adodc adoJueces 
      Height          =   495
      Left            =   2310
      Top             =   7290
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
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
      RecordSource    =   "SELECT codigo, id_juez FROM jueces, juez_categ where codigo = cod_juez"
      Caption         =   "mml_FRASE0050"
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
   Begin VB.Frame Frame3 
      Height          =   1305
      Left            =   4500
      TabIndex        =   39
      Top             =   6990
      Width           =   2055
      Begin VB.TextBox tbPanel 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   690
         TabIndex        =   44
         Top             =   870
         Width           =   735
      End
      Begin VB.TextBox tbNumJueces 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1350
         TabIndex        =   40
         Top             =   150
         Width           =   465
      End
      Begin VB.Label Label7 
         Caption         =   "mml_FRASE1094"
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
         Left            =   90
         TabIndex        =   43
         Top             =   630
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "mml_FRASE0050"
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
         Left            =   90
         TabIndex        =   41
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   7590
      TabIndex        =   26
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "mml_FRASE0372"
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
      Left            =   2685
      TabIndex        =   24
      Top             =   2490
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "mml_FRASE0373"
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
      Left            =   7860
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
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
      TabIndex        =   9
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
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmAJuezBaile"
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

Private Sub chkAutorizar_Click()
    If chkAutorizar.Value = 1 Then
        cmdImportarBailes.Enabled = True
    Else
        cmdImportarBailes.Enabled = False
    End If
End Sub

Private Sub chkJuezPorCat_Click()
    chkJuezPorComp.Value = 0
End Sub

Private Sub chkJuezPorComp_Click()
    chkJuezPorCat.Value = 0
End Sub

Public Sub cmdActualizar_Click()
Dim rs As Recordset
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    Sleep 500
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    adoBailes.ConnectionString = "DSN=Escrutinio"
    adoBailes.RecordSource = "SELECT codigo, nombre, fase, posicion FROM bailes_categ, bailes WHERE cod_baile = codigo AND cod_categoria = " & Val(tbCodCateg.Text) & " AND fase = " & cbFase.ListIndex + 1 & " ORDER BY posicion,3,2"
    adoBailes.Refresh
    adoJueces.ConnectionString = "DSN=Escrutinio"
    adoJueces.RecordSource = "SELECT codigo, id_juez,cod_categoria, nombre, pasos  FROM juez_categ, jueces WHERE cod_juez = codigo AND cod_categoria = " & Val(tbCodCateg.Text) & " ORDER BY 2"
    adoJueces.Refresh
    
    dgBailes.Columns(0).Width = 800
    dgBailes.Columns(1).Width = 800
    dgBailes.Columns(2).Width = 800
    dgBailes.Columns(3).Width = 800
    
    dgJueces.Columns(0).Width = 400
    dgJueces.Columns(1).Width = 500
    dgJueces.Columns(2).Width = 500
    dgJueces.Columns(3).Width = 1500
    dgJueces.Columns(4).Width = 400
    
    If tbCodCateg.Text <> "" Then
        Set rs = db.OpenRecordset("SELECT MAX(posicion) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & cbFase.ListIndex + 1, dbOpenSnapshot)
        tbPos.Text = "1"
        If Not IsNull(rs.Fields(0)) Then
            If rs.Fields(0) > 0 And rs.Fields(0) < 90 Then
                tbPos.Text = rs.Fields(0) + 1
            End If
        End If
        rs.Close
        
        tbPanel.Text = PanelJuecesCateg(Val(tbCodCateg.Text))
    Else
        tbPanel.Text = ""
    End If
    
    tbNumJueces.Text = dgJueces.ApproxCount
End Sub

Private Sub cmdCambiar_Click()
Dim iCodBaileNuevo As Integer
Dim sBaile As String
Dim rs As Recordset
Dim iCodBaile As Integer
Dim iCodFase As Integer

    If dgBailes.Row > 0 Or tbCodCateg.Text = "" Or tbPos.Text = "" Then
        MsgBox mml_FRASE0374, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    dgBailes.Col = 0
    iCodBaile = dgBailes.Text
    dgBailes.Col = 2
    iCodFase = dgBailes.Text
    
    iCodBaileNuevo = Val(sSeleccionar("SELECT * FROM bailes"))
    If iCodBaile > 0 Then
        If Not ExisteBaile(tbCodCateg.Text, iCodBaileNuevo, cbFase.ListIndex + 1) Then
            db.Execute "UPDATE bailes_categ SET cod_baile =" & iCodBaileNuevo & " WHERE cod_baile = " & iCodBaile & " AND cod_categoria = " & tbCodCateg.Text & " AND fase =" & iCodFase
            'Actualizamos todas las posibles puntuaciones introducidas
            db.Execute "UPDATE puntuaciones SET cod_baile =" & iCodBaileNuevo & " WHERE cod_baile = " & iCodBaile & " AND cod_categoria = " & tbCodCateg.Text & " AND (fase =" & iCodFase & " OR (fase > " & iCodFase & " AND " & iCodFase & " > 1 ))"
        Else
            MsgBox mml_FRASE1102, vbOKOnly Or vbCritical, G_MSG_ERROR
        End If
    End If
    Call cmdActualizar_Click
End Sub

Private Sub cmdCambiarJuez_Click()
Dim rs As Recordset
Dim lCodJuez As Integer
Dim lCodJuezAnt As Integer
Dim sIdJuez As String


    If dgBailes.Row > 0 Or tbCodCateg.Text = "" Or tbPos.Text = "" Then
        MsgBox mml_FRASE0374, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE1184, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If
    
    dgJueces.Col = 0
    lCodJuezAnt = dgJueces.Text
    dgJueces.Col = 1
    sIdJuez = dgJueces.Text
    
    lCodJuez = 0
    If G_ASIGNAR_AUTOMATICAMENTE_LETRA_A_JUEZ Then
        Set rs = db.OpenRecordset("SELECT TOP 1 cod_juez FROM juez_categ WHERE id_juez = '" & cbIdJuez.Text & "' AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion =  " & tbCodComp.Text & ")", dbOpenSnapshot)
        If Not rs.EOF Then
            lCodJuez = rs.Fields("cod_juez")
        End If
        rs.Close
    End If
    If lCodJuez = 0 Then
        lCodJuez = Val(sSeleccionar("SELECT * FROM jueces j WHERE codigo = (SELECT MAX(CODIGO) FROM jueces j1 WHERE j1.nombre = j.nombre)"))
    End If
    
    If lCodJuez > 0 Then
        db.Execute "UPDATE juez_categ SET cod_juez =" & lCodJuez & ", id_juez = '" & cbIdJuez.Text & "' WHERE cod_juez = " & lCodJuezAnt & " AND cod_categoria = " & tbCodCateg.Text
        'Actualizamos todas las posibles puntuaciones introducidas
        If cbIdJuez.Text <> sIdJuez Then
            db.Execute "UPDATE puntuaciones SET cod_juez ='" & cbIdJuez.Text & "' WHERE cod_juez = '" & sIdJuez & "' AND cod_categoria = " & tbCodCateg.Text
        End If
    End If
    Call cmdActualizar_Click

End Sub

Private Sub cmdCateg_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCateg.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCateg.Text = sResultado(2)
    Call cmdActualizar_Click
End Sub


Private Sub cmdCategCopia_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCategCopia.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCategCopia.Text = sResultado(2)

End Sub

Private Sub cmdCopiarDatos_Click()
Dim rs As Recordset
    If tbCodCateg.Text = "" Or tbCodCategCopia.Text = "" Then
        MsgBox mml_FRASE0264, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE1022, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        
        If chkSoloJueces.Value = 0 Then
            db.Execute "DELETE FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text
            Set rs = db.OpenRecordset("SELECT * FROM bailes_categ WHERE cod_categoria = " & tbCodCategCopia.Text, dbOpenSnapshot)
            While Not rs.EOF
                db.Execute "INSERT INTO bailes_categ VALUES (" & tbCodCateg.Text & "," & rs!cod_baile & "," & rs!fase & "," & rs!posicion & ")"
                rs.MoveNext
            Wend
            rs.Close
        End If
        
        db.Execute "DELETE FROM juez_categ WHERE cod_categoria = " & tbCodCateg.Text
        Set rs = db.OpenRecordset("SELECT * FROM juez_categ WHERE cod_categoria = " & tbCodCategCopia.Text, dbOpenSnapshot)
        While Not rs.EOF
            db.Execute "INSERT INTO juez_categ VALUES (" & rs!cod_juez & "," & tbCodCateg.Text & ",'" & rs!id_juez & "'," & rs!pasos & ")"
            rs.MoveNext
        Wend
        rs.Close
    End If
    Call cmdActualizar_Click
End Sub

Private Sub cmdImportarBailes_Click()
Dim iCodcomp As Integer
Dim rs As Recordset, rsCat As Recordset

    If Not C_DEBUG Then On Local Error GoTo error
    If MsgBox(mml_FRASE1030, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        iCodcomp = Val(sSeleccionar("SELECT * FROM competiciones ORDER BY 1"))
        If iCodcomp = 0 Then Exit Sub
        'Recuperamos la información de los bailes de la competición seleccionada
        Set rsCat = db.OpenRecordset("SELECT * FROM categorias WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
        While Not rsCat.EOF
            Set rs = db.OpenRecordset("SELECT * FROM bailes_categ WHERE cod_categoria = (select TOP 1 codigo FROM categorias WHERE cod_competicion = " & iCodcomp & " AND cod_modalidad = " & rsCat!cod_modalidad & " AND id_categoria = '" & rsCat!id_categoria & "')")
            If Not rs.EOF Then
                'Borramos los bailes que tenga asociados la categoría
                db.Execute "DELETE FROM bailes_categ WHERE cod_categoria = " & rsCat!codigo
                While Not rs.EOF
                    db.Execute "INSERT INTO bailes_categ VALUES (" & rsCat!codigo & "," & rs!cod_baile & "," & rs!fase & "," & rs!posicion & ")"
                    rs.MoveNext
                Wend
                rs.Close
            End If
            rsCat.MoveNext
        Wend
        rsCat.Close
        MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    End If
    Exit Sub
error:
    ProcesarError "cmdImportarBailes_Click"
End Sub

Private Sub cmdPoner_Click()
Dim iCodBaile As Integer
Dim rs As Recordset

    If tbCodCateg.Text = "" Or tbPos.Text = "" Then
        MsgBox mml_FRASE0374, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    iCodBaile = Val(sSeleccionar("SELECT * FROM bailes"))
    If iCodBaile > 0 Then
        If chkBailesPorCateg.Value Then
            Set rs = db.OpenRecordset("SELECT codigo FROM categorias  WHERE cod_competicion = " & tbCodComp.Text & " AND (id_categoria & cod_modalidad) = (SELECT (id_categoria & cod_modalidad) FROM categorias WHERE codigo =" & tbCodCateg.Text & ")", dbOpenSnapshot)
            While Not rs.EOF
                If Not ExisteBaile(rs!codigo, iCodBaile, cbFase.ListIndex + 1) Then
                    db.Execute ("INSERT INTO bailes_categ VALUES(" & rs!codigo & "," & iCodBaile & "," & cbFase.ListIndex + 1 & "," & tbPos.Text & ")")
                End If
                rs.MoveNext
            Wend
            rs.Close
        Else
            If Not ExisteBaile(tbCodCateg.Text, iCodBaile, cbFase.ListIndex + 1) Then
                db.Execute ("INSERT INTO bailes_categ VALUES(" & tbCodCateg.Text & "," & iCodBaile & "," & cbFase.ListIndex + 1 & "," & tbPos.Text & ")")
            Else
                MsgBox mml_FRASE1102, vbOKOnly Or vbCritical, G_MSG_ERROR
            End If
        End If
    End If
    tbPos = Val(tbPos) + 1
    Call cmdActualizar_Click
End Sub

Function ExisteBaile(lCodCat As Long, iCodBaile As Integer, iCodFase As Integer) As Boolean
Dim rsBailes As Recordset

    Set rsBailes = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & lCodCat & _
                                      " AND cod_baile = " & iCodBaile & " AND fase = " & iCodFase, dbOpenSnapshot)
    If rsBailes.Fields(0) > 0 Then
        ExisteBaile = True
    Else
        ExisteBaile = False
    End If
    rsBailes.Close
    
End Function
Function ExisteCodJuez(lCodCat As Long, sCodJuez As String) As Boolean
Dim rsJueces As Recordset

    Set rsJueces = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE cod_categoria = " & lCodCat & _
                                      " AND id_juez = '" & sCodJuez & "'", dbOpenSnapshot)
    If rsJueces.Fields(0) > 0 Then
        ExisteCodJuez = True
    Else
        ExisteCodJuez = False
    End If
    rsJueces.Close
    
End Function


Private Sub cmdPonerJuez_Click()
Dim iCodJuez As Integer
Dim sIdJuez As String
Dim i As Integer
Dim rs As Recordset
    
    If tbCodCateg.Text = "" Then
        MsgBox mml_FRASE0375, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    'Comprobamos si esta letra ya ha sido asignada en esta competición
    iCodJuez = 0
    If G_ASIGNAR_AUTOMATICAMENTE_LETRA_A_JUEZ Then
        Set rs = db.OpenRecordset("SELECT * FROM bailes", dbOpenSnapshot)
        
        Set rs = db.OpenRecordset("SELECT TOP 1 cod_juez FROM juez_categ WHERE id_juez = '" & cbIdJuez.Text & "' AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion =  " & tbCodComp.Text & ")", dbOpenSnapshot)
        If Not rs.EOF Then
            iCodJuez = rs.Fields("cod_juez")
        End If
        rs.Close
    End If
    If iCodJuez = 0 Then
        iCodJuez = Val(sSeleccionar("SELECT * FROM jueces j WHERE codigo = (SELECT MAX(CODIGO) FROM jueces j1 WHERE j1.nombre = j.nombre)"))
    End If
    
    If iCodJuez > 0 Then
        If chkJuezPorComp.Value Or chkJuezPorCat.Value Then
            If chkJuezPorComp.Value Then
                If chkPasos.Value = 1 Then
                    Set rs = db.OpenRecordset("SELECT codigo, id_categoria, descripcion FROM categorias WHERE id_categoria IN (" & VarCfg("categ_juez_pasos") & ") AND cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
                Else
                    Set rs = db.OpenRecordset("SELECT codigo, id_categoria, descripcion FROM categorias WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
                End If
            Else
                Set rs = db.OpenRecordset("SELECT codigo, descripcion FROM categorias  WHERE cod_competicion = " & tbCodComp.Text & " AND id_categoria = (SELECT id_categoria FROM categorias WHERE codigo =" & tbCodCateg.Text & ")", dbOpenSnapshot)
            End If
            While Not rs.EOF
                If InStr(rs!DESCRIPCION, cbPista.Text) > 0 Then
                    If Not ExisteCodJuez(rs!codigo, cbIdJuez.List(cbIdJuez.ListIndex)) Then
                        db.Execute ("INSERT INTO juez_categ VALUES(" & iCodJuez & ",'" & rs!codigo & "','" & cbIdJuez.List(cbIdJuez.ListIndex) & "'," & IIf(chkPasos.Value <> False, 1, 0) & ")")
                    End If
                End If
                rs.MoveNext
            Wend
            rs.Close
        Else
            If Not ExisteCodJuez(Val(tbCodCateg.Text), cbIdJuez.List(cbIdJuez.ListIndex)) Then
                db.Execute ("INSERT INTO juez_categ VALUES(" & iCodJuez & ",'" & tbCodCateg.Text & "','" & cbIdJuez.List(cbIdJuez.ListIndex) & "'," & IIf(chkPasos.Value <> False, 1, 0) & ")")
            End If
        End If
        cbIdJuez.ListIndex = cbIdJuez.ListIndex + 1
        Call cmdActualizar_Click
    End If

End Sub

Private Sub cmdQuitar_Click()
Dim iCodBaile As Integer
Dim iCodFase As Integer
Dim rs As Recordset
    
    dgBailes.Col = 0
    
    If dgBailes.Text = "" Or tbCodCateg.Text = "" Then
        MsgBox mml_FRASE0376, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    dgBailes.Col = 0
    iCodBaile = dgBailes.Text
    dgBailes.Col = 2
    iCodFase = dgBailes.Text
    
    If chkBailesPorCateg.Value Then
        Set rs = db.OpenRecordset("SELECT codigo FROM categorias  WHERE cod_competicion = " & tbCodComp.Text & " AND (id_categoria & cod_modalidad) = (SELECT (id_categoria & cod_modalidad) FROM categorias WHERE codigo =" & tbCodCateg.Text & ")", dbOpenSnapshot)
        While Not rs.EOF
            If Not ComprobarPuntuaciones(iCodBaile, rs!codigo, iCodFase) Then
                db.Execute ("DELETE FROM bailes_categ WHERE cod_baile =" & iCodBaile & " AND fase = " & iCodFase & " AND cod_categoria = " & rs!codigo)
            Else
                MsgBox mml_FRASE1164 & " CodCat: " & rs.Fields("codigo"), vbOKOnly Or vbCritical, G_MSG_ERROR
            End If
            rs.MoveNext
        Wend
        rs.Close
    Else
        If Not ComprobarPuntuaciones(iCodBaile, tbCodCateg.Text, iCodFase) Then
            db.Execute ("DELETE FROM bailes_categ WHERE cod_baile =" & iCodBaile & " AND fase = " & iCodFase & " AND cod_categoria = " & tbCodCateg.Text)
        Else
            MsgBox mml_FRASE1164, vbOKOnly Or vbCritical, G_MSG_ERROR
        End If
    End If
    Call cmdActualizar_Click
End Sub

Function ComprobarPuntuaciones(iCodBaile As Integer, lCodCateg As Long, iCodFase As Integer) As Boolean
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & lCodCateg & " AND cod_baile = " & iCodBaile & " AND (fase = " & iCodFase & " OR (fase >" & iCodFase & " AND " & iCodFase & " > 1))", dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        ComprobarPuntuaciones = True
    Else
        ComprobarPuntuaciones = False
    End If
    rs.Close
End Function
Function ComprobarPuntuacionesJuez(sCodJuez As String, lCodCateg As Long) As Boolean
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & lCodCateg & " AND cod_juez = '" & sCodJuez & "'", dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        ComprobarPuntuacionesJuez = True
    Else
        ComprobarPuntuacionesJuez = False
    End If
    rs.Close
End Function


Private Sub CommandButton1_Click()

End Sub

Private Sub cmdQuitarJuez_Click()
Dim iCodJuez As Integer
Dim sCodJuez As String
Dim rs As Recordset
    
    dgJueces.Col = 0
    If dgJueces.Row < 0 Then
        Exit Sub
    End If
    
    If dgJueces.Text = "" Or tbCodCateg.Text = "" Then
        MsgBox mml_FRASE0375, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    dgJueces.Col = 0
    iCodJuez = dgJueces.Text
    dgJueces.Col = 1
    sCodJuez = dgJueces.Text
    
    If chkJuezPorComp.Value Or chkJuezPorCat.Value Then
        If chkJuezPorComp.Value = 1 Then
            Set rs = db.OpenRecordset("SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
        Else
            Set rs = db.OpenRecordset("SELECT codigo FROM categorias  WHERE cod_competicion = " & tbCodComp.Text & " AND id_categoria = (SELECT id_categoria FROM categorias WHERE codigo =" & tbCodCateg.Text & ")", dbOpenSnapshot)
        End If
        While Not rs.EOF
            If Not ComprobarPuntuacionesJuez(sCodJuez, rs.Fields("codigo")) Then
                db.Execute ("DELETE FROM juez_categ WHERE cod_juez =" & iCodJuez & " AND cod_categoria = " & rs!codigo)
            Else
                MsgBox mml_FRASE1164 & " CodCat: " & rs.Fields("codigo"), vbOKOnly Or vbCritical, G_MSG_ERROR
            End If
            rs.MoveNext
        Wend
        rs.Close
    Else
        If Not ComprobarPuntuacionesJuez(sCodJuez, Val(tbCodCateg.Text)) Then
            db.Execute ("DELETE FROM juez_categ WHERE cod_juez =" & iCodJuez & " AND cod_categoria = " & tbCodCateg.Text)
        Else
            MsgBox mml_FRASE1164, vbOKOnly Or vbCritical, G_MSG_ERROR
        End If
    End If
    
    Call cmdActualizar_Click

End Sub

Private Sub cmdSalir_Click()
    
    If Val(tbCodCateg.Text) > 0 Then
        If Not ComprobarPosicionesConsecutivas Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCateg.Text = ""
    tbDescCateg.Text = ""
    Call cmdActualizar_Click
End Sub

Private Sub cmdSubirDatos_Click()
    If cbIdJuez.ListCount > 0 Then
        cbIdJuez.ListIndex = 0
        cbIdJuez.SetFocus
    End If
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    cbFase.ListIndex = 0
    cbIdJuez.ListIndex = 0
    CargarPistas cbPista
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
        cbFase.ListIndex = 0
    Else
        tbCodCateg.Text = ""
        tbDescCateg.Text = ""
        tbDescCateg.Text = sCateg
    End If
    Call cmdActualizar_Click
    Exit Sub
error:
    ProcesarError "tbCodCategCopia_LostFocus"
End Sub

Private Sub tbCodCategCopia_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub
Private Sub tbCodCategCopia_GotFocus()
    tbCodCategCopia.SelStart = 0
    tbCodCategCopia.SelLength = Len(tbCodCategCopia.Text)
End Sub

Private Sub tbCodCategCopia_LostFocus()
Dim sCateg As String
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    If Val(tbCodCategCopia.Text) > 0 Then
        sCateg = sDescCategoria(tbCodCategCopia.Text)
        If sCateg = "" Then
            tbCodCategCopia.Text = ""
            tbDescCategCopia.Text = ""
        Else
            tbDescCategCopia.Text = sCateg
        End If
    End If
    Call cmdActualizar_Click
    Exit Sub
error:
    ProcesarError "tbCodCategCopia_LostFocus"
End Sub


Private Sub tbPos_KeyPress(KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Function ComprobarPosicionesConsecutivas() As Boolean
    Dim rs As Recordset
    Dim iBailes As Integer
    Dim iMaxPos As Integer
    Dim iPos As Integer
    
    ComprobarPosicionesConsecutivas = True
    'Comprobamos el número de bailes de esta categoría y fase y lo comparamos con la última posición
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & (cbFase.ListIndex + 1), dbOpenSnapshot)
    iBailes = rs.Fields(0)
    rs.Close
    If iBailes = 0 Then
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT MAX(posicion) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & (cbFase.ListIndex + 1), dbOpenSnapshot)
    iMaxPos = rs.Fields(0)
    rs.Close
    
    'Comprobamos si hay asignada la posición 6
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE posicion = 6 AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & (cbFase.ListIndex + 1), dbOpenSnapshot)
    iPos = rs.Fields(0)
    rs.Close
    
    'If iBailes > 0 And iMaxPos <> iBailes Then
    '    MsgBox mml_FRASE1286, vbOKOnly Or vbCritical, "ERROR"
    '    ComprobarPosicionesConsecutivas = False
    'End If
    
    'Nunca debe asignarse la posición 6 con menos de 6 bailes
    If iPos > 0 And iBailes < 6 Then
        If MsgBox(mml_FRASE1286, vbYesNo Or vbCritical, "ERROR") = vbYes Then
            db.Execute "UPDATE bailes_categ SET posicion = " & iMaxPos + 1 & " WHERE posicion = 6 AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & (cbFase.ListIndex + 1)
            Call cmdActualizar_Click
        End If
        
        ComprobarPosicionesConsecutivas = False
    End If
    
    
End Function
