VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCfg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1249"
   ClientHeight    =   10350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   15915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame mrcDatosGen 
      BorderStyle     =   0  'None
      Height          =   10650
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15945
      Begin VB.Frame Frame6 
         Caption         =   "MatchAnalisys"
         Height          =   1695
         Index           =   3
         Left            =   5280
         TabIndex        =   16
         Top             =   8580
         Width           =   10500
         Begin MSDataGridLib.DataGrid dgMA 
            Height          =   1335
            Left            =   60
            TabIndex        =   17
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   2355
            _Version        =   393216
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
         Begin MSAdodcLib.Adodc AdoMA 
            Height          =   1470
            Left            =   10080
            Top             =   120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   2593
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   3
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   1
            Enabled         =   -1
            Connect         =   "DSN=Escrutinio"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "mml_FRASE0033"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "MA_Parametros"
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
      End
      Begin VB.Frame Frame6 
         Caption         =   "mml_FRASE1250"
         Height          =   1695
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   8520
         Width           =   5220
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmCfg.frx":0000
            Height          =   1380
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   2434
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
         Begin MSAdodcLib.Adodc AdoFasesHorario 
            Height          =   1470
            Left            =   4800
            Top             =   240
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   2593
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   3
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   1
            Enabled         =   -1
            Connect         =   "DSN=Escrutinio"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "mml_FRASE0033"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "FasesHorario"
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
      End
      Begin VB.Frame Frame6 
         Caption         =   "mml_FRASE0031"
         Height          =   1695
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   6840
         Width           =   5220
         Begin MSDataGridLib.DataGrid dgCateg 
            Bindings        =   "frmCfg.frx":001E
            Height          =   1380
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   2434
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
         Begin MSAdodcLib.Adodc adoCateg 
            Height          =   1470
            Left            =   4800
            Top             =   240
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   2593
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   3
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   1
            Enabled         =   -1
            Connect         =   "DSN=Escrutinio"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "mml_FRASE0033"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "DescCategoria"
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
      End
      Begin VB.Frame frameCfg 
         Caption         =   "mml_FRASE0078"
         Height          =   8400
         Left            =   5280
         TabIndex        =   9
         Top             =   120
         Width           =   10545
         Begin VB.CommandButton cmdActualziar 
            Height          =   315
            Left            =   10020
            Picture         =   "frmCfg.frx":0036
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   7980
            Width           =   435
         End
         Begin VB.TextBox tbFiltro 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1020
            TabIndex        =   18
            Top             =   7980
            Width           =   8955
         End
         Begin VB.CommandButton cmdModificar 
            Height          =   915
            Left            =   10080
            Picture         =   "frmCfg.frx":0318
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4440
            Width           =   315
         End
         Begin MSDataGridLib.DataGrid dgCfg 
            Bindings        =   "frmCfg.frx":05FA
            Height          =   7650
            Left            =   90
            TabIndex        =   11
            Top             =   225
            Width           =   9870
            _ExtentX        =   17410
            _ExtentY        =   13494
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
         Begin MSAdodcLib.Adodc adoCfg 
            Height          =   4065
            Left            =   10080
            Top             =   240
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   7170
            ConnectMode     =   0
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
            Orientation     =   1
            Enabled         =   -1
            Connect         =   "DSN=Escrutinio"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "mml_FRASE0033"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "SELECT * FROM cfg ORDER BY variable"
            Caption         =   "Cfg"
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
         Begin VB.Label Label1 
            Caption         =   "Filtro"
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
            Left            =   240
            TabIndex        =   19
            Top             =   8040
            Width           =   675
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "mml_FRASE0853"
         Height          =   1695
         Left            =   0
         TabIndex        =   7
         Top             =   5160
         Width           =   5205
         Begin MSDataGridLib.DataGrid DataGrid 
            Bindings        =   "frmCfg.frx":060F
            Height          =   1380
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   2434
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
         Begin MSAdodcLib.Adodc AdoEdad 
            Height          =   1470
            Left            =   4800
            Top             =   225
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   2593
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   3
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   1
            Enabled         =   -1
            Connect         =   "DSN=Escrutinio"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "mml_FRASE0033"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "GruposEdad"
            Caption         =   "mml_FRASE0385"
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
      Begin VB.Frame Frame2 
         Caption         =   "mml_FRASE0851"
         Height          =   1695
         Left            =   0
         TabIndex        =   5
         Top             =   1800
         Width           =   5220
         Begin MSDataGridLib.DataGrid DataGrid 
            Bindings        =   "frmCfg.frx":0625
            Height          =   1380
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   2434
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
         Begin MSAdodcLib.Adodc adoModalidad 
            Height          =   1470
            Left            =   4800
            Top             =   120
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   2593
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   3
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   1
            Enabled         =   -1
            Connect         =   "DSN=Escrutinio"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "mml_FRASE0033"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Modalidad"
            Caption         =   "mml_FRASE0852"
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
      Begin VB.Frame Frame6 
         Caption         =   "mml_FRASE0850"
         Height          =   1695
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   90
         Width           =   5220
         Begin MSDataGridLib.DataGrid dgBailes 
            Bindings        =   "frmCfg.frx":0640
            Height          =   1380
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   2434
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
         Begin MSAdodcLib.Adodc adoBailes 
            Height          =   1470
            Left            =   4800
            Top             =   240
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   2593
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   3
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   1
            Enabled         =   -1
            Connect         =   "DSN=Escrutinio"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "mml_FRASE0033"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Bailes"
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Comunidades "
         Height          =   1695
         Left            =   0
         TabIndex        =   1
         Top             =   3480
         Width           =   5235
         Begin MSDataGridLib.DataGrid DataGrid 
            Bindings        =   "frmCfg.frx":0658
            Height          =   1380
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   2434
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
         Begin MSAdodcLib.Adodc AdoComunidades 
            Height          =   1470
            Left            =   4800
            Top             =   225
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   2593
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   3
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   1
            Enabled         =   -1
            Connect         =   "DSN=Escrutinio"
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   "mml_FRASE0033"
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Comunidades"
            Caption         =   "Comunidades"
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
   End
End
Attribute VB_Name = "frmCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualziar_Click()
Dim filtro As String

    filtro = tbFiltro.Text
    adoCfg.RecordSource = "SELECT * FROM cfg WHERE variable LIKE '%" & filtro & "%' OR descripcion LIKE '%" & filtro & "%' OR desc_ingles LIKE '%" & filtro & "%' ORDER BY variable"
    adoCfg.Refresh
End Sub

Private Sub cmdModificar_Click()
    dgCfg.AllowUpdate = Not dgCfg.AllowUpdate

End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    
    Set dgCateg.DataSource = adoCateg
    Set dgMA.DataSource = AdoMA
End Sub

