VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmHorario 
   Caption         =   "mml_FRASE0079"
   ClientHeight    =   10395
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16275
   Icon            =   "frmTablasDatos.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   16275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameBotones 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "frm_FRASE1035"
      ForeColor       =   &H80000008&
      Height          =   10335
      Left            =   13380
      TabIndex        =   12
      Top             =   30
      Width           =   2865
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
         Height          =   315
         Left            =   270
         TabIndex        =   47
         Top             =   8430
         Width           =   2445
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
         Height          =   285
         Left            =   270
         TabIndex        =   46
         Top             =   8760
         Width           =   2445
      End
      Begin VB.CommandButton cmdCateg 
         Height          =   675
         Left            =   1560
         Picture         =   "frmTablasDatos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "mml_FRASE0040"
         Top             =   660
         Width           =   615
      End
      Begin VB.CommandButton cmdJueces 
         Height          =   705
         Left            =   2220
         Picture         =   "frmTablasDatos.frx":0EEC
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "mml_FRASE0040"
         Top             =   630
         Width           =   615
      End
      Begin VB.CommandButton cmdVerGrupos 
         Caption         =   "mml_FRASE1253"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   43
         Top             =   6420
         Width           =   2445
      End
      Begin VB.CommandButton cmdInsertarSeparadorFases 
         Caption         =   "mml_FRASE1250"
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
         Left            =   270
         TabIndex        =   42
         Top             =   9150
         Width           =   2445
      End
      Begin VB.CommandButton cmdPuntuaciones 
         Height          =   585
         Left            =   2190
         Picture         =   "frmTablasDatos.frx":2302
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "mml_FRASE0040"
         Top             =   30
         Width           =   645
      End
      Begin VB.CommandButton cmdDorsales 
         Height          =   525
         Left            =   1560
         Picture         =   "frmTablasDatos.frx":2EE4
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "mml_FRASE0028"
         Top             =   90
         Width           =   615
      End
      Begin VB.CommandButton cmdResetearGrupoBaile 
         Caption         =   "mml_FRASE1242"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   39
         Top             =   6060
         Width           =   2445
      End
      Begin VB.CommandButton cmdGrupoInicialPista 
         Caption         =   "mml_FRASE1236"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   38
         Top             =   5460
         Width           =   2445
      End
      Begin VB.CommandButton Todos1Grupo 
         Caption         =   "mml_FRASE1237"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   37
         Top             =   5760
         Width           =   2445
      End
      Begin VB.Frame Frame4 
         Height          =   1365
         Left            =   90
         TabIndex        =   32
         Top             =   0
         Width           =   1425
         Begin VB.TextBox tbDorsales 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   330
            TabIndex        =   34
            Top             =   990
            Width           =   705
         End
         Begin VB.TextBox tbDorsalesGrupo 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   330
            TabIndex        =   33
            Top             =   390
            Width           =   705
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "mml_FRASE1036"
            Height          =   255
            Left            =   90
            TabIndex        =   36
            Top             =   750
            Width           =   1305
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "mml_FRASE1035"
            Height          =   255
            Left            =   60
            TabIndex        =   35
            Top             =   150
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdAsignarPista 
         Caption         =   "mml_FRASE1176"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   270
         TabIndex        =   31
         Top             =   3690
         Width           =   1395
      End
      Begin VB.CommandButton cmdMoverDespues 
         Caption         =   "mml_FRASE1158"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   30
         Top             =   3030
         Width           =   2445
      End
      Begin VB.CommandButton cmdAsignarFase 
         Caption         =   "mml_FRASE0291"
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
         Left            =   300
         TabIndex        =   29
         Top             =   2040
         Width           =   2445
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
         ItemData        =   "frmTablasDatos.frx":3A06
         Left            =   1710
         List            =   "frmTablasDatos.frx":3A28
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3690
         Width           =   1065
      End
      Begin VB.CommandButton cmdReagrupar 
         Caption         =   "mml_FRASE1032"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   27
         Top             =   5040
         Width           =   2445
      End
      Begin VB.CommandButton cmdNuevoGrupo 
         Caption         =   "mml_FRASE1031"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   26
         Top             =   4740
         Width           =   2445
      End
      Begin VB.CommandButton cmdMover 
         Caption         =   "mml_FRASE1007"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   25
         Top             =   2310
         Width           =   2445
      End
      Begin VB.CommandButton cmdInicioSesion 
         Caption         =   "mml_FRASE0997"
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
         Left            =   270
         TabIndex        =   24
         Top             =   4140
         Width           =   2445
      End
      Begin VB.CommandButton cmdCambiarHoraORden 
         Caption         =   "mml_FRASE0996"
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
         Left            =   300
         TabIndex        =   23
         Top             =   1770
         Width           =   2445
      End
      Begin VB.CommandButton cmdRenumerar 
         Caption         =   "mml_FRASE0974"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   22
         Top             =   4410
         Width           =   2445
      End
      Begin VB.CommandButton cmdBorrarLinea 
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
         Height          =   315
         Left            =   300
         TabIndex        =   21
         Top             =   2580
         Width           =   2445
      End
      Begin VB.CommandButton cmdCopiarLinea 
         Caption         =   "mml_FRASE0958"
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
         Left            =   300
         TabIndex        =   20
         Top             =   3330
         Width           =   2445
      End
      Begin VB.CommandButton cmdBailes 
         Caption         =   "mml_FRASE0185"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   270
         TabIndex        =   19
         Top             =   8040
         Width           =   2445
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
         Height          =   315
         Left            =   300
         TabIndex        =   18
         Top             =   1410
         Width           =   2445
      End
      Begin VB.CommandButton cmdSelecHorario 
         Caption         =   "mml_FRASE0854"
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
         Left            =   270
         TabIndex        =   17
         Top             =   6780
         Width           =   2445
      End
      Begin VB.CommandButton cmdOculparPublic 
         Caption         =   "mml_FRASE0956"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   16
         Top             =   7710
         Width           =   2445
      End
      Begin VB.CommandButton cmdAsigPublicidad 
         Caption         =   "mml_FRASE0957"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   270
         TabIndex        =   15
         Top             =   7410
         Width           =   2445
      End
      Begin VB.CommandButton cmdBorrarHorario 
         Caption         =   "mml_FRASE0857"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   270
         TabIndex        =   14
         Top             =   7050
         Width           =   2445
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
         Height          =   600
         Left            =   270
         TabIndex        =   13
         Top             =   9540
         Width           =   2445
      End
   End
   Begin VB.Frame frameHorario 
      Caption         =   "mml_FRASE0067"
      Height          =   7725
      Left            =   0
      TabIndex        =   7
      Top             =   690
      Width           =   13380
      Begin VB.CommandButton cmdBajar 
         Height          =   675
         Left            =   12930
         Picture         =   "frmTablasDatos.frx":3A64
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5310
         Width           =   345
      End
      Begin VB.CommandButton cmdSubir 
         Height          =   675
         Left            =   12930
         Picture         =   "frmTablasDatos.frx":3ECE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4620
         Width           =   345
      End
      Begin MSAdodcLib.Adodc AdoHorario 
         Height          =   4170
         Left            =   12930
         Top             =   300
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   7355
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
         RecordSource    =   "select * from horario order by cod_competicion,orden"
         Caption         =   "mml_FRASE0067"
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
      Begin MSDataGridLib.DataGrid dgHorario 
         Bindings        =   "frmTablasDatos.frx":4338
         Height          =   7350
         Left            =   45
         TabIndex        =   10
         Top             =   180
         Width           =   12825
         _ExtentX        =   22622
         _ExtentY        =   12965
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
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
      Begin VB.Label lblSelec 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S E L E C T"
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   13020
         TabIndex        =   11
         Top             =   6120
         Visible         =   0   'False
         Width           =   165
      End
   End
   Begin VB.Frame frmCont 
      Caption         =   "mml_FRASE0189"
      Height          =   675
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin VB.TextBox tbOrdenAct 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2370
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   180
         Width           =   825
      End
      Begin VB.TextBox tbComp 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3210
         TabIndex        =   5
         Top             =   180
         Width           =   9705
      End
      Begin VB.TextBox tbInc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   525
      End
      Begin VB.TextBox tbOrden 
         Alignment       =   2  'Center
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
         Left            =   810
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Inc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "mml_FRASE0190"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gl_iComp As Long
Dim g_aOrden(200) As Integer
Dim g_iCPos As Integer


Sub DefinirFormato()
    dgHorario.Columns.Item(0).Width = 1000
    dgHorario.Columns.Item(1).Width = 3000
    dgHorario.Columns.Item(2).Width = 1000
    dgHorario.Columns.Item(3).Width = 800
    dgHorario.Columns.Item(4).Width = 700
    dgHorario.Columns.Item(5).Width = 700
    dgHorario.Columns.Item(6).Width = 600
    dgHorario.Columns.Item(7).Width = 700
    dgHorario.Columns.Item(8).Width = 650
    dgHorario.Columns.Item(9).Width = 700

End Sub

Private Sub cbPista_Click()
    If Val(tbComp.Tag) > 0 Then
        CargarHorario Val(tbComp.Tag), cbPista.List(cbPista.ListIndex)
    End If
End Sub

Private Sub cmdActualizar_Click()

    If Not C_DEBUG Then On Local Error GoTo error
    AdoHorario.Refresh
    Dim Pos
    Pos = dgHorario.Bookmark
    dgHorario.Bookmark = dgHorario.FirstRow

    DefinirFormato
    Exit Sub
error:
    ProcesarError "cmdActualizar_Click"
End Sub

Private Sub cmdAsignarFase_Click()
Dim iFase As Integer
Dim iOrden As Integer
Dim iCodcomp As Integer

    
    If Not C_DEBUG Then On Local Error GoTo error
    iFase = Val(InputBox(mml_FRASE1153, G_MSG_PREGUNTA, "1"))
    If sDescFase(iFase) <> "" Then
        dgHorario.Col = 8
        iCodcomp = dgHorario.Text
        dgHorario.Col = 6
        iOrden = dgHorario.Text
        db.Execute "UPDATE horario SET numfase = " & iFase & ", fase = '" & sDescFase(iFase) & "' WHERE orden = " & dgHorario.Text & " AND cod_competicion = " & iCodcomp
    
        Sleep 1000
        cmdActualizar_Click
        MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    End If
    
    Exit Sub
error:
    ProcesarError "cmdAsignarFase_Click"
End Sub

Private Sub cmdAsignarPista_Click()
Dim iPista As Integer
Dim sPista As String
Dim iGrupo As Integer
Dim alCodCateg(200) As Long
Dim aiRep(200) As Integer
Dim aiFase(200) As Integer
Dim iCodBaile As Integer
Dim asDescCateg(200) As String
Dim aOrden(200) As Integer
Dim iCPos As Integer
Dim sPosIni As String, i As Integer
Dim iCodcomp As Integer
Dim iPos

    iCodcomp = Val(tbComp.Tag)
    If Not C_DEBUG Then On Local Error GoTo error
                
    If dgHorario.Row >= 0 Then
    
        'Localizamos las lineas del horario marcadas
        For Each iPos In dgHorario.SelBookmarks
            If iCPos < 200 Then
                dgHorario.Bookmark = iPos
                dgHorario.Col = 6
                aOrden(iCPos) = Val(dgHorario.Text)
                dgHorario.Col = 1
                asDescCateg(iCPos) = dgHorario.Text
                dgHorario.Col = 3
                aiFase(iCPos) = Val(dgHorario.Text)
                dgHorario.Col = 4
                alCodCateg(iCPos) = Val(dgHorario.Text)
                dgHorario.Col = 5
                aiRep(iCPos) = Val(dgHorario.Text)
                Inc iCPos
            End If
        Next
        
    
        sPista = InputBox(mml_FRASE1177, G_MSG_MENSAJE, "")
        
        If IsNumeric(sPista) And Val(sPista) >= 0 And Val(sPista) <= 9 Then
        
            iPista = Val(sPista)
            
            iPos = 1
            For i = 0 To iCPos - 1
            
                iPos = InStr(asDescCateg(i), "(P")
                If iPista > 0 Then
                    If iPos > 0 Then
                        asDescCateg(i) = Mid$(asDescCateg(i), 1, iPos - 1) & "(P" & Trim$(Str$(iPista)) & ")"
                    Else
                        asDescCateg(i) = asDescCateg(i) & " " & "(P" & Trim$(Str$(iPista)) & ")"
                    End If
                Else
                    If iPos > 0 Then
                        asDescCateg(i) = Trim$(Mid$(asDescCateg(i), 1, iPos - 1))
                    End If
                End If
                If Val(alCodCateg(i)) > 0 Then
                    db.Execute ("UPDATE categorias SET descripcion = '" & asDescCateg(i) & "' WHERE codigo = " & alCodCateg(i))
                    db.Execute ("UPDATE horario SET grupo = '" & asDescCateg(i) & "' WHERE cod_categoria = " & alCodCateg(i) & " AND numfase = " & aiFase(i) & " AND repesca = " & aiRep(i))
                Else
                    db.Execute ("UPDATE horario SET grupo = '" & asDescCateg(i) & "' WHERE cod_competicion = " & CodCompActiva & " AND orden = " & aOrden(i) & " AND numfase = " & aiFase(i) & " AND repesca = " & aiRep(i))
                End If
            Next
            
            DoEvents
            Sleep 500
            DoEvents
            cmdActualizar_Click
            
            'MsgBox mml_FRASE0313, vbOKOnly Or vbInformation, mml_FRASE0086
        End If
    Else
        MsgBox mml_FRASE0349, vbOKOnly Or vbCritical, G_MSG_ERROR
    End If
    Exit Sub
error:
    ProcesarError "cmdAsignarPista_Click"
End Sub

Private Sub cmdAsigPublicidad_Click()
    db.Execute ("UPDATE cfg SET valor = 'S' WHERE variable ='publicidad_activa'")
End Sub

Private Sub cmdCateg_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    
    dgHorario.Col = 3
    If Val(dgHorario.Text) = 0 Then Exit Sub
    
    If dgHorario.Col >= 0 Then
        dgHorario.Col = 4
        frmACategorias.adoCat.RecordSource = "SELECT * FROM categorias WHERE codigo = " & Val(dgHorario.Text) & " ORDER BY descripcion"
        frmACategorias.adoCat.Refresh
        frmACategorias.Show vbNomodal
    End If
    Exit Sub
error:
     ProcesarError "cmdPuntuaciones_Click", False


End Sub

Private Sub cmdDorsales_Click()
    Dim i As Integer
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    dgHorario.Col = 3
    If Val(dgHorario.Text) = 0 Then Exit Sub
    
    If dgHorario.Row >= 0 Then
        With dgHorario
            .Col = 3
            If Val(.Text) = 99 Then Exit Sub
            frmADorsales.cbFase.ListIndex = Log(Val(.Text)) / Log(2) + 1
            
            frmADorsales.tbCodComp.Text = CodCompActiva
            frmADorsales.tbDescComp.Text = sDescCompeticion(CodCompActiva)
            .Col = 4
            frmADorsales.tbCodCateg.Text = .Text
            .Col = 1
            frmADorsales.tbDescCateg.Text = .Text
            .Col = 3
            For i = 0 To frmADorsales.cbFase.ListCount - 1
                If Val(frmADorsales.cbFase.List(i)) = Val(.Text) Then
                    frmADorsales.cbFase.ListIndex = i
                End If
            Next
            
            frmADorsales.Show vbNomodal
        End With
        
    End If
    Exit Sub
error:
     ProcesarError "cmdDorsales_Click", False
End Sub

Private Sub cmdGrupoInicialPista_Click()
Dim aOrden(200) As Integer
Dim iCPos As Integer, iPos
Dim sPosIni As String, i As Integer
Dim iCodcomp As Integer

    iCodcomp = CodCompActiva
    iCPos = 0
    If MsgBox(mml_FRASE1239, vbYesNo Or vbQuestion) = vbYes Then
        If iCodcomp > 0 Then
            'Localizamos las lineas del horario marcadas para agrupar
            
            For Each iPos In dgHorario.SelBookmarks
                dgHorario.Bookmark = iPos
                dgHorario.Col = 6
                aOrden(iCPos) = Val(dgHorario.Text)
                Inc iCPos
            Next
        
            If iCPos > 0 Then
                'colocamos la primera como grupo inicial, las demás con 0
                iPos = 1
                db.Execute ("UPDATE horario SET inicio_grupo = 1 WHERE cod_competicion = " & iCodcomp & " AND orden = " & aOrden(0))
                For i = 1 To iCPos - 1
                    db.Execute ("UPDATE horario SET inicio_grupo = 0 WHERE cod_competicion = " & iCodcomp & " AND orden = " & aOrden(i))
                    Inc iPos
                Next
                
                Sleep 1000
                DoEvents
                AdoHorario.Refresh
                DefinirFormato
                MsgBox G_MSG_OPERACION_OK, vbOKOnly And vbInformation, G_MSG_MENSAJE
            End If
        End If
    End If
End Sub

Private Sub cmdInicioSesion_Click()
Dim sSesion As String
Dim sOrden As String
Dim sCateg As String
    If MsgBox(mml_FRASE0998, vbYesNo Or vbQuestion) = vbYes Then
        dgHorario.Col = 4
        sCateg = dgHorario.Text
        dgHorario.Col = 6
        sOrden = dgHorario.Text
        dgHorario.Col = 7
        sSesion = dgHorario.Text
        sSesion = IIf(Val(sSesion) = 0, 1, 0)
        db.Execute "UPDATE horario SET inicio_sesion = " & sSesion & "  WHERE cod_categoria = " & sCateg & " AND orden = " & sOrden
        DoEvents
        Sleep 1000
        AdoHorario.Refresh
        cmdActualizar_Click
    End If
    
End Sub


Private Sub cmdInsertarSeparadorFases_Click()
Dim i As Integer
Dim iOrden As Integer
Dim dHora As Date
Dim gGrupo As String
Dim iDuracion As Integer
Dim iInicioSesion As Integer
Dim rs As Recordset
Dim bInsertar As Boolean

    If frmInsertarSeparadorFases.SeleccionarFasesHorario = 0 Then
        'Localizamos el último orden del horario
        Set rs = db.OpenRecordset("SELECT MAX(orden) AS maxorden FROM horario WHERE cod_competicion = " & CodCompActiva, dbOpenSnapshot)
        If Not rs.EOF Then
            If IsNull(rs.Fields("maxorden")) Then
                iOrden = 10
            Else
                iOrden = rs.Fields("maxorden") + 10
            End If
        End If
        rs.Close
        Set rs = db.OpenRecordset("SELECT MAX(hora) AS maxhora FROM horario WHERE cod_competicion = " & CodCompActiva, dbOpenSnapshot)
        If Not rs.EOF Then
            If IsNull(rs.Fields("maxhora")) Then
                dHora = CDate("10:00")
            Else
                dHora = DateAdd("n", 10, rs.Fields("maxhora"))
                dHora = Format(dHora, "hh:nn")
            End If
        End If
        rs.Close
        i = 0
        With frmInsertarSeparadorFases.lstFasesHorario
            For i = 0 To frmInsertarSeparadorFases.lstFasesHorario.ListCount - 1
                If .Selected(i) Then
                    gGrupo = Trim(Mid(.List(i), 1, InStr(.List(i), "[") - 1))
                    iDuracion = Val(Mid(.List(i), InStr(.List(i), "[") + 1))
                    iInicioSesion = .ItemData(i)
                    'Comprobamos que no está ya en el horario
                    Set rs = db.OpenRecordset("SELECT COUNT(*) as cont FROM horario WHERE cod_competicion = " & CodCompActiva & " AND grupo = '" & gGrupo & "'", dbOpenSnapshot)
                    bInsertar = False
                    If rs.Fields("cont") = 0 Then
                        bInsertar = True
                    Else
                        If MsgBox(mml_FRASE1261, vbYesNo Or vbQuestion, "") = vbYes Then
                            bInsertar = True
                        End If
                    End If
                    
                    If bInsertar Then
                        'Insertamos el registro
                        sSQL = "INSERT INTO horario VALUES (#" & dHora & "#,'" & gGrupo & "','-',99,0,0," & iOrden & "," & iInicioSesion & "," & CodCompActiva & ",0,0,0,1)"
                        dHora = DateAdd("n", iDuracion, dHora)
                        dHora = Format(dHora, "hh:nn")
                        iOrden = iOrden + 10
                        Debug.Print sSQL
                        db.Execute (sSQL)
                    End If
                    rs.Close
                End If
            Next
        End With
        Sleep 1000
        cmdActualizar_Click
    End If
End Sub

Private Sub cmdJueces_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    
    dgHorario.Col = 3
    If Val(dgHorario.Text) = 0 Then Exit Sub
    
    If dgHorario.Row >= 0 Then
        With dgHorario
            .Col = 3
            If Val(.Text) = 99 Then Exit Sub
            frmAJuezBaile.cbFase.ListIndex = IIf(Val(.Text) = 1, 0, 1)
            
            frmAJuezBaile.tbCodComp.Text = CodCompActiva
            frmAJuezBaile.tbDescComp.Text = sDescCompeticion(CodCompActiva)
            .Col = 4
            frmAJuezBaile.tbCodCateg.Text = .Text
            .Col = 1
            frmAJuezBaile.tbDescCateg.Text = .Text
            
            frmAJuezBaile.cmdActualizar_Click
            frmAJuezBaile.Show vbModal
        End With
        
    End If
    Exit Sub
error:
     ProcesarError "cmdJueces_Click", False
End Sub

Private Sub cmdMover_Click()
Dim aOrden(200) As Integer
Dim iCPos As Integer, iPos
Dim sPosIni As String, i As Integer
Dim iCodcomp As Integer

    iCodcomp = Val(VarCfg("horario_codcompeticion"))

    sPosIni = InputBox(mml_FRASE1008, G_MSG_AVISO)
    If Val(sPosIni) > 0 And iCodcomp > 0 Then
        'Localizamos las lineas del horario marcadas para desplazar
        
        For Each iPos In dgHorario.SelBookmarks
            dgHorario.Bookmark = iPos
            dgHorario.Col = 6
            aOrden(iCPos) = Val(dgHorario.Text)
            Inc iCPos
        Next
        'Comprobamos si hay hueco para desplazar las líneas
        Dim rs As Recordset
        Set rs = db.OpenRecordset("SELECT TOP 1 orden FROM horario WHERE cod_competicion = " & iCodcomp & " AND orden > " & sPosIni, dbOpenSnapshot)
        If Not rs.EOF Then
            Dim OrdenSig As Long
            OrdenSig = rs.Fields("orden")
            If OrdenSig - Val(sPosIni) - 1 < iCPos Then
                MsgBox mml_FRASE1299, vbOKOnly And vbInformation, G_MSG_MENSAJE
                Exit Sub
            End If
        End If
        rs.Close
        'Desplazamos las lineas. Se renumeran desde sPosIni con incrementos de 1
        iPos = 1
        For i = 0 To iCPos - 1
            db.Execute ("UPDATE horario SET orden = " & Val(sPosIni) + iPos & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & aOrden(i))
            Inc iPos
        Next
        
        Sleep 1000
        DoEvents
        AdoHorario.Refresh
        DefinirFormato
        MsgBox mml_FRASE1009, vbOKOnly And vbInformation, G_MSG_MENSAJE
    End If
End Sub

Private Sub cmdMoverDespues_Click()
Dim sPosIni As String, i As Integer
Dim iCodcomp As Integer
Dim iPos

    iCodcomp = Val(tbComp.Tag)

    'Recuperamos el salto entre categorias
    iPos = Val(tbInc.Text)
    
    If iCodcomp > 0 Then
        'Localizamos las lineas del horario marcadas para desplazar
        If Not lblSelec.Visible Then
            g_iCPos = 0
            For Each iPos In dgHorario.SelBookmarks
                dgHorario.Bookmark = iPos
                dgHorario.Col = 6
                g_aOrden(g_iCPos) = Val(dgHorario.Text)
                Inc g_iCPos
            Next
            lblSelec.Visible = True
            MsgBox mml_FRASE1159, vbOKOnly Or vbInformation, G_MSG_MENSAJE
        End If
    End If

End Sub

Private Sub cmdNuevoGrupo_Click()
Dim aOrden(200) As Integer
Dim iCPos As Integer, iPos
Dim sPosIni As String, i As Integer
Dim iCodcomp As Integer
Dim iMaxGrupo As Integer
Dim rs As Recordset
    iCodcomp = Val(VarCfg("horario_codcompeticion"))

    Set rs = db.OpenRecordset("SELECT MAX(num_grupo) FROM horario WHERE cod_competicion = " & iCodcomp, dbOpenSnapshot)
    
    If IsNull(rs.Fields(0)) Or (rs.Fields(0) < 0 Or rs.Fields(0) > C_MAX_INTEGER) Then
        iMaxGrupo = 0
    Else
        iMaxGrupo = rs.Fields(0)
    End If
    
    sPosIni = InputBox(mml_FRASE1033, G_MSG_AVISO, Str$(iMaxGrupo + 1))
    If sPosIni = "" Then
        Exit Sub
    End If
    If iCodcomp > 0 Then
        iMaxGrupo = Val(sPosIni)
        'Localizamos las lineas del horario marcadas para agrupar
        For Each iPos In dgHorario.SelBookmarks
            dgHorario.Bookmark = iPos
            dgHorario.Col = 6
            aOrden(iCPos) = Val(dgHorario.Text)
            Inc iCPos
        Next
    
        'Agrupamos las lineas
        iPos = 1
        For i = 0 To iCPos - 1
            db.Execute ("UPDATE horario SET num_grupo = " & iMaxGrupo & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & aOrden(i))
            Inc iPos
        Next
        
        Sleep 1000
        DoEvents
        AdoHorario.Refresh
        DefinirFormato
        MsgBox mml_FRASE1009, vbOKOnly And vbInformation, G_MSG_MENSAJE
    End If

End Sub

Private Sub cmdPuntuaciones_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    
    dgHorario.Col = 3
    If Val(dgHorario.Text) = 0 Then Exit Sub
    
    If dgHorario.Row >= 0 Then
        With dgHorario
            .Col = 3
            If Val(.Text) = 99 Then Exit Sub
            frmAPuntuacionesBaile.cbFase.ListIndex = Log(Val(.Text)) / Log(2) + 1
            
            frmAPuntuacionesBaile.tbCodComp.Text = CodCompActiva
            frmAPuntuacionesBaile.tbDescComp.Text = sDescCompeticion(CodCompActiva)
            .Col = 4
            frmAPuntuacionesBaile.tbCodCateg.Text = .Text
            .Col = 1
            frmAPuntuacionesBaile.tbDescCateg.Text = .Text
            
            frmAPuntuacionesBaile.Show vbNomodal
        End With
        
    End If
    Exit Sub
error:
     ProcesarError "cmdPuntuaciones_Click", False

End Sub

Private Sub cmdReagrupar_Click()
Dim rs As Recordset
Dim lCont As Long
Dim lCom As Long
Dim iInc As Integer
Dim bAsignarInicioGrupo As Boolean
Dim iGrupo As Integer
Dim iInicioGrupo As Integer

    bAsignarInicioGrupo = False
    iGrupo = -1
    
    If MsgBox(mml_FRASE0975, vbYesNo Or vbQuestion) = vbYes Then
        tbOrdenAct.Text = 10
        
        lCom = Val(VarCfg("horario_codcompeticion"))
        
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM horario WHERE orden >= " & C_CONT_RENUM_INI & " AND cod_competicion = " & lCom, dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            MsgBox mml_FRASE0976 & " " & C_CONT_RENUM_INI, vbOKOnly Or vbCritical
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
        
        Set rs = db.OpenRecordset("SELECT * FROM horario WHERE cod_competicion = " & lCom & " ORDER BY num_grupo, orden", dbOpenSnapshot)
        lCont = C_CONT_RENUM_INI
        While Not rs.EOF
            db.Execute "UPDATE horario SET orden = " & lCont & " WHERE cod_competicion = " & lCom & " AND orden = " & rs!orden
            Inc lCont
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
    
        If MsgBox(mml_FRASE1262, vbYesNo Or vbQuestion, "") = vbYes Then
            bAsignarInicioGrupo = True
        End If
        
        Set rs = db.OpenRecordset("SELECT * FROM horario WHERE cod_competicion = " & lCom & " ORDER BY orden", dbOpenSnapshot)
        lCont = Val(tbOrden.Text)
        iInc = Val(tbInc.Text)
        If iInc = 0 Then iInc = 1
        While Not rs.EOF
            iInicioGrupo = 0
            If bAsignarInicioGrupo Then
                If iGrupo <> rs!num_grupo Then
                    iInicioGrupo = 1
                    iGrupo = rs!num_grupo
                End If
            End If
            
            db.Execute "UPDATE horario SET inicio_grupo = " & iInicioGrupo & ", orden = " & lCont & " WHERE cod_competicion = " & lCom & " AND orden = " & rs!orden
            lCont = lCont + iInc
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        
        Sleep 1000
        DoEvents
        AdoHorario.Refresh
        DefinirFormato
        MsgBox mml_FRASE1009, vbOKOnly And vbInformation, G_MSG_MENSAJE
    End If

End Sub

Private Sub cmdRegenHorario_Click()
    frmADorsales.cmdRegenHorario_Click
    Sleep 1000
    cmdActualizar_Click
    
End Sub

Private Sub cmdRegenHoras_Click()
    frmADorsales.RegenerarHoras CodCompActiva
    Sleep 1000
    cmdActualizar_Click
    
End Sub


Private Sub cmdSubir_Click()
Dim sHora As String
Dim sHoraAnt As String
Dim sOrden As String
Dim sOrdenAnt As String
Dim iCodcomp As Integer
Dim iCompAnt As Integer
Dim SelRow As Integer
Dim iOrdenAleatorio As Integer
Dim sCategoria As String
Dim sFase As String
Dim rs As Recordset
Static bEjec As Boolean
Dim Pos

    If bEjec Then Exit Sub
    bEjec = True

    If dgHorario.Row = 0 And dgHorario.FirstRow > 1 Then
    Dim iFila As Integer
        iFila = dgHorario.FirstRow - 20
        If iFila < 0 Then iFila = 0
        dgHorario.FirstRow = dgHorario.FirstRow - 1
        dgHorario.Refresh
        DoEvents
    End If


    If dgHorario.Row > 0 Then
        
        Pos = dgHorario.Bookmark
    
        SelRow = dgHorario.Row + dgHorario.FirstRow - 1
        dgHorario.Col = 1
        sCategoria = dgHorario.Text
        dgHorario.Col = 3
        sFase = dgHorario.Text
        
        dgHorario.Col = 0
        sHora = dgHorario.Text
        dgHorario.Col = 8
        iCodcomp = Val(dgHorario.Text)
        dgHorario.Col = 6
        sOrden = dgHorario.Text
        
        dgHorario.Row = dgHorario.Row - 1
        DoEvents
        dgHorario.Col = 0
        sHoraAnt = dgHorario.Text
        dgHorario.Col = 8
        iCompAnt = Val(dgHorario.Text)
        dgHorario.Col = 6
        sOrdenAnt = dgHorario.Text
        Randomize
        iOrdenAleatorio = -Int(Rnd * 1000)
        
        If iCodcomp <> iCompAnt Then
            MsgBox mml_FRASE0993, vbCritical Or vbOKOnly, mml_FRASE0096
            Exit Sub
        End If
        If sOrden = "" Then
            bEjec = False
            Exit Sub
        End If
        
        db.Execute ("UPDATE horario SET orden = " & iOrdenAleatorio & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & sOrden)
        db.Execute ("UPDATE horario SET hora = #" & sHora & "# , orden = " & sOrden & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & sOrdenAnt)
        db.Execute ("UPDATE horario SET hora = #" & sHoraAnt & "# , orden = " & sOrdenAnt & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & iOrdenAleatorio)
        
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM horario WHERE cod_competicion = " & iCodcomp & " AND orden = " & iOrdenAleatorio, dbOpenSnapshot)
        If rs.Fields(0) > 0 Or Val(sOrdenAnt) < 0 Then
            MensajeError mml_FRASE0995
        End If
        rs.Close
        
        dgHorario.Visible = False
        AdoHorario.Refresh
        DoEvents
        dgHorario.Row = SelRow - 1
        dgHorario.Visible = True
        
        DefinirFormato
    Else
        'No puede subir más
        MsgBox mml_FRASE0992, vbOKOnly Or vbCritical, mml_FRASE0096
    End If
    bEjec = False
End Sub

Private Sub cmdBajar_Click()
Dim sHora As String
Dim sHoraSig As String
Dim sOrden As String
Dim sOrdenSig As String
Dim iCodcomp As Integer
Dim iCompSig As Integer
Dim SelRow As Integer
Dim iOrdenAleatorio As Integer
Dim sCategoria As String
Dim sFase As String
Dim rs As Recordset
Static bEjec As Boolean
Dim Pos

    If bEjec Then Exit Sub
    bEjec = True
    
    If dgHorario.Row + dgHorario.FirstRow - 1 < AdoHorario.Recordset.RecordCount - 1 Then
        Pos = dgHorario.FirstRow
        SelRow = dgHorario.Row + dgHorario.FirstRow - 1
        dgHorario.Col = 1
        sCategoria = dgHorario.Text
        dgHorario.Col = 3
        sFase = dgHorario.Text
        
        dgHorario.Col = 0
        sHora = dgHorario.Text
        dgHorario.Col = 8
        iCodcomp = Val(dgHorario.Text)
        dgHorario.Col = 6
        sOrden = dgHorario.Text
        
        dgHorario.Row = dgHorario.Row + 1
        DoEvents
        dgHorario.Col = 0
        sHoraSig = dgHorario.Text
        dgHorario.Col = 8
        iCompSig = Val(dgHorario.Text)
        dgHorario.Col = 6
        sOrdenSig = dgHorario.Text
        Randomize
        iOrdenAleatorio = -Int(Rnd * 1000)
        
        If iCodcomp <> iCompSig Then
            MsgBox mml_FRASE0993, vbCritical Or vbOKOnly, mml_FRASE0096
            bEjec = False
            Exit Sub
        End If
        If sOrden = "" Then
            bEjec = False
            Exit Sub
        End If
        
        db.Execute ("UPDATE horario SET orden = " & iOrdenAleatorio & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & sOrden)
        db.Execute ("UPDATE horario SET hora = #" & sHora & "# , orden = " & sOrden & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & sOrdenSig)
        db.Execute ("UPDATE horario SET hora = #" & sHoraSig & "# , orden = " & sOrdenSig & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & iOrdenAleatorio)
        
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM horario WHERE cod_competicion = " & iCodcomp & " AND orden = " & iOrdenAleatorio, dbOpenSnapshot)
        If rs.Fields(0) > 0 Or Val(sOrdenSig) < 0 Then
            MensajeError mml_FRASE0995
        End If
        rs.Close
        
        dgHorario.Visible = False
        AdoHorario.Refresh
        DoEvents
        dgHorario.Row = SelRow + 1
        dgHorario.Visible = True

        DefinirFormato
    Else
        'No puede bajar más
        MsgBox mml_FRASE0994, vbOKOnly Or vbCritical, mml_FRASE0096
    End If
    bEjec = False
End Sub

Private Sub cmdBorrarHorario_Click()
Dim iCom As Integer
    If MsgBox(mml_FRASE0858, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        iCom = Val(sSeleccionar("SELECT * FROM competiciones ORDER BY 1"))
        If iCom > 0 Then
            db.Execute ("DELETE FROM horario WHERE cod_competicion = " & iCom)
            MsgBox mml_FRASE0859, vbOKOnly Or vbInformation, mml_FRASE0084
        End If
    End If
End Sub

Private Sub cmdBorrarLinea_Click()
Dim aOrden(200) As Integer
Dim iCPos As Integer, iPos
Dim sPosIni As String, i As Integer
Dim iCodcomp As Integer

    iCodcomp = CodCompActiva

    If MsgBox(mml_FRASE0973, vbYesNo Or vbQuestion) = vbYes Then
        If iCodcomp > 0 Then
            'Localizamos las lineas del horario marcadas para borrar
            
            For Each iPos In dgHorario.SelBookmarks
                dgHorario.Bookmark = iPos
                dgHorario.Col = 6
                aOrden(iCPos) = Val(dgHorario.Text)
                Inc iCPos
            Next
        
            'borramos las lineas
            iPos = 1
            For i = 0 To iCPos - 1
                db.Execute ("DELETE FROM horario WHERE cod_competicion = " & iCodcomp & " AND orden = " & aOrden(i))
                Inc iPos
            Next
            
            Sleep 1000
            DoEvents
            AdoHorario.Refresh
            DefinirFormato
            MsgBox G_MSG_OPERACION_OK, vbOKOnly And vbInformation, G_MSG_MENSAJE
        End If
    End If
End Sub

Private Sub cmdCambiarHoraORden_Click()
    dgHorario_DblClick

End Sub

Private Sub cmdCopiarLinea_Click()
Dim iOrden As Integer, iCodcomp As Integer, rs As Recordset, sSQL As String
    If dgHorario.Row >= 0 Then
        dgHorario.Col = 6
        iOrden = dgHorario.Text
        dgHorario.Col = 8
        iCodcomp = dgHorario.Text
        
        Set rs = db.OpenRecordset("SELECT * FROM horario WHERE orden = " & iOrden & " AND cod_competicion = " & iCodcomp, dbOpenSnapshot)
        sSQL = "INSERT INTO horario VALUES (#" & rs!hora & "#,'" & rs!grupo & "','" & rs!fase & "'," & rs!numfase & "," & rs!cod_categoria & "," & rs!repesca & "," & rs!orden + 1 & "," & rs!inicio_sesion & "," & rs!cod_competicion & "," & rs!cod_baile & "," & Val(SinNulos(rs!num_dorsales)) & "," & Val(SinNulos(rs!num_grupo)) & ",1)"
        Debug.Print sSQL
        db.Execute sSQL
        rs.Close
        Sleep 1000
        cmdActualizar_Click
    End If
End Sub


Private Sub cmdRenumerar_Click()
Dim rs As Recordset
Dim lCont As Long
Dim lCom As Long
Dim iInc As Integer
    If MsgBox(mml_FRASE0975, vbYesNo Or vbQuestion) = vbYes Then
        tbOrdenAct.Text = 10
        
        lCom = Val(VarCfg("horario_codcompeticion"))
        
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM horario WHERE orden >= " & C_CONT_RENUM_INI & " AND cod_competicion = " & lCom, dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            MsgBox mml_FRASE0976 & " " & C_CONT_RENUM_INI, vbOKOnly Or vbCritical
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
        rs.Close
        Set rs = Nothing
        
        Set rs = db.OpenRecordset("SELECT * FROM horario WHERE cod_competicion = " & lCom & " ORDER BY orden", dbOpenSnapshot)
        lCont = C_CONT_RENUM_INI
        While Not rs.EOF
            db.Execute "UPDATE horario SET orden = " & lCont & " WHERE cod_competicion = " & lCom & " AND orden = " & rs!orden
            Inc lCont
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
    
        Set rs = db.OpenRecordset("SELECT * FROM horario WHERE cod_competicion = " & lCom & " ORDER BY orden", dbOpenSnapshot)
        lCont = Val(tbOrden.Text)
        iInc = Val(tbInc.Text)
        If iInc = 0 Then iInc = 1
        While Not rs.EOF
            db.Execute "UPDATE horario SET orden = " & lCont & " WHERE cod_competicion = " & lCom & " AND orden = " & rs!orden
            lCont = lCont + iInc
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        
        Sleep 1000
        DoEvents
        AdoHorario.Refresh
        DefinirFormato
        MsgBox mml_FRASE1009, vbOKOnly And vbInformation, G_MSG_MENSAJE
    End If
End Sub

Private Sub cmdSelecHorario_Click()
Dim lCom As Long
    lCom = Val(sSeleccionar("SELECT * FROM competiciones ORDER BY 1"))
    gl_iComp = lCom
    If lCom > 0 Then CargarHorario lCom, cbPista.List(cbPista.ListIndex)
End Sub
Sub CargarHorario(lCom As Long, Optional sPista As String = "")
    If lCom > 0 Then
        AdoHorario.RecordSource = "SELECT * FROM horario WHERE cod_competicion = " & lCom & " AND grupo LIKE '%" & sPista & "%' ORDER BY cod_competicion,orden, hora"
        AdoHorario.Refresh
        DefinirFormato
        tbComp.Text = sDescCompeticion(lCom)
        tbComp.Tag = lCom
    Else
        AdoHorario.RecordSource = "SELECT * FROM horario ORDER BY cod_competicion,orden, hora"
        AdoHorario.Refresh
        DefinirFormato
        tbComp.Text = ""
        tbComp.Tag = 0
    End If

End Sub

Private Sub cmdOculparPublic_Click()
    db.Execute ("UPDATE cfg SET valor = 'N' WHERE variable ='publicidad_activa'")
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

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

Sub PegarElementosCopiadosConMoverDespues()
Dim iPos As Integer
Dim iCodcomp As Integer
Dim lOrden As Long, lSigOrden As Long
Dim rs As Recordset
Dim i As Integer

    If MsgBox(mml_FRASE1160, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
        ' 1.Localizamos la línea a partir de la que queremos insertar
        ' 2.Comprobamos el orden de la siguiete categoria para comprobar si hay hueco
        'Desplazamos las lineas
        iCodcomp = CodCompActiva
        If iCodcomp > 0 And dgHorario.Row >= 0 Then
            dgHorario.Col = 6
            lOrden = Val(dgHorario.Text)
            iPos = 1
            Set rs = db.OpenRecordset("SELECT TOP 1 orden FROM horario WHERE orden > " & lOrden & " AND cod_competicion = " & iCodcomp & " ORDER BY orden", dbOpenSnapshot)
            If Not rs.EOF Then
                lSigOrden = rs.Fields("orden")
            Else
                lSigOrden = MAX_NUMERO
            End If
            
            If lSigOrden - lOrden - 1 >= g_iCPos Then
                For i = 0 To g_iCPos - 1
                    db.Execute ("UPDATE horario SET orden = " & lOrden + iPos & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & g_aOrden(i))
                    Inc iPos
                Next
                
                Sleep 1000
                DoEvents
                AdoHorario.Refresh
                DefinirFormato
                MsgBox mml_FRASE1009, vbOKOnly And vbInformation, G_MSG_MENSAJE
            Else
                MsgBox mml_FRASE1161, vbOKOnly Or vbCritical, G_MSG_ERROR
                lblSelec.Visible = False
            End If
        End If
    End If
    lblSelec.Visible = False

End Sub

Private Sub cmdResetearGrupoBaile_Click()
Dim iCodcomp As Integer

    iCodcomp = CodCompActiva
    If MsgBox(mml_FRASE1241, vbYesNo Or vbQuestion) = vbYes Then
        If iCodcomp > 0 Then
            db.Execute ("UPDATE horario SET inicio_grupo = 0 WHERE cod_competicion = " & iCodcomp)
        End If
    End If

    Sleep 1000
    cmdActualizar_Click
End Sub


Private Sub cmdVerGrupos_Click()
Dim i As Integer
Dim bMarca As Boolean
Dim Mark As SelBookmarks


'On Local Error GoTo error
bMarca = False

    For i = dgHorario.SelBookmarks.Count - 1 To 0 Step -1
        dgHorario.SelBookmarks.Remove i
    Next
    
    i = 0
    While i < dgHorario.VisibleRows
        dgHorario.Row = i
        dgHorario.Col = 0
        'dgHorario.Refresh
        
        If dgHorario.Text = "" Then Exit Sub
        dgHorario.Col = 12
        If dgHorario.Text = "1" Then
            bMarca = Not bMarca
        End If
        
        If bMarca Then
            dgHorario.SelBookmarks.add dgHorario.Bookmark
        End If
        Inc i
        
        
    Wend
    
error:

End Sub

Private Sub dgHorario_Click()
Dim iPos
Dim iNumDorsales As Integer
Dim rs As Recordset
Dim iCodcomp As Integer
Dim iCol As Integer

    If lblSelec.Visible Then
        PegarElementosCopiadosConMoverDespues
    Else
        If Not C_DEBUG Then On Local Error GoTo error
    
        iCodcomp = Val(tbComp.Tag)
    
        If iCodcomp > 0 Then
            iCol = dgHorario.Col
            dgHorario.Col = 11
            If Trim$(dgHorario.Text) <> "" Then
                Set rs = db.OpenRecordset("SELECT SUM(num_dorsales) FROM horario WHERE cod_competicion = " & iCodcomp & " AND num_grupo = " & dgHorario.Text, dbOpenSnapshot)
                tbDorsalesGrupo.Text = Val(SinNulos(rs.Fields(0)))
                rs.Close
            End If
            iNumDorsales = 0
            For Each iPos In dgHorario.SelBookmarks
                dgHorario.Bookmark = iPos
                dgHorario.Col = 10
                iNumDorsales = iNumDorsales + Val(dgHorario.Text)
            Next
            If iCol >= 0 Then
                dgHorario.Col = iCol
            End If
            tbDorsales.Text = iNumDorsales
        End If
    End If
    Exit Sub
error:
    ProcesarError "dgHorario_click"
End Sub

Private Sub dgHorario_DblClick()
Dim sHora As String, iOrden As Integer, iCodcomp As Integer


    sHora = InputBox(mml_FRASE0860 & Chr$(13) & Chr$(10) & mml_FRASE1037, , tbOrdenAct.Text)
    If IsDate(sHora) Then
        dgHorario.Col = 6
        iOrden = Val(dgHorario.Text)
        dgHorario.Col = 8
        iCodcomp = Val(dgHorario.Text)
        db.Execute ("UPDATE horario SET hora = #" & sHora & "# WHERE cod_competicion = " & iCodcomp & " AND orden = " & iOrden)
    ElseIf IsNumeric(sHora) Then
        dgHorario.Col = 6
        iOrden = Val(dgHorario.Text)
        dgHorario.Col = 8
        iCodcomp = Val(dgHorario.Text)
        tbOrdenAct.Text = Val(sHora) + Val(tbInc.Text)
        db.Execute ("UPDATE horario SET orden = " & sHora & " WHERE cod_competicion = " & iCodcomp & " AND orden = " & iOrden)
    End If
End Sub

Private Sub Form_Load()
Dim lCom As Long
    TraducirCadenas Me
    lCom = Val(VarCfg("horario_codcompeticion"))
    CargarHorario lCom
    tbComp.Text = sDescCompeticion(lCom)
    gl_iComp = lCom
    tbOrden.Text = 10
    tbInc.Text = 10
    tbOrdenAct.Text = 10
    
    DefinirFormato
    

End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    frameHorario.Height = Me.Height - 600 - frameHorario.Top
    'frameHorario.Width = Me.Width - frameBotones.Width
    'frameBotones.Left = Me.Width - frameBotones.Width
    
    dgHorario.Height = frameHorario.Height - 200
    
End Sub

Private Sub lblSelec_Click()
    lblSelec.Visible = False
End Sub

Private Sub tbInc_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii

End Sub

Private Sub tbOrden_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbOrdenAct_DblClick()
    tbOrdenAct.Text = 10
End Sub

Private Sub Todos1Grupo_Click()
Dim iCodcomp As Integer

    iCodcomp = CodCompActiva
    If MsgBox(mml_FRASE1240, vbYesNo Or vbQuestion) = vbYes Then
        If iCodcomp > 0 Then
            db.Execute ("UPDATE horario SET inicio_grupo = 1 WHERE cod_competicion = " & iCodcomp)
        End If
    End If
    Sleep 1000
    cmdActualizar_Click
End Sub
