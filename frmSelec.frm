VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSelec 
   Caption         =   "mml_FRASE0834"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
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
      Height          =   495
      Left            =   11310
      Picture         =   "frmSelec.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7980
      Width           =   435
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "mml_FRASE1028"
      Default         =   -1  'True
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
      Left            =   6840
      TabIndex        =   10
      Top             =   8040
      Width           =   1365
   End
   Begin VB.TextBox tbBuscar 
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
      Left            =   4020
      TabIndex        =   9
      Top             =   8040
      Width           =   2745
   End
   Begin VB.ComboBox cbOrden 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmSelec.frx":046A
      Left            =   11820
      List            =   "frmSelec.frx":0480
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   7980
      Width           =   615
   End
   Begin VB.CommandButton cmdOrden 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   10830
      TabIndex        =   7
      Top             =   7980
      Width           =   465
   End
   Begin VB.CommandButton cmdOrden 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   10350
      TabIndex        =   6
      Top             =   7980
      Width           =   465
   End
   Begin VB.CommandButton cmdOrden 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   9870
      TabIndex        =   5
      Top             =   7980
      Width           =   465
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "mml_FRASE0625"
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
      Left            =   360
      TabIndex        =   4
      Top             =   7980
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "mml_FRASE0836"
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
      Left            =   2010
      TabIndex        =   3
      Top             =   7980
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc adoSelec 
      Height          =   375
      Left            =   180
      Top             =   8085
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Escrutinio"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "mml_FRASE0033"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM bailes"
      Caption         =   "1"
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
   Begin VB.Frame Frame1 
      Height          =   7920
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12465
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgSelec 
         Bindings        =   "frmSelec.frx":0496
         Height          =   7590
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   13388
         _Version        =   393216
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLineWidthBand=   1
      End
   End
   Begin VB.Label Label1 
      Caption         =   "mml_FRASE0835"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8310
      TabIndex        =   2
      Top             =   8055
      Width           =   1485
   End
End
Attribute VB_Name = "frmSelec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function sSeleccionar(sSql As String)
    adoSelec.ConnectionString = "DSN=Escrutinio"
    adoSelec.RecordSource = sSql
    adoSelec.Refresh
    dgSelec.Refresh
    cbOrden.ListIndex = Val(VarCfg("orden_sel")) - 1
    Me.Show 1
End Function


Private Sub cmdBuscar_Click()
    ActualizarGrid

End Sub

Private Sub cmdCancel_Click()
Dim i As Integer
    For i = 0 To 50
        sResultado(i) = ""
    Next
    Unload Me
End Sub

Private Sub cmdOrden_Click(Index As Integer)

    If cbOrden.ListCount >= Index + 1 Then
        cbOrden.ListIndex = Index
    End If
    
    ActualizarGrid
    
End Sub
Sub ActualizarGrid()
Dim sBuscar As String
Dim sSql As String

    If m_sSelec <> "" And tbBuscar.Text <> "" Then
        sBuscar = " AND " & m_sSelec & " LIKE '%" & tbBuscar.Text & "%' "
    End If
    sSql = sSelecSQL & sBuscar & m_sGroupBy
    If InStr(sSql, "ORDER BY") > 0 Then
        sSql = Mid$(sSql, 1, InStr(sSql, "ORDER BY") - 1)
    End If
    If InStr(UCase(adoSelec.RecordSource), "DISTINCT FASE") > 0 Then
        adoSelec.RecordSource = sSql & " ORDER BY 1"
        cbOrden.Visible = False
    Else
        adoSelec.RecordSource = sSql & " ORDER BY " & cbOrden.ListIndex + 1
    End If
    adoSelec.Refresh
    dgSelec.Refresh

End Sub

Private Sub cmdSelCateg_Click()
    If dgSelec.Rows > 0 Then
        dgSelec.Row = dgSelec.Rows - 1
        dgSelec.TopRow = dgSelec.Row
        dgSelec.Row = dgSelec.Rows - 1
        dgSelec.RowSel = dgSelec.Rows - 1
        dgSelec.ColSel = 0
    End If
End Sub

Private Sub cmdSeleccionar_Click()
Dim i As Integer
    For i = 0 To dgSelec.Cols - 1
        dgSelec.Col = i
        If dgSelec.Row = 0 Then
            sResultado(i) = ""
            sResultado(1) = ""
            sResultado(2) = ""
            Exit For
        Else
            sResultado(i + 1) = dgSelec.Text
        End If
    Next
    AsignarParametro "orden_sel", cbOrden.ListIndex + 1
    Unload Me
End Sub

Private Sub dgSelec_DblClick()
    Call cmdSeleccionar_Click
End Sub

Private Sub Form_Activate()
    On Local Error Resume Next
    If tbBuscar.Visible Then tbBuscar.SetFocus
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    dgSelec.ColWidth(1) = 4000
    
    If m_sSelec = "" Then
        tbBuscar.Visible = False
        cmdBuscar.Visible = False
    Else
        tbBuscar.Visible = True
        cmdBuscar.Visible = True
    End If
End Sub

Private Sub tbBuscar_GotFocus()
    tbBuscar.SelStart = 0
    tbBuscar.SelLength = Len(tbBuscar.Text)

End Sub
