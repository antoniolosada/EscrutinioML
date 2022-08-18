VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPaneles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1094"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGeneraPaneles 
      Caption         =   "mml_FRASE1204"
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
      Left            =   9090
      TabIndex        =   21
      Top             =   8250
      Width           =   3105
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
      Height          =   465
      Left            =   9090
      TabIndex        =   20
      Top             =   8640
      Width           =   3105
   End
   Begin VB.CommandButton cmdGenCategPanel 
      Caption         =   "mml_FRASE1202"
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
      Left            =   9090
      TabIndex        =   19
      Top             =   7050
      Width           =   3105
   End
   Begin VB.CommandButton cmdAsignarPanel 
      Caption         =   "mml_FRASE1203"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9090
      TabIndex        =   18
      Top             =   7560
      Width           =   3105
   End
   Begin VB.Frame Frame4 
      Height          =   1125
      Left            =   8250
      TabIndex        =   13
      Top             =   60
      Width           =   4455
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
         Left            =   2190
         Picture         =   "frmPaneles.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox tbCodComp 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Height          =   375
         Left            =   2670
         TabIndex        =   15
         Top             =   240
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
         Left            =   60
         TabIndex        =   14
         Top             =   690
         Width           =   4320
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
         Left            =   510
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5775
      Left            =   8250
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
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
         Left            =   3330
         Picture         =   "frmPaneles.frx":046A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   780
         Width           =   315
      End
      Begin VB.CheckBox chkPasos 
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   1020
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
         ItemData        =   "frmPaneles.frx":08D4
         Left            =   2520
         List            =   "frmPaneles.frx":0A5E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   780
         Width           =   795
      End
      Begin VB.Frame Frame2 
         Caption         =   "mml_FRASE0050"
         Height          =   4125
         Left            =   150
         TabIndex        =   6
         Top             =   1410
         Width           =   4095
         Begin MSDataGridLib.DataGrid dgJueces 
            Bindings        =   "frmPaneles.frx":0C50
            Height          =   3795
            Left            =   120
            TabIndex        =   7
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
         Left            =   120
         TabIndex        =   5
         Top             =   780
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
         Height          =   495
         Left            =   1230
         TabIndex        =   4
         Top             =   780
         Width           =   1245
      End
      Begin VB.ComboBox cbPanel 
         BackColor       =   &H00C0FFFF&
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
         ItemData        =   "frmPaneles.frx":0C68
         Left            =   3060
         List            =   "frmPaneles.frx":0C9C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   795
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
         Left            =   3660
         TabIndex        =   12
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "mml_FRASE1094"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   450
         TabIndex        =   11
         Top             =   270
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0032"
      Height          =   9105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8115
      Begin MSDataGridLib.DataGrid dgCat 
         Bindings        =   "frmPaneles.frx":0CD6
         Height          =   8805
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   15531
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
   Begin MSAdodcLib.Adodc adoCat 
      Height          =   495
      Left            =   11760
      Top             =   720
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
      RecordSource    =   "SELECT * FROM categorias c, paneles p WHERE c.codigo = p.cod_categoria ORDER BY descripcion"
      Caption         =   "mml_FRASE0032"
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
      Left            =   11760
      Top             =   180
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
      RecordSource    =   $"frmPaneles.frx":0CEB
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
End
Attribute VB_Name = "frmPaneles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbPanel_Click()

    If cbPanel.Text <> "" And Val(tbCodComp.Text) <> 0 Then
        CargarPanel cbPanel.Text, tbCodComp.Text
    End If
End Sub

Private Sub cmdAsignarPanel_Click()
Dim iCateg
Dim sCodCat As String

    If Not C_DEBUG Then On Local Error GoTo error
    If cbPanel.Text = "" Or Val(tbCodComp.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
    For Each iCateg In dgCat.SelBookmarks
        dgCat.Col = 0
        dgCat.Bookmark = iCateg
        sCodCat = Val(dgCat.Text)
        
        db.Execute "UPDATE paneles SET cod_panel = '" & cbPanel.Text & "' WHERE cod_categoria = " & sCodCat
    Next
    db.Execute "UPDATE paneles SET cod_panel = '" & cbPanel.Text & "' WHERE cod_categoria = " & sCodCat
    
    CargarCateg Val(tbCodComp.Text)
    
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
        
    Exit Sub
error:
    ProcesarError "cmdAsignarPanel_Click"
End Sub

Private Sub cmdGenCategPanel_Click()
Dim rs As Recordset

    If Not C_DEBUG Then On Local Error GoTo error
    If Val(tbCodComp.Text) = 0 Or cbPanel.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE1201, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If
    
    db.Execute "DELETE FROM paneles WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ")"
    Set rs = db.OpenRecordset("SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    While Not rs.EOF
        db.Execute ("INSERT INTO paneles VALUES ('" & cbPanel.Text & "'," & rs.Fields("codigo") & ")")
        rs.MoveNext
    Wend
    
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    
    CargarPanel cbPanel.Text, Val(tbCodComp.Text)
    CargarCateg Val(tbCodComp.Text)
        
    Exit Sub
error:
    ProcesarError "cmdGenCategPanel_Click"

End Sub

Private Sub cmdGeneraPaneles_Click()
Dim rs As Recordset
Dim rsPanel As Recordset

    If Not C_DEBUG Then On Local Error GoTo error
    If Val(tbCodComp.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE1205, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    While Not rs.EOF
        ' Solo asignamos los paneles a las categorias que no tienen puntuaciones introducidas
        If Not TienePuntuaciones(rs.Fields("codigo")) Then
            Set rsPanel = db.OpenRecordset("SELECT * FROM juez_panel jp, paneles p WHERE jp.cod_competicion = " & tbCodComp.Text & " AND p.cod_panel = jp.cod_panel AND p.cod_categoria = " & rs.Fields("codigo"), dbOpenSnapshot)
            If Not rsPanel.EOF Then
                'Borramos el antiguo panel
                db.Execute "DELETE FROM juez_categ WHERE cod_categoria = " & rs.Fields("codigo")
                While Not rsPanel.EOF
                    db.Execute "INSERT INTO juez_categ VALUES (" & rsPanel.Fields("cod_juez") & "," & rs.Fields("codigo") & ",'" & rsPanel.Fields("id_juez") & "'," & rsPanel.Fields("pasos") & ")"
                    rsPanel.MoveNext
                Wend
            End If
            rsPanel.Close
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    
    Exit Sub
error:
    ProcesarError "cmdGeneraPaneles_Click"

End Sub

Private Sub cmdPonerJuez_Click()
Dim iCodJuez As Integer
Dim sIdJuez As String
Dim i As Integer
Dim rs As Recordset
    
    If Not C_DEBUG Then On Local Error GoTo error
    If cbPanel.Text = "" Or cbIdJuez.Text = "" Or Val(tbCodComp.Text) = 0 Then
        MsgBox mml_FRASE1200, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    'Comprobamos si esta letra ya ha sido asignada en esta competición
    iCodJuez = 0
    If G_ASIGNAR_AUTOMATICAMENTE_LETRA_A_JUEZ Then
        'Buscamos en categorias
        Set rs = db.OpenRecordset("SELECT TOP 1 cod_juez FROM juez_categ WHERE id_juez = '" & cbIdJuez.Text & "' AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion =  " & tbCodComp.Text & ")", dbOpenSnapshot)
        If Not rs.EOF Then
            iCodJuez = rs.Fields("cod_juez")
        End If
        rs.Close
        If iCodJuez = 0 Then
            'Buscamos en paneles
            Set rs = db.OpenRecordset("SELECT TOP 1 cod_juez FROM juez_panel jp WHERE id_juez = '" & cbIdJuez.Text & "' AND jp.cod_competicion =  " & tbCodComp.Text, dbOpenSnapshot)
            If Not rs.EOF Then
                iCodJuez = rs.Fields("cod_juez")
            End If
            rs.Close
        End If
    End If
    If iCodJuez = 0 Then
        iCodJuez = Val(sSeleccionar("SELECT * FROM jueces j WHERE codigo = (SELECT MAX(CODIGO) FROM jueces j1 WHERE j1.nombre = j.nombre)"))
    End If
    
    If iCodJuez > 0 Then
        If Not ExisteCodJuezPanel(Val(tbCodComp.Text), cbPanel.Text, cbIdJuez.Text) Then
            db.Execute ("INSERT INTO juez_panel VALUES(" & tbCodComp.Text & "," & iCodJuez & ",'" & cbPanel.Text & "','" & cbIdJuez.Text & "'," & IIf(chkPasos.Value <> False, 1, 0) & ")")
        End If
        cbIdJuez.ListIndex = cbIdJuez.ListIndex + 1
        DoEvents
        CargarPanel cbPanel.Text, Val(tbCodComp.Text)
        MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    End If

    Exit Sub
error:
    ProcesarError "cmdPonerJuez_Click"
End Sub

Private Sub cmdQuitarJuez_Click()
Dim iCodJuez As Integer
Dim sCodJuez As String
Dim rs As Recordset

    If Not C_DEBUG Then On Local Error GoTo error
    
    dgJueces.Col = 0
    If dgJueces.Row < 0 Then
        Exit Sub
    End If
    
    If dgJueces.Text = "" Or cbPanel.Text = "" Or Val(tbCodComp.Text) = 0 Then
        MsgBox mml_FRASE1200, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    dgJueces.Col = 0
    iCodJuez = dgJueces.Text
    dgJueces.Col = 1
    sCodJuez = dgJueces.Text
    
    db.Execute ("DELETE FROM juez_panel WHERE cod_juez =" & iCodJuez & " AND cod_competicion = " & tbCodComp.Text & " AND cod_panel = '" & cbPanel.Text & "'")
    
    CargarPanel cbPanel.Text, Val(tbCodComp.Text)
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    
    Exit Sub
error:
    ProcesarError "cmdQuitarJuez_Click"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    
    CargarPanel "0", Val(tbCodComp.Text)
    CargarCateg Val(tbCodComp.Text)
End Sub

Private Sub cmdSubirDatos_Click()
    cbIdJuez.ListIndex = 0
End Sub

Private Sub Form_Load()
Dim i As Integer
    TraducirCadenas Me
    
    tbCodComp.Text = CodCompActiva
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    cbPanel.Clear
    For i = 1 To 9
        cbPanel.AddItem i
    Next
    For i = Asc("A") To Asc("Z")
        cbPanel.AddItem Chr$(i)
    Next
    
    CargarCateg Val(tbCodComp.Text)
End Sub


Function ExisteCodJuezPanel(iCodcomp As Integer, sCodPanel As String, sCodJuez As String) As Boolean
Dim rsJueces As Recordset

    Set rsJueces = db.OpenRecordset("SELECT COUNT(*) FROM juez_panel WHERE cod_panel = '" & sCodPanel & _
                                      "' AND id_juez = '" & sCodJuez & "' AND cod_competicion = " & iCodcomp, dbOpenSnapshot)
    If rsJueces.Fields(0) > 0 Then
        ExisteCodJuezPanel = True
    Else
        ExisteCodJuezPanel = False
    End If
    rsJueces.Close
    
End Function

Sub CargarCateg(iCodcomp As Integer)
    dgCat.Visible = False
    adoCat.ConnectionString = "DSN=Escrutinio"
    adoCat.RecordSource = "SELECT c.codigo, c.descripcion, id_categoria, hora, cod_panel FROM categorias c, paneles p WHERE c.codigo = p.cod_categoria AND cod_competicion = " & tbCodComp.Text & " ORDER BY descripcion"
    adoCat.Refresh
    DoEvents

    dgCat.Columns(0).Width = 800
    dgCat.Columns(1).Width = 5400
    dgCat.Columns(2).Width = 800
    dgCat.Columns(3).Width = 800
    dgCat.Columns(4).Width = 1000
    dgCat.Refresh
    dgCat.Visible = True
    adoCat.Refresh
    dgCat.Refresh
End Sub

Sub CargarPanel(sPanel As String, iCodcomp As Integer)
    dgCat.Visible = False
    adoJueces.ConnectionString = "DSN=Escrutinio"
    adoJueces.RecordSource = "SELECT codigo, id_juez,jp.cod_panel, nombre, pasos  FROM juez_panel jp, jueces j WHERE jp.cod_competicion = " & iCodcomp & " AND jp.cod_juez = j.codigo AND jp.cod_panel = '" & sPanel & "' ORDER BY 2"
    adoJueces.Refresh
    DoEvents
    
    dgJueces.Columns(0).Width = 400
    dgJueces.Columns(1).Width = 500
    dgJueces.Columns(2).Width = 500
    dgJueces.Columns(3).Width = 1500
    dgJueces.Columns(4).Width = 400
    
    dgCat.Refresh
    dgCat.Visible = True

End Sub

