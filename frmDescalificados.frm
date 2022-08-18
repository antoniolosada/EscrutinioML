VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDescalificados 
   Caption         =   "mml_FRASE0043"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8010
   ScaleWidth      =   9810
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
      Left            =   8055
      TabIndex        =   27
      Top             =   7425
      Width           =   1695
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
      Left            =   2070
      Picture         =   "frmDescalificados.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   135
      Width           =   495
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
      Height          =   405
      Left            =   2070
      Picture         =   "frmDescalificados.frx":046A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   615
      Width           =   495
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
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   1080
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   120
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pbDesc 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   9555
      TabIndex        =   22
      Top             =   5160
      Width           =   9615
   End
   Begin VB.CommandButton cmdImpDesc 
      Caption         =   "mml_FRASE0550"
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
      TabIndex        =   21
      Top             =   7440
      Width           =   3495
   End
   Begin VB.CommandButton cmdMotivo 
      Caption         =   "mml_FRASE0551"
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
      Left            =   3720
      TabIndex        =   20
      Top             =   7440
      Width           =   4215
   End
   Begin VB.TextBox tbMotivo 
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
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   2280
      Width           =   9615
   End
   Begin VB.ComboBox cbBailes 
      BackColor       =   &H00C0FFFF&
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
      ItemData        =   "frmDescalificados.frx":08D4
      Left            =   5520
      List            =   "frmDescalificados.frx":08F0
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1680
      Width           =   4215
   End
   Begin VB.ComboBox cbDorsal 
      BackColor       =   &H00C0FFFF&
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
      ItemData        =   "frmDescalificados.frx":096F
      Left            =   6960
      List            =   "frmDescalificados.frx":0979
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cbJuez 
      BackColor       =   &H00C0FFFF&
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
      ItemData        =   "frmDescalificados.frx":0999
      Left            =   4920
      List            =   "frmDescalificados.frx":099B
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1080
      Width           =   975
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
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cbFase 
      BackColor       =   &H00C0FFFF&
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
      ItemData        =   "frmDescalificados.frx":099D
      Left            =   1080
      List            =   "frmDescalificados.frx":09B9
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Frame mrcDesc 
      Caption         =   "mml_FRASE0043"
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   9615
      Begin MSDataGridLib.DataGrid dgDesc 
         Bindings        =   "frmDescalificados.frx":0A38
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2990
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
      Left            =   960
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
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
      TabIndex        =   4
      Top             =   600
      Width           =   6375
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
      TabIndex        =   3
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin MSAdodcLib.Adodc adoDesc 
      Height          =   495
      Left            =   120
      Top             =   5280
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
      RecordSource    =   "SELECT d.codigo, num_dorsal,id_juez,b.nombre,fase,cod_categoria FROM descalificaciones d, bailes b where d.cod_baile = b.codigo"
      Caption         =   "mml_FRASE0552"
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
   Begin VB.Label Label6 
      Caption         =   "mml_FRASE0553"
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
      TabIndex        =   24
      Top             =   2010
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "mml_FRASE0300"
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
      Left            =   5940
      TabIndex        =   16
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "mml_FRASE0421"
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
      Left            =   4080
      TabIndex        =   14
      Top             =   1080
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
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   825
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
      TabIndex        =   5
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
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmDescalificados"
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

Private Sub chkPasos_Click()

End Sub

Public Sub cmdActualizar_Click()
Dim rs As Recordset
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    Sleep 500
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    adoDesc.ConnectionString = "DSN=Escrutinio"
    adoDesc.RecordSource = "SELECT d.codigo, num_dorsal,id_juez,b.nombre,fase,cod_categoria, motivo, repesca FROM descalificaciones d, bailes b where d.cod_baile = b.codigo AND cod_categoria LIKE '" & tbCodCateg.Text & "%' AND repesca=" & chkRep.Value & " AND fase = " & 2 ^ cbFase.ListIndex & " ORDER BY 1,3,2"
    adoDesc.Refresh
    
    If tbCodCateg.Text <> "" Then
        cbBailes.Clear
        Set rs = db.OpenRecordset("SELECT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & IIf(cbFase.ListIndex = 0, 1, 2) & " ORDER BY posicion", dbOpenSnapshot)
        While Not rs.EOF
            cbBailes.AddItem rs!codigo & " " & rs!Nombre
            rs.MoveNext
        Wend
        rs.Close
        cbJuez.Clear
        Set rs = db.OpenRecordset("SELECT id_juez FROM juez_categ WHERE cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rs.EOF
            cbJuez.AddItem rs!id_juez
            rs.MoveNext
        Wend
        rs.Close
        cbDorsal.Clear
        Set rs = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase=" & 2 ^ (cbFase.ListIndex) & " ORDER BY 1", dbOpenSnapshot)
        While Not rs.EOF
            cbDorsal.AddItem rs!num_dorsal
            rs.MoveNext
        Wend
        rs.Close
    End If
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



Private Sub cmdImpDesc_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim bLetraJuez As Boolean

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE0554, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If VarCfg("tipo_hoja_puntuaciones") = "hoja_rec_optico" Then
        If MsgBox(mml_FRASE0555, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            bLetraJuez = False
        Else
            bLetraJuez = True
        End If
    Else
        bLetraJuez = False
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection Or cdlPDNoPageNums
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    If Not bLetraJuez Then
        ImpDescNormales
    Else
        ImpDescGraficas
    End If
End Sub

Sub ImprimirCabeceraDesc()
Dim rs As Recordset
Dim iEscala As Integer, X As Integer, Y As Integer

    Printer.PaintPicture frmMenu.picEPA.Picture, Printer.Width - frmMenu.picEPA.Width - G_MARGEN_EPA, G_MARGEN_EPA_Y
    iEscala = Printer.Width / 10
    Printer.FontSize = 13
    Printer.Print
    Printer.CurrentX = 0
    Set rs = db.OpenRecordset(" SELECT descripcion, fecha FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    Printer.DrawWidth = 2
    X = Printer.CurrentX
    Y = Printer.CurrentY
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    SaltoLinea Printer, 4
    Printer.FontBold = True
    Centrado Printer, rs!DESCRIPCION & "  (" & rs!fecha & ")", Printer.Width
    rs.Close
    Printer.FontBold = False
    Printer.FontSize = 13
    Centrado Printer, sEscuela(tbCodComp.Text), Printer.Width
    Printer.FontBold = True
    Centrado Printer, mml_FRASE0556, Printer.Width
    Printer.FontBold = False
    SaltoLinea Printer, 4
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    Printer.DrawWidth = 1
    SaltoLinea Printer, 3
    Printer.FontSize = 10

End Sub

Private Sub ImpDescNormales()
Dim iPag As Integer
Dim iEscala As Long
Dim rs As Recordset, rsPart As Recordset

    iEscala = Printer.Width / 12
    
    iPag = 1
    ImprimirCabeceraDesc
    Printer.Print mml_FRASE0557 & iPag
        
    Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion, c.codigo FROM categorias c, gruposedad ge WHERE cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY hora", dbOpenSnapshot)
    While Not rs.EOF
    
        Set rsPart = db.OpenRecordset("SELECT d.num_dorsal, d.fase, d.repesca, d.id_juez, b.nombre, d.motivo , jc.pasos FROM bailes b, descalificaciones d, juez_categ jc  WHERE  d.id_juez = jc.id_juez AND jc.cod_categoria = d.cod_categoria AND d.cod_baile = b.codigo AND d.cod_categoria = " & rs!codigo & " ORDER BY 2,1", dbOpenSnapshot)
        If Not rsPart.EOF Then
            
            Printer.FontBold = True
            
            Printer.FontSize = 12
            Printer.CurrentX = iEscala
            Printer.Print rs!codigo;
            Printer.CurrentX = iEscala * 2
            Printer.Print rs!DESCRIPCION
            Printer.Line Step(iEscala, 0)-Step(Printer.Width - iEscala, 0)
            
            Printer.FontSize = 10
            
            Printer.CurrentX = iEscala
            Printer.Print mml_FRASE0300;
            Printer.CurrentX = iEscala * 2
            Printer.Print mml_FRASE0299;
            Printer.CurrentX = iEscala * 3
            Printer.Print mml_FRASE0421;
            Printer.CurrentX = iEscala * 4
            Printer.Print mml_FRASE0436;
            Printer.CurrentX = iEscala * 5
            Printer.Print mml_FRASE0553
            Printer.FontBold = False
            While Not rsPart.EOF
                Printer.FontSize = 10
                Printer.CurrentX = iEscala
                Printer.Print rsPart!num_dorsal;
                Printer.CurrentX = iEscala * 2
                
                Select Case rsPart!fase
                    Case 1:
                        Printer.Print mml_FRASE0329;
                    Case 2:
                        Printer.Print mml_FRASE0330;
                    Case Else
                        Printer.Print "1/" & Trim$(Str$(rsPart!fase));
                End Select
                
                If rsPart!repesca = 1 Then
                    Printer.Print mml_FRASE0558;
                End If
                
                Printer.CurrentX = iEscala * 3
                Printer.Print rsPart!id_juez;
                If rsPart!pasos = 1 Then
                    Printer.Print mml_FRASE0373;
                End If
                Printer.CurrentX = iEscala * 4
                Printer.Print rsPart!Nombre;
                Printer.CurrentX = iEscala * 5
                Printer.FontSize = 8
                Printer.Print IIf(IsNull(rsPart!motivo), "", rsPart!motivo)
                Printer.Line Step(iEscala, 0)-Step(Printer.Width - iEscala, 0)
                Printer.FontSize = 2
                Printer.Print
                If Printer.CurrentY + MARGEN_PAGINA_INF > Printer.Height Then
                    Printer.NewPage
                    Inc iPag
                    ImprimirCabeceraDesc
                    Printer.Print mml_FRASE0557 & iPag
                End If
                rsPart.MoveNext
            Wend
            rsPart.Close
        End If
        rs.MoveNext
    Wend
    rs.Close
        
        
    Printer.EndDoc

End Sub
Private Sub ImpDescGraficas()
Dim iPag As Integer
Dim iEscala As Long
Dim iBaile As Integer
Dim rs As Recordset
Dim rsPart As Recordset
    ' Solo imprime las de los jueces de pasos y figuras
    iEscala = Printer.Width / 12
    
    iPag = 1
    ImprimirCabeceraDesc
    Printer.Print mml_FRASE0557 & iPag
        
    iBaile = -1
    Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion, c.codigo FROM categorias c, gruposedad ge WHERE cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY hora", dbOpenSnapshot)
    While Not rs.EOF
    
        Debug.Print "SELECT d.num_dorsal, d.fase, d.id_juez, b.nombre, d.motivo, jc.pasos FROM bailes b, descalificaciones d, juez_categ jc WHERE d.juez = jc.id_juez AND jc.cod_categoria = d.cod_categoria AND d.cod_baile = b.codigo AND ((b.fase=1 AND d.fase=1) OR (b.fase=2 AND d.fase>1) ) AND d.cod_categoria = " & rs!codigo & " ORDER BY 1"
        Set rsPart = db.OpenRecordset("SELECT d.num_dorsal, d.fase, d.repesca, d.id_juez, b.nombre, b.codigo, d.motivo , jc.pasos, d.anotacion FROM bailes b, descalificaciones d, juez_categ jc  WHERE  jc.pasos = 1 AND d.id_juez = jc.id_juez AND jc.cod_categoria = d.cod_categoria AND d.cod_baile = b.codigo AND d.cod_categoria = " & rs!codigo & " ORDER BY 1", dbOpenSnapshot)
        If Not rsPart.EOF Then
            Printer.Print
            
            Printer.FontBold = True
            
            Printer.FontSize = 12
            Printer.CurrentX = iEscala
            Printer.Print rs!codigo;
            Printer.CurrentX = iEscala * 2
            Printer.Print rs!DESCRIPCION
            Printer.Line Step(iEscala, 0)-Step(Printer.Width - iEscala, 0)
            
            Printer.FontSize = 10
            
            ' Imprimimos la descalificación
            Printer.FontSize = 10
            Printer.CurrentX = iEscala
            Printer.Print mml_FRASE0300;
            Printer.CurrentX = iEscala * 2
            Printer.Print mml_FRASE0299;
            Printer.CurrentX = iEscala * 3
            Printer.Print mml_FRASE0421;
            Printer.CurrentX = iEscala * 4
            Printer.Print mml_FRASE0436;
            Printer.CurrentX = iEscala * 5
            Printer.Print mml_FRASE0553
            Printer.FontBold = False
            While Not rsPart.EOF
                Printer.FontSize = 10
                Printer.CurrentX = iEscala
                Printer.Print rsPart!num_dorsal;
                Printer.CurrentX = iEscala * 2
                
                Select Case rsPart!fase
                    Case 1:
                        Printer.Print mml_FRASE0329;
                    Case 2:
                        Printer.Print mml_FRASE0330;
                    Case Else
                        Printer.Print "1/" & Trim$(Str$(rsPart!fase));
                End Select
                
                If rsPart!repesca = 1 Then
                    Printer.Print mml_FRASE0558;
                End If
                            
                Printer.CurrentX = iEscala * 3
                Printer.Print rsPart!id_juez;
                If rsPart!pasos = 1 Then
                    Printer.Print mml_FRASE0373;
                End If
                Printer.CurrentX = iEscala * 4
                Printer.Print rsPart!Nombre;
                Printer.CurrentX = iEscala * 5
                Printer.FontSize = 8
                Printer.Print IIf(IsNull(rsPart!motivo), "", rsPart!motivo)
                Printer.Line Step(iEscala, 0)-Step(Printer.Width - iEscala, 0)
                Printer.FontSize = 2
                Printer.Print
                
                If Not IsNull(rsPart!anotacion) Then
                    On Error GoTo seguir
                    LeerBinary rsPart!anotacion, pbDesc
                    Printer.PaintPicture pbDesc, Printer.CurrentX + iEscala, Printer.CurrentY
seguir:
                    On Error GoTo 0
                End If
                Printer.CurrentY = Printer.CurrentY + pbDesc.Height
                If Printer.CurrentY + MARGEN_PAGINA_INF + pbDesc.Height > Printer.Height Then
                    Inc iPag
                    Printer.NewPage
                    ImprimirCabeceraDesc
                    Printer.Print mml_FRASE0557 & iPag
                End If
                rsPart.MoveNext
            Wend
            rsPart.Close
        End If
        rs.MoveNext
    Wend
    rs.Close
        
        
    Printer.EndDoc

End Sub

Private Sub cmdMotivo_Click()
Dim iCodDesc As Integer
    If Not C_DEBUG Then On Error GoTo error
    dgDesc.Col = 0
    iCodDesc = Val(dgDesc.Text)
    If tbMotivo.Text = "" Or iCodDesc = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    db.Execute ("UPDATE descalificaciones SET motivo = '" & tbMotivo.Text & "' WHERE codigo = " & iCodDesc)
error:
    ProcesarError
End Sub

Private Sub cmdPonerJuez_Click()
Dim iCodJuez As Integer
Dim sIdJuez As String
Dim i As Integer
    
    If tbMotivo.Text = "" Or tbCodCateg.Text = "" Or cbFase.ListIndex = -1 Or cbJuez.ListIndex = -1 Or cbJuez.ListIndex = -1 Or cbBailes.ListIndex = -1 Then
        CamposSinCubrir
        Exit Sub
    End If
    db.Execute ("INSERT INTO descalificaciones VALUES(" & MaxCod("descalificaciones") & ",'" & tbCodCateg.Text & "','" & 2 ^ cbFase.ListIndex & "','" & cbJuez.List(cbJuez.ListIndex) & "'," & Mid$(cbBailes.List(cbBailes.ListIndex), 1, 2) & "," & cbDorsal.List(cbDorsal.ListIndex) & ",'" & tbMotivo.Text & "',''," & chkRep.Value & ")")
    Call cmdActualizar_Click

End Sub



Private Sub CommandButton1_Click()

End Sub

Private Sub cmdQuitarJuez_Click()
Dim iCodDesc As Integer
    
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    dgDesc.Col = 0
    iCodDesc = dgDesc.Text
    
    db.Execute ("DELETE FROM descalificaciones WHERE codigo =" & iCodDesc)
    Call cmdActualizar_Click

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCateg.Text = ""
    tbDescCateg.Text = ""
    Call cmdActualizar_Click
End Sub

Private Sub dgBailes_Click()

End Sub

Private Sub dgDesc_Click()
Dim rs As Recordset
Dim iCodigo As Integer

    On Local Error GoTo salir
    dgDesc.Col = 0
    iCodigo = Val(dgDesc.Text)
    Set rs = db.OpenRecordset("SELECT anotacion FROM descalificaciones WHERE codigo = " & iCodigo, dbOpenSnapshot)
    If Not rs.EOF Then
    On Local Error GoTo seguir
        If Not IsNull(rs!anotacion) Then
            LeerBinary rs!anotacion, pbDesc
        End If
    End If
seguir:
    rs.Close
salir:
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    cbFase.ListIndex = 0
    If VarCfg("tipo_hoja_puntuaciones") <> "hoja_rec_optico" Then
        pbDesc.Visible = False
    End If
End Sub

