VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPublicar 
   Caption         =   "mml_FRASE0895"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "mml_FRASE0896"
      Height          =   3795
      Left            =   30
      TabIndex        =   10
      Top             =   3480
      Width           =   10095
      Begin MSDataGridLib.DataGrid dgPublicar 
         Bindings        =   "frmPublicar.frx":0000
         Height          =   3465
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6112
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0897"
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
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
         Height          =   585
         Left            =   9090
         TabIndex        =   26
         Top             =   2790
         Width           =   810
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "mml_FRASE0295"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   7485
         TabIndex        =   25
         Top             =   2790
         Width           =   1560
      End
      Begin VB.CommandButton cmdComentario 
         Caption         =   "mml_FRASE0898"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5460
         TabIndex        =   24
         Top             =   2790
         Width           =   1995
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "mml_FRASE0186"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   270
         TabIndex        =   23
         Top             =   2790
         Width           =   1575
      End
      Begin VB.CommandButton cmdBorrarUna 
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
         Height          =   585
         Left            =   3900
         TabIndex        =   22
         Top             =   2790
         Width           =   1560
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "mml_FRASE0899"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1860
         TabIndex        =   21
         Top             =   2790
         Width           =   2040
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
         Picture         =   "frmPublicar.frx":001A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   450
      End
      Begin VB.CommandButton CommandButton1 
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
         Picture         =   "frmPublicar.frx":0484
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   855
         Width           =   450
      End
      Begin VB.CommandButton CommandButton2 
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
         Picture         =   "frmPublicar.frx":08EE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1305
         Width           =   450
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
         Left            =   6000
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cbRep 
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
         ItemData        =   "frmPublicar.frx":0D58
         Left            =   7320
         List            =   "frmPublicar.frx":0D65
         TabIndex        =   16
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox tbDesc 
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
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   1800
         Width           =   7935
      End
      Begin VB.TextBox tbComen 
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
         Left            =   1800
         TabIndex        =   12
         Top             =   2280
         Width           =   7935
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   360
         Width           =   6615
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   840
         Width           =   6615
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0607"
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
         TabIndex        =   13
         Top             =   2280
         Width           =   1575
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc adoPublicar 
      Height          =   495
      Left            =   3120
      Top             =   5640
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Escrutinio"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "mml_FRASE0033"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Publicar"
      Caption         =   "mml_FRASE0186"
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
Attribute VB_Name = "frmPublicar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Publicar(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sDesc As String, sRep As Integer, sDescCat As String, sDescComp As String)
    tbCodComp.Text = sCodComp
    tbCodCat.Text = sCodCat
    tbCodFase.Text = sCodFase
    
    tbDescComp.Text = sDescComp
    tbDescCat.Text = sDescCat
    Select Case tbCodFase.Text
        Case 1:
            tbDescFase.Text = mml_FRASE0329
        Case 2:
            tbDescFase.Text = "SEMI-FINAL"
        Case 4:
            tbDescFase.Text = "CUARTOS DE FINAL"
        Case 8:
            tbDescFase.Text = "OCTAVOS DE FINAL"
        Case Else
            tbDescFase.Text = tbCodFase.Text & "OS DE FINAL"
    End Select
    tbDesc_DblClick
    'If sRep = 1 Then
    '    cbRep.ListIndex = 1
    '    chkRep.Value = 1
    'End If
    Me.Show 1
End Sub
Public Sub PublicacionCompleta(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sDesc As String, sRep As Integer, sDescCat As String, sDescComp As String)
    tbCodComp.Text = sCodComp
    tbCodCat.Text = sCodCat
    tbCodFase.Text = sCodFase
    
    tbDescComp.Text = sDescComp
    tbDescCat.Text = sDescCat
    Select Case tbCodFase.Text
        Case 1:
            tbDescFase.Text = mml_FRASE0329
        Case 2:
            tbDescFase.Text = "SEMI-FINAL"
        Case 4:
            tbDescFase.Text = "CUARTOS DE FINAL"
        Case 8:
            tbDescFase.Text = "OCTAVOS DE FINAL"
        Case Else
            tbDescFase.Text = tbCodFase.Text & "OS DE FINAL"
    End Select
    tbDesc_DblClick
    cmdCalcular_Click

End Sub


Private Sub cbRep_Click()
    tbComen.Text = "  " & cbRep.List(cbRep.ListIndex)
    tbDesc_DblClick
End Sub

Private Sub cbRep_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub chkRep_Click()
    If chkRep.Value = 1 Then
        MsgBox mml_FRASE0900, vbOKOnly Or vbInformation, mml_FRASE0084
        cbRep.ListIndex = 2
        cbRep.Enabled = False
    Else
        cbRep.Text = ""
        cbRep.Enabled = True
    End If
    tbDesc_DblClick
End Sub

Private Sub cmdAct_Click()
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    Sleep 500
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    adoPublicar.Refresh
    dgPublicar.Refresh
End Sub

Private Sub cmdBorrar_Click()
    If MsgBox(mml_FRASE0901, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        db.Execute ("DELETE FROM publicar")
        adoPublicar.Refresh
        dgPublicar.Refresh
    End If
End Sub

Private Sub cmdBorrarUna_Click()
Dim sCodComp As String, sCodCateg As String, sCodFase As String, sCodRep As String
    If Not C_DEBUG Then On Local Error GoTo error
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
        If dgPublicar.Row >= 0 Then
            With dgPublicar
                .Col = 0
                sCodComp = .Text
                .Col = 1
                sCodCateg = .Text
                .Col = 2
                sCodFase = .Text
                .Col = 6
                sCodRep = .Text
                
                db.Execute "DELETE FROM publicar WHERE cod_competicion = " & sCodComp & " AND cod_categoria = " & sCodCateg & " AND fase = " & sCodFase & " AND repesca = " & sCodRep
                
                cmdAct_Click
                MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
            End With
        End If
    End If
    Exit Sub
error:
    ProcesarError "cmdBorrarUna_Click"
End Sub

Private Sub cmdCalcular_Click()
Dim rs As Recordset
    On Local Error Resume Next
    If tbDesc.Text = "" Then
        tbDesc.Text = tbDescCat.Text & "-" & tbDescFase.Text & cbRep.Text
    End If
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Or tbDesc.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    ' Primero debemos publicar la fase inicial y después la repesca
    ' Solo hay publicada la última fase de cada grupo
    ' Se borran todas las fases anteriores por si hubiera una inicial y una repesca
    db.Execute ("DELETE FROM publicar WHERE cod_competicion = " & tbCodComp.Text & " AND cod_categoria =" & tbCodCat.Text & " AND fase >=" & tbCodFase.Text)
    db.Execute ("INSERT INTO publicar VALUES(" & tbCodComp.Text & "," & tbCodCat.Text & "," & tbCodFase.Text & ",'" & tbDesc.Text & "','" & tbComen & "','" & Str$(Now) & "'," & chkRep.Value & ")")
    db.Execute "UPDATE cfg SET valor = 'S' WHERE variable = 'publicacion_pendiente'"
    adoPublicar.Recordset.Requery
    adoPublicar.Refresh
    dgPublicar.Refresh
End Sub

Private Sub cmdComentario_Click()
    db.Execute ("UPDATE publicar SET comentarios = '" & tbComen.Text & "'")
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelComp_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCat.Text = ""
    tbDescCat.Text = ""
    tbCodFase.Text = ""
    tbDescFase.Text = ""
End Sub

Private Sub CommandButton1_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCat.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text, "descripcion", " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCat.Text = sResultado(2)
    tbCodFase.Text = ""
    tbDescFase.Text = ""
End Sub

Private Sub CommandButton2_Click()
    If tbCodCat.Text = "" Then
        MsgBox mml_FRASE0504, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & tbCodCat.Text)
    Select Case tbCodFase.Text
        Case 1:
            tbDescFase.Text = mml_FRASE0329
        Case 2:
            tbDescFase.Text = mml_FRASE0330
        Case 4:
            tbDescFase.Text = mml_FRASE0652
        Case 8:
            tbDescFase.Text = mml_FRASE0653
        Case Else
            tbDescFase.Text = "1/" & tbCodFase.Text & "os de Final"
    End Select
    AsignarComentario
End Sub



Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))

End Sub

Private Sub tbDesc_DblClick()
    AsignarComentario
        
End Sub
Sub AsignarComentario()
Dim rs As Recordset, iNumBailes As Integer
    tbDesc.Text = tbDescCat.Text
    If tbComen.Text = "" Then
        If BailesParciales(Val(tbCodCat.Text)) Then
            Set rs = db.OpenRecordset("SELECT nombre FROM bailes b, bailes_categ bc WHERE b.codigo IN (SELECT cod_baile FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND fase = " & tbCodFase.Text & " AND repesca = " & chkRep.Value & " ) AND b.codigo = bc.cod_baile AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase = " & IIf(tbCodFase.Text = 1, 1, 2), dbOpenSnapshot)
            tbComen.Text = mml_FRASE0473
            While Not rs.EOF
                tbComen.Text = tbComen.Text & sNombreBaileAbreviado(rs!Nombre) & " "
                rs.MoveNext
            Wend
            rs.Close
        End If
        If cbRep.Text = "" Then
            tbComen.Text = tbComen.Text & C_RESULTADOS_ELIMINATORIA
        Else
            tbComen.Text = tbComen.Text & cbRep.Text
        End If
    End If
End Sub

