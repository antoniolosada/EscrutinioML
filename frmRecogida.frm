VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmRecogida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1110"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDevolverDorsal 
      Caption         =   "mml_FRASE1115"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5160
      TabIndex        =   13
      Top             =   540
      Width           =   3495
   End
   Begin VB.CommandButton cmdRecogerDorsal 
      Caption         =   "mml_FRASE1110"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1470
      TabIndex        =   12
      Top             =   540
      Width           =   3495
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "mml_FRASE0295"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5430
      TabIndex        =   11
      Top             =   60
      Width           =   2355
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
      Left            =   3180
      Picture         =   "frmRecogida.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   60
      Width           =   495
   End
   Begin VB.TextBox tbFecha 
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
      Left            =   3660
      TabIndex        =   4
      Top             =   60
      Width           =   1725
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "mml_FRASE0886"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   8220
      Width           =   2355
   End
   Begin VB.Frame mrcMarco 
      Height          =   7185
      Left            =   30
      TabIndex        =   0
      Top             =   990
      Width           =   10275
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "mml_FRASE0251"
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
         Left            =   4560
         TabIndex        =   9
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "mml_FRASE1028"
         Default         =   -1  'True
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
         Left            =   3240
         TabIndex        =   8
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox tbTexto 
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
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   3195
      End
      Begin VB.ComboBox cbOrden 
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
         Left            =   7620
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   2535
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgParejas 
         Height          =   6525
         Left            =   60
         TabIndex        =   3
         Top             =   600
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   11509
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label25 
         Caption         =   "mml_FRASE0190"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.Label Label13 
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
      Left            =   2220
      TabIndex        =   5
      Top             =   90
      Width           =   945
   End
End
Attribute VB_Name = "frmRecogida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const C_RECOGIDA = 0
Const C_DEVOLVER = 1
Const C_CODE_CTRL = 17

Dim m_KeyCode As Long


Private Sub cbOrden_Click()
    Botones False
    CargarDatos
    Botones True

End Sub

Private Sub cmdBorrar_Click()
    tbTexto.Text = ""
    tbTexto.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    CargarDatos
    tbTexto.SetFocus
End Sub

Sub Botones(iOp As Boolean)
    cmdCargar.Enabled = iOp
    cmdRecogerDorsal.Enabled = iOp
    cmdDevolverDorsal.Enabled = iOp

End Sub

Private Sub cmdCargar_Click()
    Botones False
    CargarDatos
    Botones True
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub



Private Sub cmdDevolverDorsal_Click()
    Botones False
    Dorsal C_DEVOLVER, mml_FRASE1116
    Botones True
End Sub

Private Sub cmdSelComp_Click()
    tbFecha.Text = frmCalendario.Mostrar
End Sub

Private Sub cmdRecogerDorsal_Click()
    Botones False
    Dorsal C_RECOGIDA, mml_FRASE1114
    Botones True
End Sub

Private Sub Dorsal(iOp As Integer, sFrase As String, Optional bExacto As Boolean = False)

Dim rs As Recordset
Dim iDorsal, lCodCat As Long, lCodPareja As Long
Dim sNombreMujer As String, sNombreHombre As String
Dim sTmp As String
Dim sSQL As String

    m_KeyCode = 0
    If Not C_DEBUG Then On Local Error GoTo error

    If dgParejas.Visible And dgParejas.Row >= 0 Then
        iDorsal = dgParejas.TextMatrix(dgParejas.Row, 0)
        sNombreHombre = dgParejas.TextMatrix(dgParejas.Row, 1)
        sNombreMujer = dgParejas.TextMatrix(dgParejas.Row, 2)
        lCodCat = dgParejas.TextMatrix(dgParejas.Row, 5)
        lCodPareja = dgParejas.TextMatrix(dgParejas.Row, 6)
        
        'Localizamos todas las parejas que tengan el mismo nombre_hombre, nombre_mujer y dorsal
        sSQL = "SELECT nombre_hombre, nombre_mujer, d.num_dorsal, d.cod_categoria,d.cod_pareja, cat.descripcion AS nombre_categoria, c.descripcion AS nombre_competicion  FROM parejas p, dorsales d, " & _
            "categorias cat, competiciones c WHERE c.fecha = #" & Format$(CDate(tbFecha.Text), "mm/dd/yyyy") & _
            "# AND cat.codigo = d.cod_categoria AND cat.cod_competicion = c.codigo AND p.codigo = d.cod_pareja " & _
            " AND nombre_hombre = '" & sNombreHombre & "' AND nombre_mujer = '" & sNombreMujer & _
            "' AND d.num_dorsal = " & iDorsal
        If bExacto Then
            sSQL = sSQL & " AND p.codigo = " & lCodPareja
        End If
        Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
        sTmp = ""
        If Not rs.EOF Then
            While Not rs.EOF
                sTmp = sTmp & vbCrLf & "Dorsal/Number (" & rs.Fields("num_dorsal") & ") " & rs.Fields("nombre_hombre") & " - " & rs.Fields("nombre_mujer") & " - " & rs.Fields("nombre_categoria") & ", " & rs.Fields("nombre_competicion")
                rs.MoveNext
            Wend
            If MsgBox(sFrase & sTmp, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
                rs.MoveFirst
                While Not rs.EOF
                    If iOp = C_RECOGIDA Then
                        db.Execute "INSERT INTO recogidadorsales VALUES (" & rs.Fields("num_dorsal") & "," & rs.Fields("cod_categoria") & "," & rs.Fields("cod_pareja") & ")"
                    Else
                        db.Execute "DELETE FROM recogidadorsales WHERE num_dorsal = " & rs.Fields("num_dorsal") & " AND cod_categoria = " & rs.Fields("cod_categoria") & " AND cod_pareja = " & rs.Fields("cod_pareja")
                    End If
                    rs.MoveNext
                Wend
            End If
            rs.Close
        Else
            rs.Close
        End If
        Botones False
        CargarDatos
        Botones True
    Else
        MsgBox mml_FRASE1113, vbOKOnly Or vbCritical, G_MSG_ERROR
    End If
    Exit Sub
error:
    ProcesarError "cmdRecogerDorsal_Click"
End Sub

Private Sub dgParejas_DblClick()
Dim bOp As Boolean
    If dgParejas.Visible And dgParejas.Row >= 0 Then
        If m_KeyCode = C_CODE_CTRL Then
            bOp = True
        End If
        dgParejas.Col = 0
        If dgParejas.CellBackColor = vbRed Then
            Dorsal C_RECOGIDA, mml_FRASE1114, bOp
        Else
            Dorsal C_DEVOLVER, mml_FRASE1116, bOp
        End If
    End If
    
    Exit Sub
error:
    ProcesarError "dgParejas_DblClick"
End Sub


Private Sub dgParejas_KeyDown(KeyCode As Integer, Shift As Integer)
    m_KeyCode = KeyCode
End Sub

Private Sub dgParejas_KeyUp(KeyCode As Integer, Shift As Integer)
    m_KeyCode = 0
End Sub

Private Sub Form_Load()
    If Not C_DEBUG Then On Local Error GoTo error
    TraducirCadenas Me
    tbFecha.Text = Format$(Now, "dd/mm/yyyy")
    dgParejas.Cols = 8
    dgParejas.ColWidth(0) = 800
    dgParejas.ColWidth(1) = 2200
    dgParejas.ColWidth(2) = 2200
    dgParejas.ColWidth(3) = 2200
    dgParejas.ColWidth(4) = 2200
    dgParejas.ColWidth(5) = 600
    
    cbOrden.AddItem mml_FRASE1025
    cbOrden.AddItem mml_FRASE1026
    cbOrden.AddItem mml_FRASE1111
    cbOrden.AddItem mml_FRASE1112
    cbOrden.ListIndex = 0
    dgParejas.Visible = False
    
    Exit Sub
error:
    ProcesarError "Form_Load"
End Sub

Sub CargarDatos()
Dim rs As Recordset
Dim i As Integer
Dim iDorsal, lCodCat As Long, lCodPareja As Long
Dim sNombreMujer As String, sNombreHombre As String
Dim sTmp As String


    If Not C_DEBUG Then On Local Error GoTo error
    If tbFecha.Text = "" Then Exit Sub
    
    Set rs = db.OpenRecordset("SELECT DISTINCT d.num_dorsal, d.cod_categoria, nombre_hombre, cat.descripcion as categ, " & _
            "c.descripcion as comp, c.codigo , cat.cod_modalidad, nombre_mujer, d.cod_pareja  " & _
            " FROM parejas p, categorias cat, competiciones c, dorsales d " & _
            " WHERE c.fecha = #" & Format$(CDate(tbFecha.Text), "mm/dd/yyyy") & _
            "# AND cat.codigo = d.cod_categoria AND cat.cod_competicion = c.codigo AND p.codigo = d.cod_pareja " & _
            " AND (nombre_hombre LIKE '*" & tbTexto.Text & "*' OR nombre_mujer LIKE '*" & tbTexto.Text & _
            "*' OR cat.descripcion LIKE '*" & tbTexto.Text & "*' OR d.num_dorsal = " & Val(tbTexto.Text) & _
            ") ORDER BY " & Val(cbOrden.List(cbOrden.ListIndex)), dbOpenSnapshot)
    
    mrcMarco.Visible = False
    dgParejas.Visible = False
    dgParejas.Rows = 0
    
    If Not rs.EOF Then
        dgParejas.Visible = True
    End If
    While Not rs.EOF
        dgParejas.AddItem rs.Fields("num_dorsal") & vbTab & rs!nombre_hombre & vbTab & rs!nombre_mujer & vbTab & rs.Fields("categ") & vbTab & rs.Fields("comp") & vbTab & rs.Fields("cod_categoria") & vbTab & rs.Fields("cod_pareja")
        rs.MoveNext
    Wend
    rs.Close
    
    For i = 0 To dgParejas.Rows - 1
        iDorsal = dgParejas.TextMatrix(i, 0)
        sNombreHombre = dgParejas.TextMatrix(i, 1)
        sNombreMujer = dgParejas.TextMatrix(i, 2)
        lCodCat = dgParejas.TextMatrix(i, 5)
        lCodPareja = dgParejas.TextMatrix(i, 6)
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM recogidadorsales WHERE cod_categoria = " & lCodCat & " AND cod_pareja = " & lCodPareja & " AND num_dorsal = " & iDorsal, dbOpenSnapshot)
        
        dgParejas.Col = 0
        dgParejas.Row = i
        If rs.Fields(0) > 0 Then
            dgParejas.CellBackColor = vbGreen
        Else
            dgParejas.CellBackColor = vbRed
        End If
        rs.Close
    Next
    mrcMarco.Visible = True
    mrcMarco.Refresh
    
    Exit Sub
error:
    ProcesarError "CargarDatos"
End Sub

Private Sub tbCodComp_Change()
    CargarDatos
End Sub
