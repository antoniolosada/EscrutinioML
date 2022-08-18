VERSION 5.00
Begin VB.Form frmComprobarFases 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mml_FRASE0536"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   495
      Left            =   8190
      TabIndex        =   3
      Top             =   5490
      Width           =   1860
   End
   Begin VB.CommandButton cmdCorregirFases 
      Caption         =   "mml_FRASE0537"
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
      Left            =   4020
      TabIndex        =   2
      Top             =   5490
      Width           =   2895
   End
   Begin VB.CommandButton cmdComprobarFases 
      Caption         =   "mml_FRASE0538"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   5490
      Width           =   2715
   End
   Begin VB.TextBox tbCateg 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5385
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmComprobarFases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdComprobarFases_Click()
Dim rs As Recordset, rs1 As Recordset, iNumFases As Integer

    tbCateg.Text = mml_FRASE0539 & Chr$(13) & Chr$(10)
    
    Set rs = db.OpenRecordset("SELECT MAX(fase) as maxfase,COUNT(*) as numdorsales,cod_categoria,descripcion FROM dorsales d, categorias c WHERE d.cod_categoria = c.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & frmADorsales.tbCodComp.Text & ") GROUP BY cod_categoria, descripcion", dbOpenSnapshot)
    While Not rs.EOF
        'Contamos el número de fases
        Set rs1 = db.OpenRecordset("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & rs!cod_categoria, dbOpenSnapshot)
        iNumFases = 0
        If Not rs1.EOF Then
            rs1.MoveLast
            iNumFases = rs1.RecordCount
        End If
        rs1.Close
        
        If iNumFases > 1 Then
            tbCateg.Text = tbCateg.Text & rs!DESCRIPCION & " (" & rs!cod_categoria & mml_FRASE0540 & Chr$(13) & Chr$(10)
        Else ' Si solo hay una comprobamos si es la correcta para el número de dorsales
            If (rs!maxfase = 1 And rs!numdorsales <= 7) Or (rs!maxfase = 2 And rs!numdorsales > 7 And rs!numdorsales <= 13) Or (rs!maxfase > 2 And CalcularFase(rs!numdorsales) = rs!maxfase) Then
            'Fase correcta
            Else
                tbCateg.Text = tbCateg.Text & rs!DESCRIPCION & mml_FRASE0541 & rs!maxfase & mml_FRASE0542 & CalcularFase(rs!numdorsales) & mml_FRASE0543 & rs!numdorsales & mml_FRASE0544 & Chr$(13) & Chr$(10)
            End If
        End If
        
        rs.MoveNext
    Wend
    rs.Close

    tbCateg.Text = tbCateg.Text & mml_FRASE0545

End Sub
Private Sub cmdCorregirFases_Click()
Dim rs As Recordset, rs1 As Recordset, iNumFases As Integer

    If MsgBox(mml_FRASE0546, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then Exit Sub

    tbCateg.Text = mml_FRASE0547 & Chr$(13) & Chr$(10)
    Set rs = db.OpenRecordset("SELECT MAX(fase) as maxfase,COUNT(*) as numdorsales,cod_categoria,descripcion FROM dorsales d, categorias c WHERE d.cod_categoria = c.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & frmADorsales.tbCodComp.Text & ") GROUP BY cod_categoria, descripcion", dbOpenSnapshot)
    While Not rs.EOF
        'Contamos el número de fases
        Set rs1 = db.OpenRecordset("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & rs!cod_categoria, dbOpenSnapshot)
        iNumFases = 0
        If Not rs1.EOF Then
            rs1.MoveLast
            iNumFases = rs1.RecordCount
        End If
        rs1.Close
        
        If iNumFases > 1 Then
            tbCateg.Text = tbCateg.Text & rs!DESCRIPCION & " (" & rs!cod_categoria & mml_FRASE0548 & Chr$(13) & Chr$(10)
            tbCateg.Refresh
            CorregirFase rs!cod_categoria, CalcularFase(rs!numdorsales)
        Else ' Si solo hay una comprobamos si es la correcta para el número de dorsales
            If (rs!maxfase = 1 And rs!numdorsales <= 7) Or (rs!maxfase = 2 And rs!numdorsales > 7 And rs!numdorsales <= 13) Or (rs!maxfase > 2 And CalcularFase(rs!numdorsales) = rs!maxfase) Then
            'Fase correcta
            Else
                tbCateg.Text = tbCateg.Text & rs!DESCRIPCION & mml_FRASE0541 & rs!maxfase & mml_FRASE0542 & CalcularFase(rs!numdorsales) & mml_FRASE0549 & Chr$(13) & Chr$(10)
                tbCateg.Refresh
                CorregirFase rs!cod_categoria, CalcularFase(rs!numdorsales)
            End If
        End If
        
        rs.MoveNext
    Wend
    rs.Close
    
    tbCateg.Text = tbCateg.Text & mml_FRASE0545

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Public Sub CorregirFase(iCodCat As Long, iFase As Integer)
    db.Execute "UPDATE dorsales SET fase = " & iFase & " WHERE cod_categoria = " & iCodCat
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub
