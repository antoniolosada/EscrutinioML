VERSION 5.00
Begin VB.Form frmControlPresencia 
   Caption         =   "mml_FRASE1117"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   14895
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
      Height          =   675
      Left            =   7980
      TabIndex        =   3
      Top             =   7290
      Width           =   3615
   End
   Begin VB.CommandButton cmdEstablecerNoPresentes 
      Caption         =   "mml_FRASE1124"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2340
      TabIndex        =   2
      Top             =   7290
      Width           =   4995
   End
   Begin VB.Frame Frame1 
      Height          =   7125
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   14775
      Begin VB.ListBox lstCateg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6660
         Left            =   150
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   210
         Width           =   14535
      End
   End
End
Attribute VB_Name = "frmControlPresencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEstablecerNoPresentes_Click()
Dim i As Integer
Dim sCateg As String
Dim rs As Recordset
Dim sCad As String

    sCad = UCase(InputBox(mml_FRASE1125, G_MSG_AVISO))
    If sCad <> "SI" And sCad <> "YES" Then
        Exit Sub
    End If
    If Not C_DEBUG Then On Local Error GoTo error
    For i = 0 To lstCateg.ListCount - 1
        If lstCateg.Selected(i) Then
            sCateg = Val(lstCateg.List(i))
            Set rs = db.OpenRecordset("SELECT num_dorsal FROM recogidadorsales WHERE cod_categoria = " & sCateg, dbOpenSnapshot)
            db.Execute "UPDATE dorsales SET no_presente = 1 WHERE cod_categoria = " & sCateg
            While Not rs.EOF
                db.Execute "UPDATE dorsales SET no_presente = 0 WHERE cod_categoria = " & sCateg & " AND num_dorsal = " & rs.Fields("num_dorsal")
                rs.MoveNext
            Wend
            rs.Close
        End If
    Next
    
    MsgBox mml_FRASE1009, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    Exit Sub
error:
    ProcesarError "cmdEstablecerNoPresentes_Click"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
End Sub

Sub ControlPresencia(lCodComp As Long)
Dim rs As Recordset
Dim rsPresencia As Recordset
Dim sTmp As String, sSelect As String

    sTmp = ""
    If Not C_DEBUG Then On Local Error GoTo error
    'Localizamos las fases mas altas que no sean repesca y todavía no tengan puntuaiones y traemos el número de dorsales
    If MsgBox(mml_FRASE1119, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        sSelect = "c.cod_competicion = " & CodCompActiva
    Else
        sSelect = "co.fecha = #" & Format$(Now, "mm/dd/yyyy") & "#"
    End If
    Set rs = db.OpenRecordset("SELECT d.cod_categoria, c.descripcion AS des_comp, d.fase, co.descripcion, c.hora, COUNT(*) as numero " & _
                "FROM categorias c, dorsales d, competiciones co WHERE co.codigo = c.cod_competicion AND " & _
                "c.codigo = d.cod_categoria AND " & _
                sSelect & _
                " AND d.repesca = 0 AND d.fase = " & _
                "(SELECT MAX(d1.fase) FROM dorsales d1 WHERE d1.cod_categoria = c.codigo ) AND " & _
                "(SELECT COUNT(*) FROM puntuaciones pu WHERE pu.cod_categoria = c.codigo) = 0 " & _
                "GROUP BY d.cod_categoria, c.descripcion, d.fase, co.descripcion, c.hora ORDER BY c.hora", dbOpenSnapshot)
    While Not rs.EOF
        Set rsPresencia = db.OpenRecordset("SELECT COUNT(*) as numero FROM recogidadorsales WHERE cod_categoria = " & rs.Fields("cod_categoria"), dbOpenSnapshot)
        If rs.Fields("fase") = 1 Then
            If rsPresencia.Fields("numero") < 2 Then
                lstCateg.AddItem rs.Fields("cod_categoria") & " - Present(es): " & rsPresencia.Fields("numero") & " - Total: " & rs.Fields("numero") & " (" & sDescFase(rs.Fields("fase")) & ") " & rs.Fields("hora") & " - " & rs.Fields("descripcion") & ", " & rs.Fields("des_comp")
            End If
        ElseIf (rs.Fields("fase") / 2) * 6 + 1 >= rsPresencia.Fields("numero") Then
            lstCateg.AddItem rs.Fields("cod_categoria") & " - Present(es): " & rsPresencia.Fields("numero") & " - Total: " & rs.Fields("numero") & " (" & sDescFase(rs.Fields("fase")) & ") " & rs.Fields("hora") & " - " & rs.Fields("descripcion") & ", " & rs.Fields("des_comp")
        Else
        End If
        rsPresencia.Close
        rs.MoveNext
    Wend
    rs.Close
    
    
    Me.Show vbModal
    Exit Sub
error:
    ProcesarError "cmdComprobarRecogida_Click"

End Sub
