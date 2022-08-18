VERSION 5.00
Begin VB.Form frmTransDatosProDance 
   Caption         =   "mml_FRASE0861"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopiarEnProDane 
      Caption         =   "mml_FRASE1173"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3600
      TabIndex        =   6
      Top             =   1980
      Width           =   3135
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "mml_FRASE0029"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6780
      TabIndex        =   3
      Top             =   1980
      Width           =   3045
   End
   Begin VB.CommandButton cmdTransDatos 
      Caption         =   "mml_FRASE0862"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   780
      TabIndex        =   2
      Top             =   1980
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0863"
      Height          =   1905
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   10725
      Begin VB.TextBox tbDatos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   930
         Width           =   10635
      End
      Begin VB.Label lblFase 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2565
         TabIndex        =   5
         Top             =   585
         Width           =   5595
      End
      Begin VB.Label lblCat 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10515
      End
   End
End
Attribute VB_Name = "frmTransDatosProDance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iCodCat As Integer
Dim iCodFase As Integer

Sub RecuperarDatos(iCodCateg As Long, iFase As Integer)
Dim rs As Recordset, rs1 As Recordset, sJuez As String

    iCodCat = iCodCateg
    iCodFase = iFase
    lblCat.Caption = sDescCategoria(iCodCat)
    lblFase.Caption = sDescFase(iCodFase)
    
    Me.Show vbModal
End Sub

Private Sub cmdCopiarEnProDane_Click()
Dim i As Integer
    If Not C_DEBUG Then On Local Error GoTo error
    AppActivate "PD3.0 - Enter"
    Sleep 2000
    For i = 1 To Len(tbDatos.Text)
        AppActivate "PD3.0 - Enter"
        DoEvents
        Sleep 10
        DoEvents
        SendKeys Mid$(tbDatos.Text, i, 1)
        DoEvents
        Sleep 10
        DoEvents
    Next
    Exit Sub
error:
    ProcesarError "cmdCopiarEnProDane_Click"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTransDatos_Click()
Dim rs As Recordset, rs1 As Recordset
    If iCodFase = 1 Then
        'Generamos todos los bailes uno a uno
        Set rs1 = db.OpenRecordset("SELECT cod_baile, nombre FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & iCodCat & " AND bc.fase = " & IIf(iCodFase = 1, 1, 2) & " ORDER BY posicion", dbOpenSnapshot)
        While Not rs1.EOF
            tbDatos.Text = ""
            Set rs = db.OpenRecordset("SELECT puesto FROM puntuaciones WHERE fase = " & iCodFase & " AND cod_baile = " & rs1!cod_baile & " AND cod_categoria = " & iCodCat & " AND repesca = 0 ORDER BY num_dorsal,cod_juez", dbOpenSnapshot)
            While Not rs.EOF
                tbDatos.Text = tbDatos.Text & rs!puesto
                rs.MoveNext
            Wend
            rs.Close
            Clipboard.Clear
            Clipboard.SetText tbDatos.Text
            If MsgBox(mml_FRASE0864 & rs1!Nombre & mml_FRASE0865, vbYesNo Or vbInformation, mml_FRASE0147) = vbNo Then
                rs1.Close
                Exit Sub
            End If
            rs1.MoveNext
        Wend
        rs1.Close
    Else
        Set rs1 = db.OpenRecordset("SELECT SUM(puesto), cod_juez, num_dorsal FROM puntuaciones WHERE fase = " & iCodFase & " AND cod_categoria = " & iCodCat & " AND repesca = 0 GROUP BY num_dorsal, cod_juez ORDER BY cod_juez, num_dorsal", dbOpenSnapshot)
        tbDatos.Text = ""
        While Not rs1.EOF
            tbDatos.Text = tbDatos.Text & rs1.Fields(0)
            rs1.MoveNext
        Wend
        rs1.Close
        Clipboard.Clear
        Clipboard.SetText tbDatos.Text
        MsgBox mml_FRASE0866, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub
