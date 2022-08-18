VERSION 5.00
Begin VB.Form frmTeamMatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orden de categorías en el Team Match"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "mml_FRASE0021"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   0
      TabIndex        =   12
      Top             =   1620
      Width           =   4680
      Begin VB.Label Label4 
         Caption         =   $"frmTeamMatch.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   4440
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "mml_FRASE0021"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   4680
      Begin VB.Label Label3 
         Caption         =   $"frmTeamMatch.frx":00A7
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Width           =   4440
      End
   End
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
      Height          =   540
      Left            =   3390
      TabIndex        =   5
      Top             =   5610
      Width           =   1260
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
      Left            =   840
      TabIndex        =   4
      Top             =   0
      Width           =   3840
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdGenCat 
      Caption         =   "mml_FRASE0024"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   30
      TabIndex        =   2
      Top             =   5610
      Width           =   3285
   End
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   30
      TabIndex        =   0
      Top             =   2880
      Width           =   4635
      Begin VB.CommandButton cmdSubir 
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
         Left            =   4095
         Picture         =   "frmTeamMatch.frx":0144
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   180
         Width           =   435
      End
      Begin VB.CommandButton cmdBajar 
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
         Left            =   4095
         Picture         =   "frmTeamMatch.frx":05AE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   630
         Width           =   435
      End
      Begin VB.TextBox tbMinBaile 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   3495
         TabIndex        =   7
         Text            =   "1"
         Top             =   2175
         Width           =   510
      End
      Begin VB.TextBox tbHoraIni 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   1320
         TabIndex        =   6
         Text            =   "12:00"
         Top             =   2175
         Width           =   960
      End
      Begin VB.ListBox lstCat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   3900
      End
      Begin VB.Label Label2 
         Caption         =   "min/baile"
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
         Left            =   2475
         TabIndex        =   9
         Top             =   2235
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "mml_FRASE0023"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   135
         TabIndex        =   8
         Top             =   2145
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmTeamMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBajar_Click()
Dim sTemp As String
    If lstCat.ListIndex < lstCat.ListCount - 1 Then
        sTemp = lstCat.List(lstCat.ListIndex + 1)
        lstCat.List(lstCat.ListIndex + 1) = lstCat.List(lstCat.ListIndex)
        lstCat.List(lstCat.ListIndex) = sTemp
        lstCat.ListIndex = lstCat.ListIndex + 1
    End If

End Sub

Private Sub cmdGenCat_Click()
Dim bComienzoEstandar As Boolean, iOrden As Integer
Dim rs1 As Recordset, rs2 As Recordset, rsBailesEst As Recordset, rsBailesLat As Recordset

    If Not IsDate(tbHoraIni.Text) Or Val(tbMinBaile.Text) = 0 Then
        MsgBox mml_FRASE0264, vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    End If
    
    If lstCat.ListCount < 2 Then
        MsgBox "Como mínimo debe haber dos categorias", vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    End If

    If MsgBox("¿Desea borrar el horario de la competición?", vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        db.Execute "DELETE FROM horario where cod_competicion = " & tbCodComp.Text
    End If
    
    Set rs1 = db.OpenRecordset("SELECT cod_modalidad FROM categorias WHERE codigo = " & Val(lstCat.List(0)), dbOpenSnapshot)
        bComienzoEstandar = False
        If Not rs1.EOF Then
            If rs1!cod_modalidad = 1 Then
                bComienzoEstandar = True
            End If
        Else
        End If
    rs1.Close
    Set rs1 = Nothing
    
    Set rs1 = db.OpenRecordset("SELECT MAX(orden) FROM horario WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    If Not IsNull(rs1.Fields(0)) Then
        iOrden = rs1.Fields(0)
    Else
        iOrden = 10
    End If
    rs1.Close
    Set rs1 = Nothing
    
    Set rsBailesEst = db.OpenRecordset("SELECT DISTINCT cod_baile,posicion FROM bailes_categ bc, categorias c WHERE c.cod_competicion = " & tbCodComp.Text & " AND c.cod_modalidad = 1 AND c.codigo = bc.cod_categoria AND c.descripcion LIKE '*" & C_TEAM_MATCH & "*' ORDER BY posicion", dbOpenSnapshot)
    Set rsBailesLat = db.OpenRecordset("SELECT DISTINCT cod_baile,posicion FROM bailes_categ bc, categorias c WHERE c.cod_competicion = " & tbCodComp.Text & " AND  c.cod_modalidad = 2 AND c.codigo = bc.cod_categoria AND c.descripcion LIKE '*" & C_TEAM_MATCH & "*' ORDER BY posicion", dbOpenSnapshot)
    
    If bComienzoEstandar Then
        Set rs1 = rsBailesEst
        Set rs2 = rsBailesLat
    Else
        Set rs2 = rsBailesEst
        Set rs1 = rsBailesLat
    End If
    
    dHora = tbHoraIni.Text
    iOrden = iOrden + 10
    While Not rs1.EOF And Not rs2.EOF
        If Not rs1.EOF Then
            For i = 0 To lstCat.ListCount / 2 - 1
                sSQL = "INSERT INTO horario VALUES (#" & dHora & "#,'" & sDescCategoria(Val(lstCat.List(i))) & "','TeamMatch',1," & Val(lstCat.List(i)) & ",0," & iOrden & ",0," & tbCodComp.Text & "," & rs1!cod_Baile & ",0,0,1)"
                Debug.Print sSQL
                db.Execute sSQL
                iOrden = iOrden + 10
                dHora = DateAdd("n", Val(tbMinBaile.Text), dHora)
            Next
            rs1.MoveNext
        End If
        If Not rs2.EOF Then
            For i = lstCat.ListCount / 2 To lstCat.ListCount - 1
                db.Execute "INSERT INTO horario VALUES (#" & dHora & "#,'" & sDescCategoria(Val(lstCat.List(i))) & "','TeamMatch',1," & Val(lstCat.List(i)) & ",0," & iOrden & ",0," & tbCodComp.Text & "," & rs2!cod_Baile & ",0,0,1)"
                iOrden = iOrden + 10
            Next
            rs2.MoveNext
        End If
    Wend
    
    rs1.Close
    rs2.Close
    
    MsgBox "Horario generado.", vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSubir_Click()
Dim sTemp As String
    If lstCat.ListIndex > 0 Then
        sTemp = lstCat.List(lstCat.ListIndex - 1)
        lstCat.List(lstCat.ListIndex - 1) = lstCat.List(lstCat.ListIndex)
        lstCat.List(lstCat.ListIndex) = sTemp
        lstCat.ListIndex = lstCat.ListIndex - 1
    End If
End Sub

Public Sub GenerarCategorias(iCodcomp As Long)
Dim rs As Recordset
    tbCodComp.Text = iCodcomp
    tbDescComp.Text = sDescCompeticion(iCodcomp)
    Set rs = db.OpenRecordset("SELECT * FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " AND descripcion LIKE '*" & C_TEAM_MATCH & "*'", dbOpenSnapshot)
    While Not rs.EOF
        lstCat.AddItem rs!codigo & " - " & rs!DESCRIPCION
        rs.MoveNext
    Wend
    rs.Close

    Me.Show vbModal

End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub
