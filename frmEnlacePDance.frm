VERSION 5.00
Begin VB.Form frmEnlaceProDance 
   Caption         =   "mml_FRASE0559"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0559"
      Height          =   2970
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   10635
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
         Height          =   480
         Left            =   6705
         TabIndex        =   24
         Top             =   2385
         Width           =   1965
      End
      Begin VB.CommandButton cmdPublicar 
         Caption         =   "mml_FRASE0560"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4365
         TabIndex        =   23
         Top             =   2385
         Width           =   2280
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "mml_FRASE0058"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1890
         TabIndex        =   22
         Top             =   2385
         Width           =   2415
      End
      Begin VB.CommandButton cmdRecDatos 
         Caption         =   "mml_FRASE0561"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   6435
         TabIndex        =   21
         Top             =   1845
         Width           =   2235
      End
      Begin VB.CommandButton cmdTransDatos 
         Caption         =   "mml_FRASE0563"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4140
         TabIndex        =   20
         Top             =   1845
         Width           =   2235
      End
      Begin VB.CommandButton cmdGenDatos 
         Caption         =   "mml_FRASE0562"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1845
         TabIndex        =   19
         Top             =   1845
         Width           =   2235
      End
      Begin VB.CommandButton cmdSelFase 
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
         Picture         =   "frmEnlacePDance.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1305
         Width           =   450
      End
      Begin VB.CommandButton cmdSelCat 
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
         Picture         =   "frmEnlacePDance.frx":046A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   855
         Width           =   450
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
         Picture         =   "frmEnlacePDance.frx":08D4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   450
      End
      Begin VB.Frame Frame3 
         Caption         =   "mml_FRASE0437"
         Height          =   735
         Left            =   135
         TabIndex        =   14
         Top             =   1845
         Width           =   1305
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
            ItemData        =   "frmEnlacePDance.frx":0D3E
            Left            =   150
            List            =   "frmEnlacePDance.frx":0D60
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdDescalif 
         Height          =   495
         Left            =   9900
         Picture         =   "frmEnlacePDance.frx":0D8A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "mml_FRASE0043"
         Top             =   765
         Width           =   615
      End
      Begin VB.CommandButton cmdDorsales 
         Height          =   495
         Left            =   9900
         Picture         =   "frmEnlacePDance.frx":196C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "mml_FRASE0028"
         Top             =   270
         Width           =   615
      End
      Begin VB.CommandButton cmdCategAct 
         Height          =   495
         Left            =   9135
         Picture         =   "frmEnlacePDance.frx":248E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "mml_FRASE0428"
         Top             =   300
         Width           =   615
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   360
         Width           =   5895
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   1320
         Width           =   3960
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
         Left            =   7170
         TabIndex        =   1
         Top             =   1305
         Width           =   1950
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmEnlaceProDance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBorrarFicheros_Click()
    AppActivate "C:\WINDOWS"
    Sleep 1000
    SendKeys Chr$(65)
    Sleep 1000
    SendKeys mml_FRASE0564, 100
End Sub

Private Sub cmdCalcular_Click()
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        frmCalcular.tbCodComp.Text = tbCodComp.Text
        frmCalcular.tbDescComp.Text = tbDescComp.Text
        frmCalcular.tbCodCat.Text = tbCodCat.Text
        frmCalcular.tbDescCat.Text = tbDescCat.Text
        frmCalcular.tbCodFase.Text = Val(tbCodFase.Text)
        frmCalcular.tbDescFase.Text = sDescFase(Val(tbCodFase.Text))
    
        frmCalcular.Show vbModal
    End If

End Sub

Private Sub cmdCategAct_Click()
Dim rs As Recordset, sMsj As String
    Set rs = db.OpenRecordset("SELECT TOP 1 * FROM horario h WHERE  grupo LIKE '*" & cbPista.Text & "*' AND numfase <> " & C_FASE_GENERAL_LOOK & " AND cod_competicion = " & Val(VarCfg("horario_codcompeticion")) & " AND (SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = h.cod_categoria AND fase = h.numfase AND repesca = h.repesca) = 0 ORDER BY orden", dbOpenSnapshot)
    If Not rs.EOF Then
        tbCodComp.Text = rs!cod_competicion
        tbDescComp.Text = sDescCompeticion(rs!cod_competicion)
        tbCodCat.Text = rs!cod_categoria
        tbDescCat.Text = sDescCategoria(rs!cod_categoria)
        tbCodFase.Text = rs!numfase
        DescFase
    Else
        MsgBox mml_FRASE0442, vbOKOnly Or vbInformation, mml_FRASE0147
    End If
    rs.Close

End Sub

Private Sub cmdDescalif_Click()
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        frmDescalificados.tbCodComp.Text = tbCodComp.Text
        frmDescalificados.tbDescComp.Text = tbDescComp.Text
        frmDescalificados.tbCodCateg.Text = tbCodCat.Text
        frmDescalificados.tbDescCateg.Text = tbDescCat.Text
        frmDescalificados.cbFase.ListIndex = Log(Val(tbCodFase.Text)) / Log(2)
        frmDescalificados.cmdActualizar_Click

        frmDescalificados.Show vbNomodal
    End If

End Sub

Private Sub cmdDorsales_Click()
    If tbCodComp.Text <> "" And tbCodCat.Text <> "" And Val(tbCodFase.Text) > 0 Then
        frmADorsales.tbCodComp.Text = tbCodComp.Text
        frmADorsales.tbDescComp.Text = tbDescComp.Text
        frmADorsales.tbCodCateg.Text = tbCodCat.Text
        frmADorsales.tbDescCateg.Text = tbDescCat.Text
        frmADorsales.cbFase.ListIndex = Log(Val(tbCodFase.Text)) / Log(2) + 1
        
        frmADorsales.Show vbNomodal
        
    End If

End Sub

Private Sub cmdGenDatos_Click()
Dim i As Integer, iFile As Long
Dim rs As Recordset, rs1 As Recordset, rs2 As Recordset
Dim bDatGen As Boolean, dFecha As Date, sFichero As String
Dim sExt As String

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    If tbCodCat.Text = "" Then
        If MsgBox(mml_FRASE0565, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            Exit Sub
        End If
         'Fichero CLASS.XX
         bDatGen = False
         Set rs = db.OpenRecordset("SELECT * FROM PD_Secciones", dbOpenSnapshot)
         While Not rs.EOF
             iFile = FreeFile
             Set rs1 = db.OpenRecordset("SELECT * FROM PD_NumBailes WHERE seccion = '" & rs!abrev & "'", dbOpenSnapshot)
             If Not rs1.EOF Then
                 Open G_DIR_PRODANCE & "\DATA\CLASS." & rs!abrev For Output As #iFile
                 While Not rs1.EOF
                     Print #iFile, rs1!letra & " " & rs1!nobailes
                     rs1.MoveNext
                 Wend
             End If
             Close #iFile
             If Not bDatGen Then
                 FileCopy G_DIR_PRODANCE & "\DATA\CLASS." & rs!abrev, G_DIR_PRODANCE & "\DATA\CLASS.DAT"
                 bDatGen = True
             End If
             rs1.Close
             rs.MoveNext
         Wend
         rs.Close
         
         'Fichero DANCE.XX
         bDatGen = False
         Set rs = db.OpenRecordset("SELECT * FROM PD_Secciones", dbOpenSnapshot)
         While Not rs.EOF
             iFile = FreeFile
             Set rs1 = db.OpenRecordset("SELECT * FROM PD_Bailes WHERE seccion = '" & rs!abrev & "'", dbOpenSnapshot)
             If Not rs1.EOF Then
                 Open G_DIR_PRODANCE & "\DATA\DANCE." & rs!abrev For Output As #iFile
                 While Not rs1.EOF
                     Print #iFile, rs1!abrev & " " & rs1!baile
                     rs1.MoveNext
                 Wend
             End If
             Close #iFile
             If Not bDatGen Then
                 FileCopy G_DIR_PRODANCE & "\DATA\DANCE." & rs!abrev, G_DIR_PRODANCE & "\DATA\DANCE.DAT"
                 bDatGen = True
             End If
             rs1.Close
             rs.MoveNext
         Wend
         rs.Close
         
         'Fichero DANCESEQ.XX
         Set rs = db.OpenRecordset("SELECT * FROM PD_Disciplinas", dbOpenSnapshot)
         While Not rs.EOF
             iFile = FreeFile
             Set rs1 = db.OpenRecordset("SELECT * FROM PD_seqBailes WHERE seccion = '" & rs!codigo & rs!seccion & "'", dbOpenSnapshot)
             If Not rs1.EOF Then
                 Open G_DIR_PRODANCE & "\DATA\DANCESEQ." & rs!codigo & rs!seccion For Output As #iFile
                 While Not rs1.EOF
                     Print #iFile, rs1![1baile] & " " & rs1![2bailes] & " " & rs1![3bailes] & " " & rs1![4bailes] & " " & rs1![5bailes]
                     rs1.MoveNext
                 Wend
             End If
             Close #iFile
             rs1.Close
             rs.MoveNext
         Wend
         rs.Close
        
         'Fichero DISCIPL.XX
         Set rs = db.OpenRecordset("SELECT * FROM PD_Secciones", dbOpenSnapshot)
         While Not rs.EOF
             iFile = FreeFile
             Set rs1 = db.OpenRecordset("SELECT * FROM PD_Disciplinas WHERE seccion = '" & rs!abrev & "'", dbOpenSnapshot)
             If Not rs1.EOF Then
                 Open G_DIR_PRODANCE & "\DATA\DISCIPL." & rs!abrev For Output As #iFile
                 While Not rs1.EOF
                     Print #iFile, rs1!codigo & " " & rs1!abrev & " " & rs1!DESCRIPCION
                     rs1.MoveNext
                 Wend
             End If
             Close #iFile
             rs1.Close
             rs.MoveNext
         Wend
         rs.Close
         
         'Fichero GROUP.XX
         bDatGen = False
         Set rs = db.OpenRecordset("SELECT * FROM PD_Secciones", dbOpenSnapshot)
         While Not rs.EOF
             iFile = FreeFile
             Set rs1 = db.OpenRecordset("SELECT * FROM PD_GruposEdad WHERE seccion = '" & rs!abrev & "'", dbOpenSnapshot)
             If Not rs1.EOF Then
                 Open G_DIR_PRODANCE & "\DATA\GROUP." & rs!abrev For Output As #iFile
                 While Not rs1.EOF
                     Print #iFile, rs1!codigo & " " & rs1!abrev1 & " " & rs1!abrev2 & " " & rs1!DESCRIPCION
                     rs1.MoveNext
                 Wend
             End If
             Close #iFile
             If Not bDatGen Then
                 FileCopy G_DIR_PRODANCE & "\DATA\GROUP." & rs!abrev, G_DIR_PRODANCE & "\DATA\GROUP.DAT"
                 bDatGen = True
             End If
             rs1.Close
             rs.MoveNext
         Wend
         rs.Close
         
         'Fichero SECTION.DAT
         Set rs = db.OpenRecordset("SELECT * FROM PD_Secciones", dbOpenSnapshot)
         iFile = FreeFile
         Open G_DIR_PRODANCE & "\DATA\SECTION.DAT" For Output As #iFile
         While Not rs.EOF
             Print #iFile, rs!codigo & rs!abrev & rs!DESCRIPCION
             rs.MoveNext
         Wend
         Close #iFile
         rs.Close
   End If
         
    'Recuperamos la fecha de la competición
    Set rs = db.OpenRecordset("SELECT fecha FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    If rs.EOF Then
       MsgBox mml_FRASE0566, vbOKOnly Or vbCritical, mml_FRASE0096
       Exit Sub
    End If
    dFecha = rs!fecha
    rs.Close
    sFichero = "1" & Format$(dFecha, "ddmmyy")
    'Comprobamos si la competición está en el fichero de competiciones
    Dim sCad As String, bPresente As Boolean
    bPresente = False
    iFile = FreeFile
    Open G_DIR_PRODANCE & "\DATA\TURNIERE.DAT" For Input As #iFile
    While Not EOF(iFile)
        Line Input #iFile, sCad
        If sFichero = sCad Then
            bPresente = True
        End If
    Wend
    Close #iFile
    If Not bPresente Then
        Open G_DIR_PRODANCE & "\DATA\TURNIERE.DAT" For Append As #iFile
        Print #iFile, sFichero
        Close #iFile
    End If
     
    'Primero creamos el directorio
    On Local Error Resume Next
    MkDir G_DIR_PRODANCE & "\" & sFichero
    ChDir G_DIR_PRODANCE & "\" & sFichero
    If Dir$("*.*") <> "" Then
        If MsgBox(mml_FRASE0567, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            Exit Sub
        End If
        Kill "*.*"
    End If
    On Local Error GoTo 0
    
    If tbCodCat.Text = "" Then
        'Ahora generamos la información de cada categoría IDSF o Prof
        Set rs = db.OpenRecordset("SELECT * FROM categorias WHERE (descripcion LIKE '*Prof*' OR descripcion LIKE '*IDSF*' OR descripcion LIKE '*ProD*') AND cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    Else
        If InStr(tbDescCat.Text, mml_FRASE0568) = 0 And InStr(tbDescCat.Text, mml_FRASE0569) = 0 Then
            If MsgBox(mml_FRASE0570, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                Exit Sub
            End If
        End If
        Set rs = db.OpenRecordset("SELECT * FROM categorias WHERE codigo = " & tbCodCat.Text & " AND cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    End If
     While Not rs.EOF
        tbCodCat.Text = rs!codigo
        tbDescCat.Text = rs!DESCRIPCION
        'Calculamos la extensión
        'Grupo de edad
        Set rs1 = db.OpenRecordset("SELECT codigo FROM PD_GruposEdad WHERE cod_grupoedad = " & rs!cod_grupoedad, dbOpenSnapshot)
        sExt = rs1.Fields(0)
        rs1.Close
        'Bailes
        Set rs1 = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE fase = 1 AND cod_categoria = " & rs!codigo, dbOpenSnapshot)
        sExt = sExt & rs1.Fields(0)
        rs1.Close
        'Modalidad
        Set rs1 = db.OpenRecordset("SELECT codigo FROM PD_Disciplinas WHERE cod_modalidad = " & rs!cod_modalidad, dbOpenSnapshot)
        sExt = sExt & rs1.Fields(0)
        rs1.Close
        
        'Fichero K general
        iFile = FreeFile
        Open "K" & sFichero & ".DAT" For Append As #iFile
        Print #iFile, sFichero & "." & sExt
        Close #iFile
        
        'Datos organización
        iFile = FreeFile
        Open "P" & sFichero & "." & sExt For Output As #iFile
        Print #iFile, G_DATOS_ORG_PRODANCE
        Close #iFile
        FileCopy "P" & sFichero & "." & sExt, "P" & sFichero & ".DAT"
        
        'Datos del evento
        iFile = FreeFile
        Open "V" & sFichero & "." & sExt For Output As #iFile
        Set rs1 = db.OpenRecordset("SELECT * FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
        Print #iFile, rs1!DESCRIPCION
        Print #iFile, rs1!direccion
        Print #iFile, rs1!escuela
        Print #iFile, rs1!DESCRIPCION
        i = InStr(rs1!direccion, mml_FRASE0571)
        If i > 0 Then
            Print #iFile, Mid$(rs1!direccion, i - 6, 2)
            Print #iFile, Mid$(rs1!direccion, i - 3, 2)
        Else
            Print #iFile, "10"
            Print #iFile, "50"
        End If
        rs1.Close
        Close #iFile
        FileCopy "V" & sFichero & "." & sExt, "V" & sFichero & ".DAT"
     
        'Jueces
        iFile = FreeFile
        Set rs1 = db.OpenRecordset("SELECT nombre, direccion FROM juez_categ, jueces WHERE cod_juez = jueces.codigo AND cod_categoria = " & rs!codigo, dbOpenSnapshot)
        Open "R" & sFichero & "." & sExt For Output As #iFile
        Dim aNombre(10) As String, iSP As Integer, j As Integer
        While Not rs1.EOF
            iSP = 20
            i = DividirCampo(rs1!Nombre, aNombre(), " ")
            For j = 0 To IIf(i > 2, 1, i - 1)
                Print #iFile, Left$(aNombre(j) & Space(25), iSP);
                iSP = iSP + 5
            Next
            If j < 2 Then
                Print #iFile, Space(25);
            End If
            Print #iFile, Left$(rs1!direccion & Space(40), 40)
            
            rs1.MoveNext
        Wend
        Close #iFile
        rs1.Close
        
        'Dorsales y parejas
        iFile = FreeFile
        Set rs1 = db.OpenRecordset("SELECT DISTINCT d.num_dorsal, p.nombre_hombre, p.nombre_mujer, p.provincia FROM dorsales d, parejas p WHERE p.codigo = d.cod_pareja AND cod_categoria = " & rs!codigo & " ORDER BY num_dorsal", dbOpenSnapshot)
        Open "T" & sFichero & "." & sExt For Output As #iFile
        While Not rs1.EOF
            Print #iFile, Right$(Space(3) & rs1!num_dorsal, 3) & "1";
            i = DividirCampo(rs1!nombre_hombre, aNombre(), " ")
            iSP = 20
            For j = 0 To IIf(i > 2, 1, i - 1)
                Print #iFile, Left$(aNombre(j) & Space(25), iSP);
                iSP = iSP + 5
            Next
            If j < 2 Then
                Print #iFile, Space(25);
            End If
            i = DividirCampo(rs1!nombre_mujer, aNombre(), " ")
            iSP = 20
            For j = 0 To IIf(i > 2, 1, i - 1)
                Print #iFile, Left$(aNombre(j) & Space(25), iSP);
                iSP = iSP + 5
            Next
            If j < 2 Then
                Print #iFile, Space(25);
            End If
            Print #iFile, Left$(rs1!provincia & Space(40), 40)
            
            rs1.MoveNext
        Wend
        Close #iFile
        rs1.Close
        
        'Fichero inicial de puntuaciones y dorsales
        iFile = FreeFile
        Set rs1 = db.OpenRecordset("SELECT DISTINCT d.num_dorsal, p.nombre_hombre, p.nombre_mujer FROM dorsales d, parejas p WHERE p.codigo = d.cod_pareja AND cod_categoria = " & rs!codigo & " ORDER BY num_dorsal", dbOpenSnapshot)
        Open "W" & sFichero & "." & sExt For Output As #iFile
        While Not rs1.EOF
            Print #iFile, Right$(Space(3) & rs1!num_dorsal, 3) & "1100000000----------------"
            rs1.MoveNext
        Wend
        Close #iFile
        rs1.Close
        
        'Grabamos la información en la lista de categorias asociadas a prodance
        db.Execute "DELETE FROM PD_Categorias WHERE cod_categoria = " & rs!codigo
        db.Execute "INSERT INTO PD_Categorias VALUES (" & rs!codigo & ",'" & sExt & "',#01/01/1999 11:00:00#,'" & sFichero & "')"
        
        rs.MoveNext
     Wend
     rs.Close
     MsgBox mml_FRASE0572, vbOKOnly Or vbInformation, mml_FRASE0086
    Exit Sub
error:
    ProcesarError "cmdGenDatos_Click"
End Sub


Private Sub cmdPublicar_Click()
Dim iFaseSig As Integer
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    iFaseSig = tbCodFase.Text
    
    frmPublicar.Publicar tbCodComp.Text, tbCodCat.Text, iFaseSig, tbDescCat.Text, chkRep.Value, tbDescCat.Text, tbDescComp.Text

End Sub

Private Sub cmdRecDatos_Click()
Dim rs1 As Recordset, rs As Recordset, sExt As String, iFile As Integer, sCad As String
Dim iFaseIni As Integer, iFaseAct As Integer, sNumDorsal As String

    If tbCodFase.Text = "" Or tbCodCat.Text = "" Or tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If Val(tbCodFase.Text) > 1 Then
        If MsgBox(mml_FRASE0573, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
            tbCodFase.Text = Val(tbCodFase.Text) / 2
            tbDescFase.Text = sDescFase(Val(tbCodFase.Text))
            DoEvents
        End If
    Else
        If MsgBox(mml_FRASE0574, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            Exit Sub
        End If
    End If
    
    If MsgBox(mml_FRASE0575, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If

    Set rs1 = db.OpenRecordset("SELECT * FROM PD_Categorias WHERE cod_categoria =" & tbCodCat.Text, dbOpenSnapshot)
    If rs1.EOF Then
        MsgBox mml_FRASE0576, vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    Else
        Set rs = db.OpenRecordset("SELECT MAX(fase) FROM dorsales WHERE cod_categoria = " & tbCodCat.Text, dbOpenSnapshot)
        iFaseIni = rs.Fields(0)
        rs.Close
        
        iFaseAct = Val(tbCodFase.Text)
        iFaseIni = Log(iFaseIni) / Log(2)
        iFaseAct = Log(iFaseAct) / Log(2)
        
        iFile = FreeFile
        db.Execute "DELETE FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca = 0 AND fase = " & tbCodFase.Text
        Open G_DIR_PRODANCE & "\" & rs1!Directorio & "\W" & rs1!Directorio & "." & rs1!Ext For Input As #iFile
        While Not EOF(iFile)
            Line Input #iFile, sCad
            sNumDorsal = Mid$(sCad, 1, 3)
            sCad = Mid$(sCad, 5, 10)
            If InStr(sCad, Trim$(Str$(iFaseIni - iFaseAct + 1))) > 0 Then
                'El dorsal está en la fase, lo incorporamos a dorsales
                Set rs = db.OpenRecordset("SELECT cod_pareja FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND num_dorsal = " & sNumDorsal, dbOpenSnapshot)
                db.Execute "INSERT INTO dorsales VALUES (" & MaxCod("Dorsales") & "," & sNumDorsal & "," & tbCodCat.Text & "," & tbCodFase.Text & "," & rs!cod_pareja & ",0,0)"
                rs.Close
            End If
        Wend
        Close #iFile
    End If
    rs1.Close

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCat_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCat.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCat.Text = sResultado(2)
    tbCodFase.Text = ""
    tbDescFase.Text = ""
    
End Sub
Function RecuperarJueces(iCodCateg As Integer, cbJuez As ComboBox) As Integer
Dim rs As Recordset, i As Integer
    ' Recuperamos los jueces
    cbJuez.Clear
    Set rs = db.OpenRecordset("SELECT id_juez FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & iCodCateg & " ORDER BY 1", dbOpenSnapshot)
        cbJuez.Tag = 0
        i = 0
        While Not rs.EOF
            i = i + 1
            cbJuez.Tag = 1
            cbJuez.AddItem rs!id_juez
            rs.MoveNext
        Wend
    rs.Close
    cbJuez.Refresh
    RecuperarJueces = i
    cbJuez.ListIndex = -1
End Function

Private Sub cmdSelComp_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCat.Text = ""
    tbDescCat.Text = ""
    tbCodFase.Text = ""
    tbDescFase.Text = ""

End Sub

Private Sub cmdSelFase_Click()
    If tbCodCat.Text = "" Then
        MsgBox mml_FRASE0504, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1")
    DescFase

End Sub

Sub DescFase()
    Select Case tbCodFase.Text
        Case 1:
            tbDescFase.Text = mml_FRASE0329
        Case 2:
            tbDescFase.Text = "SEMI-FINAL"
        Case 4:
            tbDescFase.Text = "CUARTOS DE FINAL"
        Case 8:
            tbDescFase.Text = "OCTAVOS DE FINAL"
        Case "":
            tbDescFase.Text = ""
        Case Else
            tbDescFase.Text = tbCodFase.Text & "OS DE FINAL"
    End Select

End Sub

Private Sub cmdTransDatos_Click()
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    frmTransDatosProDance.RecuperarDatos Val(tbCodCat.Text), Val(tbCodFase.Text)
End Sub


Private Sub Form_Load()
    TraducirCadenas Me
    CargarPistas cbPista

End Sub
