VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAgrupar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mml_FRASE1013"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3240
      TabIndex        =   9
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox tbDescComp 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   60
      Width           =   8445
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
      Height          =   375
      Left            =   2790
      Picture         =   "frmAgrupar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   60
      Width           =   465
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
      Height          =   645
      Left            =   7860
      TabIndex        =   6
      Top             =   9540
      Width           =   2325
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "mml_FRASE1016"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3180
      TabIndex        =   5
      Top             =   9540
      Width           =   3255
   End
   Begin VB.Frame frmEmparejar 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   14685
      Begin VB.CommandButton cmdBorrarGrupos 
         Caption         =   "mml_FRASE1015"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6120
         TabIndex        =   13
         Top             =   6270
         Width           =   2115
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gNoAgrupados 
         Height          =   8025
         Left            =   120
         TabIndex        =   11
         Top             =   810
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   14155
         _Version        =   393216
         Rows            =   6
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdDesagrupar 
         Caption         =   "mml_FRASE1014"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6120
         TabIndex        =   2
         Top             =   2130
         Width           =   2085
      End
      Begin VB.CommandButton cmdAgrupar 
         Caption         =   "mml_FRASE1013"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6120
         TabIndex        =   1
         Top             =   1410
         Width           =   2085
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gGrupos 
         Height          =   8085
         Left            =   8640
         TabIndex        =   12
         Top             =   690
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   14261
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblNumPar 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   6600
         TabIndex        =   16
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "mml_FRASE1244"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9030
         TabIndex        =   15
         Top             =   420
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE1243"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   540
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "mml_FRASE1018"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9030
         TabIndex        =   4
         Top             =   150
         Width           =   3285
      End
      Begin VB.Label Label1 
         Caption         =   "mml_FRASE1017"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   210
         Width           =   3285
      End
   End
   Begin VB.Label Label13 
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
      Left            =   1080
      TabIndex        =   10
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmAgrupar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sGruposSel As String


Private Sub cmdAgrupar_Click()
Dim iFilaIni As Integer, iFilaFin As Integer
Dim sNombre As String, sGrupos As String, sMod As String, sCat As String, sCatGrupo As String
Dim sCompMod As String, sCompCat As String
Dim iCont As Integer, iCodMod As Integer
Dim sDenomGrupo As String, sDenomCateg As String
Dim sGrupoAnt As String, sCategAnt As String
Dim sGrupo As String, sCateg As String
Dim sNombreGrupo As String

sDenomGrupo = ""
sDenomCateg = ""
sGrupoAnt = ""
sCategAnt = ""
sGrupo = ""
sCateg = ""

If Not C_DEBUG Then On Local Error GoTo error
    gNoAgrupados.Visible = False
    With gNoAgrupados
    
        iCont = 0
        For i = 0 To gNoAgrupados.Rows - 1
            .Row = i
            .Col = 6
            If .Text = "   X" Then
                .Col = 1 ' Categor?a
                sCateg = .Text
                If sCatGrupo <> "" Then sCatGrupo = sCatGrupo & "+"
                sCatGrupo = sCatGrupo & .Text
                sCat = .Text
                .Col = 2 ' Grupo de edad
                sGrupo = .Text
                If sNombre <> "" Then sNombre = sNombre & "+"
                sNombre = sNombre & .Text
                .Col = 5
                If sGrupos <> "" Then sGrupos = sGrupos & ","
                sGrupos = sGrupos & "|" & sCat & "@" & .Text & "|"
                .Col = 3
                iCont = iCont + Val(.Text)
            
                'Descripci?n unificada
                'Si no tenemos categoria la grabamos
                'en caso contrario
                ' Si la categor?a es la misma que la anterior
                '  Comprobamos si el grupo de edad es el mismo
                '   En caso afirmativo continuamos
                '  Si el grupo es distinto anexamos el grupo
                ' Si la categor?a es distinta
                '  Comprobamos si el grupo es el mismo
                '   En este caso solo anexamos la categor?a
                ' Si el grupo es distinto
                '  Generamos la cadena y comenzamos de nuevo con el nuevo grupo y categor?a
                If sCategAnt = "" Then
                    sDenomCateg = sCateg
                    sDenomGrupo = sGrupo
                ElseIf sCategAnt = sCateg Then
                    If sGrupoAnt <> sGrupo Then
                        'Anexamos el grupo
                        If sDenomGrupo <> "" Then sDenomGrupo = sDenomGrupo & "+"
                        sDenomGrupo = sDenomGrupo & sGrupo
                    End If
                Else
                    If sGrupoAnt = sGrupo Then 'Categor?a distinta pero mismo grupo
                        'Solo anexamos la categor?a si en el grupo anterior solo hay un grupo
                        If InStr(sDenomGrupo, "+") = 0 Then
                            If sDenomCateg <> "" Then sDenomCateg = sDenomCateg & "+"
                            sDenomCateg = sDenomCateg & sCateg
                        Else
                            'Inicializamos las variables
                            GoTo inicializar
                        End If
                    Else 'Categor?a distinta y distinto grupo
inicializar:
                        'Componemos parte del nombre y asociamos grupo y categor?a de nuevo
                        sDenom = "(" & sDenomCateg & " " & sDenomGrupo & ")"
                        sDenomGrupo = sGrupo
                        sDenomCateg = sCateg
                    
                        If sNombreGrupo <> "" Then sNombreGrupo = sNombreGrupo & " + "
                        sNombreGrupo = sNombreGrupo & sDenom
                        sDenom = ""
                    End If
                End If
                sCategAnt = sCateg
                sGrupoAnt = sGrupo
                
                'Comprobamos si queremos agrupar categorias o modalidades distintas
                .Col = 0
                sMod = .Text
                If sCompMod = "" Then
                    sCompMod = sMod
                ElseIf sCompMod <> sMod Then
                    If MsgBox(mml_FRASE1011, vbYesNo Or vbCritical, G_MSG_ERROR) = vbNo Then
                        gNoAgrupados.Visible = True
                        Exit Sub
                    End If
                End If
                .Col = 1
                sCat = .Text
                If sCompCat = "" Then
                    sCompCat = sCat
                ElseIf sCompCat <> sCat Then
                    'If MsgBox(mml_FRASE1011, vbYesNo Or vbCritical, G_MSG_ERROR) = vbNo Then
                    '    gNoAgrupados.Visible = True
                    '    Exit Sub
                    'End If
                End If
                .Col = 4
                iCodMod = Val(.Text)
            End If
        Next
        If sDenomCateg <> "" Then
            sDenom = "(" & sDenomCateg & " " & sDenomGrupo & ")"
            
            If sNombreGrupo <> "" Then sNombreGrupo = sNombreGrupo & " + "
            sNombreGrupo = sNombreGrupo & sDenom
        End If
        
        
        If iCont > 0 Then
            'gGrupos.AddItem sMod & vbTab & sCat & vbTab & sNombre & vbTab & iCont & vbTab & sGrupos
            db.Execute "INSERT INTO agrupaciones VALUES (" & tbCodComp.Text & ",""" & sMod & """,""" & sCatGrupo & """,""" & sNombre & """," & iCont & ",""" & sGrupos & """," & iCodMod & ",""" & sNombreGrupo & """)"
        End If
    End With
    gNoAgrupados.Visible = False
    gGrupos.Visible = False
    'LeerNoAgrupados
    LeerGrupos
    gNoAgrupados.Visible = True
    gNoAgrupados.Visible = True
    gGrupos.Visible = True
    Exit Sub
error:
    ProcesarError "cmdAgrupar_Click"
End Sub

Private Sub cmdBorrarGrupos_Click()
If C_DEBUG Then On Local Error GoTo error
    If MsgBox(mml_FRASE1012, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        If tbCodComp.Text <> "" Then
            db.Execute "DELETE FROM agrupaciones WHERE cod_competicion = " & tbCodComp.Text
            
            MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
        End If
    End If
    LeerGrupos
    Exit Sub
error:
    ProcesarError "cmdBorrarGrupos_Click"
End Sub

Private Sub cmdDesagrupar_Click()
Dim sMod As String, sCat As String, sGrupos As String
Dim iFilaIni As Integer, iFilaFin As Integer

If Not C_DEBUG Then On Local Error GoTo error
    If gGrupos.Rows > 0 Then
        If gGrupos.RowSel >= 0 And Val(tbCodComp.Text) > 0 Then
            If gGrupos.Row > gGrupos.RowSel Then
                iFilaIni = gGrupos.RowSel
                iFilaFin = gGrupos.Row
            Else
                iFilaIni = gGrupos.Row
                iFilaFin = gGrupos.RowSel
            End If
            
            For i = iFilaIni To iFilaFin
                gGrupos.Row = i
                gGrupos.Col = 0
                sMod = gGrupos.Text
                gGrupos.Col = 1
                sCat = gGrupos.Text
                gGrupos.Col = 2
                sGrupos = gGrupos.Text
                
                db.Execute "DELETE FROM agrupaciones WHERE cod_competicion = " & tbCodComp.Text & " AND modalidad= """ & sMod & """ AND categoria = """ & sCat & """ AND grupos = """ & sGrupos & """"
            Next
            gNoAgrupados.Visible = False
            gGrupos.Visible = False
            LeerGrupos
            gNoAgrupados.Visible = True
            gGrupos.Visible = True
            
        End If
    End If
    Exit Sub
error:
    ProcesarError "cmdDesagrupar_Click"
End Sub

Private Sub cmdGenerar_Click()
Dim rsGrupos As Recordset, rsPar As Recordset, rsDorsal As Recordset
Dim sNombre As String
Dim iMod As Integer
Dim lCodGrupo As Long, lCodDorsal As Long
Dim iNumDorsal As Integer, iNumDorsales As Integer, iDorsal As Integer
Dim sSQL As String
Dim aGrupos(30) As String, iNumGrupos As Integer
Dim aCat(2) As String

If Not C_DEBUG Then On Local Error GoTo error
    If Val(tbCodComp.Text) > 0 Then
        If MsgBox(mml_FRASE0322, vbYesNo Or vbCritical, mml_FRASE0084) = vbYes Then
            BorrarCompeticion Val(tbCodComp.Text), False, False, False
            
            iNumDorsal = iMinDorsalOficial(tbCodComp.Text)
            Set rsGrupos = db.OpenRecordset("SELECT * FROM agrupaciones WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY cod_modalidad", dbOpenSnapshot)
            While Not rsGrupos.EOF
                'sNombre = Left$(rsGrupos!modalidad, 3) & " " & Trim$(rsGrupos!categoria) & " " & Trim$(rsGrupos!grupos)
                sNombre = Left$(rsGrupos!modalidad, 3) & " " & Trim$(rsGrupos!DESCRIPCION)
                'Insertamos la nueva categoria
                lCodGrupo = MaxCod("categorias")
                db.Execute "INSERT INTO categorias VALUES (" & lCodGrupo & ", """ & sNombre & """, """ & PrimeraCategoria(rsGrupos!categoria) & """," & PrimerGrupo(rsGrupos!cod_grupos) & "," & tbCodComp.Text & "," & rsGrupos!cod_modalidad & ",""10:00"",0,0," & VarCfg("max_dorsales_tanda") & ",0," & ImpHojaUnica & ")"
                'Calculamos la fase
                iNumDorsales = rsGrupos.Fields("cont")
                If iNumDorsales <= 7 Then
                    iFase = 1
                ElseIf iNumDorsales <= 13 Then
                    iFase = 2
                Else
                    iFase = 2 ^ (Int(Log((iNumDorsales - 1) / 6) / Log(2)) + 1)
                End If
                'Ahora insertamos todas las parejas de la categoria
                iNumGrupos = DividirCampo(rsGrupos!cod_grupos, aGrupos, ",", "|")
                Dim i As Integer, iTemp As Integer
                For i = 0 To iNumGrupos - 1
                    iTemp = DividirCampo(aGrupos(i), aCat, "@")
                    Set rsPar = db.OpenRecordset("SELECT * FROM parejas WHERE cod_competicion = " & tbCodComp.Text & " AND cod_modalidad = " & rsGrupos!cod_modalidad & " AND categoria = '" & aCat(0) & "' AND cod_grupoedad = " & aCat(1) & " ORDER BY cod_grupoedad, nombre_hombre", dbOpenSnapshot)
                        If Not rsPar.EOF Then
                            While Not rsPar.EOF
                                lCodDorsal = MaxCod("dorsales")
                                
                                ' Si esta pareja ya tiene dorsal en otra modalidad
                                sSQL = "SELECT num_dorsal FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND cod_categoria < " & lCodGrupo & " AND cod_modalidad <= " & rsGrupos!cod_modalidad & " AND cod_competicion = " & tbCodComp.Text & " AND ((nif_hombre <> '' AND nif_hombre = '" & rsPar!nif_hombre & "' AND nif_mujer = '" & rsPar!nif_mujer & "') OR (nombre_hombre <> '' AND nombre_hombre ='" & rsPar!nombre_hombre & "' AND nombre_mujer = '" & rsPar!nombre_mujer & "'))"
                                Debug.Print sSQL
                                Set rsDorsal = db.OpenRecordset(sSQL, dbOpenSnapshot)
                                If Not rsDorsal.EOF Then
                                    iDorsal = rsDorsal!num_dorsal
                                Else
                                    iDorsal = iNumDorsal
                                    Inc iNumDorsal
                                End If
                                    
                                rsDorsal.Close
                                
                                db.Execute "INSERT INTO dorsales VALUES (" & lCodDorsal & "," & iDorsal & "," & lCodGrupo & "," & iFase & "," & rsPar!codigo & ",0,0)"
                                rsPar.MoveNext
                            Wend
                        End If
                    rsPar.Close
                Next
                rsGrupos.MoveNext
            Wend
            rsGrupos.Close
            MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
       End If
    End If
    
    Exit Sub
error:
    ProcesarError "cmdGenerar_Click"
End Sub

Function PrimeraCategoria(sCad) As String
Dim i As Integer

    i = InStr(sCad, "+")
    If i > 0 Then
        PrimeraCategoria = Mid(sCad, 1, i - 1)
    Else
        PrimeraCategoria = sCad
    End If
    
End Function
Function PrimerGrupo(sCad) As Integer
Dim i As Integer

    i = InStr(sCad, "@")
    sCad = Mid(sCad, i + 1)
        
    PrimerGrupo = Val(sCad)
    
End Function


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)

End Sub

Private Sub Form_Initialize()
    gNoAgrupados.Refresh
    gGrupos.Refresh
    frmEmparejar.Refresh
End Sub

Private Sub Form_Load()
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    TraducirCadenas Me
End Sub


Private Sub gNoAgrupados_Click()
Dim iNumParejas As Integer
    With gNoAgrupados
        If .Rows > 0 Then
            .Col = 3
            iNumParejas = Val(.Text)
            .Col = 6
            If .Text = "   X" Then
                .Text = ""
                lblNumPar.Caption = Val(lblNumPar.Caption) - iNumParejas
            Else
                .Text = "   X"
                lblNumPar.Caption = Val(lblNumPar.Caption) + iNumParejas
            End If
        End If
    End With
End Sub

Private Sub tbCodComp_Change()

    If Val(tbCodComp.Text) > 0 Then
        gGrupos.Cols = 7
        gGrupos.ColWidth(0) = 1000
        gGrupos.ColWidth(1) = 1500
        gGrupos.ColWidth(2) = 2400
        gGrupos.ColWidth(3) = 500
        gGrupos.ColWidth(4) = 500
        gGrupos.ColWidth(5) = 500
        gGrupos.ColWidth(6) = 4000
        
        With gNoAgrupados
            .Clear
            .Cols = 7
            .Rows = 0
            .ColWidth(0) = 1000
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 500
            .ColWidth(4) = 500
            .ColWidth(5) = 500
            .ColWidth(6) = 500
            LeerGrupos
        End With
    End If
End Sub

Private Sub LeerNoAgrupados()
Dim sSQL As String
Dim rs As Recordset
Dim bAgrupada As Boolean

    lblNumPar.Caption = "0"
    If Val(tbCodComp.Text) > 0 Then
        
        
        With gNoAgrupados
            .Clear
            .Rows = 0

            sSQL = "SELECT DISTINCT COUNT(*) as cont, grupoedad, cod_grupoedad, cod_modalidad, categoria, m.nombre, g.abreviatura FROM parejas p, modalidad m, gruposedad g WHERE g.codigo = p.cod_grupoedad AND p.cod_modalidad = m.codigo AND cod_competicion = " & Val(tbCodComp.Text) & " GROUP BY grupoedad, abreviatura, cod_grupoedad, cod_modalidad, categoria, m.nombre ORDER BY cod_modalidad, categoria, cod_grupoedad"
            Set rs = db.OpenRecordset("SELECT DISTINCT COUNT(*) as cont, grupoedad, cod_grupoedad, cod_modalidad, categoria, m.nombre, g.abreviatura, cod_modalidad FROM parejas p, modalidad m, gruposedad g WHERE g.codigo = p.cod_grupoedad AND p.cod_modalidad = m.codigo AND cod_competicion = " & Val(tbCodComp.Text) & " GROUP BY grupoedad, abreviatura, cod_grupoedad, cod_modalidad, categoria, m.nombre ORDER BY cod_modalidad, categoria, cod_grupoedad", dbOpenSnapshot)
            While Not rs.EOF
                bAgrupada = False
                'Comprobamos que no se encuentre agrupado
                For i = 0 To gGrupos.Rows - 1
                    With gGrupos
                        .Row = i
                        .Col = 0
                        If Trim$(.Text) <> rs!Nombre Then GoTo seguir
                        .Col = 4
                        If InStr(Trim$(.Text), "|" & rs!categoria & "@" & rs!cod_grupoedad & "|") = 0 Then GoTo seguir
                        'La categoria esta agrupada
                        bAgrupada = True
                    End With
seguir:
                Next
                If Not bAgrupada Then
                    .AddItem rs!Nombre & vbTab & rs!categoria & vbTab & rs!abreviatura & vbTab & rs!Cont & vbTab & rs!cod_modalidad & vbTab & rs!cod_grupoedad
                    .Row = .Rows - 1
                    .Col = 0
                    If rs!Cont >= 5 Then
                        .CellBackColor = vbGreen
                    Else
                        .CellBackColor = vbYellow
                    End If
                End If
                rs.MoveNext
            Wend
            rs.Close
        End With
    End If
End Sub

Private Sub LeerGrupos()
Dim rs As Recordset

    If Val(tbCodComp.Text) > 0 Then
        
        With gGrupos
        .Clear
        .Rows = 0

            
            Set rs = db.OpenRecordset("SELECT * FROM agrupaciones WHERE cod_competicion =" & tbCodComp.Text & " ORDER BY modalidad, categoria, grupos", dbOpenSnapshot)
            While Not rs.EOF
                    .AddItem rs!modalidad & vbTab & rs!categoria & vbTab & rs!grupos & vbTab & rs!Cont & vbTab & rs!cod_grupos & vbTab & rs!cod_modalidad & vbTab & rs!DESCRIPCION
                    .Row = .Rows - 1
                    .Col = 0
                    If rs!Cont >= 5 Then
                        .CellBackColor = vbGreen
                    Else
                        .CellBackColor = vbYellow
                    End If
                
                rs.MoveNext
            Wend
            rs.Close
        End With
    End If
    LeerNoAgrupados
End Sub
