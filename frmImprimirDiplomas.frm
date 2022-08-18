VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimirDiplomas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1076"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   11100
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   270
      Top             =   1110
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   10905
      Begin VB.CheckBox chkImpGrupo 
         Caption         =   "mml_FRASE1276"
         Height          =   315
         Left            =   1830
         TabIndex        =   15
         Top             =   2520
         Width           =   8655
      End
      Begin VB.CheckBox chkModClasificacion 
         Caption         =   "mml_FRASE1234"
         Height          =   315
         Left            =   1830
         TabIndex        =   14
         Top             =   2210
         Width           =   8655
      End
      Begin VB.CheckBox chkOrdenarModalidad 
         Caption         =   "mml_FRASE1226"
         Height          =   315
         Left            =   1830
         TabIndex        =   13
         Top             =   1900
         Width           =   2235
      End
      Begin VB.CheckBox chkUnoPorFichero 
         Caption         =   "mml_FRASE1138"
         Height          =   315
         Left            =   1830
         TabIndex        =   12
         Top             =   1590
         Width           =   2235
      End
      Begin VB.TextBox tbCodCateg 
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   5
         TabIndex        =   10
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox tbDescCateg 
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
         MaxLength       =   35
         TabIndex        =   9
         Top             =   630
         Width           =   7650
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
         Height          =   375
         Left            =   1830
         Picture         =   "frmImprimirDiplomas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   630
         Width           =   465
      End
      Begin VB.ComboBox tbCat 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "frmImprimirDiplomas.frx":046A
         Left            =   1830
         List            =   "frmImprimirDiplomas.frx":04B0
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   4845
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
         Height          =   360
         Left            =   1830
         Picture         =   "frmImprimirDiplomas.frx":0523
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   450
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
         TabIndex        =   4
         Top             =   240
         Width           =   7680
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
         TabIndex        =   3
         Top             =   240
         Width           =   855
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
         Left            =   120
         TabIndex        =   11
         Top             =   630
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
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
      Height          =   495
      Left            =   5250
      TabIndex        =   1
      Top             =   3360
      Width           =   2085
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "mml_FRASE1076"
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
      Left            =   2610
      TabIndex        =   0
      Top             =   3360
      Width           =   2385
   End
End
Attribute VB_Name = "frmImprimirDiplomas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ImprimirCabecera()
Dim rs As Recordset
Dim X As Integer, Y As Integer

    Printer.PaintPicture frmMenu.picEPA.Picture, Printer.Width - frmMenu.picEPA.Width - G_MARGEN_EPA, G_MARGEN_EPA_Y
    If G_LOGO_ESCUELA <> "" Then
        Printer.PaintPicture LoadPicture(G_LOGO_ESCUELA), G_MARGEN_EPA, G_MARGEN_EPA_Y
    End If
    Printer.CurrentX = 0
    Set rs = db.OpenRecordset(" SELECT descripcion, fecha FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    Printer.DrawWidth = 2
    X = Printer.CurrentX
    Y = Printer.CurrentY
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    SaltoLinea Printer, 4
    Printer.FontBold = True
    Printer.FontSize = 13
    Centrado Printer, rs!DESCRIPCION & "  (" & rs!fecha & ")", Printer.Width
    rs.Close
    Printer.FontBold = False
    Printer.FontSize = 13
    Centrado Printer, sEscuela(tbCodComp.Text), Printer.Width
    SaltoLinea Printer, 4
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    Printer.DrawWidth = 1
    SaltoLinea Printer, 10
End Sub

Private Sub ImprimirDiplomasConTabla()

Dim rs As Recordset
Dim rsCont As Recordset
Dim rsBailes As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iNumDorsales As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(6) As TCelda
Dim iPag As Integer
Dim bDescanso As Boolean
Dim iCols As Integer
Dim sPart As String, sPart1 As String
Dim sHombre As String
Dim sMujer As String

    If MsgBox(mml_FRASE0670, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    If Not C_DEBUG Then On Local Error GoTo error
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        icFilasTabla = 1
        iCols = 3
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 3)
        
        aTabla(0, 0) = mml_FRASE0673 ' Categoria
        aTabla(0, 1) = mml_FRASE0300 ' Dorsal
        aTabla(0, 2) = mml_FRASE0372 ' posición
        
        aDefCelda(0).Ancho = 5000
        aDefCelda(0).Justificado = ecizquierda
        aDefCelda(1).Ancho = 1000
        aDefCelda(1).Justificado = ecizquierda
        aDefCelda(2).Ancho = 1000
        aDefCelda(2).Justificado = eccentro
        
        Set rs = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer , c.descripcion, rs.dorsal, rs.posicion, m.nombre, rs.cod_categoria FROM modalidad m, dorsales d, parejas p, resumenfinales rs, categorias c WHERE c.cod_modalidad = m.codigo AND c.codigo = d.cod_categoria AND p.cod_competicion = " & tbCodComp.Text & " AND  p.codigo = d.cod_pareja AND d.cod_categoria = rs.cod_categoria AND d.num_dorsal = rs.dorsal ORDER BY 1,2", dbOpenSnapshot)
        While Not rs.EOF
            sHombre = rs.Fields("nombre_hombre")
            sMujer = rs.Fields("nombre_mujer")
            On Local Error Resume Next
            Printer.PaintPicture LoadPicture(G_IMAGEN_FONDO_DIPLOMA), G_MARGEN_IZQ_IMAGEN_DIPLOMAS, G_MARGEN_SUP_IMAGEN_DIPLOMAS, G_ANCHO_IMAGEN_DIPLOMAS, G_ALTO_IMAGEN_DIPLOMAS
            If Not C_DEBUG Then On Local Error GoTo error
            Printer.CurrentY = G_MARGEN_SUP_TABLA_DIPLOMAS
            Printer.CurrentX = G_MARGEN_IZQ_TABLA_DIPLOMAS
            Printer.FontName = G_DIPLOMA_TITULO_FUENTE
            Printer.FontSize = 24
            Printer.Print mml_FRASE1078
            Printer.Print
            Printer.FontName = "Times New Roman"
            Printer.FontSize = 16
            Printer.FontBold = True
            Printer.CurrentX = G_MARGEN_IZQ_TABLA_DIPLOMAS
            Printer.Print sDescCompeticion(CodCompActiva)
            Printer.CurrentX = G_MARGEN_IZQ_TABLA_DIPLOMAS
            Printer.FontBold = False
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.CurrentX = G_MARGEN_IZQ_TABLA_DIPLOMAS
            Printer.FontBold = True
            Printer.Print mml_FRASE1079
            Printer.FontBold = False
            Printer.CurrentX = G_MARGEN_IZQ_TABLA_DIPLOMAS
            If Len(sHombre) > 3 Then
                Printer.CurrentX = G_MARGEN_IZQ_TABLA_DIPLOMAS
                Printer.Print sHombre
            End If
            If Len(sMujer) > 3 Then
                Printer.CurrentX = G_MARGEN_IZQ_TABLA_DIPLOMAS
                Printer.Print sMujer
            End If
            Printer.Print
            Printer.Print
            Printer.Print
            
            aTabla(icFilasTabla, 0) = ".12b" & rs.Fields("nombre")
            aTabla(icFilasTabla, 1) = ""
            aTabla(icFilasTabla, 2) = ""
            icFilasTabla = icFilasTabla + 1
            aTabla(icFilasTabla, 0) = rs.Fields("descripcion")
            aTabla(icFilasTabla, 1) = rs.Fields("dorsal")
            aTabla(icFilasTabla, 2) = rs.Fields("posicion")
                        
            'Ahora recuperamos la información de los bailes
            Set rsBailes = db.OpenRecordset("SELECT r.dorsal, r.posicion, b.nombre FROM resultadosfinales r, bailes b WHERE r.cod_baile = b.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_categoria = " & rs.Fields("cod_categoria") & " AND dorsal = " & rs.Fields("dorsal") & " ORDER BY cod_baile", dbOpenSnapshot)
            While Not rsBailes.EOF
                icFilasTabla = icFilasTabla + 1
                aTabla(icFilasTabla, 0) = "      -> " & rsBailes.Fields("nombre")
                aTabla(icFilasTabla, 1) = rsBailes.Fields("dorsal")
                aTabla(icFilasTabla, 2) = rsBailes.Fields("posicion")
                rsBailes.MoveNext
            Wend
            rsBailes.Close
            
            DibujarTablaA Printer, Printer.CurrentX + G_MARGEN_IZQ_TABLA_DIPLOMAS, Printer.CurrentY, icFilasTabla + 1, iCols, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
            Printer.NewPage
            icFilasTabla = 1
            
            rs.MoveNext
        Wend
        rs.Close
            
        Printer.EndDoc
        MsgBox mml_FRASE1009, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    Next iCCopias
    Exit Sub
error:
    ProcesarError "cmdCambiar_Click-frmImprimirDiplomas"
End Sub

Private Sub ImprimirDiplomasSinTabla()

Dim rs As Recordset
Dim rsCont As Recordset
Dim rsBailes As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iNumDorsales As Integer
Dim icFilasTabla As Integer
Dim iPag As Integer
Dim bDescanso As Boolean
Dim iCols As Integer
Dim sPart As String, sPart1 As String
Dim sHombre As String
Dim sMujer As String
Dim sModalidad As String
Dim sGrupoEdad As String
Dim sIdCategoria As String
Dim sCompeticion As String
Dim sCategoria As String
Dim dFecha As Date
Dim i As Integer
Dim iPos As Integer
Dim sCad As String
Dim sSelec As String
Dim sPuesto As String
Dim sOrden As String
Dim sSQL As String

    If MsgBox(mml_FRASE0670, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    If Not C_DEBUG Then On Local Error GoTo error
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        
        If Val(tbCodCateg.Text) > 0 Then
            sSelec = "c.codigo = " & tbCodCateg.Text & " AND "
        End If
        If tbCat.Text <> "" Then
            sSelec = sSelec & "c.descripcion LIKE '*" & tbCat.Text & "*' AND "
        End If
        
        Dim sSalon As String
        
        If C_COUNTRY Then
            sSalon = ""
        Else
            sSalon = " AND d.fase = 1 "
        End If
        
        Dim sImpGrupo As String
        Dim iNumGrupo As Integer
        
        sImpGrupo = ""
        If chkImpGrupo.Value Then
            iNumGrupo = Val(InputBox(mml_FRASE1277))
            If iNumGrupo > 0 Then
                sImpGrupo = " AND " & iNumGrupo & " = (SELECT MAX(num_grupo) FROM horario h WHERE d.cod_categoria = h.cod_categoria and h.numfase = 1 AND h.cod_competicion = c.cod_competicion) "
            End If
        End If
        
        If chkModClasificacion.Value Then ' Imprimir por modalidad y clasificación final
            Dim sDesc As String
            'Primero comprobamos que todas las categorías están en la tabla de categorías
            
            Set rs = db.OpenRecordset("SELECT DISTINCT id_categoria " & _
                    "FROM categorias c , competiciones co " & _
                    "WHERE co.codigo = " & tbCodComp.Text & " AND c.cod_competicion = co.codigo AND NOT (c.id_categoria IN (SELECT descripcion FROM DescCategoria))", dbOpenSnapshot)
            While Not rs.EOF
                sDesc = sDesc & vbCrLf & ", " & rs.Fields(0)
                rs.MoveNext
            Wend
            If sDesc <> "" Then
                MsgBox mml_FRASE1235 & vbCrLf & sDesc, vbOKOnly Or vbCritical, "ERROR"
                Exit Sub
            End If
            
            sSQL = "SELECT DISTINCT m.orden, dcat.orden, ge.orden, p.nombre_hombre, p.nombre_mujer, c.descripcion, rs.dorsal, rs.posicion, m.nombre, " & _
                " rs.cod_categoria, ge.nombre AS grupo_edad, c.id_categoria, co.fecha " & _
                " FROM gruposedad AS ge, modalidad AS m, dorsales AS d, parejas AS p, resumenfinales AS rs, categorias AS c, DescCategoria AS dcat, " & _
                " competiciones AS co " & _
                " Where co.codigo = C.cod_competicion And C.cod_grupoedad = ge.codigo And C.cod_modalidad = m.codigo And C.codigo = d.cod_categoria " & _
                " And p.cod_competicion = " & tbCodComp.Text & " And p.codigo = d.cod_pareja And d.cod_categoria = rs.cod_categoria And C.id_categoria = dcat.DESCRIPCION " & _
                " And d.num_dorsal = rs.Dorsal AND rs.posicion <> 'NO PRESENTADO'" & _
                sSalon & sImpGrupo & " ORDER BY m.orden, dcat.orden, ge.orden, c.descripcion, rs.posicion"
            Debug.Print sSQL
            
            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
        Else
            sOrden = "c.descripcion, rs.posicion"
            
            Set rs = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer , c.descripcion, rs.dorsal, " & _
                    "rs.posicion, m.nombre, rs.cod_categoria, ge.nombre AS grupo_edad, c.id_categoria, co.fecha " & _
                    "FROM gruposedad ge, modalidad m, dorsales d, parejas p, resumenfinales rs, categorias c, " & _
                    "competiciones co " & _
                    "WHERE " & sSelec & " co.codigo = c.cod_competicion AND c.cod_grupoedad = ge.codigo AND " & _
                    "c.cod_modalidad = m.codigo AND c.codigo = d.cod_categoria AND p.cod_competicion = " & tbCodComp.Text & _
                    " AND p.codigo = d.cod_pareja AND d.cod_categoria = rs.cod_categoria AND " & _
                    "d.num_dorsal = rs.dorsal " & sSalon & sImpGrupo & _
                    "ORDER BY " & sOrden, dbOpenSnapshot)
        End If
        
        Dim iNumeroPuestos As Integer
        Dim iPrimerPuesto As Integer
        
        'En salón se imprimen diplomas de todos los participantes de la final
        If C_COUNTRY Then
            iNumeroPuestos = G_PUESTOS_CON_DIPLOMA
            iPrimerPuesto = 1
        Else
            iNumeroPuestos = 99
            iPrimerPuesto = G_DIPLOMAS_PRIMER_DORSAL_SALON
        End If
        
        While Not rs.EOF
            sPuesto = rs.Fields("posicion")
            If Val(sPuesto) >= iPrimerPuesto And Val(sPuesto) <= iNumeroPuestos Then
                sHombre = rs.Fields("nombre_hombre")
                sMujer = rs.Fields("nombre_mujer")
                On Local Error Resume Next
                Printer.Orientation = vbPRORLandscape
    
                Printer.PaintPicture LoadPicture(G_IMAGEN_FONDO_DIPLOMA), G_MARGEN_IZQ_IMAGEN_DIPLOMAS, G_MARGEN_SUP_IMAGEN_DIPLOMAS, G_ANCHO_IMAGEN_DIPLOMAS, G_ALTO_IMAGEN_DIPLOMAS
                If Not C_DEBUG Then On Local Error GoTo error
                Printer.CurrentY = G_MARGEN_SUP_TABLA_DIPLOMAS
                Printer.FontName = G_DIPLOMA_TITULO_FUENTE
                Printer.FontName = "Times New Roman"
                sCompeticion = sDescCompeticion(CodCompActiva)
                dFecha = rs.Fields("fecha")
                'Printer.FontBold = True
                'sCad = sDescCompeticion(CodCompActiva)
                'i = Printer.Width / 2 - Printer.TextWidth(sCad) / 2
                'Printer.CurrentX = i
                'Printer.Print sCad
                'Printer.Print
                'Printer.FontBold = True
                
                'sCad = mml_FRASE1079
                'i = Printer.Width / 2 - Printer.TextWidth(sCad) / 2
                'Printer.CurrentX = i
                'Printer.Print sCad
                Printer.FontSize = 22
                Printer.FontBold = False
                If Len(sHombre) > 3 Then
                    sCad = sHombre
                    i = Printer.Width / 2 - Printer.TextWidth(sCad) / 2
                    Printer.CurrentX = i
                    Printer.Print sCad
                End If
                If Len(sMujer) > 3 Then
                    sCad = sMujer
                    i = Printer.Width / 2 - Printer.TextWidth(sCad) / 2
                    Printer.CurrentX = i
                    Printer.Print sCad
                End If
                Printer.Print
                
                sModalidad = rs.Fields("nombre")
                i = InStr(sModalidad, " ")
                If i > 0 Then
                    sModalidad = Mid$(sModalidad, i + 1)
                End If
                sGrupoEdad = rs.Fields("grupo_edad")
                sIdCategoria = rs.Fields("id_categoria")
                sCad = Trim$(rs.Fields("posicion"))
                sCategoria = Trim$(rs.Fields("descripcion"))
                If C_COUNTRY Or G_IDIOMA = C_INGLES Then
                    sCad = sCad & "th Place "
                Else
                    sCad = sCad & "ª Posición en "
                End If
                If C_COUNTRY Then
                    sCad = sCad & sModalidad & " " & sIdCategoria & " " & sGrupoEdad
                Else
                    sCad = sCad & sCategoria
                End If
                
                Printer.FontSize = 20
                Printer.FontBold = True
                i = Printer.Width / 2 - Printer.TextWidth(sCad) / 2
                Printer.CurrentX = i
                Printer.Print sCad
                                        
                Printer.FontSize = 17
                Printer.FontBold = False
                'Ahora recuperamos la información de los bailes
                Set rsBailes = db.OpenRecordset("SELECT r.dorsal, r.posicion, b.nombre FROM resultadosfinales r, bailes b WHERE r.cod_baile = b.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_categoria = " & rs.Fields("cod_categoria") & " AND dorsal = " & rs.Fields("dorsal") & " ORDER BY cod_baile", dbOpenSnapshot)
                i = 0
                sCad = ""
                While Not rsBailes.EOF
                    If sCad <> "" Then
                        sCad = sCad & " - "
                    End If
                    sCad = sCad & Trim$(rsBailes.Fields("posicion"))
                    If C_COUNTRY Or G_IDIOMA = C_INGLES Then
                        sCad = sCad & "th Place in Dance "
                    Else
                        sCad = sCad & "ª Posición en el baile "
                    End If
                    sCad = sCad & rsBailes.Fields("nombre")
                    i = i + 1
                    If i = 3 Then
                        iPos = Printer.Width / 2 - Printer.TextWidth(sCad) / 2
                        Printer.CurrentX = iPos
                        Printer.Print sCad
                        i = 0
                        sCad = ""
                    End If
                    rsBailes.MoveNext
                Wend
                rsBailes.Close
                If sCad <> "" Then
                    iPos = Printer.Width / 2 - Printer.TextWidth(sCad) / 2
                    Printer.CurrentX = iPos
                    Printer.Print sCad
                    sCad = ""
                End If
                
                ' Imprimimos el nombre de la competición
                Printer.FontBold = False
                Printer.FontSize = 18
                
                If G_MARGEN_SUP_NOMBRE_COMP > 0 Then
                    Printer.CurrentY = G_MARGEN_SUP_NOMBRE_COMP
                    
                    Printer.CurrentX = G_MARGEN_IZQ_NOMBRE_COMP
                    Printer.Print sCompeticion
                    Printer.CurrentX = G_MARGEN_IZQ_NOMBRE_COMP
                    Printer.Print sMes(Month(dFecha)) & ", " & Year(dFecha)
                End If
                If chkUnoPorFichero.Value = 1 Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
            rs.MoveNext
        Wend
        rs.Close
            
        Printer.EndDoc
        MsgBox mml_FRASE1009, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    Next iCCopias
    Exit Sub
error:
    ProcesarError "ImprimirDiplomasSinTabla"
End Sub



Private Sub cmdCambiar_Click()
    ImprimirDiplomasSinTabla
End Sub

Private Sub cmdCateg_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCateg.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCateg.Text = sResultado(2)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
End Sub

Private Sub Form_Load()
Dim rs As Recordset

    'Actualizamos las categorías
    Set rs = db.OpenRecordset("SELECT * FROM desccategoria ORDER BY 1", dbOpenSnapshot)
    tbCat.Clear
    While Not rs.EOF
        tbCat.AddItem rs!DESCRIPCION
        rs.MoveNext
    Wend
    rs.Close
    
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
End Sub

Private Sub tbCodCateg_GotFocus()
    tbCodCateg.SelStart = 0
    tbCodCateg.SelLength = Len(tbCodCateg.Text)
End Sub

Private Sub tbCodCateg_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbCodCateg_LostFocus()
Dim sCateg As String
    
    If Not C_DEBUG Then On Local Error GoTo error
    If Val(tbCodCateg.Text) > 0 Then
        sCateg = sDescCategoria(tbCodCateg.Text)
        If sCateg = "" Then
            tbCodCateg.Text = ""
            tbDescCateg.Text = ""
        Else
            tbDescCateg.Text = sCateg
        End If
    Else
        tbCodCateg.Text = ""
        tbDescCateg.Text = ""
    End If
    Exit Sub
error:
    ProcesarError "tbCodCategCopia_LostFocus"
End Sub

