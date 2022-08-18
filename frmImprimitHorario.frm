VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimirHorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0712"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "mml_FRASE0437"
      Height          =   780
      Left            =   7065
      TabIndex        =   7
      Top             =   90
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
         ItemData        =   "frmImprimitHorario.frx":0000
         Left            =   150
         List            =   "frmImprimitHorario.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdHTML 
      Caption         =   "-> HTML"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   960
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6915
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
         Left            =   1845
         Picture         =   "frmImprimitHorario.frx":004C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
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
         Width           =   3720
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
      Left            =   4890
      TabIndex        =   1
      Top             =   960
      Width           =   2085
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "mml_FRASE0712"
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
      TabIndex        =   0
      Top             =   960
      Width           =   2385
   End
End
Attribute VB_Name = "frmImprimirHorario"
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

Private Sub cmdCambiar_Click()
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
    
    If G_COUNTRY Then
        iCols = 7
    Else
        iCols = 6
    End If
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 7)
        
        aTabla(0, 0) = mml_FRASE0671
        aTabla(0, 1) = mml_FRASE0672
        aTabla(0, 2) = mml_FRASE0673
        aTabla(0, 3) = mml_FRASE0674
        aTabla(0, 4) = mml_FRASE0713
        aTabla(0, 5) = mml_FRASE0675
        aTabla(0, 6) = mml_FRASE0739
        
        If MsgBox(mml_FRASE0676, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
            aDefCelda(0).Ancho = 800
            aDefCelda(2).Ancho = 4500
        Else
            aDefCelda(0).Ancho = 0
            aDefCelda(2).Ancho = 5300
        End If
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 800
        aDefCelda(1).Justificado = eccentro
        aDefCelda(2).Justificado = ecizquierda
        aDefCelda(3).Ancho = 800
        aDefCelda(3).Justificado = eccentro
        aDefCelda(4).Ancho = 800
        aDefCelda(4).Justificado = eccentro
        aDefCelda(5).Ancho = 3500
        aDefCelda(5).Justificado = ecizquierda
        aDefCelda(6).Ancho = 800
        aDefCelda(6).Justificado = eccentro
        
        'Set rs = db.OpenRecordset("SELECT h.numfase, h.grupo, c.codigo, h.hora, id_categoria, ge.nombre, descripcion, m.nombre as mod, h.num_grupo FROM horario h, categorias c, gruposedad ge, modalidad m WHERE c.descripcion LIKE '*" & cbPista.Text & "*' AND h.cod_categoria = c.codigo AND c.cod_modalidad = m.codigo AND h.cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY h.hora, 5", dbOpenSnapshot)
        Set rs = db.OpenRecordset(" (SELECT h.numfase, h.grupo, c.codigo, h.hora, id_categoria, ge.nombre, descripcion, m.nombre as mod, h.num_grupo FROM horario h, categorias c, gruposedad ge, modalidad m " & _
                                  " WHERE  c.descripcion LIKE '*" & cbPista.Text & "*' AND h.cod_categoria = c.codigo AND c.cod_modalidad = m.codigo AND h.cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo) " & _
                                  " UNION " & _
                                  " (SELECT h.numfase, h.grupo, 0, h.hora, '', '', '', '', h.num_grupo FROM horario h " & _
                                  " WHERE  h.grupo LIKE '*" & cbPista.Text & "*' AND h.cod_categoria = 0 AND h.cod_competicion = " & tbCodComp.Text & ") " & _
                                  " ORDER BY 4, 5", dbOpenSnapshot)
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = rs!codigo
            aTabla(icFilasTabla, 1) = Format(CDate(rs!hora), "hh:mm")
            
            If Left$(rs!grupo, 1) = "*" Then
                aTabla(icFilasTabla, 2) = Left(Mid$(rs!grupo, 2), MAX_GRUPO_HORARIO) 'rs!DESCRIPCION
                aTabla(icFilasTabla, 3) = ""
                aTabla(icFilasTabla, 5) = ""
                aTabla(icFilasTabla, 4) = ""
                bDescanso = True
            Else
                aTabla(icFilasTabla, 2) = Left(rs!grupo, MAX_GRUPO_HORARIO) 'rs!DESCRIPCION
                bDescanso = False
                
                If rs!codigo > 0 Then
                    Set rsCont = db.OpenRecordset(" SELECT DISTINCT num_dorsal FROM dorsales WHERE fase = " & rs!numfase & " AND cod_categoria = " & rs!codigo, dbOpenSnapshot)
                    iNumDorsales = 0
                    If Not rsCont.EOF Then
                        rsCont.MoveLast
                        iNumDorsales = rsCont.RecordCount
                    End If
                    aTabla(icFilasTabla, 3) = IIf(iNumDorsales = 0, "", Str$(iNumDorsales))
                    rsCont.Close
                    
                    aTabla(icFilasTabla, 5) = ""
                    
                    aTabla(icFilasTabla, 4) = " (" & sDescCortaFase(rs!numfase) & ")"
                End If
                
                If rs!numfase > 1 Then
                    Set rsBailes = db.OpenRecordset("SELECT b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & rs!codigo & " AND fase = 2 ORDER BY posicion", dbOpenSnapshot)
                    While Not rsBailes.EOF
                        aTabla(icFilasTabla, 5) = aTabla(icFilasTabla, 5) & sNombreBaileAbreviado(rsBailes!Nombre)
                        rsBailes.MoveNext
                        If Not rsBailes.EOF Then
                            aTabla(icFilasTabla, 5) = aTabla(icFilasTabla, 5) & ", "
                        End If
                    Wend
                    rsBailes.Close
                Else
                    Set rsBailes = db.OpenRecordset("SELECT b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & rs!codigo & " AND fase = 1 ORDER BY posicion", dbOpenSnapshot)
                    While Not rsBailes.EOF
                        aTabla(icFilasTabla, 5) = aTabla(icFilasTabla, 5) & sNombreBaileAbreviado(rsBailes!Nombre)
                        rsBailes.MoveNext
                        If Not rsBailes.EOF Then
                            aTabla(icFilasTabla, 5) = aTabla(icFilasTabla, 5) & ", "
                        End If
                    Wend
                    rsBailes.Close
                End If
            End If
            If G_COUNTRY Then aTabla(icFilasTabla, 6) = rs!num_grupo
            Inc icFilasTabla
            If icFilasTabla >= G_MAX_FILAS_POR_PAG - 2 Then
                Printer.FontBold = False
                Printer.FontSize = 10
                Printer.Print mml_FRASE0659 & iPag
                ImprimirCabecera
                DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, iCols, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                Printer.NewPage
                Inc iPag
                icFilasTabla = 1
            End If
            rs.MoveNext
        Wend
        rs.Close
            
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, iCols, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias
End Sub


Private Sub cmdHTML_Click()
Dim rs As Recordset
Dim rsCont As Recordset
Dim rsBailes As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iNumDorsales As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(5) As TCelda
Dim iPag As Integer
Dim i As Integer, j As Integer
Dim bDescanso As Boolean


    If MsgBox(mml_FRASE0714, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(500, 7)
        
        aTabla(0, 0) = mml_FRASE0261
        aTabla(0, 1) = mml_FRASE0259
        aTabla(0, 2) = mml_FRASE0301
        aTabla(0, 3) = mml_FRASE0715
        aTabla(0, 4) = mml_FRASE0299
        aTabla(0, 5) = mml_FRASE0185
        aTabla(0, 6) = mml_FRASE0739
        
        'Set rs = db.OpenRecordset(" SELECT h.numfase, h.grupo, c.codigo, h.hora, id_categoria, ge.nombre, descripcion, m.nombre as mod, h.num_grupo FROM horario h, categorias c, gruposedad ge, modalidad m WHERE h.cod_categoria = c.codigo AND c.cod_modalidad = m.codigo AND h.cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY h.hora, 5", dbOpenSnapshot)
        Set rs = db.OpenRecordset(" (SELECT h.numfase, h.grupo, c.codigo, h.hora, id_categoria, ge.nombre, descripcion, m.nombre as mod, h.num_grupo FROM horario h, categorias c, gruposedad ge, modalidad m " & _
                                  " WHERE  c.descripcion LIKE '*" & cbPista.Text & "*'  AND h.cod_categoria = c.codigo AND c.cod_modalidad = m.codigo AND h.cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo) " & _
                                  " UNION " & _
                                  " (SELECT h.numfase, h.grupo, 0, h.hora, '', '', '', '', h.num_grupo FROM horario h " & _
                                  " WHERE  h.grupo LIKE '*" & cbPista.Text & "*'  AND h.cod_categoria = 0 AND h.cod_competicion = " & tbCodComp.Text & ") " & _
                                  " ORDER BY 4, 5", dbOpenSnapshot)
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = rs!codigo
            aTabla(icFilasTabla, 1) = Format(CDate(rs!hora), "hh:mm")
            
            If Left$(rs!grupo, 1) = "*" Then
                aTabla(icFilasTabla, 2) = Mid$(rs!grupo, 2)
                bDescanso = True
            Else
                aTabla(icFilasTabla, 2) = rs!grupo
                bDescanso = False
                
                If rs!codigo > 0 Then
                    Set rsCont = db.OpenRecordset(" SELECT DISTINCT num_dorsal FROM dorsales WHERE fase = " & rs!numfase & " AND cod_categoria = " & rs!codigo, dbOpenSnapshot)
                    iNumDorsales = 0
                    If Not rsCont.EOF Then
                        rsCont.MoveLast
                        iNumDorsales = rsCont.RecordCount
                    End If
                    aTabla(icFilasTabla, 3) = IIf(iNumDorsales = 0, "", Str$(iNumDorsales))
                    rsCont.Close
                
                    aTabla(icFilasTabla, 4) = sDescFase(rs!numfase)
                End If
                
                'aTabla(icFilasTabla, 4) = "(dT " & iDorsalesPorTandaCateg(rs!codigo) & ") "
                
                aTabla(icFilasTabla, 5) = ""
                If rs!numfase > 1 Then
                    Set rsBailes = db.OpenRecordset("SELECT b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & rs!codigo & " AND fase = 2 ORDER BY posicion", dbOpenSnapshot)
                    'If Not rsBailes.EOF Then aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & mml_FRASE0677
                    While Not rsBailes.EOF
                        aTabla(icFilasTabla, 5) = aTabla(icFilasTabla, 5) & sNombreBaileAbreviado(rsBailes!Nombre)
                        rsBailes.MoveNext
                        If Not rsBailes.EOF Then
                            aTabla(icFilasTabla, 5) = aTabla(icFilasTabla, 5) & ", "
                        End If
                    Wend
                    rsBailes.Close
                Else
                    Set rsBailes = db.OpenRecordset("SELECT b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & rs!codigo & " AND fase = 1 ORDER BY posicion", dbOpenSnapshot)
                    'If Not rsBailes.EOF Then aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & mml_FRASE0679
                    While Not rsBailes.EOF
                        aTabla(icFilasTabla, 5) = aTabla(icFilasTabla, 5) & sNombreBaileAbreviado(rsBailes!Nombre)
                        rsBailes.MoveNext
                        If Not rsBailes.EOF Then
                            aTabla(icFilasTabla, 5) = aTabla(icFilasTabla, 5) & ", "
                        End If
                    Wend
                    rsBailes.Close
                End If
            End If
            If G_COUNTRY Then aTabla(icFilasTabla, 6) = rs!num_grupo
            Inc icFilasTabla
            rs.MoveNext
        Wend
        rs.Close
        
        Open G_ARCH_HORARIO For Output As #100
        Print #100, mml_FRASE0326
        Dim iCols As Integer
        If G_COUNTRY Then
            iCols = 6
        Else
            iCols = 5
        End If
        For i = 0 To icFilasTabla - 1
            Print #100, "<tr>"
            For j = 0 To iCols
                Print #100, "<td>"
                Print #100, aTabla(i, j)
                Print #100, "</td>"
            Next
            Print #100, "</tr>"
        Next
        Print #100, "</TABLE></CENTER></BODY></HTML>"
        Close #100
    
        MsgBox mml_FRASE0327 & G_ARCH_HORARIO, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    CargarPistas cbPista
End Sub
