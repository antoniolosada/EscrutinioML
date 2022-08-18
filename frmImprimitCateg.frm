VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimirCateg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0066"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImpJueces 
      Caption         =   "mml_FRASE0982"
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
      Left            =   2940
      TabIndex        =   9
      Top             =   1470
      Width           =   2385
   End
   Begin VB.Frame Frame3 
      Caption         =   "mml_FRASE0437"
      Height          =   780
      Left            =   7020
      TabIndex        =   6
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
         ItemData        =   "frmImprimitCateg.frx":0000
         Left            =   150
         List            =   "frmImprimitCateg.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   1065
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   600
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame cmdImprimircmdImprimir 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6870
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   630
         Width           =   3660
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
         Left            =   1860
         Picture         =   "frmImprimitCateg.frx":004C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   630
         Width           =   435
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
         Left            =   1845
         Picture         =   "frmImprimitCateg.frx":04B6
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Width           =   3675
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
         TabIndex        =   13
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
      Left            =   6600
      TabIndex        =   1
      Top             =   1470
      Width           =   1725
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "mml_FRASE0017"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1470
      Width           =   2655
   End
End
Attribute VB_Name = "frmImprimirCateg"
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
Dim aDefCelda(5) As TCelda
Dim iPag As Integer
Dim sSelecCateg As String

    If Not C_DEBUG Then On Local Error GoTo error

    If Val(tbCodComp.Text) = 0 Then
        CamposSinCubrir
    End If

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
    
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 5)
        
        aTabla(0, 0) = mml_FRASE0671
        aTabla(0, 1) = mml_FRASE0672
        aTabla(0, 2) = mml_FRASE0673
        aTabla(0, 3) = mml_FRASE0674
        aTabla(0, 4) = mml_FRASE0675
        
        If MsgBox(mml_FRASE0676, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
            aDefCelda(0).Ancho = 800
        Else
            aDefCelda(0).Ancho = 0
        End If
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 800
        aDefCelda(1).Justificado = eccentro
        aDefCelda(2).Ancho = 3000
        aDefCelda(2).Justificado = ecizquierda
        aDefCelda(3).Ancho = 800
        aDefCelda(3).Justificado = eccentro
        aDefCelda(4).Ancho = 5500
        aDefCelda(4).Justificado = ecizquierda
        
        If Val(tbCodCateg.Text) = 0 Then
            sSelecCateg = ""
        Else
            sSelecCateg = " AND c.codigo = " & tbCodCateg.Text
        End If
        
        Set rs = db.OpenRecordset(" SELECT c.codigo, hora, id_categoria, ge.nombre, descripcion, m.nombre as mod FROM categorias c, gruposedad ge, modalidad m WHERE c.descripcion LIKE '*" & cbPista.Text & "*' AND c.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo " & sSelecCateg & " ORDER BY hora, 5", dbOpenSnapshot)
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = rs!codigo
            aTabla(icFilasTabla, 1) = Format(CDate(rs!hora), "hh:mm")
            aTabla(icFilasTabla, 2) = rs!DESCRIPCION
            
            Set rsCont = db.OpenRecordset(" SELECT DISTINCT num_dorsal FROM dorsales WHERE cod_categoria = " & rs!codigo, dbOpenSnapshot)
            iNumDorsales = 0
            If Not rsCont.EOF Then
                rsCont.MoveLast
                iNumDorsales = rsCont.RecordCount
            End If
            aTabla(icFilasTabla, 3) = iNumDorsales
            rsCont.Close
            Set rsCont = db.OpenRecordset(" SELECT MAX(fase) FROM dorsales WHERE cod_categoria = " & rs!codigo, dbOpenSnapshot)
            If Not rsCont.EOF Then
                If rsCont.Fields(0) >= 1 And rsCont.Fields(0) <= 256 Then
                    aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & " (" & sDescCortaFase(rsCont.Fields(0)) & ")"
                End If
            End If
            rsCont.Close
            
            aTabla(icFilasTabla, 4) = "(dT " & iDorsalesPorTandaCateg(rs!codigo) & ") "
            Dim iCBailes As Integer
            iCBailes = 0
            Set rsBailes = db.OpenRecordset("SELECT b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & rs!codigo & " AND fase = 2 ORDER BY posicion", dbOpenSnapshot)
            If Not rsBailes.EOF Then aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & mml_FRASE0677
            While Not rsBailes.EOF
                Inc iCBailes
                aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & sNombreBaileAbreviado(rsBailes!Nombre)
                rsBailes.MoveNext
                If Not rsBailes.EOF Then
                    aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & ", "
                Else
                    aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & " - "
                End If
            Wend
            rsBailes.Close
            If iCBailes > 5 Then
                aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & "= Final"
            Else
                Set rsBailes = db.OpenRecordset("SELECT b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & rs!codigo & " AND fase = 1 ORDER BY posicion", dbOpenSnapshot)
                If Not rsBailes.EOF Then aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & mml_FRASE0679
                While Not rsBailes.EOF
                    aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & sNombreBaileAbreviado(rsBailes!Nombre)
                    rsBailes.MoveNext
                    If Not rsBailes.EOF Then
                        aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & ", "
                    End If
                Wend
                rsBailes.Close
            End If
            Inc icFilasTabla
            If icFilasTabla >= G_MAX_FILAS_POR_PAG - 2 Then
                Printer.FontBold = False
                Printer.FontSize = 10
                Printer.Print mml_FRASE0659 & iPag
                ImprimirCabecera
                DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 5, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
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
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 5, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias
    Exit Sub
error:
    ProcesarError "cmdCambiar_Click"
End Sub


Private Sub cmdCateg_Click()
Dim rs As Recordset
   If Not C_DEBUG Then On Local Error GoTo error
   If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
   End If
    'JOIN para que aparezcan categ con 0 dorsales marcando que tienen 1 dorsal
    tbCodCateg.Text = sSeleccionar("SELECT c.codigo, c.descripcion, COUNT(*) AS NumDors, c.id_categoria, c.hora FROM (Categorias c LEFT JOIN Dorsales d ON d.cod_categoria = c.codigo) WHERE cod_competicion = " & tbCodComp.Text, "descripcion", " GROUP BY c.codigo, c.descripcion, c.id_categoria, c.hora  ORDER BY " & G_ORDEN_CATEGORIAS)
    ' No muestra las categorias vacias
    ' tbCodCateg.Text = sSeleccionar("SELECT c.codigo, c.descripcion, COUNT(*) AS NumDors, c.id_categoria, c.hora FROM Categorias c, Dorsales d WHERE d.cod_categoria = c.codigo AND cod_competicion = " & tbCodComp.Text & " GROUP BY c.codigo, c.descripcion, c.id_categoria, c.hora")
    tbDescCateg.Text = sResultado(2)
    
    Exit Sub
error:
    ProcesarError

End Sub

Private Sub cmdImpJueces_Click()
Dim bImpCodigo As Boolean
    If Not C_DEBUG Then On Local Error GoTo error
    
    If Val(tbCodComp.Text) = 0 Then
        CamposSinCubrir
    End If

    If MsgBox(mml_FRASE0670, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0676, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
        bImpCodigo = True
    Else
        bImpCodigo = False
    End If
    
    ComprobarImpresoraPorDefecto
    If MsgBox(mml_FRASE1122, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        ImprimirPanelJueces
    ElseIf MsgBox(mml_FRASE1186, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbNo Then
        ImprimirJuecesCateg
    Else
        ImprimirCategoriasQueJuzgaUnJuez
    End If
    Exit Sub
error:
    ProcesarError "cmdImpJueces_Click"
End Sub
Sub ImprimirCategoriasQueJuzgaUnJuez()
Dim rs As Recordset

    Set rs = db.OpenRecordset("SELECT DISTINCT id_juez FROM juez_categ WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & Val(tbCodComp.Text) & ") ORDER BY 1", dbOpenSnapshot)
    While Not rs.EOF
        ImprimirJuecesCateg rs.Fields("id_juez")
        rs.MoveNext
    Wend
    rs.Close
End Sub
Private Sub ImprimirJuecesCateg(Optional sJuez As String = "*", Optional bImpCodigo = False)
Dim rs As Recordset
Dim rsCont As Recordset
Dim rsJueces As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iNumDorsales As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(5) As TCelda
Dim iPag As Integer
Dim sSelecCateg As String


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
    
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 5)
        
        aTabla(0, 0) = mml_FRASE0671
        aTabla(0, 1) = mml_FRASE0672
        aTabla(0, 2) = mml_FRASE0673
        aTabla(0, 3) = mml_FRASE0674
        aTabla(0, 4) = mml_FRASE1123
        
        If bImpCodigo Then
            aDefCelda(0).Ancho = 800
        Else
            aDefCelda(0).Ancho = 0
        End If
        
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 800
        aDefCelda(1).Justificado = eccentro
        aDefCelda(2).Ancho = 3000
        aDefCelda(2).Justificado = ecizquierda
        aDefCelda(3).Ancho = 800
        aDefCelda(3).Justificado = eccentro
        aDefCelda(4).Ancho = 5500
        aDefCelda(4).Justificado = ecizquierda
        
        If Val(tbCodCateg.Text) = 0 Then
            sSelecCateg = ""
        Else
            sSelecCateg = " AND c.codigo = " & tbCodCateg.Text
        End If
        
        Set rs = db.OpenRecordset(" SELECT c.codigo, hora, id_categoria, ge.nombre, descripcion, m.nombre as mod FROM categorias c, gruposedad ge, modalidad m WHERE c.descripcion LIKE '*" & cbPista.Text & "*' AND c.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo " & sSelecCateg & " ORDER BY hora, 5", dbOpenSnapshot)
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = rs!codigo
            aTabla(icFilasTabla, 1) = Format(CDate(rs!hora), "hh:mm")
            aTabla(icFilasTabla, 2) = rs!DESCRIPCION
            
            Set rsCont = db.OpenRecordset(" SELECT DISTINCT num_dorsal FROM dorsales WHERE cod_categoria = " & rs!codigo, dbOpenSnapshot)
            iNumDorsales = 0
            If Not rsCont.EOF Then
                rsCont.MoveLast
                iNumDorsales = rsCont.RecordCount
            End If
            aTabla(icFilasTabla, 3) = iNumDorsales
            rsCont.Close
            Set rsCont = db.OpenRecordset(" SELECT MAX(fase) FROM dorsales WHERE cod_categoria = " & rs!codigo, dbOpenSnapshot)
            If Not rsCont.EOF Then
                If rsCont.Fields(0) >= 1 And rsCont.Fields(0) <= 256 Then
                    aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & " (" & sDescCortaFase(rsCont.Fields(0)) & ")"
                End If
            End If
            rsCont.Close
            
            aTabla(icFilasTabla, 4) = "(dT " & iDorsalesPorTandaCateg(rs!codigo) & ") "
            Dim iCBailes As Integer
            iCBailes = 0
            Set rsJueces = db.OpenRecordset("SELECT jc.id_juez FROM juez_categ jc WHERE jc.id_juez LIKE '" & sJuez & "' AND jc.cod_categoria = " & rs!codigo & " ORDER BY 1", dbOpenSnapshot)
            If rsJueces.EOF Then
                aTabla(icFilasTabla, 4) = mml_FRASE1185
            Else
                While Not rsJueces.EOF
                    aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & rsJueces.Fields("id_juez")
                    rsJueces.MoveNext
                    If Not rsJueces.EOF Then
                        aTabla(icFilasTabla, 4) = aTabla(icFilasTabla, 4) & ", "
                    End If
                Wend
            End If
            rsJueces.Close
            Inc icFilasTabla
            If icFilasTabla >= G_MAX_FILAS_POR_PAG - 2 Then
                Printer.FontBold = False
                Printer.FontSize = 10
                Printer.Print mml_FRASE0659 & iPag
                ImprimirCabecera
                DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 5, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
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
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 5, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias
End Sub

Private Sub ImprimirPanelJueces()

Dim rs As Recordset
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(3) As TCelda
Dim sSelecCateg As String

    If tbCodComp.Text = "" Then
        CamposSinCubrir
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
    
    iLineas = 4
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 2)
        
        aTabla(0, 0) = mml_FRASE0983
        aTabla(0, 1) = mml_FRASE0984
        
        aDefCelda(0).Ancho = 1000
        aDefCelda(0).Justificado = ecizquierda
        aDefCelda(1).Ancho = 5500
        aDefCelda(1).Justificado = ecizquierda
        
        icFilasTabla = 1
        iEscala = Printer.Width / 10
        
        If Val(tbCodCateg.Text) = 0 Then
            sSelecCateg = ""
        Else
            sSelecCateg = " AND c.codigo = " & tbCodCateg.Text
        End If
        
        Set rs = db.OpenRecordset("SELECT DISTINCT jc.id_juez, j.nombre FROM juez_categ jc, jueces j, categorias c WHERE  jc.cod_juez = j.codigo AND jc.cod_categoria = c.codigo AND c.codigo IN (SELECT codigo FROM categorias WHERE cod_competicion = " & Val(tbCodComp.Text) & ") " & sSelecCateg, dbOpenSnapshot)
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = rs.Fields(0)
            aTabla(icFilasTabla, 1) = rs.Fields(1)
            rs.MoveNext
            
            Inc icFilasTabla
            If icFilasTabla >= G_MAX_FILAS_POR_PAG Then
                Printer.FontBold = False
                Printer.FontSize = 10
                Printer.Print mml_FRASE0659 & iPag
                ImprimirCabecera
                DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                Printer.NewPage
                Inc iPag
                icFilasTabla = 1
            End If
        Wend
        rs.Close
        
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        If sSelecCateg <> "" Then
            Printer.Print "     " & tbCodCateg.Text & " - " & tbDescCateg.Text & " - " & mml_FRASE1094 & " " & PanelJuecesCateg(Val(tbCodCateg.Text))
            Printer.Print
        End If
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 2, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias


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
        sCateg = sDescCategoria(tbCodCateg.Text, tbCodComp.Text)
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
    ProcesarError "tbCodCateg_LostFocus"
End Sub

