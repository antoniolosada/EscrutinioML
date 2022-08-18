VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimirParticipantes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0745"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCoordinadorHTML 
      Caption         =   "->HTML"
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
      Left            =   5400
      TabIndex        =   26
      Top             =   3000
      Width           =   1515
   End
   Begin VB.Frame Frame3 
      Caption         =   "mml_FRASE0437"
      Height          =   780
      Left            =   7650
      TabIndex        =   20
      Top             =   0
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
         ItemData        =   "frmImprimitParticipantes.frx":0000
         Left            =   150
         List            =   "frmImprimitParticipantes.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdDorsalCatHorario 
      Caption         =   "mml_FRASE1047"
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
      Left            =   60
      TabIndex        =   19
      Top             =   3000
      Width           =   5310
   End
   Begin VB.CommandButton cmdPorCatTXT 
      Caption         =   "Txt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2670
      TabIndex        =   17
      Top             =   1350
      Width           =   555
   End
   Begin VB.CommandButton cmdPorNombreTXT 
      Caption         =   "Txt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6180
      TabIndex        =   16
      Top             =   2445
      Width           =   795
   End
   Begin VB.CommandButton cmdPorNombreHtml 
      Caption         =   "Html"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   15
      Top             =   2445
      Width           =   795
   End
   Begin VB.CommandButton cmdPorCategParaHtml 
      Caption         =   "Html"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   14
      Top             =   1350
      Width           =   765
   End
   Begin VB.CommandButton cmdPorCodigo 
      Caption         =   "mml_FRASE0746"
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
      Left            =   7050
      TabIndex        =   13
      Top             =   2445
      Width           =   1770
   End
   Begin VB.CommandButton cmdPorEscuela 
      Caption         =   "mml_FRASE0747"
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
      Left            =   6525
      TabIndex        =   12
      Top             =   1875
      Width           =   2310
   End
   Begin VB.CommandButton cmdPorProvincia 
      Caption         =   "mml_FRASE0748"
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
      Left            =   3870
      TabIndex        =   11
      Top             =   1890
      Width           =   2580
   End
   Begin VB.CommandButton cmdCategPag 
      Caption         =   "mml_FRASE0749"
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
      Left            =   3270
      TabIndex        =   10
      Top             =   1350
      Width           =   3750
   End
   Begin VB.CommandButton cmdListParejas 
      Caption         =   "mml_FRASE0750"
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
      Left            =   75
      TabIndex        =   9
      Top             =   2445
      Width           =   3570
   End
   Begin VB.CommandButton cmdDorsalCat 
      Caption         =   "mml_FRASE0751"
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
      Left            =   60
      TabIndex        =   8
      Top             =   1890
      Width           =   3720
   End
   Begin VB.CommandButton cmdDorsal 
      Caption         =   "mml_FRASE0752"
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
      Left            =   7080
      TabIndex        =   7
      Top             =   1350
      Width           =   1740
   End
   Begin VB.CommandButton cmdApellido 
      Caption         =   "mml_FRASE0753"
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
      Left            =   3735
      TabIndex        =   6
      Top             =   2445
      Width           =   1665
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7590
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
         Picture         =   "frmImprimitParticipantes.frx":005E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   660
         Width           =   435
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
         Left            =   3150
         MaxLength       =   35
         TabIndex        =   23
         Top             =   660
         Width           =   4320
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
         TabIndex        =   22
         Top             =   660
         Width           =   855
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
         Picture         =   "frmImprimitParticipantes.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   3150
         TabIndex        =   4
         Top             =   240
         Width           =   4320
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
         Left            =   150
         TabIndex        =   25
         Top             =   660
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
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7050
      TabIndex        =   1
      Top             =   3000
      Width           =   1785
   End
   Begin VB.CommandButton cmdPorCategoria 
      Caption         =   "mml_FRASE0369"
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
      Left            =   75
      TabIndex        =   0
      Top             =   1350
      Width           =   1845
   End
End
Attribute VB_Name = "frmImprimirParticipantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const C_MARGEN_LST_NORMAL = 200
Const MAX_CAR_CAB_POR_PAGINA = 41

Private Sub cmdApellido_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim rsDorsal As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(7) As TCelda
Dim bPagado As Boolean
Dim sHNombre As String, sMNombre As String

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If


    If MsgBox(mml_FRASE0754, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    If MsgBox(mml_FRASE0755, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        bPagado = False
    Else
        bPagado = True
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
    
    Printer.Orientation = vbPRORLandscape
    
    iLineas = 4
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 7)
        
        aTabla(0, 0) = mml_FRASE0756
        aTabla(0, 1) = mml_FRASE0757
        aTabla(0, 2) = mml_FRASE0671
        aTabla(0, 3) = mml_FRASE0758
        aTabla(0, 4) = mml_FRASE0759
        aTabla(0, 5) = mml_FRASE0760
        aTabla(0, 6) = mml_FRASE0673
        
        If bPagado Then
            aDefCelda(0).Ancho = 400
            aDefCelda(1).Ancho = 500
            aDefCelda(2).Ancho = 800
            aDefCelda(5).Ancho = 800
        Else
            aDefCelda(0).Ancho = 0
            aDefCelda(1).Ancho = 0
            aDefCelda(2).Ancho = 0
            aDefCelda(5).Ancho = 0
       End If
       
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Justificado = eccentro
        aDefCelda(2).Justificado = eccentro
        aDefCelda(3).Ancho = 3450
        aDefCelda(3).Justificado = ecizquierda
        aDefCelda(4).Ancho = 3450
        aDefCelda(4).Justificado = ecizquierda
        aDefCelda(5).Justificado = eccentro
        aDefCelda(6).Ancho = 5150
        aDefCelda(6).Justificado = eccentro
        
        Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, p.codigo, p.categoria, cod_grupoedad, grupoedad, m.nombre, p.pagado, p.pareja_adicional FROM modalidad m, parejas p WHERE  p.cod_modalidad = m.codigo AND p.cod_competicion = " & tbCodComp.Text & " ORDER BY p.nombre_hombre", dbOpenSnapshot)
        While Not rsPart.EOF
            Printer.FontSize = 10
            Printer.CurrentX = iEscala
            aTabla(icFilasTabla, 0) = IIf(rsPart!pareja_adicional = 1, mml_FRASE0129, mml_FRASE0761)
            aTabla(icFilasTabla, 1) = IIf(rsPart!pagado = 1, mml_FRASE0129, mml_FRASE0761)
            aTabla(icFilasTabla, 2) = rsPart!codigo
            If rsPart.Fields(0) = sHNombre And rsPart.Fields(1) = sMNombre Then
                aTabla(icFilasTabla, 3) = "                         """
                aTabla(icFilasTabla, 4) = "                         """
            Else
                aTabla(icFilasTabla, 3) = rsPart.Fields(0)
                aTabla(icFilasTabla, 4) = rsPart.Fields(1)
                sHNombre = rsPart.Fields(0)
                sMNombre = rsPart.Fields(1)
            End If
            
            Set rsDorsal = db.OpenRecordset("SELECT num_dorsal FROM dorsales d WHERE  d.cod_pareja =" & rsPart!codigo, dbOpenSnapshot)
            If Not rsDorsal.EOF Then
                aTabla(icFilasTabla, 5) = rsDorsal!num_dorsal
            Else
                aTabla(icFilasTabla, 5) = mml_FRASE0762
            End If
            aTabla(icFilasTabla, 6) = Mid$(rsPart!Nombre, 1, IIf(Len(rsPart!categoria) <= 3, 3, 2)) & " " & rsPart!categoria & " " & sDescCortaGrupoEdad(rsPart!cod_grupoedad)
            rsPart.MoveNext
            
            Inc icFilasTabla
            If icFilasTabla >= G_MAX_FILAS_POR_PAG_HOR - 2 Then
                Printer.FontBold = False
                Printer.FontSize = 10
                Printer.Print mml_FRASE0659 & iPag;
                Printer.FontBold = True
                Printer.FontSize = 8
                Printer.Print IIf(iPag = 1, mml_FRASE0763 & Format$(Now, "dd/mm/yyyy") & mml_FRASE0764 & Format$(Time(), "hh:mm"), "")
                Printer.FontBold = False
                Printer.FontSize = 10
                ImprimirCabecera
                DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA + IIf(bPagado, 0, C_MARGEN_LST_NORMAL), Printer.CurrentY, icFilasTabla, 7, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                Printer.NewPage
                Inc iPag
                icFilasTabla = 1
            End If
        Wend
        rsPart.Close
        
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA + IIf(bPagado, 0, C_MARGEN_LST_NORMAL), Printer.CurrentY, icFilasTabla, 7, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias
End Sub

Sub ImprimirCabecera()
Dim rs As Recordset
Dim X As Integer, Y As Integer

    On Local Error GoTo error
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
    Centrado Printer, rs!DESCRIPCION & "  (" & rs!fecha & ")" & " " & cbPista.Text, Printer.Width
    rs.Close
    Printer.FontBold = False
    Printer.FontSize = 13
    Centrado Printer, sEscuela(tbCodComp.Text), Printer.Width
    SaltoLinea Printer, 4
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    Printer.DrawWidth = 1
    SaltoLinea Printer, 10
error:
    ProcesarError
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

Private Sub cmdPorCategoria_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim X As Integer, Y As Integer
Dim aDefCelda(6) As TCelda
Dim aTabla(C_MAX_PART_CATEG, 5) As String
Dim icFilasTabla As Integer
Dim sSelecCateg As String

    If Val(tbCodCateg.Text) > 0 Then
        sSelecCateg = " AND c.codigo = " & tbCodCateg & " "
    Else
        sSelecCateg = ""
    End If
    
     
     If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE0765, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0766, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
        aDefCelda(0).Ancho = 1000
        aDefCelda(1).Ancho = 4100
        aDefCelda(2).Ancho = 4100
        aDefCelda(3).Ancho = 700
        aDefCelda(4).Ancho = 800
        aDefCelda(5).Ancho = 800
    Else
        aDefCelda(0).Ancho = 0
        aDefCelda(1).Ancho = 4900
        aDefCelda(2).Ancho = 4900
        aDefCelda(3).Ancho = 0
        aDefCelda(4).Ancho = 0
        aDefCelda(5).Ancho = 1000
    End If
    
    aDefCelda(0).Justificado = eccentro
    aDefCelda(1).Justificado = ecizquierda
    aDefCelda(2).Justificado = ecizquierda
    aDefCelda(3).Justificado = eccentro
    aDefCelda(4).Justificado = eccentro
    aDefCelda(5).Justificado = eccentro
    
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
    
    iLineas = 6
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        iEscala = Printer.Width / 10
                
        icFilasTabla = 0
        
        Dim sOrden As String
        If MsgBox(mml_FRASE0767, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
            'Ordenar por la hora de salida de la categoria según la tabla de categorías
            sOrden = "hora"
            Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo " & sSelecCateg & " ORDER BY " & sOrden, dbOpenSnapshot)
        Else
            'Ordenado por descripción
            If C_COUNTRY Then
                ' Solo para country puede ordenarse por el orden introducido en las tablas de descripción de códigos
                sOrden = "m.orden, dc.orden, ge.orden, c.descripcion"
                Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, c.descripcion FROM categorias c, gruposedad ge, desccategoria dc, modalidad m WHERE m.codigo = c.cod_modalidad AND dc.descripcion = c.id_categoria AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo " & sSelecCateg & " ORDER BY " & sOrden, dbOpenSnapshot)
            Else
                sOrden = "descripcion"
                Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo " & sSelecCateg & " ORDER BY " & sOrden, dbOpenSnapshot)
            End If
        End If
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = mml_FRASE0768 & Format(CDate(rs!hora), "hh:mm")
            aTabla(icFilasTabla, 1) = mml_FRASE0768 & Left(rs!DESCRIPCION, MAX_CAR_CAB_POR_PAGINA)
            aTabla(icFilasTabla, 3) = mml_FRASE1251
            aTabla(icFilasTabla, 4) = mml_FRASE0769
            aTabla(icFilasTabla, 5) = mml_FRASE0654
            Set rsPart = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " AND fase =(SELECT MAX(fase) FROM dorsales WHERE cod_categoria = " & rs.Fields(2) & ")", dbOpenSnapshot)
            aTabla(icFilasTabla, 2) = mml_FRASE0768 & "(" & rsPart.Fields(0) & mml_FRASE0770
            rsPart.Close
            Inc icFilasTabla
            Inc iLineas
            
            Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, num_dorsal, p.codigo, p.cod_grupoedad, p.categoria FROM parejas p, dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " AND d.cod_pareja = p.codigo ORDER BY num_dorsal", dbOpenSnapshot)
            While Not rsPart.EOF
                aTabla(icFilasTabla, 0) = rsPart!codigo
                aTabla(icFilasTabla, 1) = rsPart.Fields(0)
                aTabla(icFilasTabla, 2) = rsPart.Fields(1)
                aTabla(icFilasTabla, 3) = rsPart.Fields("categoria")
                aTabla(icFilasTabla, 4) = sDescCortaGrupoEdad(rsPart!cod_grupoedad)
                aTabla(icFilasTabla, 5) = rsPart!num_dorsal
                rsPart.MoveNext
                
                Inc icFilasTabla
                If icFilasTabla >= G_MAX_FILAS_POR_PAG Then
                    Printer.FontBold = False
                    Printer.FontSize = 10
                    Printer.Print mml_FRASE0659 & iPag;
                    Printer.FontBold = True
                    Printer.FontSize = 8
                    Printer.Print IIf(iPag = 1, mml_FRASE0763 & Format$(Now, "dd/mm/yyyy") & mml_FRASE0764 & Format$(Time(), "hh:mm"), "")
                    Printer.FontBold = False
                    Printer.FontSize = 10
                    ImprimirCabecera
                    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 6, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                    Printer.NewPage
                    Inc iPag
                    icFilasTabla = 0
                End If
                
            Wend
            rsPart.Close
            rs.MoveNext
        Wend
        rs.Close
            
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 6, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.NewPage
        Printer.EndDoc
    Next iCCopias
End Sub

Private Sub cmdPorCategParaHtml_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim X As Integer, Y As Integer
Dim aDefCelda(4) As TCelda
Dim aTabla(C_MAX_PART_CATEG, 4) As String
Dim icFilasTabla As Integer
Dim iFichero As Integer
Dim bConProvincia As Boolean

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE0771, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0772, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        bConProvincia = False
    Else
        bConProvincia = True
    End If
    
    
    iFichero = FreeFile
    Open VarCfg("fichero_part_por_categorias") For Output As #iFichero
        
        Print #iFichero, mml_FRASE0773
        
        Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY descripcion", dbOpenSnapshot)
        While Not rs.EOF
            
            Print #iFichero, "<tr>"
            Print #iFichero, "<td><b>" & Format(CDate(rs!hora), "hh:mm") & "</b></td>"
            Print #iFichero, "<td><b>" & rs!DESCRIPCION & "</b></td>"
            
            Set rsPart = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " AND fase =(SELECT MAX(fase) FROM dorsales WHERE cod_categoria = " & rs.Fields(2) & ")", dbOpenSnapshot)
            Print #iFichero, "<td><b>" & "(" & rsPart.Fields(0) & mml_FRASE0770 & "</b></td>"
            Print #iFichero, mml_FRASE0774
            Print #iFichero, "</tr>"
            rsPart.Close
            Inc icFilasTabla
            Inc iLineas
            
            Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, num_dorsal, p.codigo, p.provincia FROM parejas p, dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " AND d.cod_pareja = p.codigo ORDER BY num_dorsal", dbOpenSnapshot)
            While Not rsPart.EOF
                Print #iFichero, "<tr>"
                Print #iFichero, "<td></td>"
                If bConProvincia Then
                    Dim sHProv As String, sMProv As String, i As Integer
                    sHProv = RTrim$(LTrim$(rsPart!provincia))
                    If sHProv <> "" Then
                        i = InStr(sHProv, "-")
                        If i > 0 Then
                            sMProv = LTrim$(RTrim$(Mid$(sHProv, i + 1)))
                            sHProv = LTrim$(RTrim$(Mid$(sHProv, 1, i - 1)))
                        End If
                    End If
                    If sHProv <> "" Then
                        sHProv = " (" & sHProv & ")"
                    End If
                    If sMProv <> "" Then
                        sMProv = " (" & sMProv & ")"
                    End If
                    Print #iFichero, "<td>" & rsPart.Fields(0) & sHProv & "</td>"
                    Print #iFichero, "<td>" & rsPart.Fields(1) & sMProv & "</td>"
                Else
                    Print #iFichero, "<td>" & rsPart.Fields(0) & "</td>"
                    Print #iFichero, "<td>" & rsPart.Fields(1) & "</td>"
                End If
                Print #iFichero, "<td>" & rsPart!num_dorsal & "</td>"
                Print #iFichero, "</tr>"
                rsPart.MoveNext
                
            Wend
            rsPart.Close
            rs.MoveNext
        Wend
        rs.Close
            
    Print #iFichero, "</table></BODY></HTML>"
    Close iFichero
    
    MsgBox mml_FRASE0340 & VarCfg("fichero_part_por_categorias") & mml_FRASE0775, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Private Sub cmdPorCategParaWord_Click()

End Sub

Private Sub cmdPorCatTXT_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim X As Integer, Y As Integer
Dim aDefCelda(4) As TCelda
Dim aTabla(C_MAX_PART_CATEG, 4) As String
Dim icFilasTabla As Integer
Dim iFichero As Integer
Dim bConProvincia As Boolean
Dim bConEscuela As Boolean

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE0776, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0772, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        bConProvincia = False
    Else
        bConProvincia = True
    End If
    If MsgBox(mml_FRASE1230, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        bConEscuela = False
    Else
        bConEscuela = True
    End If
    
    
    iFichero = FreeFile
    Open VarCfg("fichero_part_por_categorias") & ".TXT" For Output As #iFichero
        
        Print #iFichero, mml_FRASE0777 & CR & LF
        
        Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY descripcion", dbOpenSnapshot)
        While Not rs.EOF
            
            Print #iFichero, ""
            Print #iFichero, mml_FRASE0300 & vbTab & rs!DESCRIPCION & vbTab & vbTab & vbTab;
            
            Set rsPart = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " AND fase =(SELECT MAX(fase) FROM dorsales WHERE cod_categoria = " & rs.Fields(2) & ")", dbOpenSnapshot)
            Print #iFichero, "(" & rsPart.Fields(0) & mml_FRASE0770;
            rsPart.Close
            
            If bConEscuela Then
                Print #iFichero, vbTab & vbTab & mml_FRASE0395
            Else
                Print #iFichero, ""
            End If
            
            Inc icFilasTabla
            Inc iLineas
            
            Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, num_dorsal, p.codigo, p.provincia, p.escuelas FROM parejas p, dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " AND d.cod_pareja = p.codigo ORDER BY num_dorsal", dbOpenSnapshot)
            While Not rsPart.EOF
                Print #iFichero, rsPart!num_dorsal & vbTab;
                If bConProvincia Then
                    Dim sHProv As String, sMProv As String, i As Integer
                    sHProv = RTrim$(LTrim$(rsPart!provincia))
                    If sHProv <> "" Then
                        i = InStr(sHProv, "-")
                        If i > 0 Then
                            sMProv = LTrim$(RTrim$(Mid$(sHProv, i + 1)))
                            sHProv = LTrim$(RTrim$(Mid$(sHProv, 1, i - 1)))
                        End If
                    End If
                    If sHProv <> "" Then
                        sHProv = " (" & sHProv & ")"
                    End If
                    If sMProv <> "" Then
                        sMProv = " (" & sMProv & ")"
                    End If
                    Print #iFichero, rsPart.Fields(0) & sHProv & vbTab & vbTab;
                    Print #iFichero, rsPart.Fields(1) & sMProv;
                Else
                    Print #iFichero, rsPart.Fields(0) & vbTab & vbTab;
                    Print #iFichero, rsPart.Fields(1);
                End If
                If bConEscuela Then
                    Print #iFichero, vbTab & vbTab & rsPart.Fields("escuelas");
                End If
                Print #iFichero, ""
                rsPart.MoveNext
                
            Wend
            rsPart.Close
            rs.MoveNext
        Wend
        rs.Close
            
    Close iFichero
    
    MsgBox mml_FRASE0340 & VarCfg("fichero_part_por_categorias") & mml_FRASE0778, vbOKOnly Or vbInformation, mml_FRASE0086

End Sub

Private Sub cmdPorCodigo_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim rsDorsal As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(7) As TCelda


    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If MsgBox(mml_FRASE0779, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
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
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 7)
        
        aTabla(0, 0) = mml_FRASE0756
        aTabla(0, 1) = mml_FRASE0757
        aTabla(0, 2) = mml_FRASE0671
        aTabla(0, 3) = mml_FRASE0758
        aTabla(0, 4) = mml_FRASE0759
        aTabla(0, 5) = mml_FRASE0760
        aTabla(0, 6) = mml_FRASE0673
        
        aDefCelda(0).Ancho = 500
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 500
        aDefCelda(1).Justificado = eccentro
        aDefCelda(2).Ancho = 800
        aDefCelda(2).Justificado = eccentro
        aDefCelda(3).Ancho = 3500
        aDefCelda(3).Justificado = ecizquierda
        aDefCelda(4).Ancho = 3500
        aDefCelda(4).Justificado = ecizquierda
        aDefCelda(5).Ancho = 800
        aDefCelda(5).Justificado = eccentro
        aDefCelda(6).Ancho = 1000
        aDefCelda(6).Justificado = eccentro
        
        Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, p.codigo, p.categoria, grupoedad, m.nombre, p.pagado, p.pareja_adicional FROM modalidad m, parejas p WHERE  p.cod_modalidad = m.codigo AND p.cod_competicion = " & tbCodComp.Text & " ORDER BY p.codigo", dbOpenSnapshot)
        While Not rsPart.EOF
            Printer.FontSize = 10
            Printer.CurrentX = iEscala
            aTabla(icFilasTabla, 0) = IIf(rsPart!pareja_adicional = 1, mml_FRASE0129, mml_FRASE0761)
            aTabla(icFilasTabla, 1) = IIf(rsPart!pagado = 1, mml_FRASE0129, mml_FRASE0761)
            aTabla(icFilasTabla, 2) = rsPart!codigo
            aTabla(icFilasTabla, 3) = rsPart.Fields(0)
            aTabla(icFilasTabla, 4) = rsPart.Fields(1)
            
            Set rsDorsal = db.OpenRecordset("SELECT num_dorsal FROM dorsales d WHERE  d.cod_pareja =" & rsPart!codigo, dbOpenSnapshot)
            If Not rsDorsal.EOF Then
                aTabla(icFilasTabla, 5) = rsDorsal!num_dorsal
            Else
                aTabla(icFilasTabla, 5) = mml_FRASE0762
            End If
            aTabla(icFilasTabla, 6) = Mid$(rsPart!Nombre, 1, 1) & " " & rsPart!categoria & " " & Mid$(rsPart!grupoedad, 1, 3)
            rsPart.MoveNext
            
            Inc icFilasTabla
            If icFilasTabla >= G_MAX_FILAS_POR_PAG - 2 Then
                Printer.FontBold = False
                Printer.FontSize = 10
                Printer.Print mml_FRASE0659 & iPag
                ImprimirCabecera
                DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 7, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                Printer.NewPage
                Inc iPag
                icFilasTabla = 1
            End If
        Wend
        rsPart.Close
        
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 7, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias

End Sub

Private Sub cmdPorEscuela_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim X As Integer, Y As Integer
Dim aDefCelda(4) As TCelda
Dim aTabla(C_MAX_PART_CATEG, 4) As String
Dim icFilasTabla As Integer

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If


    If MsgBox(mml_FRASE0765, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
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
    
    iLineas = 6
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        iEscala = Printer.Width / 10
                
        aDefCelda(0).Ancho = 1500
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 4500
        aDefCelda(1).Justificado = ecizquierda
        aDefCelda(2).Ancho = 4500
        aDefCelda(2).Justificado = ecizquierda
        icFilasTabla = 0
        
        Set rs = db.OpenRecordset("SELECT DISTINCT escuelas  FROM parejas WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = mml_FRASE0768 & mml_FRASE0395
            If Trim$(rs!Escuelas) = "" Then
                aTabla(icFilasTabla, 1) = mml_FRASE0780
            Else
                aTabla(icFilasTabla, 1) = mml_FRASE0768 & rs!Escuelas
            End If
            Set rsPart = db.OpenRecordset("SELECT COUNT(*) FROM parejas WHERE cod_competicion = " & tbCodComp.Text & " AND escuelas = '" & rs!Escuelas & "'", dbOpenSnapshot)
            aTabla(icFilasTabla, 2) = mml_FRASE0768 & "(" & rsPart.Fields(0) & mml_FRASE0770
            rsPart.Close
            Inc icFilasTabla
            Inc iLineas
            
            Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer FROM parejas p  WHERE cod_competicion=" & tbCodComp.Text & " AND escuelas = '" & rs!Escuelas & "' ORDER BY 1", dbOpenSnapshot)
            While Not rsPart.EOF
                aTabla(icFilasTabla, 0) = ""
                aTabla(icFilasTabla, 1) = rsPart.Fields(0)
                aTabla(icFilasTabla, 2) = rsPart.Fields(1)
                rsPart.MoveNext
                
                Inc icFilasTabla
                If icFilasTabla >= G_MAX_FILAS_POR_PAG Then
                    Printer.FontSize = 10
                    Printer.Print mml_FRASE0659 & iPag
                    ImprimirCabecera
                    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                    Printer.NewPage
                    Inc iPag
                    icFilasTabla = 0
                End If
                
            Wend
            rsPart.Close
            rs.MoveNext
        Wend
        rs.Close
            
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 3, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.NewPage
        Printer.EndDoc
    Next iCCopias
End Sub

Private Sub cmdPorNombreHtml_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim rsDorsal As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(7) As TCelda
Dim iFichero As Integer

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If


    If MsgBox(mml_FRASE0781, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 7)
        
        On Local Error Resume Next
        CDialog.FileName = VarCfg("fichero_part_por_nombre") & ".TXT"
        CDialog.ShowSave
        On Local Error GoTo 0
        
        iFichero = FreeFile
        Open CDialog.FileName For Output As #iFichero
            
        Print #iFichero, mml_FRASE0782
        
        Print #iFichero, mml_FRASE0783
            
        Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, p.codigo, p.categoria, grupoedad, m.nombre, p.pagado, p.pareja_adicional FROM modalidad m, parejas p WHERE  p.cod_modalidad = m.codigo AND p.cod_competicion = " & tbCodComp.Text & " ORDER BY p.nombre_hombre", dbOpenSnapshot)
        While Not rsPart.EOF
        
            Print #iFichero, "<tr>"
            
            Printer.FontSize = 10
            Printer.CurrentX = iEscala
            Print #iFichero, "<td>" & IIf(rsPart!pareja_adicional = 1, mml_FRASE0129, mml_FRASE0761) & "</td>"
            Print #iFichero, "<td>" & IIf(rsPart!pagado = 1, mml_FRASE0129, mml_FRASE0761) & "</td>"
            Print #iFichero, "<td>" & rsPart!codigo & "</td>"
            Print #iFichero, "<td>" & rsPart.Fields(0) & "</td>"
            Print #iFichero, "<td>" & rsPart.Fields(1) & "</td>"
            
            Set rsDorsal = db.OpenRecordset("SELECT num_dorsal FROM dorsales d WHERE  d.cod_pareja =" & rsPart!codigo, dbOpenSnapshot)
            If Not rsDorsal.EOF Then
                Print #iFichero, "<td>" & rsDorsal!num_dorsal & "</td>"
            Else
                Print #iFichero, "<td>" & mml_FRASE0762 & "</td>"
            End If
            Print #iFichero, "<td>" & Mid$(rsPart!Nombre, 1, 3) & " " & rsPart!categoria & " " & Mid$(rsPart!grupoedad, 1, 3) & "</td>"
            
            Print #iFichero, "</tr>"
            rsPart.MoveNext
            
        Wend
        rsPart.Close
        
    Print #iFichero, "</table></BODY></HTML>"
    Close iFichero
    
    MsgBox mml_FRASE0340 & CDialog.FileName & mml_FRASE0775, vbOKOnly Or vbInformation, mml_FRASE0086
        
End Sub

Private Sub cmdPorNombreTXT_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim rsDorsal As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(7) As TCelda
Dim iFichero As Integer
Dim bDorsal As Boolean

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If


    If MsgBox(mml_FRASE0781, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0784, vbYesNo Or vbQuestion, mml_FRASE0086) = vbYes Then
        bDorsal = True
    Else
        bDorsal = False
    End If
    
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 7)
        
        On Local Error Resume Next
        CDialog.FileName = VarCfg("fichero_part_por_nombre") & ".TXT"
        CDialog.ShowSave
        On Local Error GoTo 0
        
        iFichero = FreeFile
        Open CDialog.FileName For Output As #iFichero
            
        Print #iFichero, mml_FRASE0785
        
        Print #iFichero, mml_FRASE0389 & vbTab & vbTab & vbTab & mml_FRASE0384 & vbTab & vbTab & vbTab & IIf(bDorsal, mml_FRASE0300, "") & vbTab & vbTab & vbTab & mml_FRASE0301
            
        Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, p.codigo, p.categoria, grupoedad, m.nombre, p.pagado, p.pareja_adicional FROM modalidad m, parejas p WHERE  p.cod_modalidad = m.codigo AND p.cod_competicion = " & tbCodComp.Text & " ORDER BY p.nombre_hombre", dbOpenSnapshot)
        While Not rsPart.EOF
        
            Printer.FontSize = 10
            Printer.CurrentX = iEscala
            Print #iFichero, rsPart.Fields(0) & vbTab;
            Print #iFichero, rsPart.Fields(1) & vbTab;
            
            Set rsDorsal = db.OpenRecordset("SELECT num_dorsal FROM dorsales d WHERE  d.cod_pareja =" & rsPart!codigo, dbOpenSnapshot)
            If bDorsal Then
                If Not rsDorsal.EOF Then
                    Print #iFichero, rsDorsal!num_dorsal & vbTab;
                Else
                    Print #iFichero, mml_FRASE0762 & vbTab;
                End If
            End If
            Print #iFichero, Mid$(rsPart!Nombre, 1, 3) & " " & rsPart!categoria & " " & Mid$(rsPart!grupoedad, 1, 3)
            
            rsPart.MoveNext
            
        Wend
        rsPart.Close
        
    Close iFichero
    
    MsgBox mml_FRASE0340 & CDialog.FileName & mml_FRASE0778, vbOKOnly Or vbInformation, mml_FRASE0086

End Sub

Private Sub cmdPorProvincia_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim X As Integer, Y As Integer
Dim aDefCelda(4) As TCelda
Dim aTabla(C_MAX_PART_CATEG, 4) As String
Dim icFilasTabla As Integer

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If


    If MsgBox(mml_FRASE0765, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
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
    
    iLineas = 6
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        iEscala = Printer.Width / 10
                
        aDefCelda(0).Ancho = 1500
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 4500
        aDefCelda(1).Justificado = ecizquierda
        aDefCelda(2).Ancho = 4500
        aDefCelda(2).Justificado = ecizquierda
        icFilasTabla = 0
        
        Set rs = db.OpenRecordset("SELECT DISTINCT provincia  FROM parejas WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = mml_FRASE0768 & mml_FRASE0786
            If Trim$(rs!provincia) = "" Then
                aTabla(icFilasTabla, 1) = mml_FRASE0780
            Else
                aTabla(icFilasTabla, 1) = mml_FRASE0768 & rs!provincia
            End If
            Set rsPart = db.OpenRecordset("SELECT COUNT(*) FROM parejas WHERE cod_competicion = " & tbCodComp.Text & " AND provincia = '" & rs!provincia & "'", dbOpenSnapshot)
            aTabla(icFilasTabla, 2) = mml_FRASE0768 & "(" & rsPart.Fields(0) & mml_FRASE0770
            rsPart.Close
            Inc icFilasTabla
            Inc iLineas
            
            Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer FROM parejas p  WHERE cod_competicion=" & tbCodComp.Text & " AND provincia = '" & rs!provincia & "' ORDER BY 1", dbOpenSnapshot)
            While Not rsPart.EOF
                aTabla(icFilasTabla, 0) = ""
                aTabla(icFilasTabla, 1) = rsPart.Fields(0)
                aTabla(icFilasTabla, 2) = rsPart.Fields(1)
                rsPart.MoveNext
                
                Inc icFilasTabla
                If icFilasTabla >= G_MAX_FILAS_POR_PAG Then
                    Printer.FontSize = 10
                    Printer.Print mml_FRASE0659 & iPag
                    ImprimirCabecera
                    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                    Printer.NewPage
                    Inc iPag
                    icFilasTabla = 0
                End If
                
            Wend
            rsPart.Close
            rs.MoveNext
        Wend
        rs.Close
            
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 3, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.NewPage
        Printer.EndDoc
    Next iCCopias
End Sub


Private Sub cmdCategPag_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim X As Integer, Y As Integer
Dim aDefCelda(4) As TCelda
Dim aTabla(C_MAX_PART_CATEG, 4) As String
Dim icFilasTabla As Integer
Dim sSelecCateg As String

    If Val(tbCodCateg.Text) > 0 Then
        sSelecCateg = " AND c.codigo = " & tbCodCateg & " "
    Else
        sSelecCateg = ""
    End If

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If


    If MsgBox(mml_FRASE0765, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
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
    
    iLineas = 6
    For iCCopias = 1 To CDialog.Copies
        iPag = 1
        iEscala = Printer.Width / 10
                
        aDefCelda(0).Ancho = 1000
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 4400
        aDefCelda(1).Justificado = ecizquierda
        aDefCelda(2).Ancho = 4400
        aDefCelda(2).Justificado = ecizquierda
        aDefCelda(3).Ancho = 1100
        aDefCelda(3).Justificado = eccentro
        icFilasTabla = 0
        
        Dim sOrden As String
        If MsgBox(mml_FRASE0787, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
            sOrden = mml_FRASE0259
        Else
            sOrden = mml_FRASE0267
        End If
        Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo " & sSelecCateg & " ORDER BY " & sOrden, dbOpenSnapshot)
        While Not rs.EOF
            aTabla(icFilasTabla, 0) = ".12b" & Format(CDate(rs!hora), "hh:mm")
            aTabla(icFilasTabla, 1) = ".12b" & rs!DESCRIPCION
            aTabla(icFilasTabla, 3) = mml_FRASE0788
            Set rsPart = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " AND fase =(SELECT MAX(fase) FROM dorsales WHERE cod_categoria = " & rs.Fields(2) & ")", dbOpenSnapshot)
            aTabla(icFilasTabla, 2) = ".12b" & "(" & rsPart.Fields(0) & mml_FRASE0770
            rsPart.Close
            Inc icFilasTabla
            Inc iLineas
            
            Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, num_dorsal FROM parejas p, dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " AND d.cod_pareja = p.codigo ORDER BY num_dorsal", dbOpenSnapshot)
            While Not rsPart.EOF
                aTabla(icFilasTabla, 0) = ""
                aTabla(icFilasTabla, 1) = ".12n" & rsPart.Fields(0)
                aTabla(icFilasTabla, 2) = ".12n" & rsPart.Fields(1)
                aTabla(icFilasTabla, 3) = ".12n" & rsPart!num_dorsal
                rsPart.MoveNext
                
                Inc icFilasTabla
                If icFilasTabla >= G_MAX_FILAS_POR_PAG Then
                    Printer.FontSize = 12
                    Printer.Print mml_FRASE0659 & iPag
                    ImprimirCabecera
                    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                    Printer.NewPage
                    Inc iPag
                    icFilasTabla = 0
                End If
                
            Wend
            ' Nuevo página por categoría
            Printer.FontSize = 10
            Printer.Print mml_FRASE0659 & iPag
            Printer.FontSize = 14
            ImprimirCabecera
            DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
            Printer.NewPage
            Inc iPag
            icFilasTabla = 0
            
            rsPart.Close
            rs.MoveNext
        Wend
        rs.Close
            
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.NewPage
        Printer.EndDoc
    Next iCCopias

End Sub

Private Sub cmdDorsal_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(3) As TCelda

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If


    If MsgBox(mml_FRASE0870, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
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
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 3)
        
        aTabla(0, 0) = mml_FRASE0758
        aTabla(0, 1) = mml_FRASE0759
        aTabla(0, 2) = mml_FRASE0760
        
        aDefCelda(0).Ancho = 4500
        aDefCelda(0).Justificado = ecizquierda
        aDefCelda(1).Ancho = 5500
        aDefCelda(1).Justificado = ecizquierda
        aDefCelda(2).Ancho = 800
        aDefCelda(2).Justificado = eccentro
        
        icFilasTabla = 1
        iEscala = Printer.Width / 10
        
        Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, num_dorsal, provincia FROM parejas p, dorsales d WHERE  d.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ") AND d.cod_pareja = p.codigo ORDER BY num_dorsal", dbOpenSnapshot)
        While Not rsPart.EOF
            aTabla(icFilasTabla, 0) = rsPart.Fields(0)
            aTabla(icFilasTabla, 1) = rsPart.Fields(1)
            If Not IsNull(rsPart!provincia) Then
                If Trim$(rsPart!provincia) <> "" Then
                    aTabla(icFilasTabla, 1) = aTabla(icFilasTabla, 1) & " (" & rsPart!provincia & ")"
                End If
            End If
            aTabla(icFilasTabla, 2) = rsPart!num_dorsal
            
            rsPart.MoveNext
            
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
        rsPart.Close
        
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias

End Sub

Private Sub cmdDorsalCat_Click()
Dim rs As Recordset
Dim iContDorsales As Integer
Dim rsPart As Recordset
Dim rsFase As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim aDefCelda(4) As TCelda
Dim aTabla() As String
Dim icFilasTabla As Integer
Dim iDorsalesTanda As Integer
Dim iNumTandas As Integer
Dim iTandas As Integer
Dim sSelecCateg As String

    If Val(tbCodCateg.Text) > 0 Then
        sSelecCateg = " AND c.codigo = " & tbCodCateg & " "
    Else
        sSelecCateg = ""
    End If

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0765, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
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
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 3)
        
        aTabla(0, 0) = mml_FRASE0672
        aTabla(0, 1) = mml_FRASE0713
        aTabla(0, 2) = mml_FRASE0673
        aTabla(0, 3) = mml_FRASE0871
        
        aDefCelda(0).Ancho = 1000
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 1000
        aDefCelda(1).Justificado = eccentro
        aDefCelda(2).Ancho = 3000
        aDefCelda(2).Justificado = ecizquierda
        aDefCelda(3).Ancho = 5500
        aDefCelda(3).Justificado = ecizquierda
        icFilasTabla = 1
        
        iPag = 1
        iEscala = Printer.Width / 40
            
        Set rs = db.OpenRecordset("SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE c.descripcion LIKE '*" & cbPista.List(cbPista.ListIndex) & "*' AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo " & sSelecCateg & " ORDER BY hora, descripcion", dbOpenSnapshot)
        
        While Not rs.EOF
            Set rsFase = db.OpenRecordset("SELECT DISTINCT fase FROM  dorsales d WHERE d.cod_categoria = " & rs.Fields(2) & " ORDER BY d.fase DESC", dbOpenSnapshot)
            While Not rsFase.EOF
                aTabla(icFilasTabla, 0) = Format(CDate(rs!hora), "hh:mm")
                If Not rsFase.EOF Then
                    Select Case rsFase!fase
                        Case "1":
                            aTabla(icFilasTabla, 1) = mml_FRASE0329
                        Case "2":
                            aTabla(icFilasTabla, 1) = mml_FRASE0872
                        Case "4":
                            aTabla(icFilasTabla, 1) = mml_FRASE0652
                        Case Else
                            aTabla(icFilasTabla, 1) = rsFase!fase
                    End Select
                Else
                    aTabla(icFilasTabla, 1) = ""
                End If
                
                'Inumtandas es cubierta por la función
                aTabla(icFilasTabla, 2) = rs!DESCRIPCION
                aTabla(icFilasTabla, 3) = ""
                Set rsPart = db.OpenRecordset("SELECT DISTINCT d.num_dorsal FROM parejas p, dorsales d WHERE d.fase = " & rsFase!fase & " AND d.cod_categoria = " & rs.Fields(2) & " AND d.cod_pareja = p.codigo ORDER BY num_dorsal", dbOpenSnapshot)
                iContDorsales = 0
                iTandas = 1
                iNumTandas = 0
                iDorsalesTanda = CalcularDorsalesPorTandaCat(rs!codigo, rsFase!fase, 0, iTandas, iNumTandas)
                While Not rsPart.EOF
                    If iContDorsales >= C_MAX_DORSALES_LIST_PART Then
                        Inc icFilasTabla
                        aTabla(icFilasTabla, 0) = ""
                        aTabla(icFilasTabla, 1) = ""
                        aTabla(icFilasTabla, 2) = ""
                        aTabla(icFilasTabla, 3) = ""
                        iContDorsales = 0
                    End If
                    aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & Trim$(Str$(rsPart.Fields(0)))
                    rsPart.MoveNext
                    If (iContDorsales = iDorsalesTanda - 1 And iTandas < iNumTandas) Then
                        Inc iTandas
                        iDorsalesTanda = CalcularDorsalesPorTandaCat(rs!codigo, rsFase!fase, 0, iTandas, iNumTandas)
                        iContDorsales = C_MAX_DORSALES_LIST_PART
                        If Not rsPart.EOF Then
                            aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & " |"
                        End If
                    Else
                        If Not rsPart.EOF Then
                            aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & ","
                        End If
                    End If
                    Inc iContDorsales
                Wend
                rsPart.Close
                
                Inc icFilasTabla
                If icFilasTabla >= G_MAX_FILAS_POR_PAG - 1 Then
                    Printer.FontBold = False
                    Printer.FontSize = 10
                    Printer.Print mml_FRASE0659 & iPag
                    ImprimirCabecera
                    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                    Printer.NewPage
                    Inc iPag
                    icFilasTabla = 1
                End If
                
            
                rsFase.MoveNext
            Wend
            rsFase.Close
            
            rs.MoveNext
            
        Wend
        rs.Close
            
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias
End Sub
Private Sub cmdDorsalCatHorario_Click()
Dim rs As Recordset
Dim iContDorsales As Integer
Dim rsPart As Recordset
Dim rsFase As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim aDefCelda(4) As TCelda
Dim aTabla() As String
Dim icFilasTabla As Integer
Dim iDorsalesTanda As Integer
Dim iNumTandas As Integer
Dim iTandas As Integer
Dim iCopias As Integer
Dim bImprimirNombres As Boolean
Dim G_MAX_FILAS_POR_PAG_HOR As Integer

    G_MAX_FILAS_POR_PAG_HOR = 35
    bImprimirNombres = False
    
    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0765, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If chkCoordinadorHTML.Value = 0 Then
        ComprobarImpresoraPorDefecto
        CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
        CDialog.CancelError = True
        CDialog.FileName = "HorarioCoordinador.htm"
        On Local Error GoTo Pcancelar
        If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
        On Local Error GoTo 0
        CDialog.CancelError = False
        GoTo Pseguir
Pcancelar:
        Exit Sub
Pseguir:
        iCopias = CDialog.Copies
    Else
        CDialog.CancelError = True
        On Local Error GoTo Pcancelar1
        CDialog.FileName = "HorarioCoordinador"
        CDialog.DefaultExt = "HTML"
        CDialog.ShowSave
        On Local Error GoTo 0
        CDialog.CancelError = False
        GoTo Pseguir1
Pcancelar1:
        Exit Sub
Pseguir1:
        iCopias = 1
        If MsgBox(mml_FRASE1252, vbQuestion Or vbYesNo, "Question") = vbYes Then
            bImprimirNombres = True
        End If
        Open CDialog.FileName For Output As #100
        Print #100, "<HTML><BODY><CENTER><TABLE BORDER = 1>"
    End If
    
    iLineas = 4
    For iCCopias = 1 To iCopias
        ReDim aTabla(G_MAX_FILAS_POR_PAG_HOR + 4, 3)
        
        Printer.Orientation = vbPRORLandscape
        
        If chkCoordinadorHTML.Value = 0 Then
            aTabla(0, 0) = mml_FRASE0672
            aTabla(0, 1) = mml_FRASE0713
            aTabla(0, 2) = mml_FRASE0673
            aTabla(0, 3) = mml_FRASE0871
        Else
            aTabla(0, 0) = Mid(mml_FRASE0672, 5)
            aTabla(0, 1) = Mid(mml_FRASE0713, 5)
            aTabla(0, 2) = Mid(mml_FRASE0673, 5)
            aTabla(0, 3) = Mid(mml_FRASE0871, 5)
        End If
        
        aDefCelda(0).Ancho = 1000
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 1000
        aDefCelda(1).Justificado = eccentro
        aDefCelda(2).Ancho = 9500
        aDefCelda(2).Justificado = ecizquierda
        aDefCelda(3).Ancho = 4000
        aDefCelda(3).Justificado = ecizquierda
        icFilasTabla = 1
        
        iPag = 1
        iEscala = Printer.Width / 40
        
        'Imprimimos las categorias, junto con las entradas ficticias separadoras
        Set rs = db.OpenRecordset("(SELECT h.hora, id_categoria, c.codigo, ge.nombre, h.grupo, numfase, h.orden FROM categorias c, gruposedad ge, horario h " & _
                "WHERE h.cod_categoria = c.codigo AND cod_grupoedad = ge.codigo AND h.cod_competicion = " & tbCodComp.Text & _
                " AND grupo LIKE '*" & cbPista.Text & "*')" & _
                "UNION" & _
                "(SELECT h.hora, '', 0, '', h.grupo, 0, h.orden FROM horario h " & _
                "WHERE h.cod_categoria = 0 AND h.cod_competicion = " & tbCodComp.Text & _
                "AND grupo LIKE '*" & cbPista.Text & "*')" & _
                " ORDER BY h.orden, h.hora, h.grupo", dbOpenSnapshot)
        
        While Not rs.EOF
                aTabla(icFilasTabla, 0) = Format(CDate(rs!hora), "hh:mm")
                Select Case rs!numfase
                    Case "0":
                        aTabla(icFilasTabla, 1) = ""
                    Case "1":
                        aTabla(icFilasTabla, 1) = mml_FRASE0329
                    Case "2":
                        aTabla(icFilasTabla, 1) = mml_FRASE0872
                    Case "4":
                        aTabla(icFilasTabla, 1) = mml_FRASE0652
                    Case "99":
                        aTabla(icFilasTabla, 1) = "GenLook"
                    Case Else
                        aTabla(icFilasTabla, 1) = "1/" & Trim(rs!numfase)
                End Select
                
                'Inumtandas es cubierta por la función
                aTabla(icFilasTabla, 2) = Left(rs!grupo, MAX_GRUPO_HORARIO_COORD)
                aTabla(icFilasTabla, 3) = ""
                Set rsPart = db.OpenRecordset("SELECT DISTINCT d.num_dorsal, p.nombre_hombre, p.nombre_mujer FROM parejas p, dorsales d WHERE (d.fase = " & rs!numfase & " OR " & rs!numfase & " = " & C_FASE_GENERAL_LOOK & ") AND d.cod_categoria = " & rs.Fields(2) & " AND d.cod_pareja = p.codigo ORDER BY num_dorsal", dbOpenSnapshot)
                iContDorsales = 0
                iTandas = 1
                iNumTandas = 0
                iDorsalesTanda = CalcularDorsalesPorTandaCat(rs!codigo, rs!numfase, 0, iTandas, iNumTandas)
                While Not rsPart.EOF
                    If iContDorsales >= C_MAX_DORSALES_LIST_PART Then
                        Inc icFilasTabla
                        aTabla(icFilasTabla, 0) = ""
                        aTabla(icFilasTabla, 1) = ""
                        aTabla(icFilasTabla, 2) = ""
                        aTabla(icFilasTabla, 3) = ""
                        iContDorsales = 0
                    End If
                    If Not bImprimirNombres Then
                        aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & Trim$(Str$(rsPart.Fields(0)))
                    Else
                        aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & "<b>" & Trim$(Str$(rsPart.Fields(0))) & "</b>"
                        aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & " (" & Trim$(rsPart.Fields("nombre_hombre")) & "-" & Trim$(rsPart.Fields("nombre_mujer")) & ")<br>"
                    End If
                    rsPart.MoveNext
                    If (iContDorsales = iDorsalesTanda - 1 And iTandas < iNumTandas) Then
                        Inc iTandas
                        iDorsalesTanda = CalcularDorsalesPorTandaCat(rs!codigo, rs!numfase, 0, iTandas, iNumTandas)
                        iContDorsales = C_MAX_DORSALES_LIST_PART
                        If Not rsPart.EOF Then
                            If Not bImprimirNombres Then
                                aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & " |"
                            Else
                                aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & "<br>-------- New Round ---------<br>"
                            End If
                        End If
                    Else
                        If Not rsPart.EOF Then
                            If Not bImprimirNombres Then
                                aTabla(icFilasTabla, 3) = aTabla(icFilasTabla, 3) & ","
                            End If
                        End If
                    End If
                    Inc iContDorsales
                Wend
                rsPart.Close
                
                Inc icFilasTabla
                If icFilasTabla >= G_MAX_FILAS_POR_PAG_HOR - 1 Then
                    If chkCoordinadorHTML.Value = 0 Then
                        Printer.FontBold = False
                        Printer.FontSize = 10
                        Printer.Print mml_FRASE0659 & iPag
                        ImprimirCabecera
                        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                        Printer.NewPage
                    Else
                        Dim i As Integer
                        Dim j As Integer
                        Dim iCols As Integer
                        iCols = 4
                        
                        For i = 0 To icFilasTabla - 1
                            Print #100, "<tr>"
                            For j = 0 To iCols - 1
                                Print #100, "<td>"
                                Print #100, aTabla(i, j)
                                Print #100, "</td>"
                            Next
                            Print #100, "</tr>"
                        Next
                    End If
                    Inc iPag
                    icFilasTabla = 1
                End If
                rs.MoveNext
        Wend
        rs.Close
            
            
        If chkCoordinadorHTML.Value = 0 Then
            Printer.FontBold = False
            Printer.FontSize = 10
            Printer.Print mml_FRASE0659 & iPag
            ImprimirCabecera
            DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
            Printer.EndDoc
        Else
            iCols = 4
            
            For i = 0 To icFilasTabla - 1
                Print #100, "<tr>"
                For j = 0 To iCols - 1
                    Print #100, "<td>"
                    Print #100, aTabla(i, j)
                    Print #100, "</td>"
                Next
                Print #100, "</tr>"
            Next
            Print #100, "</TABLE></CENTER></BODY></HTML>"
            Close #100
        End If
    Next iCCopias
End Sub


Private Sub cmdListParejas_Click()
Dim rs As Recordset
Dim rsPart As Recordset
Dim rsDorsal As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iLineas As Integer
Dim iPag As Integer
Dim iNumero As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(4) As TCelda

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If


    If MsgBox(mml_FRASE0754, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
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
        
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 3)
        
        aTabla(0, 0) = mml_FRASE0873
        aTabla(0, 1) = mml_FRASE0874
        aTabla(0, 2) = mml_FRASE0875
        aTabla(0, 3) = mml_FRASE0760
        
        aDefCelda(0).Ancho = 1000
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 1000
        aDefCelda(1).Justificado = eccentro
        aDefCelda(2).Ancho = 7500
        aDefCelda(2).Justificado = ecizquierda
        aDefCelda(3).Ancho = 1000
        aDefCelda(3).Justificado = eccentro
        
        icFilasTabla = 1
        
        Set rsPart = db.OpenRecordset("SELECT DISTINCT p.nombre_hombre, p.nombre_mujer, p.codigo, nif_hombre, nif_mujer, fecha_nac_hombre, fecha_nac_mujer, num_socio_hombre, num_socio_mujer, grupoedad, cod_grupoedad, categoria, cod_modalidad, m.nombre, cod_competicion, combinar_edad  FROM parejas p, modalidad m WHERE  p.cod_modalidad = m.codigo AND p.cod_competicion = " & tbCodComp.Text & " ORDER BY cod_modalidad, categoria, cod_grupoedad", dbOpenSnapshot)
        While Not rsPart.EOF
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM parejas WHERE cod_competicion = " & rsPart!cod_competicion & " AND  categoria = '" & rsPart!categoria & "' AND cod_grupoedad=" & rsPart!cod_grupoedad & " AND cod_modalidad=" & rsPart!cod_modalidad, dbOpenSnapshot)
            Debug.Print "SELECT COUNT(*) FROM parejas WHERE cod_competicion = " & rsPart!cod_competicion & " AND  categoria = '" & rsPart!categoria & "' AND cod_grupoedad=" & rsPart!cod_grupoedad & " AND cod_modalidad=" & rsPart!cod_modalidad
            iNumero = rs.Fields(0)
            rs.Close
            
            aTabla(icFilasTabla, 0) = iNumero
            aTabla(icFilasTabla, 1) = rsPart!codigo
            aTabla(icFilasTabla, 2) = ".11n" & rsPart.Fields(0) & mml_FRASE0035 & rsPart.Fields(1)
            Set rsDorsal = db.OpenRecordset("SELECT num_dorsal FROM dorsales d WHERE  d.cod_pareja =" & rsPart!codigo, dbOpenSnapshot)
            If Not rsDorsal.EOF Then
                aTabla(icFilasTabla, 3) = rsDorsal!num_dorsal
            Else
                aTabla(icFilasTabla, 3) = mml_FRASE0762
            End If
            Inc icFilasTabla
            aTabla(icFilasTabla, 0) = ""
            aTabla(icFilasTabla, 1) = ""
            aTabla(icFilasTabla, 3) = ""
            aTabla(icFilasTabla, 2) = ".09n"
            
            If rsPart!combinar_edad = 1 Then
                aTabla(icFilasTabla, 2) = aTabla(icFilasTabla, 2) & mml_FRASE0876
            Else
                aTabla(icFilasTabla, 2) = aTabla(icFilasTabla, 2) & mml_FRASE0877
            End If
            aTabla(icFilasTabla, 2) = aTabla(icFilasTabla, 2) & mml_FRASE0878 & Mid$(rsPart!Nombre, 1, 3) & " " & rsPart!categoria & " " & rsPart!grupoedad
            Inc icFilasTabla
            aTabla(icFilasTabla, 2) = mml_FRASE0879 & rsPart!nif_hombre & mml_FRASE0880 & rsPart!fecha_nac_hombre & mml_FRASE0881 & rsPart!num_socio_hombre & ")"
            Inc icFilasTabla
            aTabla(icFilasTabla, 2) = mml_FRASE0882 & rsPart!nif_mujer & mml_FRASE0880 & rsPart!fecha_nac_mujer & mml_FRASE0881 & rsPart!num_socio_mujer & ")"
            
            rsPart.MoveNext
            
            Inc icFilasTabla
            If icFilasTabla >= G_MAX_FILAS_POR_PAG - 2 Then
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
        rsPart.Close
        
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
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

