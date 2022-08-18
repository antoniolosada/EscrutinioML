VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImprimirFinal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0680"
   ClientHeight    =   3015
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   8235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImpNoPasaron 
      Caption         =   "mml_FRASE1294"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4500
      TabIndex        =   19
      Top             =   2040
      Width           =   1845
   End
   Begin VB.CommandButton cmdImpPart 
      Caption         =   "mml_FRASE0978"
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
      Left            =   2460
      TabIndex        =   18
      Top             =   1920
      Width           =   1845
   End
   Begin VB.CommandButton cmdImpOrden 
      Caption         =   "mml_FRASE0681"
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
      Left            =   2460
      TabIndex        =   13
      Top             =   2460
      Width           =   1845
   End
   Begin VB.CommandButton cmdImprimirPart 
      Caption         =   "mml_FRASE0682"
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
      TabIndex        =   12
      Top             =   2460
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8175
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
         Picture         =   "frmImprimitFinal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   270
         Width           =   450
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
         Height          =   360
         Left            =   1845
         Picture         =   "frmImprimitFinal.frx":046A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   765
         Width           =   450
      End
      Begin VB.CommandButton CommandButton2 
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
         Picture         =   "frmImprimitFinal.frx":08D4
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1215
         Width           =   450
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
         Left            =   6600
         TabIndex        =   14
         Top             =   1200
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   1080
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         TabIndex        =   10
         Top             =   1200
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
         TabIndex        =   9
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox tbCodCateg 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbDescCateg 
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
         Top             =   720
         Width           =   4935
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
         Width           =   4935
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
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
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
         TabIndex        =   8
         Top             =   720
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
      Height          =   735
      Left            =   6960
      TabIndex        =   1
      Top             =   2040
      Width           =   1125
   End
   Begin VB.CommandButton cmdImpTablas 
      Caption         =   "mml_FRASE0683"
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
      Left            =   30
      TabIndex        =   0
      Top             =   1920
      Width           =   2265
   End
End
Attribute VB_Name = "frmImprimirFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Imprimir(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, Optional iRepesca As Integer = 0)
    tbCodComp.Text = sCodComp
    tbDescComp.Text = Buscar("competiciones", "descripcion", tbCodComp.Text)
    tbCodCateg.Text = sCodCat
    tbDescCateg.Text = Buscar("categorias", "descripcion", tbCodCateg.Text)
    tbCodFase.Text = sCodFase
    chkRep.Value = iRepesca
    tbDescFase.Text = LiteralFase(tbCodFase.Text)
    Me.Show 1
End Sub
Public Sub ImpresionCompleta(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, Optional iRepesca As Integer = 0)
    tbCodComp.Text = sCodComp
    tbDescComp.Text = Buscar("competiciones", "descripcion", tbCodComp.Text)
    tbCodCateg.Text = sCodCat
    tbDescCateg.Text = Buscar("categorias", "descripcion", tbCodCateg.Text)
    tbCodFase.Text = sCodFase
    chkRep.Value = iRepesca
    tbDescFase.Text = LiteralFase(tbCodFase.Text)
    
    If sCodFase > 1 Then
        If VersionActiva("IMPRESION_DORSALES_ELIMINADOS") Then
            ImprimirResumen Val(tbCodFase.Text)
        End If
        CDialog.Copies = VarCfg("no_copias_tablas")
        ImprimirResultados
        CDialog.Copies = VarCfg("no_copias_sig_fase")
        tbCodFase.Text = Val(tbCodFase.Text) / 2
        ImprimirSigFase
    Else
        CDialog.Copies = VarCfg("no_copias_tablas")
        ImprimirResultados
        CDialog.Copies = VarCfg("no_copias_resumen")
                
        frmImprimirInternet.gl_bGenerarInet = True
        ImprimirResumen
    End If
End Sub

Public Sub ImprimirResultados()
Dim rsBailes As Recordset
Dim iMaxHojas As Integer
Dim i As Integer
Dim iCCopias As Integer
Dim bTeamMatch As Boolean
Dim bCountryAmateur As Boolean

    If Not C_DEBUG Then On Local Error GoTo error
    
    bTeamMatch = ComprobarSiTeamMatch(tbCodCateg.Text)
    bCountryAmateur = ComprobarSiCountryAmateur(tbCodCateg.Text)
    
    For iCCopias = 1 To CDialog.Copies
        Set rsBailes = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
        rsBailes.Close
        
        If tbCodFase.Text = "1" Then
            Set rsBailes = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
            iMaxHojas = rsBailes.Fields(0)
            rsBailes.Close
            If iMaxHojas > C_BAILES_POR_PAG_FINAL Then
                iMaxHojas = 2
            Else
                iMaxHojas = 1
            End If
            For i = 1 To iMaxHojas
                If UCase(VarCfg("calcular_todos_los_puestos")) = "S" Then
                    If bTeamMatch Then
                        ImprimirTeamMatch i
                    ElseIf bCountryAmateur Then
                        ImprimirCountryAmateur i
                    Else
                        ImprimirFinal (i)
                    End If
                    If i < iMaxHojas Then
                        Printer.NewPage
                    End If
                Else
                    ImprimirFinal1 (i)
                End If
            Next
            Printer.EndDoc
        Else
            Set rsBailes = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_NO_FINAL, dbOpenSnapshot)
            iMaxHojas = rsBailes.Fields(0)
            rsBailes.Close
            If iMaxHojas > C_BAILES_POR_PAG_NO_FINAL Then
                iMaxHojas = 2
            Else
                iMaxHojas = 1
            End If
            For i = 1 To iMaxHojas
                ImprimirNoFinal (i)
                If i < iMaxHojas Then
                    Printer.NewPage
                End If
            Next
            Printer.EndDoc
        End If
    Next iCCopias
    Exit Sub
error:
    ProcesarError "ImprimirResultados"
End Sub

Private Sub cmdImpNoPasaron_Click()
    
    If tbCodComp.Text = "" Or tbCodCateg.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    CDialog.Copies = VarCfg("no_copias_resumen")
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    ImprimirResumen Val(tbCodFase.Text)
    If Not frmImprimirInternet.gl_bGenerarInet Then
        Printer.EndDoc
    End If
    
End Sub

Private Sub cmdImpTablas_Click()
Dim rsBailes As Recordset
Dim iMaxHojas As Integer
Dim i As Integer
Dim iCCopias As Integer
    
    ComprobarImpresoraPorDefecto
    
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    CDialog.Copies = VarCfg("no_copias_tablas")
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:

    
    ImprimirResultados
    
End Sub

Sub ImprimirCabecera(sCad As String, Optional bAviso As Boolean = True)
Dim iEscala As Integer
Dim rs As Recordset
Dim X As Integer, Y As Integer

    Printer.PaintPicture frmMenu.picEPA.Picture, Printer.Width - frmMenu.picEPA.Width - G_MARGEN_EPA, G_MARGEN_EPA_Y
    If G_LOGO_ESCUELA <> "" Then
        Printer.PaintPicture LoadPicture(G_LOGO_ESCUELA), G_MARGEN_EPA, G_MARGEN_EPA_Y + C_MARGEN_RESULTADOS_Y
    End If
    iEscala = Printer.Width / 10
    If bAviso Then
        Printer.FontSize = 8
        If G_COUNTRY Then
            If VersionActiva("MENSAJE CABECERA FEBD") Then
                Printer.Print mml_FRASE0684;
            Else
                Printer.Print mml_FRASE1034;
            End If
        Else
            Printer.Print mml_FRASE0684;
        End If
        Printer.Print " (" & Format$(Now, "dd/mm/yyyy") & "-" & Format$(Time, "hh:mm:ss") & ")"
        Printer.FontSize = 13
        Printer.Print
    End If
    Printer.CurrentX = 0
    Set rs = db.OpenRecordset(" SELECT descripcion, fecha FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    Printer.DrawWidth = 2
    X = Printer.CurrentX
    Y = Printer.CurrentY
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    SaltoLinea Printer, 4
    Printer.FontBold = True
    Centrado Printer, rs!DESCRIPCION & "  (" & rs!fecha & ")", Printer.Width
    rs.Close
    Printer.FontBold = False
    Printer.FontSize = 13
    Centrado Printer, sEscuela(tbCodComp.Text), Printer.Width
    Printer.FontBold = True
    Printer.FontSize = 11
    Set rs = db.OpenRecordset("SELECT id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE c.codigo = " & tbCodCateg.Text & " AND cod_grupoedad = ge.codigo ORDER BY hora", dbOpenSnapshot)
    Centrado Printer, mml_FRASE0685 & "(" & rs.Fields("codigo") & ") " & rs!DESCRIPCION & "  (" & sCad & ")", Printer.Width
    rs.Close
    Printer.FontBold = False
    SaltoLinea Printer, 4
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    Printer.DrawWidth = 1
    SaltoLinea Printer, 3
    Printer.FontSize = 10

End Sub

Private Sub ImprimirFinal(iHoja As Integer)
Dim iJueces As Integer
Dim iBailes As Integer
Dim iParejas As Integer
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim aBailes() As String
Dim aJueces() As String
Dim aDorsales() As String
Dim iCBailes As Integer
Dim iCJueces As Integer
Dim iCPuestos As Integer
Dim iCDorsales As Integer
Dim iEscala As Integer
Dim iTablas As Integer
Dim iPosY As Integer, iPosX As Integer, iPosFinBailesY As Integer
Dim dPuntosPuesto As Double
Dim iNumDorsales As Integer
Dim iNumRepPuesto As Integer
Dim X As Integer, Y As Integer

    If Not C_DEBUG Then On Local Error GoTo error

    If tbCodCateg.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    ImprimirCabecera mml_FRASE0329
        
    Printer.CurrentX = iEscala * 8
    Printer.Print mml_FRASE0557 & iHoja
    
    ' Comprobamos el número de jueces
    'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    'iJueces = rs.Fields(0)
    'rs.Close
    ' Comprobamos el número de jueces en las puntuaciones
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iJueces = 0
    If Not rs.EOF Then
        rs.MoveLast
        iJueces = rs.RecordCount
    End If
    rs.Close
    
    ' Comprobamos el número de bailes
    If BailesParciales(Val(tbCodCateg.Text)) Then
        ' Comprobamos el número de bailes que hay en las puntuaciones
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ")", dbOpenSnapshot)
    Else
        ' Comprobamos el número de bailes
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
    End If
    iBailes = rs.Fields(0)
    rs.Close
    
    'Comprobamos si tiene definidos bailes y jueces
    If iJueces = 0 Or iBailes = 0 Then
        MsgBox mml_FRASE0686, vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    End If
    
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    rs.Close
    
    ReDim aTabla(iParejas + 1, iJueces + iParejas + 2)
    'Cargamos las etiquetas de los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM juez_categ jc, jueces j WHERE pasos = 0 AND j.codigo = jc.cod_juez AND cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        ReDim aJueces(rs.RecordCount, 2)
        i = 0
        rs.MoveFirst
        While Not rs.EOF
            aJueces(i, 0) = rs!id_juez
            If G_COUNTRY Then
                aJueces(i, 1) = mml_FRASE0421 & " " & rs!id_juez
            Else
                aJueces(i, 1) = rs!Nombre
            End If
            rs.MoveNext
            i = i + 1
        Wend
    End If
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    If iHoja = 2 Then
        While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
            Inc iCBailes
            rsBailes.MoveNext
        Wend
    End If
    iCBailes = 0
    
    iTablas = 1
    aTabla(0, 0) = "Num"
    While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos los puestos
            Set rsPuestos = db.OpenRecordset("SELECT cod_juez, puesto FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " ORDER BY 1", dbOpenSnapshot)
            While Not rsPuestos.EOF
                aTabla(0, iCPuestos) = rsPuestos!cod_juez
                aTabla(iCDorsales, iCPuestos) = CadPuesto(rsPuestos!Puesto)
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            'Cargamos los calculos
            Set rsPuestos = db.OpenRecordset("SELECT puesto, posiciones_mi, suma_posmi FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " ORDER BY 1", dbOpenSnapshot)
            iCPuestos = 0
            While Not rsPuestos.EOF
                If rsPuestos!Puesto > 0 Then
                    aTabla(0, iJueces + rsPuestos!Puesto) = Trim$(Str$(rsPuestos!Puesto)) & "º"
                Else
                    aTabla(0, iJueces + iParejas + 1) = "Pos"
                End If
                If rsPuestos!posiciones_mi > 0 And rsPuestos!suma_posmi > 0 Then
                    aTabla(iCDorsales, iJueces + rsPuestos!Puesto) = Trim$(Str$(rsPuestos!posiciones_mi)) & "(" & Trim$(Str$(rsPuestos!suma_posmi)) & ")"
                ElseIf rsPuestos!posiciones_mi > 0 Then
                    'Posición total
                    aTabla(iCDorsales, iJueces + iParejas + 1) = Trim$(Str$(rsPuestos!posiciones_mi))
                    If InStr(aTabla(iCDorsales, 1), "d") > 0 Then
                        aTabla(iCDorsales, iJueces + iParejas + 1) = aTabla(iCDorsales, iJueces + iParejas + 1) + "d"
                    ElseIf InStr(aTabla(iCDorsales, 1), mml_FRASE0687) > 0 Then
                        aTabla(iCDorsales, iJueces + iParejas + 1) = mml_FRASE0687
                    End If
                Else
                    aTabla(iCDorsales, iJueces + rsPuestos!Puesto) = "-"
                End If
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
                Exit Sub
            End If
        Wend
        rsDorsales.Close
        
        If Printer.CurrentY + (iParejas + 1) * 400 + MARGEN_PAGINA > Printer.Height Then
            Printer.NewPage
        End If
        
        If (iTablas Mod 2 = 0) And iJueces <= G_LIM_JUECES_PARA_TABLAS_FINAL_DOBLES Then
            iPosX = C_POS_COLUMNA_2
        Else
            iPosY = Printer.CurrentY
            iPosX = 100
        End If
        Printer.CurrentY = iPosY
        Printer.Print
        Printer.CurrentX = iPosX
        Printer.FontSize = 10
        Printer.Print rsBailes.Fields(1)
        Printer.Print
        Printer.FontSize = 7
        Printer.FontName = "Arial"
        'Imprimimos los bailes
        DibujarTablaExt Printer, iPosX, Printer.CurrentY, iParejas + 1, iJueces + iParejas + 2, 350, 250, 350, 0, 0, iJueces + iParejas + 1, IIf(iJueces + iParejas > Val(C_LIM_TABLA_CON_PUESTOS), True, False)
        rsBailes.MoveNext
        Inc iCBailes
        Inc iTablas
    Wend
    
    
    If iJueces > G_LIM_JUECES_PARA_TABLAS_FINAL_DOBLES Then
        Printer.NewPage
        
        ImprimirCabecera mml_FRASE0329
    End If
    
    iPosFinBailesY = Printer.CurrentY
        
    'Imprimimos la tabla de los jueces
    ReDim aTabla(iJueces + 1, 2)
    
    If (iTablas Mod 2 = 0) And iJueces <= G_LIM_JUECES_PARA_TABLAS_FINAL_DOBLES Then
        iPosX = C_POS_COLUMNA_2
    Else
        iPosY = Printer.CurrentY
        iPosX = 100
    End If
    Printer.CurrentY = iPosY
    Printer.Print
    Printer.CurrentX = iPosX
    Printer.FontSize = 10
    Printer.Print mml_FRASE0689
    Printer.Print
    Printer.FontSize = 7
    Printer.FontName = "Arial"
    For i = 0 To iJueces
        aTabla(i, 0) = aJueces(i, 0)
        aTabla(i, 1) = aJueces(i, 1)
    Next
    'DibujarTabla Printer, iPosX, Printer.CurrentY, iJueces, 2, 2500, 250
    DibujarTablaExt Printer, iPosX, Printer.CurrentY, iJueces, 2, 3500, 250, 300, 1, 0, 0
        
    If iPosFinBailesY > Printer.CurrentY Then
        Printer.CurrentY = iPosFinBailesY
    End If
    
    ReDim aTabla(iParejas, 50)
    aTabla(0, 0) = "Num"
    'Ahora dibujamos las tablas de totales
    'Calculamos la diferencia de puntos entre puestos ignorando a las parejas adicionales
    'Set rs = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM dorsales WHERE num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    'Set rs = db.OpenRecordset("SELECT DISTINCT d.num_dorsal FROM puntuaciones p, dorsales d, parejas pa WHERE p.num_dorsal = d.num_dorsal AND p.cod_categoria = d.cod_categoria AND d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
    Set rs = db.OpenRecordset("SELECT DISTINCT d.num_dorsal FROM dorsales d, parejas pa WHERE d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
        dPuntosPuesto = 1000
        If Not rs.EOF Then
            rs.MoveLast
            iNumDorsales = rs.RecordCount
            If rs.RecordCount > 1 Then
                dPuntosPuesto = 1000 / (rs.RecordCount - 1)
            End If
        End If
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    While Not rsBailes.EOF
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos posiciones por baile
            Set rsPuestos = db.OpenRecordset("SELECT posiciones_mi FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " AND puesto=" & BD_POS_POSICION & " ORDER BY 1", dbOpenSnapshot)
            iCPuestos = 0
            aTabla(0, iCBailes + 1) = Mid$(rsBailes.Fields(1), 1, 1)
            aTabla(0, iBailes + 3) = mml_FRASE0690
            While Not rsPuestos.EOF
                If aTabla(iCDorsales, iBailes + 3) = "" Then
                    aTabla(iCDorsales, iBailes + 3) = "0"
                End If
                If NoPresente(rsDorsales!num_dorsal, tbCodCateg.Text, tbCodFase.Text, chkRep.Value) Then
                    aTabla(iCDorsales, iBailes + 3) = mml_FRASE0687
                    aTabla(iCDorsales, iCBailes + 1) = mml_FRASE0687
                Else
                    aTabla(iCDorsales, iBailes + 3) = Trim(Str$(Val(aTabla(iCDorsales, iBailes + 3)) + _
                                                    rsPuestos!posiciones_mi))
                    aTabla(iCDorsales, iCBailes + 1) = Trim$(Str$(rsPuestos!posiciones_mi))
                End If
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            
        Dim iNumParAdicPuestosInferiores As Integer
        Dim rsPuestosInf As Recordset
        Dim rsParAdic As Recordset
        Dim iPuesto As Integer
        Dim iPuntos As Integer
            'Cargamos la posición final
            Set rsPuestos = db.OpenRecordset("SELECT puesto FROM cal_conjunto WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND regla='FIN' ORDER BY 1", dbOpenSnapshot)
            'Si solo hay un dorsal no hay reglas
            aTabla(0, iBailes + 1) = mml_FRASE0433
            aTabla(0, iBailes + 2) = mml_FRASE0691
            If G_COUNTRY And iNumDorsales = 1 Then
                aTabla(1, iBailes + 1) = "1"
                aTabla(1, iBailes + 2) = "200"
            End If
If rsPuestos.EOF Then GoTo continuar
            'Comprobamos el número de dorsales no adicionales que tienen ese puesto
            'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto WHERE puesto = " & rsPuestos!puesto & " AND cod_categoria = " & tbCodCateg.Text & " AND regla='FIN' ORDER BY 1", dbOpenSnapshot)
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto c, dorsales d, parejas p WHERE c.num_dorsal = d.num_dorsal AND d.cod_pareja = p.codigo AND c.cod_Categoria = d.cod_categoria AND d.fase = 1 AND p.pareja_adicional = 0 AND c.puesto = " & rsPuestos!Puesto & " AND c.cod_categoria = " & tbCodCateg.Text & " AND c.regla='FIN' ORDER BY 1", dbOpenSnapshot)
            'Comprobamos el número de parejas adicionales que tienen puestos inferiores
            Set rsPuestosInf = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto c, dorsales d, parejas p WHERE c.num_dorsal = d.num_dorsal AND d.cod_pareja = p.codigo AND c.cod_Categoria = d.cod_categoria AND d.fase = 1 AND p.pareja_adicional = 1 AND c.cod_categoria = " & tbCodCateg.Text & " AND c.regla='FIN' AND c.puesto < " & rsPuestos!Puesto & " ORDER BY 1", dbOpenSnapshot)
            iNumParAdicPuestosInferiores = rsPuestosInf.Fields(0)
            rsPuestosInf.Close
            iNumRepPuesto = rs.Fields(0)
            rs.Close
            iCPuestos = 0
            While Not rsPuestos.EOF
                iPuesto = rsPuestos!Puesto - iNumParAdicPuestosInferiores
                aTabla(iCDorsales, iBailes + 1) = iPuesto
                ' Comprobamos si la pareja es adicional y no debe tener puntos
                Set rsParAdic = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND p.pareja_adicional = 1 AND d.cod_categoria = " & tbCodCateg.Text & " AND d.num_dorsal = " & rsDorsales!num_dorsal, dbOpenSnapshot)
                If rsParAdic.Fields(0) = 0 Then
                    ' Puntos
                    If G_MOSTRAR_PUNTOS Then
                        If G_COUNTRY Then
                            iPuntos = Redondea((iNumDorsales - (iPuesto + (iNumRepPuesto - 1) / 2)) / (iNumDorsales - 1) * 1000, 0)
                            Select Case iNumDorsales
                            Case 1
                                iPuntos = 200
                            Case 2
                                iPuntos = iPuntos / 3
                            Case 3
                                iPuntos = iPuntos / 2
                            Case 4
                                iPuntos = Redondea(iPuntos / 1.35, 0)
                            Case 5
                                iPuntos = Redondea(iPuntos / 1.2, 0)
                            End Select
                            If iPuntos = 0 Then iPuntos = 10
                            aTabla(iCDorsales, iBailes + 2) = iPuntos
                        Else
                            aTabla(iCDorsales, iBailes + 2) = Redondea((iNumDorsales - (iPuesto + (iNumRepPuesto - 1) / 2)) / (iNumDorsales - 1) * 1000, 0)
                        End If
                    Else
                        aTabla(iCDorsales, iBailes + 2) = ""
                    End If
                Else
                    aTabla(iCDorsales, iBailes + 2) = mml_FRASE0692
                End If
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
                rsParAdic.Close
            Wend
continuar:
            rsPuestos.Close
            
            'Cargamos las reglas
            Set rsPuestos = db.OpenRecordset("SELECT posicion ,regla, puesto FROM cal_conjunto WHERE cod_categoria = " _
                 & tbCodCateg.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & _
                 " AND regla<>'FIN' AND regla <> 'R09' ORDER BY 1,2,3", dbOpenSnapshot)
            iCPuestos = 0
    Dim iDatosRegla As Integer
    Dim sDatosRegla As String
    Dim iCalcPuesto As Integer
            While Not rsPuestos.EOF
                iCalcPuesto = rsPuestos!Posicion
                aTabla(0, iBailes + 4 + iCPuestos) = Mid$(rsPuestos!regla, 1, 3) & "(" & rsPuestos!Posicion & ")"
                iDatosRegla = InStr(rsPuestos!regla, "-")
                If iDatosRegla > 0 Then
                    sDatosRegla = Mid$(rsPuestos!regla, iDatosRegla + 1)
                Else
                    If rsPuestos!Puesto >= iCalcPuesto Then
                        sDatosRegla = ">=" & rsPuestos!Puesto
                    Else
                        sDatosRegla = rsPuestos!Puesto
                    End If
                End If
                aTabla(iCDorsales, iBailes + 4 + iCPuestos) = sDatosRegla
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            
            
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
            End If
        Wend
        rsDorsales.Close
        rsBailes.MoveNext
        iCBailes = iCBailes + 1
    Wend
    
    If Printer.CurrentY + (iParejas + 1) * 400 + MARGEN_PAGINA > Printer.Height Then
        Printer.NewPage
    End If
    
    iCPuestos = iCPuestos + 2
    Printer.Print
    Printer.FontSize = 14
    Printer.Print mml_FRASE0693
    Printer.FontSize = 10
    Printer.Print
    'DibujarTabla Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 2 + iCPuestos, 650, 350
    DibujarTablaExt Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 2 + iCPuestos, 650, 350, 500, iBailes + 1, 0, iBailes + 1
    Printer.Print
    On Local Error Resume Next
    If Mid$(C_LOGO_PATH, 1, 1) = "*" Then
        Printer.PaintPicture LoadPicture(Mid$(C_LOGO_PATH, 2)), (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    Else
        Printer.PaintPicture frmMenu.picEPADance.Picture, (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    End If
    On Local Error GoTo error
    Printer.Print
    
    ' {·M} Imprimimos el hueco para la firma
    iPosY = Printer.CurrentY
    Printer.CurrentX = 0
    iPosY = Printer.CurrentY
    Printer.FontSize = 10
    Printer.Print mml_FRASE0694
    Printer.Line (0, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.CurrentX = 1800
    Printer.CurrentY = iPosY
    Printer.Print mml_FRASE0033
    Printer.Line (1800, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    'Printer.EndDoc
    
    Exit Sub

error:
    Dim Msj As String
    If Err.Number <> 0 Then
       Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
             & Err.Source & Chr(13) & Err.Description
       MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
    End If

End Sub

Private Sub ImprimirCountryAmateur(iHoja As Integer)
Dim iJueces As Integer
Dim iBailes As Integer
Dim iParejas As Integer
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim aBailes() As String
Dim aJueces() As String
Dim aDorsales() As String
Dim iCBailes As Integer
Dim iCJueces As Integer
Dim iCPuestos As Integer
Dim iCDorsales As Integer
Dim iEscala As Integer
Dim iTablas As Integer
Dim iPosY As Integer, iPosX As Integer, iPosFinBailesY As Integer
Dim dPuntosPuesto As Double
Dim iNumDorsales As Integer
Dim iNumRepPuesto As Integer
Dim X As Integer, Y As Integer

    If Not C_DEBUG Then On Local Error GoTo error

    If tbCodCateg.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    ImprimirCabecera mml_FRASE0329
        
    Printer.CurrentX = iEscala * 8
    Printer.Print mml_FRASE0557 & iHoja
    
    ' Comprobamos el número de jueces
    'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    'iJueces = rs.Fields(0)
    'rs.Close
    ' Comprobamos el número de jueces en las puntuaciones
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iJueces = 0
    If Not rs.EOF Then
        rs.MoveLast
        iJueces = rs.RecordCount
    End If
    rs.Close
    
    ' Comprobamos el número de bailes
    If BailesParciales(Val(tbCodCateg.Text)) Then
        ' Comprobamos el número de bailes que hay en las puntuaciones
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ")", dbOpenSnapshot)
    Else
        ' Comprobamos el número de bailes
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
    End If
    iBailes = rs.Fields(0)
    rs.Close
    
    'Comprobamos si tiene definidos bailes y jueces
    If iJueces = 0 Or iBailes = 0 Then
        MsgBox mml_FRASE0686, vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    End If
    
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    rs.Close
    
    ReDim aTabla(iParejas + 1, iJueces + iParejas + 2)
    'Cargamos las etiquetas de los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM juez_categ jc, jueces j WHERE pasos = 0 AND j.codigo = jc.cod_juez AND cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        ReDim aJueces(rs.RecordCount, 2)
        i = 0
        rs.MoveFirst
        While Not rs.EOF
            aJueces(i, 0) = rs!id_juez
            If G_COUNTRY Then
                aJueces(i, 1) = mml_FRASE0421 & " " & rs!id_juez
            Else
                aJueces(i, 1) = rs!Nombre
            End If
            rs.MoveNext
            i = i + 1
        Wend
    End If
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    If iHoja = 2 Then
        While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
            Inc iCBailes
            rsBailes.MoveNext
        Wend
    End If
    iCBailes = 0
    
    iTablas = 1
    aTabla(0, 0) = "Num"
    While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos los puestos
            Set rsPuestos = db.OpenRecordset("SELECT cod_juez, puesto FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " ORDER BY 1", dbOpenSnapshot)
            While Not rsPuestos.EOF
                aTabla(0, iCPuestos) = rsPuestos!cod_juez
                aTabla(iCDorsales, iCPuestos) = CadPuesto(rsPuestos!Puesto)
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            'Cargamos los resultados totales por baile
            aTabla(0, iJueces + 1) = "Pos"
            Dim rsTotal As Recordset
            Set rsTotal = db.OpenRecordset("SELECT SUM(posiciones_mi) AS suma FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales.Fields("num_dorsal") & " AND cod_baile = " & rsBailes.Fields("cod_baile") & "  ORDER BY 1", dbOpenSnapshot)
            If Not rsTotal.EOF Then
                aTabla(iCDorsales, iJueces + 1) = rsTotal.Fields("suma")
            End If
            rsTotal.Close
            
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
                Exit Sub
            End If
        Wend
        rsDorsales.Close
        
        If Printer.CurrentY + (iParejas + 1) * 400 + MARGEN_PAGINA > Printer.Height Then
            Printer.NewPage
        End If
        
        If (iTablas Mod 2 = 0) And iJueces <= G_LIM_JUECES_PARA_TABLAS_FINAL_DOBLES Then
            iPosX = C_POS_COLUMNA_2
        Else
            iPosY = Printer.CurrentY
            iPosX = 100
        End If
        Printer.CurrentY = iPosY
        Printer.Print
        Printer.CurrentX = iPosX
        Printer.FontSize = 10
        Printer.Print rsBailes.Fields(1)
        Printer.Print
        Printer.FontSize = 7
        Printer.FontName = "Arial"
        'Imprimimos los bailes
        DibujarTablaExt Printer, iPosX, Printer.CurrentY, iParejas + 1, iJueces + 2, 350, 250, 350, 0, 0, iJueces + 1, IIf(iJueces + iParejas > Val(C_LIM_TABLA_CON_PUESTOS), True, False)
        rsBailes.MoveNext
        Inc iCBailes
        Inc iTablas
    Wend
    
    
    If iJueces > G_LIM_JUECES_PARA_TABLAS_FINAL_DOBLES Then
        Printer.NewPage
        
        ImprimirCabecera mml_FRASE0329
    End If
    
    iPosFinBailesY = Printer.CurrentY
        
    'Imprimimos la tabla de los jueces
    ReDim aTabla(iJueces + 1, 2)
    
    If (iTablas Mod 2 = 0) And iJueces <= G_LIM_JUECES_PARA_TABLAS_FINAL_DOBLES Then
        iPosX = C_POS_COLUMNA_2
    Else
        iPosY = Printer.CurrentY
        iPosX = 100
    End If
    Printer.CurrentY = iPosY
    Printer.Print
    Printer.CurrentX = iPosX
    Printer.FontSize = 10
    Printer.Print mml_FRASE0689
    Printer.Print
    Printer.FontSize = 7
    Printer.FontName = "Arial"
    For i = 0 To iJueces
        aTabla(i, 0) = aJueces(i, 0)
        aTabla(i, 1) = aJueces(i, 1)
    Next
    'DibujarTabla Printer, iPosX, Printer.CurrentY, iJueces, 2, 2500, 250
    DibujarTablaExt Printer, iPosX, Printer.CurrentY, iJueces, 2, 3500, 250, 300, 1, 0, 0
        
    If iPosFinBailesY > Printer.CurrentY Then
        Printer.CurrentY = iPosFinBailesY
    End If
    
    ReDim aTabla(iParejas, 50)
    aTabla(0, 0) = "Num"
    'Ahora dibujamos las tablas de totales
    'Calculamos la diferencia de puntos entre puestos ignorando a las parejas adicionales
    'Set rs = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM dorsales WHERE num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    'Set rs = db.OpenRecordset("SELECT DISTINCT d.num_dorsal FROM puntuaciones p, dorsales d, parejas pa WHERE p.num_dorsal = d.num_dorsal AND p.cod_categoria = d.cod_categoria AND d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
    Set rs = db.OpenRecordset("SELECT DISTINCT d.num_dorsal FROM dorsales d, parejas pa WHERE d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
        dPuntosPuesto = 1000
        If Not rs.EOF Then
            rs.MoveLast
            iNumDorsales = rs.RecordCount
            If rs.RecordCount > 1 Then
                dPuntosPuesto = 1000 / (rs.RecordCount - 1)
            End If
        End If
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    While Not rsBailes.EOF
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos posiciones por baile
            Set rsPuestos = db.OpenRecordset("SELECT posiciones_mi FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " AND puesto=" & BD_POS_POSICION & " ORDER BY 1", dbOpenSnapshot)
            iCPuestos = 0
            aTabla(0, iCBailes + 1) = Mid$(rsBailes.Fields(1), 1, 1)
            aTabla(0, iBailes + 3) = mml_FRASE0690
            While Not rsPuestos.EOF
                If aTabla(iCDorsales, iBailes + 3) = "" Then
                    aTabla(iCDorsales, iBailes + 3) = "0"
                End If
                If NoPresente(rsDorsales!num_dorsal, tbCodCateg.Text, tbCodFase.Text, chkRep.Value) Then
                    aTabla(iCDorsales, iBailes + 3) = mml_FRASE0687
                    aTabla(iCDorsales, iCBailes + 1) = mml_FRASE0687
                Else
                    aTabla(iCDorsales, iBailes + 3) = Trim(Str$(Val(aTabla(iCDorsales, iBailes + 3)) + _
                                                    rsPuestos!posiciones_mi))
                    aTabla(iCDorsales, iCBailes + 1) = Trim$(Str$(rsPuestos!posiciones_mi))
                End If
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            
        Dim iNumParAdicPuestosInferiores As Integer
        Dim rsPuestosInf As Recordset
        Dim rsParAdic As Recordset
        Dim iPuesto As Integer
        Dim iPuntos As Integer
            'Cargamos la posición final
            Set rsPuestos = db.OpenRecordset("SELECT puesto FROM cal_conjunto WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND regla='FIN' ORDER BY 1", dbOpenSnapshot)
            'Si solo hay un dorsal no hay reglas
            aTabla(0, iBailes + 1) = mml_FRASE0433
            aTabla(0, iBailes + 2) = mml_FRASE0691
            If G_COUNTRY And iNumDorsales = 1 Then
                aTabla(1, iBailes + 1) = "1"
                aTabla(1, iBailes + 2) = "200"
            End If
If rsPuestos.EOF Then GoTo continuar
            'Comprobamos el número de dorsales no adicionales que tienen ese puesto
            'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto WHERE puesto = " & rsPuestos!puesto & " AND cod_categoria = " & tbCodCateg.Text & " AND regla='FIN' ORDER BY 1", dbOpenSnapshot)
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto c, dorsales d, parejas p WHERE c.num_dorsal = d.num_dorsal AND d.cod_pareja = p.codigo AND c.cod_Categoria = d.cod_categoria AND d.fase = 1 AND p.pareja_adicional = 0 AND c.puesto = " & rsPuestos!Puesto & " AND c.cod_categoria = " & tbCodCateg.Text & " AND c.regla='FIN' ORDER BY 1", dbOpenSnapshot)
            'Comprobamos el número de parejas adicionales que tienen puestos inferiores
            Set rsPuestosInf = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto c, dorsales d, parejas p WHERE c.num_dorsal = d.num_dorsal AND d.cod_pareja = p.codigo AND c.cod_Categoria = d.cod_categoria AND d.fase = 1 AND p.pareja_adicional = 1 AND c.cod_categoria = " & tbCodCateg.Text & " AND c.regla='FIN' AND c.puesto < " & rsPuestos!Puesto & " ORDER BY 1", dbOpenSnapshot)
            iNumParAdicPuestosInferiores = rsPuestosInf.Fields(0)
            rsPuestosInf.Close
            iNumRepPuesto = rs.Fields(0)
            rs.Close
            iCPuestos = 0
            While Not rsPuestos.EOF
                iPuesto = rsPuestos!Puesto - iNumParAdicPuestosInferiores
                aTabla(iCDorsales, iBailes + 1) = iPuesto
                ' Comprobamos si la pareja es adicional y no debe tener puntos
                Set rsParAdic = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND p.pareja_adicional = 1 AND d.cod_categoria = " & tbCodCateg.Text & " AND d.num_dorsal = " & rsDorsales!num_dorsal, dbOpenSnapshot)
                If rsParAdic.Fields(0) = 0 Then
                    ' Puntos
                    If G_MOSTRAR_PUNTOS Then
                        If G_COUNTRY Then
                            iPuntos = Redondea((iNumDorsales - (iPuesto + (iNumRepPuesto - 1) / 2)) / (iNumDorsales - 1) * 1000, 0)
                            Select Case iNumDorsales
                            Case 1
                                iPuntos = 200
                            Case 2
                                iPuntos = iPuntos / 3
                            Case 3
                                iPuntos = iPuntos / 2
                            Case 4
                                iPuntos = Redondea(iPuntos / 1.35, 0)
                            Case 5
                                iPuntos = Redondea(iPuntos / 1.2, 0)
                            End Select
                            If iPuntos = 0 Then iPuntos = 10
                            aTabla(iCDorsales, iBailes + 2) = iPuntos
                        Else
                            aTabla(iCDorsales, iBailes + 2) = Redondea((iNumDorsales - (iPuesto + (iNumRepPuesto - 1) / 2)) / (iNumDorsales - 1) * 1000, 0)
                        End If
                    Else
                        aTabla(iCDorsales, iBailes + 2) = ""
                    End If
                Else
                    aTabla(iCDorsales, iBailes + 2) = mml_FRASE0692
                End If
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
                rsParAdic.Close
            Wend
continuar:
            rsPuestos.Close
                        
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
            End If
        Wend
        rsDorsales.Close
        rsBailes.MoveNext
        iCBailes = iCBailes + 1
    Wend
    
    If Printer.CurrentY + (iParejas + 1) * 400 + MARGEN_PAGINA > Printer.Height Then
        Printer.NewPage
    End If
    
    iCPuestos = iCPuestos + 2
    Printer.Print
    Printer.FontSize = 14
    Printer.Print mml_FRASE0693
    Printer.FontSize = 10
    Printer.Print
    'DibujarTabla Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 2 + iCPuestos, 650, 350
    DibujarTablaExt Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 2 + iCPuestos, 650, 350, 500, iBailes + 1, 0, iBailes + 1
    Printer.Print
    On Local Error Resume Next
    If Mid$(C_LOGO_PATH, 1, 1) = "*" Then
        Printer.PaintPicture LoadPicture(Mid$(C_LOGO_PATH, 2)), (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    Else
        Printer.PaintPicture frmMenu.picEPADance.Picture, (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    End If
    On Local Error GoTo error
    Printer.Print
    
    ' {·M} Imprimimos el hueco para la firma
    iPosY = Printer.CurrentY
    Printer.CurrentX = 0
    iPosY = Printer.CurrentY
    Printer.FontSize = 10
    Printer.Print mml_FRASE0694
    Printer.Line (0, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.CurrentX = 1800
    Printer.CurrentY = iPosY
    Printer.Print mml_FRASE0033
    Printer.Line (1800, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    'Printer.EndDoc
    
    Exit Sub

error:
    Dim Msj As String
    If Err.Number <> 0 Then
       Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
             & Err.Source & Chr(13) & Err.Description
       MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
    End If

End Sub


Private Sub ImprimirTeamMatch(iHoja As Integer)
Dim iJueces As Integer
Dim iBailes As Integer
Dim iParejas As Integer
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim aBailes() As String
Dim aJueces() As String
Dim aDorsales() As String
Dim iCBailes As Integer
Dim iCJueces As Integer
Dim iCPuestos As Integer
Dim iCDorsales As Integer
Dim iEscala As Integer
Dim iTablas As Integer
Dim iPosY As Integer, iPosX As Integer, iPosFinBailesY As Integer
Dim dPuntosPuesto As Double
Dim iNumDorsales As Integer
Dim iNumRepPuesto As Integer
Dim X As Integer, Y As Integer

    If Not C_DEBUG Then On Local Error GoTo error

    If tbCodCateg.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    ImprimirCabecera "TeamMatch"
        
    Printer.CurrentX = iEscala * 8
    Printer.Print mml_FRASE0557 & iHoja
    
    ' Comprobamos el número de jueces
    'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    'iJueces = rs.Fields(0)
    'rs.Close
    ' Comprobamos el número de jueces en las puntuaciones
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iJueces = 0
    If Not rs.EOF Then
        rs.MoveLast
        iJueces = rs.RecordCount
    End If
    rs.Close
    
    If iJueces = 0 Then
        MsgBox mml_FRASE0686, vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    End If
    
    ' Comprobamos el número de bailes
    If BailesParciales(Val(tbCodCateg.Text)) Then
        ' Comprobamos el número de bailes que hay en las puntuaciones
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ")", dbOpenSnapshot)
    Else
        ' Comprobamos el número de bailes
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
    End If
    iBailes = rs.Fields(0)
    rs.Close
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    rs.Close
    
    ReDim aTabla(iParejas + 1, iJueces + 2)
    'Cargamos las etiquetas de los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM juez_categ jc, jueces j WHERE pasos = 0 AND j.codigo = jc.cod_juez AND cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        ReDim aJueces(rs.RecordCount, 2)
        i = 0
        rs.MoveFirst
        While Not rs.EOF
            aJueces(i, 0) = rs!id_juez
            If G_COUNTRY Then
                aJueces(i, 1) = mml_FRASE0421 & " " & rs!id_juez
            Else
                aJueces(i, 1) = rs!Nombre
            End If
            rs.MoveNext
            i = i + 1
        Wend
    End If
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    If iHoja = 2 Then
        While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
            Inc iCBailes
            rsBailes.MoveNext
        Wend
    End If
    iCBailes = 0
    
    iTablas = 1
    aTabla(0, 0) = "Num"
    While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos los puestos
            Set rsPuestos = db.OpenRecordset("SELECT cod_juez, puesto FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " ORDER BY 1", dbOpenSnapshot)
            aTabla(0, iJueces + 1) = "Tot"
            While Not rsPuestos.EOF
                aTabla(0, iCPuestos) = rsPuestos!cod_juez
                aTabla(iCDorsales, iCPuestos) = CadPuestoDbl(rsPuestos!Puesto)
                aTabla(iCDorsales, iJueces + 1) = Val(aTabla(iCDorsales, iJueces + 1)) + rsPuestos!Puesto
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
                Exit Sub
            End If
        Wend
        rsDorsales.Close
        
        If Printer.CurrentY + (iParejas + 1) * 400 + MARGEN_PAGINA > Printer.Height Then
            Printer.NewPage
        End If
        
        If iTablas Mod 2 = 0 Then
            iPosX = C_POS_COLUMNA_2
        Else
            iPosY = Printer.CurrentY
            iPosX = 100
        End If
        Printer.CurrentY = iPosY
        Printer.Print
        Printer.CurrentX = iPosX
        Printer.FontSize = 10
        Printer.Print rsBailes.Fields(1)
        Printer.Print
        Printer.FontSize = 7
        Printer.FontName = "Arial"
        'Imprimimos los bailes
        'DibujarTabla Printer, iPosX, Printer.CurrentY, iParejas + 1, iJueces + iParejas + 1, 350, 250
        DibujarTablaExt Printer, iPosX, Printer.CurrentY, iParejas + 1, iJueces + 2, 350, 250, 350, 0, 0, iJueces + 1, IIf(iJueces + iParejas > Val(C_LIM_TABLA_CON_PUESTOS), True, False)
        rsBailes.MoveNext
        Inc iCBailes
        Inc iTablas
    Wend
    
    iPosFinBailesY = Printer.CurrentY
    
    'Imprimimos la tabla de los jueces
    ReDim aTabla(iJueces + 1, 2)
    
    If iTablas Mod 2 = 0 Then
        iPosX = C_POS_COLUMNA_2
    Else
        iPosY = Printer.CurrentY
        iPosX = 100
    End If
    Printer.CurrentY = iPosY
    Printer.Print
    Printer.CurrentX = iPosX
    Printer.FontSize = 10
    Printer.Print mml_FRASE0689
    Printer.Print
    Printer.FontSize = 7
    Printer.FontName = "Arial"
    For i = 0 To iJueces
        aTabla(i, 0) = aJueces(i, 0)
        aTabla(i, 1) = aJueces(i, 1)
    Next
    'DibujarTabla Printer, iPosX, Printer.CurrentY, iJueces, 2, 2500, 250
    DibujarTablaExt Printer, iPosX, Printer.CurrentY, iJueces, 2, 3500, 250, 300, 1, 0, 0
        
    If iPosFinBailesY > Printer.CurrentY Then
        Printer.CurrentY = iPosFinBailesY
    End If
    
    Printer.CurrentY = Printer.CurrentY + 300
    
    iPosFinBailesY = Printer.CurrentY
    
    'Imprimimos las puntuaciones totales
    Dim iNumCom As Integer, iPos As Integer
    iPosX = 100
    'Calculamos el sumatorio de todos los bailes de la categoría actual
    Set rs = db.OpenRecordset("SELECT d.num_dorsal, provincia, SUM(puesto) FROM puntuaciones pu, parejas pa, dorsales d WHERE pu.cod_categoria = d.cod_categoria AND d.num_dorsal = pu.num_dorsal AND d.cod_pareja = pa.codigo AND d.fase = pu.fase AND pu.cod_categoria = " & tbCodCateg.Text & " AND pu.fase = " & tbCodFase.Text & " GROUP BY provincia, d.num_dorsal ORDER BY provincia", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        iNumCom = rs.RecordCount
        rs.MoveFirst
        ReDim aTabla(iNumCom, 2)
        aTabla(0, 0) = "Comunidad"
        aTabla(0, 1) = "Ptos. Grupo"
        iPos = 1
        While Not rs.EOF
            aTabla(iPos, 0) = mml_FRASE0300 & rs!num_dorsal & " - " & rs!provincia
            aTabla(iPos, 1) = rs.Fields(2)
            Inc iPos
            rs.MoveNext
        Wend
    End If
    rs.Close
    DibujarTablaExt Printer, iPosX, iPosFinBailesY, iNumCom + 1, 2, 1800, 250, 1800, 0, 0, 0, True
    
    iPosX = 4000
    'Calculamos el sumatorio de todos los bailes de todas las categorias del TeamMatch
    Set rs = db.OpenRecordset("SELECT provincia, SUM(puesto) FROM puntuaciones pu, parejas pa, dorsales d WHERE pu.cod_categoria = d.cod_categoria AND d.num_dorsal = pu.num_dorsal AND d.cod_pareja = pa.codigo AND d.fase = pu.fase AND pu.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " AND descripcion LIKE '*" & C_TEAM_MATCH & "*') AND pu.fase = 1 GROUP BY provincia ORDER BY provincia", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        iNumCom = rs.RecordCount
        rs.MoveFirst
        ReDim aTabla(iNumCom, 2)
        aTabla(0, 0) = "Comunidad"
        aTabla(0, 1) = "PTOS. TOTALES"
        iPos = 1
        While Not rs.EOF
            aTabla(iPos, 0) = rs!provincia
            aTabla(iPos, 1) = rs.Fields(1)
            Inc iPos
            rs.MoveNext
        Wend
    End If
    rs.Close
    DibujarTablaExt Printer, iPosX, iPosFinBailesY, iNumCom + 1, 2, 1500, 250, 1500, 0, 0, 0, True
    
    Printer.CurrentY = Printer.CurrentY + 300
    On Local Error Resume Next
    If Mid$(C_LOGO_PATH, 1, 1) = "*" Then
        Printer.PaintPicture LoadPicture(Mid$(C_LOGO_PATH, 2)), (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    Else
        Printer.PaintPicture frmMenu.picEPADance.Picture, (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    End If
    On Local Error GoTo error
    Printer.Print
    
    ' {·M} Imprimimos el hueco para la firma
    iPosY = Printer.CurrentY
    Printer.CurrentX = 0
    iPosY = Printer.CurrentY
    Printer.FontSize = 10
    Printer.Print mml_FRASE0694
    Printer.Line (0, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.CurrentX = 1800
    Printer.CurrentY = iPosY
    Printer.Print mml_FRASE0033
    Printer.Line (1800, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    'Printer.EndDoc
    
    Exit Sub

error:
    Dim Msj As String
    If Err.Number <> 0 Then
       Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
             & Err.Source & Chr(13) & Err.Description
       MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
    End If

End Sub


Private Sub cmdCateg_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCateg.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text, "descripcion", " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCateg.Text = sResultado(2)
    tbCodFase.Text = ""
    tbDescFase.Text = ""

End Sub

Private Sub cmdImpOrden_Click()
Dim iEscala As Integer
Dim rs As Recordset
Dim rsDorsal As Recordset
Dim rsElim As Recordset
Dim rsMarcas As Recordset
Dim aPosicion() As Integer
Dim iParejas As Long
Dim iParejasPart As Long
Dim iCPareja As Long
Dim iUltimaFase As Integer
Dim iPosValorMax As Integer
Dim iCValorMax As Integer
Dim lValorMax As Long
Dim lValor As Long
Dim iCParejas As Long
Dim iCPuestos As Integer
Dim aTabla() As Integer
Dim iPuestoIni As Integer
Dim iRepesca As Integer
Dim iCCopias As Integer

Const C_NUM_DORSAL = 0
Const C_ELIM_SUP = 1
Const C_CONT_MARCAS = 2
Const C_PUESTO = 3
Const C_PUESTO_FIN = 4
Const C_PUNTOS = 5
Const C_COD_PAREJA = 6
Const C_PUESTOS_FINAL = 7
Const C_MARCAS_ELIM = 8
Const C_MAX_DIM = 9
Const C_NO_PRESENTADO = -1



    If tbCodComp.Text = "" Or tbCodCateg.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    CDialog.Copies = VarCfg("no_copias_resumen")
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    ImprimirResumen
    If Not frmImprimirInternet.gl_bGenerarInet Then
        Printer.EndDoc
    End If

End Sub

Sub ImprimirCabeceraResumen(Optional iFase As Integer = 0)
Dim rs As Recordset
Dim iEscala As Integer, X As Integer, Y As Integer

        Printer.PaintPicture frmMenu.picEPA.Picture, Printer.Width - frmMenu.picEPA.Width - G_MARGEN_EPA, G_MARGEN_EPA_Y
    If G_LOGO_ESCUELA <> "" Then
        Printer.PaintPicture LoadPicture(G_LOGO_ESCUELA), G_MARGEN_EPA, G_MARGEN_EPA_Y + C_MARGEN_RESULTADOS_Y
    End If
        iEscala = Printer.Width / 10
        Printer.FontSize = 8
        
        If G_COUNTRY Then
            If VersionActiva("MENSAJE CABECERA FEBD") Then
                Printer.Print mml_FRASE0684;
            Else
                Printer.Print mml_FRASE1034;
            End If
        Else
            Printer.Print mml_FRASE0684;
        End If
        
        Printer.Print " (" & Format$(Now, "dd/mm/yyyy") & "-" & Format$(Time, "hh:mm:ss") & ")"
        Printer.FontSize = 13
        Printer.Print
        Printer.CurrentX = 0
        Set rs = db.OpenRecordset(" SELECT descripcion, fecha FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
        Printer.DrawWidth = 2
        X = Printer.CurrentX
        Y = Printer.CurrentY
        Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
        SaltoLinea Printer, 4
        Printer.FontBold = True
        Centrado Printer, rs!DESCRIPCION & "  (" & rs!fecha & ")", Printer.Width
        rs.Close
        Printer.FontBold = False
        Printer.FontSize = 13
        Centrado Printer, sEscuela(tbCodComp.Text), Printer.Width
        Printer.FontBold = True
        Printer.FontSize = 11
        Set rs = db.OpenRecordset(" SELECT id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE c.codigo = " & tbCodCateg.Text & " AND cod_grupoedad = ge.codigo ORDER BY hora", dbOpenSnapshot)
        Dim sCad As String
        If iFase = 0 Then
        Centrado Printer, mml_FRASE0685 & rs!DESCRIPCION & mml_FRASE0695, Printer.Width
        Else
        Centrado Printer, mml_FRASE0685 & rs!DESCRIPCION & mml_FRASE1294, Printer.Width
        End If
        rs.Close
        Printer.FontBold = False
        SaltoLinea Printer, 4
        Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
        Printer.DrawWidth = 1
        SaltoLinea Printer, 3

End Sub

Public Sub ImprimirResumen(Optional iFase As Integer = 0)
Dim iEscala As Integer
Dim rs As Recordset
Dim rsDorsal As Recordset
Dim rsElim As Recordset
Dim rsMarcas As Recordset
Dim aPosicion() As Integer
Dim iParejas As Long
Dim iParejasPart As Long
Dim iCPareja As Long
Dim iUltimaFase As Integer
Dim iPosValorMax As Integer
Dim iCValorMax As Integer
Dim lValorMax As Long
Dim lValor As Long
Dim iCParejasNP As Integer
Dim iCParejas As Long
Dim iCPuestos As Integer
Dim aTabla() As Long
Dim iPuestoIni As Integer
Dim iRepesca As Integer
Dim iCCopias As Integer
Dim sPuesto As String
Dim sPuntos As String
Dim sDorsal As String
Dim X As Integer, Y As Integer
Dim aTablaResumen() As String
Dim aDefCelda(4) As TCelda
Dim icParejasTabla As Integer
Dim iPagina As Integer
Dim sBailes As String
Dim iCBailes As Integer
Dim sSQL As String

Const C_NUM_DORSAL = 0
Const C_ELIM_SUP = 1
Const C_CONT_MARCAS = 2
Const C_PUESTO = 3
Const C_PUESTO_FIN = 4
Const C_PUNTOS = 5
Const C_COD_PAREJA = 6
Const C_ULT_ELIM = 7
Const C_MAX_DIM = 8
Const C_NO_PRESENTADO = -1
Const C_MAX_FILAS_POR_PAG_RESUMEN = 30

Dim aTablasRes() As String
Const C_PUESTOS_FINAL = 0
Const C_MARCAS_ELIM = 1

    Printer.Orientation = vbPRORLandscape
    
    If Not C_DEBUG Then On Local Error GoTo error
    For iCCopias = 1 To CDialog.Copies
        iPagina = 1
        icParejasTabla = 1
        Printer.FontBold = False
        'Borramos la información del ResumenFinal de esta categoría
        db.Execute ("DELETE FROM resumenfinales WHERE cod_categoria = " & tbCodCateg.Text)
        db.Execute ("DELETE FROM Resultadosfinales WHERE cod_categoria = " & tbCodCateg.Text)
        
        ' Seleccionamos las parejas que tienen puntuación por ser oficiales
        ' Comprobamos el número de parejas anotadas
        Set rsDorsal = db.OpenRecordset("SELECT DISTINCT d.num_dorsal FROM dorsales d, parejas pa WHERE d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
        If Not rsDorsal.EOF Then
            rsDorsal.MoveLast
            iParejas = rsDorsal.RecordCount
            rsDorsal.MoveFirst
        Else
            MsgBox mml_FRASE0696, vbOKOnly Or vbCritical, mml_FRASE0096
            rsDorsal.Close
            Exit Sub
        End If
        ' Comprobamos el número de parejas oficiales que participaron (bailaron) (No se cuentan parejas adicionales ni especiales -> dorsal bajo)
        Set rs = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM dorsales d, parejas pa WHERE d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria=" & tbCodCateg.Text, dbOpenSnapshot)
        If Not rs.EOF Then
            rs.MoveLast
            iParejasPart = rs.RecordCount
            rs.Close
        Else
            MsgBox mml_FRASE0697, vbOKOnly Or vbCritical, mml_FRASE0096
            rs.Close
            Exit Sub
        End If
        ReDim aTabla(iParejas, C_MAX_DIM)
        ReDim aTablasRes(iParejas, 2)
        iCPareja = 0
        iCPuestos = 1
        'Recorremos todas las parejas anotadas
        While Not rsDorsal.EOF
            aTabla(iCPareja, C_NUM_DORSAL) = rsDorsal!num_dorsal
            
            'localizamos el código de pareja
            Set rs = db.OpenRecordset("SELECT cod_pareja FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If Not rs.EOF Then
                aTabla(iCPareja, C_COD_PAREJA) = rs!cod_pareja
            Else
                aTabla(iCPareja, C_COD_PAREJA) = 1
            End If
            rs.Close
            ' Primero comprobamos si tenemos un puesto en la final
            Set rsElim = db.OpenRecordset("SELECT puesto FROM cal_conjunto WHERE regla = 'FIN' AND cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            If Not rsElim.EOF Then
                aTabla(iCPareja, C_CONT_MARCAS) = 10 - rsElim!Puesto
                ' Si lo tenemos guardamos las posiciones por baile
                sBailes = ""
                
                'Set rs = db.OpenRecordset("SELECT cod_baile, posiciones_mi, b.nombre FROM cal_baile, bailes b WHERE cod_baile = b.codigo AND fase = 1 AND cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " AND puesto = 0 ORDER BY cod_baile", dbOpenSnapshot)
                
                Set rs = db.OpenRecordset("SELECT cb.cod_baile, posiciones_mi, b.nombre FROM cal_baile cb, bailes b, bailes_Categ bc WHERE bc.cod_categoria = cb.cod_categoria AND bc.cod_baile = cb.cod_baile AND bc.fase = 1 AND cb.cod_baile = b.codigo AND cb.fase = 1 AND cb.cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " AND puesto = 0 ORDER BY bc.posicion", dbOpenSnapshot)
                iCBailes = 0
                While Not rs.EOF
                    Inc iCBailes
                    If sBailes <> "" Then sBailes = sBailes & ", "
                    sBailes = sBailes & rs!Nombre
                    If aTablasRes(iCPareja, C_PUESTOS_FINAL) <> "" Then
                        aTablasRes(iCPareja, C_PUESTOS_FINAL) = aTablasRes(iCPareja, C_PUESTOS_FINAL) & "-"
                    End If
                    aTablasRes(iCPareja, C_PUESTOS_FINAL) = aTablasRes(iCPareja, C_PUESTOS_FINAL) & rs!posiciones_mi
                    If G_IMP_POS_BAILES_DIPLOMAS Then
                        db.Execute "INSERT INTO resultadosfinales VALUES (" & tbCodComp.Text & "," & tbCodCateg.Text & "," & rsDorsal!num_dorsal & "," & rs.Fields("cod_baile") & ",'" & rs.Fields("posiciones_mi") & "')"
                    End If
                    rs.MoveNext
                Wend
                rs.Close
            Else
                aTabla(iCPareja, C_CONT_MARCAS) = 0
            End If
            rsElim.Close
            
            'Contamos el número de marcas
            sSQL = "SELECT COUNT(*) FROM puntuaciones WHERE puesto > 0 AND cod_categoria=" & tbCodCateg.Text & " AND fase > 1 AND num_dorsal=" & rsDorsal!num_dorsal
            Debug.Print sSQL
            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
            aTablasRes(iCPareja, C_MARCAS_ELIM) = rs.Fields(0)
            rs.Close
            
            ' Por cada uno de los dorsales comprobamos el número de eliminatorias superadas
            Set rsElim = db.OpenRecordset("SELECT DISTINCT fase FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            If Not rsElim.EOF Then
                iUltimaFase = rsElim.Fields(0)
                rsElim.MoveLast
                aTabla(iCPareja, C_ELIM_SUP) = rsElim.RecordCount - 1
                aTabla(iCPareja, C_ULT_ELIM) = iUltimaFase
            Else
                aTabla(iCPareja, C_ELIM_SUP) = C_NO_PRESENTADO
                aTabla(iCPareja, C_PUESTO) = C_NO_PRESENTADO
                iUltimaFase = -1
            End If
            rsElim.Close
            ' y el número de marcas de la última eliminatoria
            ' Primero comprobamos si en la última eliminatoria contamos con una repesca, ya que las marcas serán las de la repesca
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE repesca = 1 AND fase = " & iUltimaFase & " AND cod_categoria=" & tbCodCateg.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If rs.Fields(0) > 0 Then
                iRepesca = 1
            Else
                iRepesca = 0
            End If
            rs.Close
            Set rsMarcas = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE puesto > 0 AND repesca=" & iRepesca & "AND fase = " & iUltimaFase & " AND cod_categoria=" & tbCodCateg.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If aTabla(iCPareja, C_CONT_MARCAS) = 0 Then
                If Not rsMarcas.EOF Then
                    aTabla(iCPareja, C_CONT_MARCAS) = rsMarcas.Fields(0)
                Else
                    aTabla(iCPareja, C_CONT_MARCAS) = 0
                End If
            End If
            rsMarcas.Close
            
            rsDorsal.MoveNext
            Inc iCPareja
        Wend
        rsDorsal.Close
        
        ReDim aTablaResumen(iParejasPart, 5)
        aTablaResumen(0, 0) = mml_FRASE0654
        aTablaResumen(0, 1) = mml_FRASE0698
        aTablaResumen(0, 2) = mml_FRASE0699
        aTablaResumen(0, 3) = mml_FRASE0700
        If G_COUNTRY Then
            aTablaResumen(0, 4) = mml_FRASE1045
        Else
            aTablaResumen(0, 4) = mml_FRASE0701
        End If
        
        If iCBailes <= C_BAILES_POR_HOJA Then
            aDefCelda(0).Ancho = 1000
            aDefCelda(0).Justificado = eccentro
            aDefCelda(1).Ancho = 1000
            aDefCelda(1).Justificado = eccentro
            aDefCelda(2).Ancho = 1500
            aDefCelda(2).Justificado = eccentro
            aDefCelda(3).Ancho = 1000
            aDefCelda(3).Justificado = eccentro
            aDefCelda(4).Ancho = 11400
            aDefCelda(4).Justificado = ecizquierda
        Else
            aDefCelda(0).Ancho = 1000
            aDefCelda(0).Justificado = eccentro
            aDefCelda(1).Ancho = 1000
            aDefCelda(1).Justificado = eccentro
            aDefCelda(2).Ancho = 2300
            aDefCelda(2).Justificado = eccentro
            aDefCelda(3).Ancho = 1000
            aDefCelda(3).Justificado = eccentro
            aDefCelda(4).Ancho = 10600
            aDefCelda(4).Justificado = ecizquierda
        End If
        
        ' Ahora asignamos los puestos
        Do While iCPuestos <= iParejasPart
            lValorMax = 0
            iCValorMax = 0
            'Recorremos todas las parejas
            For iCParejas = 0 To iParejas - 1
                ' Solo con los puestos no asignados
                If aTabla(iCParejas, C_PUESTO) = 0 Then
                    lValor = aTabla(iCParejas, C_ELIM_SUP) * 1000 + aTabla(iCParejas, C_CONT_MARCAS)
                    If lValor > lValorMax Then
                        lValorMax = lValor
                        iPosValorMax = iCParejas
                        iCValorMax = 1
                    ElseIf lValor = lValorMax Then
                        Inc iCValorMax
                    End If
                End If
            Next iCParejas
            
            ' Asignamos el puesto y los puntos
            'Recorremos todas las parejas
            For iCParejas = 0 To iParejas - 1
                ' Solo con los puestos no asignados
                If aTabla(iCParejas, C_PUESTO) = 0 Then
                    lValor = aTabla(iCParejas, C_ELIM_SUP) * 1000 + aTabla(iCParejas, C_CONT_MARCAS)
                    If lValor = lValorMax Then
                        aTabla(iCParejas, C_PUESTO) = iCPuestos
                        aTabla(iCParejas, C_PUESTO_FIN) = iCPuestos + iCValorMax - 1
                        'aTabla(iCParejas, C_PUNTOS) = Round((1000 / (iParejasPart - 1)) * (iParejasPart - iCPuestos - (iCValorMax - 1) / 2))
                        'Puntuación entre los participantes
                        'aTabla(iCParejas, C_PUNTOS) = Round((iParejasPart - iCPuestos) / (iParejasPart - 1) / iCValorMax * 1000)
                        'Puntuación entre los inscritos al cierre
                        ' Puntos
                        If G_MOSTRAR_PUNTOS Then
                            If G_COUNTRY Then
                            Dim iPuntos As Integer
                                If iParejas = 1 Then
                                    iPuntos = 200
                                Else
                                    iPuntos = Redondea((iParejas - (iCPuestos + (iCValorMax - 1) / 2)) / (iParejas - 1) * 1000, 0)
                                    Select Case iParejas
                                    Case 1
                                        iPuntos = 200
                                    Case 2
                                        iPuntos = iPuntos / 3
                                    Case 3
                                        iPuntos = iPuntos / 2
                                    Case 4
                                        iPuntos = Redondea(iPuntos / 1.35, 0)
                                    Case 5
                                        iPuntos = Redondea(iPuntos / 1.2, 0)
                                    End Select
                                    If iPuntos = 0 Then iPuntos = 10
                                End If
                                aTabla(iCParejas, C_PUNTOS) = iPuntos
                            Else
                                aTabla(iCParejas, C_PUNTOS) = Redondea((iParejas - (iCPuestos + (iCValorMax - 1) / 2)) / (iParejas - 1) * 1000, 0)
                            End If
                        End If
                        sDorsal = aTabla(iCParejas, C_NUM_DORSAL)
                        
                        aTablaResumen(icParejasTabla, 0) = sDorsal
                        
                        sPuesto = Trim$(aTabla(iCParejas, C_PUESTO))
                        If aTabla(iCParejas, C_PUESTO) <> aTabla(iCParejas, C_PUESTO_FIN) Then
                            sPuesto = sPuesto & "-" & aTabla(iCParejas, C_PUESTO_FIN)
                        End If
                        
                        aTablaResumen(icParejasTabla, 1) = sPuesto
                        
                        If aTablasRes(iCParejas, C_PUESTOS_FINAL) <> "" Then
                            aTablaResumen(icParejasTabla, 2) = aTablasRes(iCParejas, C_PUESTOS_FINAL)
                        Else
                            aTablaResumen(icParejasTabla, 2) = aTablasRes(iCParejas, C_MARCAS_ELIM) & " (" & sDescCortaFase(Val(aTabla(iCParejas, C_ULT_ELIM))) & ")"
                        End If
                        
                        sPuntos = aTabla(iCParejas, C_PUNTOS)
                        If G_MOSTRAR_PUNTOS Then
                            aTablaResumen(icParejasTabla, 3) = sPuntos
                        Else
                            aTablaResumen(icParejasTabla, 3) = ""
                        End If
                        Set rs = db.OpenRecordset("SELECT nombre_hombre, nombre_mujer, escuelas FROM parejas WHERE codigo =" & aTabla(iCParejas, C_COD_PAREJA), dbOpenSnapshot)
                        If Not rs.EOF Then
                            Dim Escuelas As String
                            Escuelas = rs.Fields("escuelas")
                            If Trim(Escuelas) <> "" Then
                                Escuelas = " (" & Escuelas & ")"
                            End If
                            aTablaResumen(icParejasTabla, 4) = rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer & Escuelas
                        End If
                        
                        If icParejasTabla >= C_MAX_FILAS_POR_PAG_RESUMEN Then
                            ImprimirCabeceraResumen
                            Printer.FontSize = 9
                            Printer.Print mml_FRASE0702 & iParejasPart & mml_FRASE0472 & iParejas & mml_FRASE0703 & iPagina
                            SaltoLinea Printer, 4
                            DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icParejasTabla, 5, aTablaResumen(), aDefCelda(), C_ALTO_CELDA_TABLAA
                            Printer.NewPage
                            iPagina = iPagina + 1
                            icParejasTabla = 1
                        Else
                            icParejasTabla = icParejasTabla + 1
                        End If
                        
                        rs.Close
                        db.Execute ("INSERT INTO resumenfinales VALUES (" & tbCodCateg.Text & ",'" & sDorsal & "','" & sPuesto & "','" & sPuntos & "')")
                    End If
                End If
            Next iCParejas
            
            If iCValorMax = 0 Then
                Exit Do
            End If
            iCPuestos = iCPuestos + iCValorMax
        Loop
        For iCParejasNP = 0 To iParejas - 1
            ' Solo con los N.P.
            If aTabla(iCParejasNP, C_PUESTO) = C_NO_PRESENTADO Then
                sDorsal = aTabla(iCParejasNP, C_NUM_DORSAL)
                aTablaResumen(icParejasTabla, 0) = sDorsal
                aTablaResumen(icParejasTabla, 1) = mml_FRASE0704
                aTablaResumen(icParejasTabla, 2) = "0"
                aTablaResumen(icParejasTabla, 3) = C_PUNTOS_NO_PRESENTADO
                Set rs = db.OpenRecordset("SELECT nombre_hombre, nombre_mujer FROM parejas WHERE codigo =" & aTabla(iCParejasNP, C_COD_PAREJA), dbOpenSnapshot)
                If Not rs.EOF Then
                    aTablaResumen(icParejasTabla, 4) = rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer
                End If
                If icParejasTabla >= C_MAX_FILAS_POR_PAG_RESUMEN Then
                    ImprimirCabeceraResumen
                    Printer.FontSize = 9
                    Printer.Print mml_FRASE0702 & iParejasPart & mml_FRASE0472 & iParejas & mml_FRASE0703 & iPagina
                    SaltoLinea Printer, 4
                    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icParejasTabla, 5, aTablaResumen(), aDefCelda(), C_ALTO_CELDA_TABLAA
                    Printer.NewPage
                    iPagina = iPagina + 1
                    icParejasTabla = 1
                Else
                    icParejasTabla = icParejasTabla + 1
                End If
                rs.Close
                db.Execute ("INSERT INTO resumenfinales VALUES (" & tbCodCateg.Text & ",'" & sDorsal & "','NO PRESENTADO','POR DETERMINAR')")
            End If
        Next iCParejasNP
        
        'Comprobamos si estamos emitiendo el listado de posiciones de los dorsales que no posaron
        'En este caso eliminamos a los primeros que pasaron de fase
        Dim iDorsalesSigFase As Integer
        iDorsalesSigFase = 0
        If iFase > 0 Then
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca = 0 AND fase = " & Int(2 ^ ((Log(iFase) / Log(2)) - 1)), dbOpenSnapshot)
            iDorsalesSigFase = rs.Fields(0)
            rs.Close
        End If
        
    ImprimirCabeceraResumen iFase
    
    Printer.FontSize = 9
    Printer.Print mml_FRASE0702 & iParejasPart & mml_FRASE0472 & iParejas & mml_FRASE0703 & iPagina
    SaltoLinea Printer, 4
    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icParejasTabla, 5, aTablaResumen(), aDefCelda(), C_ALTO_CELDA_TABLAA, iDorsalesSigFase
    Printer.Print
    Printer.FontBold = True
    Printer.Print mml_FRASE0706;
    Printer.FontBold = False
    Printer.Print sBailes
    If frmImprimirInternet.gl_bGenerarInet Then
        Printer.EndDoc
    Else
        Printer.NewPage
    End If
    
    Next iCCopias
    Exit Sub
error:
    ProcesarError "ImprimirResumen"
End Sub

Private Sub cmdImpPart_Click()
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rs As Recordset
Dim iLineas As Integer
Dim iEscala As Integer
Dim iMaxLineasPorPag As Integer
Dim iCCopias As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(2) As TCelda
Dim sCateg As String
Dim sFase As String

    
    If tbCodCateg.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    CDialog.Copies = VarCfg("no_copias_sig_fase")
    On Local Error GoTo Pcancelar
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    Select Case Val(tbCodFase.Text)
        Case 2:
            sFase = mml_FRASE0330
        Case 4:
            sFase = "Cuartos de Final"
        Case Else
            sFase = "1/" & tbCodFase.Text & mml_FRASE0708
    End Select
    
    ImprimirSigFase chkRep.Value

End Sub

Private Sub cmdImprimirPart_Click()
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rs As Recordset
Dim iLineas As Integer
Dim iEscala As Integer
Dim iMaxLineasPorPag As Integer
Dim iCCopias As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(2) As TCelda
Dim sCateg As String
Dim sFase As String


    If tbCodCateg.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    If Val(tbCodFase.Text) = 1 Then
        MsgBox mml_FRASE0707, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    Else
        'Solo imprimimos los de la siguiente fase si no es una repesca
        tbCodFase.Text = Val(tbCodFase.Text) / 2
    End If
    
    ComprobarImpresoraPorDefecto
    
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    CDialog.Copies = VarCfg("no_copias_sig_fase")
    On Local Error GoTo Pcancelar
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    Select Case Val(tbCodFase.Text)
        Case 2:
            sFase = mml_FRASE0330
        Case 4:
            sFase = "Cuartos de Final"
        Case Else
            sFase = "1/" & tbCodFase.Text & mml_FRASE0708
    End Select
    
    ImprimirSigFase
    
    chkRep.Value = 0
End Sub
Private Sub ImprimirSigFase(Optional iRepesca As Integer = 0)
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rs As Recordset
Dim iLineas As Integer
Dim iEscala As Integer
Dim iMaxLineasPorPag As Integer
Dim iCCopias As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(2) As TCelda
Dim sCateg As String
Dim sFase As String
    
    cmdImpTablas.Enabled = False
    For iCCopias = 1 To CDialog.Copies
        Select Case tbCodFase.Text
            Case 1:
                tbDescFase.Text = mml_FRASE0329
            Case 2:
                tbDescFase.Text = "SEMI-FINAL"
            Case 4:
                tbDescFase.Text = mml_FRASE0652
            Case 8:
                tbDescFase.Text = mml_FRASE0653
            Case Else
                tbDescFase.Text = tbCodFase.Text & "OS DE FINAL"
        End Select
        
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 3)
        
        aTabla(0, 0) = mml_FRASE0654
        aTabla(0, 1) = mml_FRASE0655
        aTabla(0, 2) = mml_FRASE0656
        
        aDefCelda(0).Ancho = 1200
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 4500
        aDefCelda(1).Justificado = ecizquierda
        aDefCelda(2).Ancho = 4500
        aDefCelda(2).Justificado = ecizquierda
                
        Set rs = db.OpenRecordset("SELECT descripcion FROM categorias WHERE codigo = " & tbCodCateg.Text, dbOpenSnapshot)
        sCateg = rs!DESCRIPCION & " (" & tbDescFase.Text & ")"
        If chkRep.Value = 1 Then
            If iRepesca = 1 Then
                'Imprimirmos la repesca real
                sCateg = sCateg & " (Rep)"
                sFase = tbDescFase.Text & " (Rep)"
            Else
                'Imprimimos la fase sig a la repesca
                sCateg = sCateg & mml_FRASE0658
                sFase = tbDescFase.Text & mml_FRASE0658
            End If
        Else
            sFase = tbDescFase.Text
        End If
        rs.Close
        
        ' Ahora recuperamos a los participantes de cada ronda
        iLineas = 0
        Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, p.nombre_hombre, p.nombre_mujer FROM dorsales d, parejas p WHERE repesca = " & iRepesca & " AND d.cod_pareja = p.codigo AND d.cod_categoria = " & tbCodCateg.Text & " AND d.fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            aTabla(icFilasTabla, 0) = ".12n" & rsDorsales!num_dorsal
            aTabla(icFilasTabla, 1) = ".12n" & rsDorsales!nombre_hombre
            aTabla(icFilasTabla, 2) = ".12n" & rsDorsales!nombre_mujer
            rsDorsales.MoveNext
            
            Inc icFilasTabla
            If icFilasTabla >= G_MAX_FILAS_POR_PAG - 3 Then
                Printer.FontBold = False
                Printer.FontSize = 10
                Printer.Print mml_FRASE0659 & iPag
                ImprimirCabecera sFase, False
                DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 3, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                Printer.NewPage
                Inc iPag
                icFilasTabla = 1
            End If
        Wend
        
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera sFase, False
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 3, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc
    Next iCCopias
    cmdImpTablas.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCateg.Text = ""
    tbDescCateg.Text = ""
    tbCodFase.Text = ""
    tbDescFase.Text = ""
End Sub


Private Sub CommandButton2_Click()
    If tbCodCateg.Text = "" Then
        MsgBox mml_FRASE0504, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " ORDER BY 1")

End Sub


Sub ImprimirNoFinal_PorBaile(iHoja As Integer)
Dim iJueces As Integer
Dim iBailes As Integer
Dim iParejas As Integer
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim aBailes() As String
Dim aJueces() As String
Dim aDorsales() As String
Dim iCBailes As Integer
Dim iCJueces As Integer
Dim iCPuestos As Integer
Dim iCDorsales As Integer
Dim iEscala As Integer
Dim iPosX As Integer, iPosY As Integer
Dim X As Integer, Y As Integer
Dim sFase As String

    If tbCodCateg.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    If Not C_DEBUG Then On Local Error GoTo error
        
    Select Case Val(tbCodFase.Text)
        Case 2:
            sFase = mml_FRASE0330
        Case 4:
            sFase = "Cuartos de Final"
        Case Else
            sFase = "1/" & tbCodFase.Text & mml_FRASE0708
    End Select
    
    'Imprimimos la cabecera
    ImprimirCabecera sFase
    
    Printer.Print mml_FRASE0557 & iHoja
    Printer.Print
    iPosY = Printer.CurrentY
    
    ' Comprobamos el número de jueces
    'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    'iJueces = rs.Fields(0)
    'rs.Close
    ' Comprobamos el número de jueces en las puntuaciones
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iJueces = 0
    If Not rs.EOF Then
        rs.MoveLast
        iJueces = rs.RecordCount
    End If
    rs.Close
    ' Comprobamos el número de bailes
    If BailesParciales(Val(tbCodCateg.Text)) Then
        ' Comprobamos el número de bailes que hay en las puntuaciones
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ")", dbOpenSnapshot)
    Else
        ' Comprobamos el número de bailes
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_NO_FINAL, dbOpenSnapshot)
    End If
    iBailes = rs.Fields(0)
    rs.Close
    
    If iJueces = 0 Or iBailes = 0 Then
        MsgBox mml_FRASE0981, vbOKOnly Or vbCritical, "ERROR"
        Exit Sub
    End If
    
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    rs.Close
    
    'Redimensionamos la tabla
    ReDim aTabla(iParejas, iJueces)
    
    'Cargamos las etiquetas de los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM juez_categ jc, jueces j WHERE pasos = 0 AND j.codigo = jc.cod_juez AND cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        ReDim aJueces(rs.RecordCount, 2)
        i = 0
        rs.MoveFirst
        While Not rs.EOF
            aJueces(i, 0) = rs!id_juez
            If G_COUNTRY Then
                aJueces(i, 1) = mml_FRASE0421 & " " & rs!id_juez
            Else
                aJueces(i, 1) = rs!Nombre
            End If
            rs.MoveNext
            i = i + 1
        Wend
    End If
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_NO_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_NO_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    If iHoja = 2 Then
        While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
            Inc iCBailes
            rsBailes.MoveNext
        Wend
    End If
    
    'Calculamos el tamaño de la fuente
    Dim iAnchoCol As Integer, iAltoCol As Integer

        G_TAM_FUENTE_TABLA_SEMI = Val(VarCfg("tam_fuente_tabla_semi"))
        Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI
        iAltoCol = Printer.TextHeight("|MM|") + 4
        iAnchoCol = Printer.TextWidth("|2399")
        If G_ANCHO_COL_JUEZ = 0 And VarCfg("ajusta_fuente_tabla_auto") = "S" And (iJueces + 1.00001) * iBailes * iAnchoCol > Printer.Width - G_MARGEN_ANCHO_TABLAS Then
            G_TAM_FUENTE_TABLA_SEMI = Printer.FontSize * (Printer.Width - G_MARGEN_ANCHO_TABLAS) / ((iJueces + 1.0000001) * iBailes * iAnchoCol)
        End If
        
        Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI
        If G_ANCHO_COL_JUEZ = 0 Then
            iAnchoCol = Printer.TextWidth("|239")
        Else
            iAnchoCol = G_ANCHO_COL_JUEZ
        End If
    
    iCBailes = 0
    While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        aTabla(0, 0) = "Nu."
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos los puestos
            Set rsPuestos = db.OpenRecordset("SELECT cod_juez, puesto FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " ORDER BY 1", dbOpenSnapshot)
            While Not rsPuestos.EOF
                aTabla(0, iCPuestos) = rsPuestos!cod_juez
                If rsPuestos!Puesto > 0 Then
                    aTabla(iCDorsales, iCPuestos) = "X"
                ElseIf rsPuestos!Puesto < 0 Then
                    aTabla(iCDorsales, iCPuestos) = "d"
                Else
                    aTabla(iCDorsales, iCPuestos) = ""
                End If
                
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
                Exit Sub
            End If
        Wend
        rsDorsales.Close
        
        Printer.CurrentX = iPosX
        Printer.CurrentY = iPosY
        Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI + 1
        Printer.Print rsBailes.Fields(1)
        Printer.Print
        Printer.FontName = "Arial"
        Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI
        DibujarTabla Printer, iPosX, Printer.CurrentY, iParejas + 1, iJueces + 1, iAnchoCol, iAltoCol
        rsBailes.MoveNext
        Inc iCBailes
        iPosX = iPosX + iAnchoCol * (iJueces + 1)
    Wend
    
    ReDim aTabla(iParejas, iBailes + 2)
    
    
    'Ahora dibujamos las tablas de totales
    '*********************************************************************************************
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_NO_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_NO_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 1
    While Not rsBailes.EOF
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        aTabla(0, 0) = "Nu."
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos posiciones por baile
            Set rsPuestos = db.OpenRecordset("SELECT posiciones_mi FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " AND puesto=" & BD_POS_POSICION & " ORDER BY 1", dbOpenSnapshot)
            iCPuestos = 0
            aTabla(0, iCBailes) = Mid$(rsBailes.Fields(1), 1, 1)
            aTabla(0, iBailes + 1) = mml_FRASE0709
            aTabla(0, iBailes + 2) = mml_FRASE0710
            'Comprobamos los que pasaron
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND repesca=0 AND fase = " & Val(tbCodFase.Text) / 2 & " AND cod_categoria=" & tbCodCateg.Text, dbOpenSnapshot)
            If rs.Fields(0) > 0 Then
                aTabla(iCDorsales, iBailes + 2) = "X"
            Else
                aTabla(iCDorsales, iBailes + 2) = ""
            End If
            rs.Close
            While Not rsPuestos.EOF
                aTabla(iCDorsales, iBailes + 1) = Val(aTabla(iCDorsales, iBailes + 1)) + _
                                                    rsPuestos!posiciones_mi
                aTabla(iCDorsales, iCBailes) = Val(rsPuestos!posiciones_mi)
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
            End If
        Wend
        rsDorsales.Close
        rsBailes.MoveNext
        Inc iCBailes
    Wend
    
    If Printer.CurrentY + (iParejas + 2) * 250 + MARGEN_PAGINA > Printer.Height Then
        Printer.NewPage
        ImprimirCabecera sFase
        Printer.FontSize = 12
        Printer.Print mml_FRASE0557 + Str$(iHoja + 1)
    End If
    
    Printer.Print
    Printer.FontSize = 12
    iPosY = Printer.CurrentY
    Printer.Print mml_FRASE0693
    Printer.FontSize = 10
    Printer.Print
    'Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI
    'DibujarTabla Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 3, iAnchoCol, iAltoCol
    Printer.FontSize = 8
    DibujarTabla Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 3, 300, 250
    
    'Imprimimos la tabla de los jueces
    ReDim aTabla(iJueces + 1, 2)
    Printer.CurrentX = C_POS_TABLA_JUECES_NO_FINAL
    Printer.CurrentY = iPosY
    Printer.FontSize = 12
    Printer.Print mml_FRASE0689
    Printer.Print
    Printer.FontSize = 8
    Printer.FontName = "Arial"
    For i = 0 To iJueces
        aTabla(i, 0) = aJueces(i, 0)
        aTabla(i, 1) = aJueces(i, 1)
    Next
    'DibujarTabla Printer, C_POS_TABLA_JUECES_NO_FINAL, Printer.CurrentY, iJueces, 2, 2500, 250
    DibujarTablaExt Printer, C_POS_TABLA_JUECES_NO_FINAL, Printer.CurrentY, iJueces, 2, 3500, 250, 300, 1, 0, 0
    
    Printer.Print
    On Local Error Resume Next
    If Mid$(C_LOGO_PATH, 1, 1) = "*" Then
        Printer.PaintPicture LoadPicture(Mid$(C_LOGO_PATH, 2)), C_POS_TABLA_JUECES_NO_FINAL, Printer.CurrentY
    Else
        Printer.PaintPicture frmMenu.picEPADance.Picture, C_POS_TABLA_JUECES_NO_FINAL, Printer.CurrentY
    End If
    On Local Error GoTo error
    
    ' {·M} Imprimimos el hueco para la firma
    iPosX = C_POS_TABLA_JUECES_NO_FINAL
    iPosY = Printer.CurrentY + 2200
    Printer.CurrentY = iPosY
    Printer.CurrentX = iPosX
    Printer.FontSize = 10
    Printer.Print mml_FRASE0694
    Printer.Line (iPosX, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.CurrentX = iPosX + 1800
    Printer.CurrentY = iPosY
    Printer.Print mml_FRASE0033
    Printer.Line (iPosX + 1800, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Exit Sub
error:
    Dim Msj As String
    If Err.Number <> 0 Then
       Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
             & Err.Source & Chr(13) & Err.Description
       MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
    End If
End Sub

Sub ImprimirNoFinal(iHoja As Integer)
    If G_ELIMINATORIAS_PAGINADAS Then
        ImprimirNoFinal_Paginado iHoja
    Else
        ImprimirNoFinal_PorBaile iHoja
    End If
End Sub

Sub ImprimirNoFinal_Paginado(iHoja As Integer)
Dim iJueces As Integer
Dim iBailes As Integer
Dim iParejas As Integer
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim aBailes() As String
Dim aJueces() As String
Dim aDorsales() As String
Dim iCBailes As Integer
Dim iCJueces As Integer
Dim iCPuestos As Integer
Dim iCDorsales As Integer
Dim iEscala As Integer
Dim iPosX As Integer, iPosY As Integer
Dim X As Integer, Y As Integer
Dim sFase As String
Dim bMasHojas As Boolean
Dim iNumDorsalComienzo As Long
Dim iPrimerBaile As Long

    If tbCodCateg.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    If Not C_DEBUG Then On Local Error GoTo error
        
    sFase = sDescFase(Val(tbCodFase.Text))
    
    ' Comprobamos el número de jueces
    'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    'iJueces = rs.Fields(0)
    'rs.Close
    ' Comprobamos el número de jueces en las puntuaciones
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iJueces = 0
    If Not rs.EOF Then
        rs.MoveLast
        iJueces = rs.RecordCount
    End If
    rs.Close
    
    ' Comprobamos el número de bailes
    'If BailesParciales(Val(tbCodCateg.Text)) Then
    '    ' Comprobamos el número de bailes que hay en los cálculos
    '    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ")", dbOpenSnapshot)
    'Else
    '    ' Comprobamos el número de bailes
    '    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_NO_FINAL, dbOpenSnapshot)
    'End If
    
    ' Comprobamos el número de bailes que hay en los cálculos
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ")", dbOpenSnapshot)
    iBailes = rs.Fields(0)
    rs.Close
    
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    rs.Close
    
    'Redimensionamos la tabla
    ReDim aTabla(iParejas, iJueces)
    
    'Cargamos las etiquetas de los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM juez_categ jc, jueces j WHERE pasos = 0 AND j.codigo = jc.cod_juez AND cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        ReDim aJueces(rs.RecordCount, 2)
        i = 0
        rs.MoveFirst
        While Not rs.EOF
            aJueces(i, 0) = rs!id_juez
            If G_COUNTRY Then
                aJueces(i, 1) = mml_FRASE0421 & " " & rs!id_juez
            Else
                aJueces(i, 1) = rs!Nombre
            End If
            rs.MoveNext
            i = i + 1
        Wend
    End If
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_NO_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_NO_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    If iHoja = 2 Then
        While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
            Inc iCBailes
            rsBailes.MoveNext
        Wend
    End If
    
    'Calculamos el tamaño de la fuente
    Dim iAnchoCol As Integer, iAltoCol As Integer

        G_TAM_FUENTE_TABLA_SEMI = Val(VarCfg("tam_fuente_tabla_semi"))
        Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI
        iAltoCol = Printer.TextHeight("|MM|") + 4
        iAnchoCol = Printer.TextWidth("|2399")
        If G_ANCHO_COL_JUEZ = 0 And VarCfg("ajusta_fuente_tabla_auto") = "S" And (iJueces + 1) * IIf(iBailes > C_BAILES_POR_HOJA, C_BAILES_POR_HOJA, iBailes) * iAnchoCol > Printer.Width - G_MARGEN_ANCHO_TABLAS Then
            G_TAM_FUENTE_TABLA_SEMI = Printer.FontSize * (Printer.Width - G_MARGEN_ANCHO_TABLAS) / ((iJueces + 1) * IIf(iBailes > C_BAILES_POR_HOJA, C_BAILES_POR_HOJA, iBailes) * iAnchoCol)
        End If
        
        Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI
        If G_ANCHO_COL_JUEZ = 0 Then
            iAnchoCol = Printer.TextWidth("|239")
        Else
            iAnchoCol = G_ANCHO_COL_JUEZ
        End If
    
    iCBailes = 0
    Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
    bMasHojas = True
    'Imprimimos la cabecera
    ImprimirCabecera sFase
    
    Printer.Print mml_FRASE0557 & iHoja
    Printer.Print
    iPosY = Printer.CurrentY
    iPosX = 0
    iPrimerBaile = rsBailes.AbsolutePosition
    While Not rsDorsales.EOF
        
        iNumDorsalComienzo = rsDorsales.AbsolutePosition
        
        rsBailes.AbsolutePosition = iPrimerBaile
        iCBailes = 0
        While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
            rsDorsales.AbsolutePosition = iNumDorsalComienzo
            
            'Cargamos los dorsales
            iCDorsales = 1
            aTabla(0, 0) = "Nu."
            bMasHojas = True
            While Not rsDorsales.EOF
                aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
                iCPuestos = 1
                'Cargamos los puestos
                Set rsPuestos = db.OpenRecordset("SELECT cod_juez, puesto FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " ORDER BY 1", dbOpenSnapshot)
                While Not rsPuestos.EOF
                    aTabla(0, iCPuestos) = rsPuestos!cod_juez
                    If rsPuestos!Puesto > 0 Then
                        aTabla(iCDorsales, iCPuestos) = "X"
                    ElseIf rsPuestos!Puesto < 0 Then
                        aTabla(iCDorsales, iCPuestos) = "d"
                    Else
                        aTabla(iCDorsales, iCPuestos) = ""
                    End If
                    
                    rsPuestos.MoveNext
                    iCPuestos = iCPuestos + 1
                Wend
                rsPuestos.Close
                rsDorsales.MoveNext
                iCDorsales = iCDorsales + 1
                If iCDorsales > G_MAX_DORSALES_HOJA_SEMI Then
                    GoTo continuar
                End If
                
                If Not rsDorsales.EOF And iCDorsales > iParejas Then
                    MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
                    Exit Sub
                End If
            Wend
continuar:
            Printer.CurrentX = iPosX
            Printer.CurrentY = iPosY
            Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI + 1
            Printer.Print rsBailes.Fields(1)
            Printer.Print
            Printer.FontName = "Arial"
            Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI
            'DibujarTabla Printer, iPosX, Printer.CurrentY, iParejas + 1, iJueces + 1, iAnchoCol, iAltoCol
            DibujarTabla Printer, iPosX, Printer.CurrentY, iCDorsales, iJueces + 1, iAnchoCol, iAltoCol
            rsBailes.MoveNext
            Inc iCBailes
            iPosX = iPosX + iAnchoCol * (iJueces + 1)
        Wend
        'Si ya no quedan dorsales no efectuamos el salto de página
        If Not rsDorsales.EOF Then
            Printer.NewPage
            Inc iHoja
            'Imprimimos la cabecera
            ImprimirCabecera sFase
            
            Printer.Print mml_FRASE0557 & iHoja
            Printer.Print
            iPosY = Printer.CurrentY
            iPosX = 0
        End If
    Wend 'MasHojas
    rsDorsales.Close
    
    ReDim aTabla(iParejas, iBailes + 2)
    
    
    'Ahora dibujamos las tablas de totales
    '*********************************************************************************************
    
    'If BailesParciales(Val(tbCodCateg.Text)) Then
    '    sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_NO_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
    '    Debug.Print sExecSQL
    '    Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    'Else
    '    Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_NO_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    'End If
    
    sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_NO_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
    Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    'Se calcula cada baile de modo independiente
    iCBailes = 1
    While Not rsBailes.EOF
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        aTabla(0, 0) = "Nu."
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos posiciones por baile
            Set rsPuestos = db.OpenRecordset("SELECT posiciones_mi FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " AND puesto=" & BD_POS_POSICION & " ORDER BY 1", dbOpenSnapshot)
            iCPuestos = 0
            aTabla(0, iCBailes) = Mid$(rsBailes.Fields(1), 1, 1)
            aTabla(0, iBailes + 1) = mml_FRASE0709
            aTabla(0, iBailes + 2) = mml_FRASE0710
            'Comprobamos los que pasaron
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND repesca=0 AND fase = " & Val(tbCodFase.Text) / 2 & " AND cod_categoria=" & tbCodCateg.Text, dbOpenSnapshot)
            If rs.Fields(0) > 0 Then
                aTabla(iCDorsales, iBailes + 2) = "X"
            Else
                aTabla(iCDorsales, iBailes + 2) = ""
            End If
            rs.Close
            While Not rsPuestos.EOF
                aTabla(iCDorsales, iBailes + 1) = Val(aTabla(iCDorsales, iBailes + 1)) + _
                                                    rsPuestos!posiciones_mi
                aTabla(iCDorsales, iCBailes) = Val(rsPuestos!posiciones_mi)
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
            End If
        Wend
        rsDorsales.Close
        rsBailes.MoveNext
        Inc iCBailes
    Wend
    
    If Printer.CurrentY + CLng(iParejas + 2) * 250 + MARGEN_PAGINA > Printer.Height Then
        Printer.NewPage
        ImprimirCabecera sFase
        Printer.FontSize = 12
        Inc iHoja
        Printer.Print mml_FRASE0557 + Str$(iHoja)
    End If
    
    Printer.Print
    Printer.FontSize = 12
    iPosY = Printer.CurrentY
    Printer.Print mml_FRASE0693
    Printer.FontSize = 10
    Printer.Print
    'Printer.FontSize = G_TAM_FUENTE_TABLA_SEMI
    'DibujarTabla Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 3, iAnchoCol, iAltoCol
    Printer.FontSize = 8
    
    Dim iFilas As Integer
    Dim iPaginas As Integer
    
    iPaginas = iParejas \ G_MAX_DORSALES_HOJA_SEMI
    If iParejas Mod G_MAX_DORSALES_HOJA_SEMI > 0 Then
        Inc iPaginas
    End If
    For i = 1 To iPaginas
        Printer.FontSize = 8
        If i = iPaginas And iParejas Mod G_MAX_DORSALES_HOJA_SEMI > 0 Then
            iFilas = iParejas Mod G_MAX_DORSALES_HOJA_SEMI
        Else
            iFilas = G_MAX_DORSALES_HOJA_SEMI
        End If
        DibujarTabla Printer, 100, Printer.CurrentY, iFilas + 1, iBailes + 3, 300, 250, True, (i - 1) * G_MAX_DORSALES_HOJA_SEMI
        If i < iPaginas Then
            Printer.NewPage
            ImprimirCabecera sFase
            Printer.FontSize = 12
            Inc iHoja
            Printer.Print mml_FRASE0557 + Str$(iHoja)
        End If
    Next
    
    'DibujarTabla Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 3, 300, 250
    
    'Imprimimos la tabla de los jueces
    ReDim aTabla(iJueces + 1, 2)
    Printer.CurrentX = C_POS_TABLA_JUECES_NO_FINAL
    Printer.CurrentY = iPosY
    Printer.FontSize = 12
    Printer.Print mml_FRASE0689
    Printer.Print
    Printer.FontSize = 8
    Printer.FontName = "Arial"
    For i = 0 To iJueces
        aTabla(i, 0) = aJueces(i, 0)
        aTabla(i, 1) = aJueces(i, 1)
    Next
    'DibujarTabla Printer, C_POS_TABLA_JUECES_NO_FINAL, Printer.CurrentY, iJueces, 2, 2500, 250
    DibujarTablaExt Printer, C_POS_TABLA_JUECES_NO_FINAL, Printer.CurrentY, iJueces, 2, 3500, 250, 300, 1, 0, 0
    
    Printer.Print
    On Local Error Resume Next
    If Mid$(C_LOGO_PATH, 1, 1) = "*" Then
        Printer.PaintPicture LoadPicture(Mid$(C_LOGO_PATH, 2)), C_POS_TABLA_JUECES_NO_FINAL, Printer.CurrentY
    Else
        Printer.PaintPicture frmMenu.picEPADance.Picture, C_POS_TABLA_JUECES_NO_FINAL, Printer.CurrentY
    End If
    On Local Error GoTo error
    
    ' {·M} Imprimimos el hueco para la firma
    iPosX = C_POS_TABLA_JUECES_NO_FINAL
    iPosY = Printer.CurrentY + 2200
    Printer.CurrentY = iPosY
    Printer.CurrentX = iPosX
    Printer.FontSize = 10
    Printer.Print mml_FRASE0694
    Printer.Line (iPosX, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.CurrentX = iPosX + 1800
    Printer.CurrentY = iPosY
    Printer.Print mml_FRASE0033
    Printer.Line (iPosX + 1800, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Exit Sub
error:
    Dim Msj As String
    If Err.Number <> 0 Then
       Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
             & Err.Source & Chr(13) & Err.Description
       MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
    End If
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    If Not VersionActiva("IMPRESION_DORSALES_ELIMINADOS") Then
        cmdImpNoPasaron.Visible = False
    End If
End Sub

Private Sub tbCodCateg_GotFocus()
    tbCodCateg.SelStart = 0
    tbCodCateg.SelLength = Len(tbCodCateg.Text)
End Sub

Private Sub tbCodCateg_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbCodCateg_LostFocus()
    ComprobarCategyFase tbCodCateg, tbDescCateg, tbCodFase, tbDescFase
End Sub
Private Sub tbCodFase_Change()
    Select Case tbCodFase.Text
        Case 1:
            tbDescFase.Text = mml_FRASE0329
        Case 2:
            tbDescFase.Text = "SEMI-FINAL"
        Case 4:
            tbDescFase.Text = "CUARTOS DE FINAL"
        Case 8:
            tbDescFase.Text = "OCTAVOS DE FINAL"
        Case Else
            tbDescFase.Text = tbCodFase.Text & "OS DE FINAL"
    End Select

End Sub

Function CadPuesto(iPuesto As Integer) As String
    If iPuesto = C_MAX_PUESTO_FINAL_NO_PRESENTADO Then
        CadPuesto = mml_FRASE0687
    ElseIf iPuesto > C_PUESTO_NEG Then
        CadPuesto = Trim$(Str$(iPuesto - C_PUESTO_NEG)) & "d"
    Else
        CadPuesto = iPuesto
    End If
End Function

Function CadPuestoDbl(iPuesto As Double) As String
    If iPuesto = C_MAX_PUESTO_FINAL_NO_PRESENTADO Then
        CadPuestoDbl = mml_FRASE0687
    ElseIf iPuesto > C_PUESTO_NEG Then
        CadPuestoDbl = Trim$(Str$(iPuesto - C_PUESTO_NEG)) & "d"
    Else
        CadPuestoDbl = iPuesto
    End If
End Function

































Private Sub ImprimirFinal1(iHoja As Integer)
Dim iJueces As Integer
Dim iBailes As Integer
Dim iParejas As Integer
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim aBailes() As String
Dim aJueces() As String
Dim aDorsales() As String
Dim iCBailes As Integer
Dim iCJueces As Integer
Dim iCPuestos As Integer
Dim iCDorsales As Integer
Dim iEscala As Integer
Dim iTablas As Integer
Dim iPosY As Integer, iPosX As Integer, iPosFinBailesY As Integer
Dim dPuntosPuesto As Double
Dim iNumDorsales As Integer
Dim iNumRepPuesto As Integer
Dim X As Integer, Y As Integer

    'On Local Error GoTo Error

    If tbCodCateg.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    Printer.PaintPicture frmMenu.picEPA.Picture, Printer.Width - frmMenu.picEPA.Width - G_MARGEN_EPA, G_MARGEN_EPA_Y
    iEscala = Printer.Width / 10
    Printer.FontSize = 8
    If G_COUNTRY Then
        Printer.Print mml_FRASE1034;
    Else
        Printer.Print mml_FRASE0684;
    End If
    Printer.FontSize = 13
    Printer.Print
    Printer.CurrentX = 0
    Set rs = db.OpenRecordset(" SELECT descripcion, fecha FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    Printer.DrawWidth = 2
    X = Printer.CurrentX
    Y = Printer.CurrentY
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    SaltoLinea Printer, 4
    Printer.FontBold = True
    Centrado Printer, rs!DESCRIPCION & "  (" & rs!fecha & ")", Printer.Width
    rs.Close
    Printer.FontBold = False
    Printer.FontSize = 13
    Centrado Printer, sEscuela(tbCodComp.Text), Printer.Width
    Printer.FontBold = True
    Printer.FontSize = 11
    Set rs = db.OpenRecordset(" SELECT id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE c.codigo = " & tbCodCateg.Text & " AND cod_grupoedad = ge.codigo ORDER BY hora", dbOpenSnapshot)
    Centrado Printer, mml_FRASE0685 & rs!DESCRIPCION & mml_FRASE0711, Printer.Width
    rs.Close
    Printer.FontBold = False
    SaltoLinea Printer, 4
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    Printer.DrawWidth = 1
    SaltoLinea Printer, 3
    Printer.FontSize = 10
    
    Printer.CurrentX = iEscala * 8
    Printer.Print mml_FRASE0557 & iHoja
    
    ' Comprobamos el número de jueces
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    iJueces = rs.Fields(0)
    rs.Close
    ' Comprobamos el número de bailes
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
    iBailes = rs.Fields(0)
    rs.Close
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    rs.Close
    
    ReDim aTabla(iParejas + 1, iJueces + iParejas + 1)
    'Cargamos las etiquetas de los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM juez_categ jc, jueces j WHERE pasos = 0 AND j.codigo = jc.cod_juez AND cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
    ReDim aJueces(iJueces, 2)
    i = 0
    While Not rs.EOF
        aJueces(i, 0) = rs!id_juez
        If G_COUNTRY Then
            aJueces(i, 1) = mml_FRASE0421 & " " & rs!id_juez
        Else
            aJueces(i, 1) = rs!Nombre
        End If
        rs.MoveNext
        i = i + 1
    Wend
    rs.Close
    Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    If iHoja = 2 Then
        While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
            Inc iCBailes
            rsBailes.MoveNext
        Wend
    End If
    iCBailes = 0
    
    iTablas = 1
    aTabla(0, 0) = "Num"
    While Not rsBailes.EOF And iCBailes < C_BAILES_POR_PAG_FINAL
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos los puestos
            Set rsPuestos = db.OpenRecordset("SELECT cod_juez, puesto FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " ORDER BY 1", dbOpenSnapshot)
            While Not rsPuestos.EOF
                aTabla(0, iCPuestos) = rsPuestos!cod_juez
                aTabla(iCDorsales, iCPuestos) = CadPuesto(rsPuestos!Puesto)
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            'Cargamos los calculos
            Set rsPuestos = db.OpenRecordset("SELECT puesto, posiciones_mi, suma_posmi FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " ORDER BY 1", dbOpenSnapshot)
            iCPuestos = 0
            While Not rsPuestos.EOF
                If rsPuestos!Puesto > 0 Then
                    aTabla(0, iJueces + rsPuestos!Puesto) = Trim$(Str$(rsPuestos!Puesto)) & "º"
                Else
                    aTabla(0, iJueces + iParejas) = "F"
                End If
                If rsPuestos!posiciones_mi > 0 And rsPuestos!suma_posmi > 0 Then
                    aTabla(iCDorsales, iJueces + rsPuestos!Puesto) = Trim$(Str$(rsPuestos!posiciones_mi)) & "(" & Trim$(Str$(rsPuestos!suma_posmi)) & ")"
                ElseIf rsPuestos!posiciones_mi > 0 Then
                    'Posición total
                    aTabla(iCDorsales, iJueces + iParejas) = Trim$(Str$(rsPuestos!posiciones_mi))
                    If InStr(aTabla(iCDorsales, 1), "d") > 0 Then
                        aTabla(iCDorsales, iJueces + iParejas) = aTabla(iCDorsales, iJueces + iParejas) + "d"
                    End If
                Else
                    aTabla(iCDorsales, iJueces + rsPuestos!Puesto) = "-"
                End If
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
                Exit Sub
            End If
        Wend
        rsDorsales.Close
        
        If Printer.CurrentY + (iParejas + 1) * 400 + MARGEN_PAGINA > Printer.Height Then
            Printer.NewPage
        End If
        
        If iTablas Mod 2 = 0 Then
            iPosX = C_POS_COLUMNA_2
        Else
            iPosY = Printer.CurrentY
            iPosX = 100
        End If
        Printer.CurrentY = iPosY
        Printer.Print
        Printer.CurrentX = iPosX
        Printer.FontSize = 10
        Printer.Print rsBailes.Fields(1)
        Printer.Print
        Printer.FontSize = 7
        Printer.FontName = "Arial"
        'Imprimimos los bailes
        'DibujarTabla Printer, iPosX, Printer.CurrentY, iParejas + 1, iJueces + iParejas + 1, 350, 250
        DibujarTablaExt Printer, iPosX, Printer.CurrentY, iParejas + 1, iJueces + iParejas + 1, 350, 250, 350, 0, 0, iJueces + iParejas
        rsBailes.MoveNext
        Inc iCBailes
        Inc iTablas
    Wend
    
    iPosFinBailesY = Printer.CurrentY
        
    'Imprimimos la tabla de los jueces
    ReDim aTabla(iJueces + 1, 2)
    
    If iTablas Mod 2 = 0 Then
        iPosX = C_POS_COLUMNA_2
    Else
        iPosY = Printer.CurrentY
        iPosX = 100
    End If
    Printer.CurrentY = iPosY
    Printer.Print
    Printer.CurrentX = iPosX
    Printer.FontSize = 10
    Printer.Print mml_FRASE0689
    Printer.Print
    Printer.FontSize = 7
    Printer.FontName = "Arial"
    For i = 0 To iJueces
        aTabla(i, 0) = aJueces(i, 0)
        aTabla(i, 1) = aJueces(i, 1)
    Next
    'DibujarTabla Printer, iPosX, Printer.CurrentY, iJueces, 2, 2500, 250
    DibujarTablaExt Printer, iPosX, Printer.CurrentY, iJueces, 2, 3500, 250, 300, 1, 0, 0
        
    If iPosFinBailesY > Printer.CurrentY Then
        Printer.CurrentY = iPosFinBailesY
    End If
    
    ReDim aTabla(iParejas, 50)
    aTabla(0, 0) = "Num"
    'Ahora dibujamos las tablas de totales
    'Calculamos la diferencia de puntos entre puestos ignorando a las parejas adicionales
    'Set rs = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM dorsales WHERE num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    Set rs = db.OpenRecordset("SELECT DISTINCT d.num_dorsal FROM puntuaciones p, dorsales d, parejas pa WHERE p.num_dorsal = d.num_dorsal AND p.cod_categoria = d.cod_categoria AND d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
        If Not rs.EOF Then
            rs.MoveLast
            iNumDorsales = rs.RecordCount
            dPuntosPuesto = 1000 / (rs.RecordCount - 1)
        Else
            dPuntosPuesto = 1000
        End If
    rs.Close
    Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    While Not rsBailes.EOF
        'Cargamos los dorsales
        iCDorsales = 1
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            aTabla(iCDorsales, 0) = rsDorsales!num_dorsal
            iCPuestos = 1
            'Cargamos posiciones por baile
            Set rsPuestos = db.OpenRecordset("SELECT posiciones_mi FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes.Fields(0) & " AND puesto=" & BD_POS_POSICION & " ORDER BY 1", dbOpenSnapshot)
            iCPuestos = 0
            aTabla(0, iCBailes + 1) = Mid$(rsBailes.Fields(1), 1, 1)
            aTabla(0, iBailes + 3) = mml_FRASE0690
            While Not rsPuestos.EOF
                If aTabla(iCDorsales, iBailes + 3) = "" Then
                    aTabla(iCDorsales, iBailes + 3) = "0"
                End If
                aTabla(iCDorsales, iBailes + 3) = Trim(Str$(Val(aTabla(iCDorsales, iBailes + 3)) + _
                                                    rsPuestos!posiciones_mi))
                aTabla(iCDorsales, iCBailes + 1) = Trim$(Str$(rsPuestos!posiciones_mi))
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            
        Dim iNumParAdicPuestosInferiores As Integer
        Dim rsPuestosInf As Recordset
        Dim rsParAdic As Recordset
        Dim iPuesto As Integer
            'Cargamos la posición final
            Set rsPuestos = db.OpenRecordset("SELECT puesto FROM cal_conjunto WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND regla='FIN' ORDER BY 1", dbOpenSnapshot)
            'Comprobamos el número de dorsales no adicionales que tienen ese puesto
            'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto WHERE puesto = " & rsPuestos!puesto & " AND cod_categoria = " & tbCodCateg.Text & " AND regla='FIN' ORDER BY 1", dbOpenSnapshot)
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto c, dorsales d, parejas p WHERE c.num_dorsal = d.num_dorsal AND d.cod_pareja = p.codigo AND c.cod_Categoria = d.cod_categoria AND d.fase = 1 AND p.pareja_adicional = 0 AND c.puesto = " & rsPuestos!Puesto & " AND c.cod_categoria = " & tbCodCateg.Text & " AND c.regla='FIN' ORDER BY 1", dbOpenSnapshot)
            'Comprobamos el número de parejas adicionales que tienen puestos inferiores
            Set rsPuestosInf = db.OpenRecordset("SELECT COUNT(*) FROM cal_conjunto c, dorsales d, parejas p WHERE c.num_dorsal = d.num_dorsal AND d.cod_pareja = p.codigo AND c.cod_Categoria = d.cod_categoria AND d.fase = 1 AND p.pareja_adicional = 1 AND c.cod_categoria = " & tbCodCateg.Text & " AND c.regla='FIN' AND c.puesto < " & rsPuestos!Puesto & " ORDER BY 1", dbOpenSnapshot)
            iNumParAdicPuestosInferiores = rsPuestosInf.Fields(0)
            rsPuestosInf.Close
            iNumRepPuesto = rs.Fields(0)
            rs.Close
            iCPuestos = 0
            aTabla(0, iBailes + 1) = mml_FRASE0433
            aTabla(0, iBailes + 2) = mml_FRASE0691
            While Not rsPuestos.EOF
                iPuesto = rsPuestos!Puesto - iNumParAdicPuestosInferiores
                aTabla(iCDorsales, iBailes + 1) = iPuesto
                ' Comprobamos si la pareja es adicional y no debe tener puntos
                Set rsParAdic = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND p.pareja_adicional = 1 AND d.cod_categoria = " & tbCodCateg.Text & " AND d.num_dorsal = " & rsDorsales!num_dorsal, dbOpenSnapshot)
                If rsParAdic.Fields(0) = 0 Then
                    ' Puntos
                    aTabla(iCDorsales, iBailes + 2) = Redondea((iNumDorsales - (iPuesto + (iNumRepPuesto - 1) / 2)) / (iNumDorsales - 1) * 1000, 0)
                Else
                    aTabla(iCDorsales, iBailes + 2) = mml_FRASE0692
                End If
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
                rsParAdic.Close
            Wend
            rsPuestos.Close
            
            'Cargamos las reglas
            Set rsPuestos = db.OpenRecordset("SELECT posicion ,regla, puesto FROM cal_conjunto WHERE cod_categoria = " _
                 & tbCodCateg.Text & " AND num_dorsal = " & rsDorsales!num_dorsal & _
                 " AND regla<>'FIN' AND regla <> 'R09' ORDER BY 1,2,3", dbOpenSnapshot)
            iCPuestos = 0
    Dim iDatosRegla As Integer
    Dim sDatosRegla As String
    Dim iCalcPuesto As Integer
            While Not rsPuestos.EOF
                iCalcPuesto = rsPuestos!Posicion
                aTabla(0, iBailes + 4 + iCPuestos) = Mid$(rsPuestos!regla, 1, 3) & "(" & rsPuestos!Posicion & ")"
                iDatosRegla = InStr(rsPuestos!regla, "-")
                If iDatosRegla > 0 Then
                    sDatosRegla = Mid$(rsPuestos!regla, iDatosRegla + 1)
                Else
                    If rsPuestos!Puesto >= iCalcPuesto Then
                        sDatosRegla = ">=" & rsPuestos!Puesto
                    Else
                        sDatosRegla = rsPuestos!Puesto
                    End If
                End If
                aTabla(iCDorsales, iBailes + 4 + iCPuestos) = sDatosRegla
                rsPuestos.MoveNext
                iCPuestos = iCPuestos + 1
            Wend
            rsPuestos.Close
            
            
            rsDorsales.MoveNext
            iCDorsales = iCDorsales + 1
            If Not rsDorsales.EOF And iCDorsales > iParejas Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
            End If
        Wend
        rsDorsales.Close
        rsBailes.MoveNext
        iCBailes = iCBailes + 1
    Wend
    
    If Printer.CurrentY + (iParejas + 1) * 400 + MARGEN_PAGINA > Printer.Height Then
        Printer.NewPage
    End If
    
    iCPuestos = iCPuestos + 2
    Printer.Print
    Printer.FontSize = 14
    Printer.Print mml_FRASE0693
    Printer.FontSize = 10
    Printer.Print
    'DibujarTabla Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 2 + iCPuestos, 650, 350
    DibujarTablaExt Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 2 + iCPuestos, 650, 350, 500, iBailes + 1, 0, iBailes + 1
    Printer.Print
    On Local Error Resume Next
    Printer.PaintPicture LoadPicture(C_LOGO_PATH), (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    On Local Error GoTo error
    Printer.Print
    
    ' {·M} Imprimimos el hueco para la firma
    iPosY = Printer.CurrentY
    Printer.CurrentX = 0
    iPosY = Printer.CurrentY
    Printer.FontSize = 10
    Printer.Print mml_FRASE0694
    Printer.Line (0, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.CurrentX = 1800
    Printer.CurrentY = iPosY
    Printer.Print mml_FRASE0033
    Printer.Line (1800, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.EndDoc
    
    Exit Sub

error:
    Dim Msj As String
    If Err.Number <> 0 Then
       Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
             & Err.Source & Chr(13) & Err.Description
       MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
    End If

End Sub


'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
'******************                            COPIAS                                      ******************************
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************

Public Sub COPIA_ImprimirResumen(Optional iFase As Integer = 0)
Dim iEscala As Integer
Dim rs As Recordset
Dim rsDorsal As Recordset
Dim rsElim As Recordset
Dim rsMarcas As Recordset
Dim aPosicion() As Integer
Dim iParejas As Integer
Dim iParejasPart As Integer
Dim iCPareja As Integer
Dim iUltimaFase As Integer
Dim iPosValorMax As Integer
Dim iCValorMax As Integer
Dim lValorMax As Long
Dim lValor As Long
Dim iCParejasNP As Integer
Dim iCParejas As Integer
Dim iCPuestos As Integer
Dim aTabla() As Integer
Dim iPuestoIni As Integer
Dim iRepesca As Integer
Dim iCCopias As Integer
Dim sPuesto As String
Dim sPuntos As String
Dim sDorsal As String
Dim X As Integer, Y As Integer
Dim aTablaResumen() As String
Dim aDefCelda(4) As TCelda
Dim icParejasTabla As Integer
Dim iPagina As Integer
Dim sBailes As String
Dim iCBailes As Integer
Dim sSQL As String

Const C_NUM_DORSAL = 0
Const C_ELIM_SUP = 1
Const C_CONT_MARCAS = 2
Const C_PUESTO = 3
Const C_PUESTO_FIN = 4
Const C_PUNTOS = 5
Const C_COD_PAREJA = 6
Const C_ULT_ELIM = 7
Const C_MAX_DIM = 8
Const C_NO_PRESENTADO = -1

Dim aTablasRes() As String
Const C_PUESTOS_FINAL = 0
Const C_MARCAS_ELIM = 1

    If Not C_DEBUG Then On Local Error GoTo error
    For iCCopias = 1 To CDialog.Copies
        iPagina = 1
        icParejasTabla = 1
        Printer.FontBold = False
        'Borramos la información del ResumenFinal de esta categoría
        db.Execute ("DELETE FROM resumenfinales WHERE cod_categoria = " & tbCodCateg.Text)
        db.Execute ("DELETE FROM Resultadosfinales WHERE cod_categoria = " & tbCodCateg.Text)
        
        ' Seleccionamos las parejas que tienen puntuación por ser oficiales
        ' Comprobamos el número de parejas anotadas
        Set rsDorsal = db.OpenRecordset("SELECT DISTINCT d.num_dorsal FROM dorsales d, parejas pa WHERE d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
        If Not rsDorsal.EOF Then
            rsDorsal.MoveLast
            iParejas = rsDorsal.RecordCount
            rsDorsal.MoveFirst
        Else
            MsgBox mml_FRASE0696, vbOKOnly Or vbCritical, mml_FRASE0096
            rsDorsal.Close
            Exit Sub
        End If
        ' Comprobamos el número de parejas oficiales que participaron (bailaron) (No se cuentan parejas adicionales ni especiales -> dorsal bajo)
        Set rs = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM dorsales d, parejas pa WHERE d.cod_pareja = pa.codigo AND pa.pareja_adicional = 0 AND d.num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND d.cod_categoria=" & tbCodCateg.Text, dbOpenSnapshot)
        If Not rs.EOF Then
            rs.MoveLast
            iParejasPart = rs.RecordCount
            rs.Close
        Else
            MsgBox mml_FRASE0697, vbOKOnly Or vbCritical, mml_FRASE0096
            rs.Close
            Exit Sub
        End If
        ReDim aTabla(iParejas, C_MAX_DIM)
        ReDim aTablasRes(iParejas, 2)
        iCPareja = 0
        iCPuestos = 1
        'Recorremos todas las parejas anotadas
        While Not rsDorsal.EOF
            aTabla(iCPareja, C_NUM_DORSAL) = rsDorsal!num_dorsal
            
            'localizamos el código de pareja
            Set rs = db.OpenRecordset("SELECT cod_pareja FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If Not rs.EOF Then
                aTabla(iCPareja, C_COD_PAREJA) = rs!cod_pareja
            Else
                aTabla(iCPareja, C_COD_PAREJA) = 1
            End If
            rs.Close
            ' Primero comprobamos si tenemos un puesto en la final
            Set rsElim = db.OpenRecordset("SELECT puesto FROM cal_conjunto WHERE regla = 'FIN' AND cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            If Not rsElim.EOF Then
                aTabla(iCPareja, C_CONT_MARCAS) = 10 - rsElim!Puesto
                ' Si lo tenemos guardamos las posiciones por baile
                sBailes = ""
                
                'Set rs = db.OpenRecordset("SELECT cod_baile, posiciones_mi, b.nombre FROM cal_baile, bailes b WHERE cod_baile = b.codigo AND fase = 1 AND cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " AND puesto = 0 ORDER BY cod_baile", dbOpenSnapshot)
                
                Set rs = db.OpenRecordset("SELECT cb.cod_baile, posiciones_mi, b.nombre FROM cal_baile cb, bailes b, bailes_Categ bc WHERE bc.cod_categoria = cb.cod_categoria AND bc.cod_baile = cb.cod_baile AND bc.fase = 1 AND cb.cod_baile = b.codigo AND cb.fase = 1 AND cb.cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " AND puesto = 0 ORDER BY bc.posicion", dbOpenSnapshot)
                iCBailes = 0
                While Not rs.EOF
                    Inc iCBailes
                    If sBailes <> "" Then sBailes = sBailes & ", "
                    sBailes = sBailes & rs!Nombre
                    If aTablasRes(iCPareja, C_PUESTOS_FINAL) <> "" Then
                        aTablasRes(iCPareja, C_PUESTOS_FINAL) = aTablasRes(iCPareja, C_PUESTOS_FINAL) & "-"
                    End If
                    aTablasRes(iCPareja, C_PUESTOS_FINAL) = aTablasRes(iCPareja, C_PUESTOS_FINAL) & rs!posiciones_mi
                    If G_IMP_POS_BAILES_DIPLOMAS Then
                        db.Execute "INSERT INTO resultadosfinales VALUES (" & tbCodComp.Text & "," & tbCodCateg.Text & "," & rsDorsal!num_dorsal & "," & rs.Fields("cod_baile") & ",'" & rs.Fields("posiciones_mi") & "')"
                    End If
                    rs.MoveNext
                Wend
                rs.Close
            Else
                aTabla(iCPareja, C_CONT_MARCAS) = 0
            End If
            rsElim.Close
            
            'Contamos el número de marcas
            sSQL = "SELECT COUNT(*) FROM puntuaciones WHERE puesto > 0 AND cod_categoria=" & tbCodCateg.Text & " AND fase > 1 AND num_dorsal=" & rsDorsal!num_dorsal
            Debug.Print sSQL
            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
            aTablasRes(iCPareja, C_MARCAS_ELIM) = rs.Fields(0)
            rs.Close
            
            ' Por cada uno de los dorsales comprobamos el número de eliminatorias superadas
            Set rsElim = db.OpenRecordset("SELECT DISTINCT fase FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            If Not rsElim.EOF Then
                iUltimaFase = rsElim.Fields(0)
                rsElim.MoveLast
                aTabla(iCPareja, C_ELIM_SUP) = rsElim.RecordCount - 1
                aTabla(iCPareja, C_ULT_ELIM) = iUltimaFase
            Else
                aTabla(iCPareja, C_ELIM_SUP) = C_NO_PRESENTADO
                aTabla(iCPareja, C_PUESTO) = C_NO_PRESENTADO
                iUltimaFase = -1
            End If
            rsElim.Close
            ' y el número de marcas de la última eliminatoria
            ' Primero comprobamos si en la última eliminatoria contamos con una repesca, ya que las marcas serán las de la repesca
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE repesca = 1 AND fase = " & iUltimaFase & " AND cod_categoria=" & tbCodCateg.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If rs.Fields(0) > 0 Then
                iRepesca = 1
            Else
                iRepesca = 0
            End If
            rs.Close
            Set rsMarcas = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE puesto > 0 AND repesca=" & iRepesca & "AND fase = " & iUltimaFase & " AND cod_categoria=" & tbCodCateg.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If aTabla(iCPareja, C_CONT_MARCAS) = 0 Then
                If Not rsMarcas.EOF Then
                    aTabla(iCPareja, C_CONT_MARCAS) = rsMarcas.Fields(0)
                Else
                    aTabla(iCPareja, C_CONT_MARCAS) = 0
                End If
            End If
            rsMarcas.Close
            
            rsDorsal.MoveNext
            Inc iCPareja
        Wend
        rsDorsal.Close
        
        ReDim aTablaResumen(iParejasPart, 5)
        aTablaResumen(0, 0) = mml_FRASE0654
        aTablaResumen(0, 1) = mml_FRASE0698
        aTablaResumen(0, 2) = mml_FRASE0699
        aTablaResumen(0, 3) = mml_FRASE0700
        If G_COUNTRY Then
            aTablaResumen(0, 4) = mml_FRASE1045
        Else
            aTablaResumen(0, 4) = mml_FRASE0701
        End If
        
        If iCBailes <= C_BAILES_POR_HOJA Then
            aDefCelda(0).Ancho = 1000
            aDefCelda(0).Justificado = eccentro
            aDefCelda(1).Ancho = 1000
            aDefCelda(1).Justificado = eccentro
            aDefCelda(2).Ancho = 1500
            aDefCelda(2).Justificado = eccentro
            aDefCelda(3).Ancho = 1000
            aDefCelda(3).Justificado = eccentro
            aDefCelda(4).Ancho = 6400
            aDefCelda(4).Justificado = ecizquierda
        Else
            aDefCelda(0).Ancho = 1000
            aDefCelda(0).Justificado = eccentro
            aDefCelda(1).Ancho = 1000
            aDefCelda(1).Justificado = eccentro
            aDefCelda(2).Ancho = 2300
            aDefCelda(2).Justificado = eccentro
            aDefCelda(3).Ancho = 1000
            aDefCelda(3).Justificado = eccentro
            aDefCelda(4).Ancho = 5600
            aDefCelda(4).Justificado = ecizquierda
        End If
        
        ' Ahora asignamos los puestos
        Do While iCPuestos <= iParejasPart
            lValorMax = 0
            iCValorMax = 0
            'Recorremos todas las parejas
            For iCParejas = 0 To iParejas - 1
                ' Solo con los puestos no asignados
                If aTabla(iCParejas, C_PUESTO) = 0 Then
                    lValor = aTabla(iCParejas, C_ELIM_SUP) * 1000 + aTabla(iCParejas, C_CONT_MARCAS)
                    If lValor > lValorMax Then
                        lValorMax = lValor
                        iPosValorMax = iCParejas
                        iCValorMax = 1
                    ElseIf lValor = lValorMax Then
                        Inc iCValorMax
                    End If
                End If
            Next iCParejas
            
            ' Asignamos el puesto y los puntos
            'Recorremos todas las parejas
            For iCParejas = 0 To iParejas - 1
                ' Solo con los puestos no asignados
                If aTabla(iCParejas, C_PUESTO) = 0 Then
                    lValor = aTabla(iCParejas, C_ELIM_SUP) * 1000 + aTabla(iCParejas, C_CONT_MARCAS)
                    If lValor = lValorMax Then
                        aTabla(iCParejas, C_PUESTO) = iCPuestos
                        aTabla(iCParejas, C_PUESTO_FIN) = iCPuestos + iCValorMax - 1
                        'aTabla(iCParejas, C_PUNTOS) = Round((1000 / (iParejasPart - 1)) * (iParejasPart - iCPuestos - (iCValorMax - 1) / 2))
                        'Puntuación entre los participantes
                        'aTabla(iCParejas, C_PUNTOS) = Round((iParejasPart - iCPuestos) / (iParejasPart - 1) / iCValorMax * 1000)
                        'Puntuación entre los inscritos al cierre
                        ' Puntos
                        If G_MOSTRAR_PUNTOS Then
                            If G_COUNTRY Then
                            Dim iPuntos As Integer
                                If iParejas = 1 Then
                                    iPuntos = 200
                                Else
                                    iPuntos = Redondea((iParejas - (iCPuestos + (iCValorMax - 1) / 2)) / (iParejas - 1) * 1000, 0)
                                    Select Case iParejas
                                    Case 1
                                        iPuntos = 200
                                    Case 2
                                        iPuntos = iPuntos / 3
                                    Case 3
                                        iPuntos = iPuntos / 2
                                    Case 4
                                        iPuntos = Redondea(iPuntos / 1.35, 0)
                                    Case 5
                                        iPuntos = Redondea(iPuntos / 1.2, 0)
                                    End Select
                                    If iPuntos = 0 Then iPuntos = 10
                                End If
                                aTabla(iCParejas, C_PUNTOS) = iPuntos
                            Else
                                aTabla(iCParejas, C_PUNTOS) = Redondea((iParejas - (iCPuestos + (iCValorMax - 1) / 2)) / (iParejas - 1) * 1000, 0)
                            End If
                        End If
                        sDorsal = aTabla(iCParejas, C_NUM_DORSAL)
                        
                        aTablaResumen(icParejasTabla, 0) = sDorsal
                        
                        sPuesto = Trim$(aTabla(iCParejas, C_PUESTO))
                        If aTabla(iCParejas, C_PUESTO) <> aTabla(iCParejas, C_PUESTO_FIN) Then
                            sPuesto = sPuesto & "-" & aTabla(iCParejas, C_PUESTO_FIN)
                        End If
                        
                        aTablaResumen(icParejasTabla, 1) = sPuesto
                        
                        If aTablasRes(iCParejas, C_PUESTOS_FINAL) <> "" Then
                            aTablaResumen(icParejasTabla, 2) = aTablasRes(iCParejas, C_PUESTOS_FINAL)
                        Else
                            aTablaResumen(icParejasTabla, 2) = aTablasRes(iCParejas, C_MARCAS_ELIM) & " (" & sDescCortaFase(Val(aTabla(iCParejas, C_ULT_ELIM))) & ")"
                        End If
                        
                        sPuntos = aTabla(iCParejas, C_PUNTOS)
                        If G_MOSTRAR_PUNTOS Then
                            aTablaResumen(icParejasTabla, 3) = sPuntos
                        Else
                            aTablaResumen(icParejasTabla, 3) = ""
                        End If
                        Set rs = db.OpenRecordset("SELECT nombre_hombre, nombre_mujer FROM parejas WHERE codigo =" & aTabla(iCParejas, C_COD_PAREJA), dbOpenSnapshot)
                        If Not rs.EOF Then
                            aTablaResumen(icParejasTabla, 4) = rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer
                        End If
                        
                        If icParejasTabla >= G_MAX_FILAS_POR_PAG Then
                            ImprimirCabeceraResumen
                            Printer.FontSize = 9
                            Printer.Print mml_FRASE0702 & iParejasPart & mml_FRASE0472 & iParejas & mml_FRASE0703 & iPagina
                            SaltoLinea Printer, 4
                            DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icParejasTabla, 5, aTablaResumen(), aDefCelda(), C_ALTO_CELDA_TABLAA
                            Printer.NewPage
                            iPagina = iPagina + 1
                            icParejasTabla = 1
                        Else
                            icParejasTabla = icParejasTabla + 1
                        End If
                        
                        rs.Close
                        db.Execute ("INSERT INTO resumenfinales VALUES (" & tbCodCateg.Text & ",'" & sDorsal & "','" & sPuesto & "','" & sPuntos & "')")
                    End If
                End If
            Next iCParejas
            
            If iCValorMax = 0 Then
                Exit Do
            End If
            iCPuestos = iCPuestos + iCValorMax
        Loop
        For iCParejasNP = 0 To iParejas - 1
            ' Solo con los N.P.
            If aTabla(iCParejasNP, C_PUESTO) = C_NO_PRESENTADO Then
                sDorsal = aTabla(iCParejasNP, C_NUM_DORSAL)
                aTablaResumen(icParejasTabla, 0) = sDorsal
                aTablaResumen(icParejasTabla, 1) = mml_FRASE0704
                aTablaResumen(icParejasTabla, 2) = "0"
                aTablaResumen(icParejasTabla, 3) = C_PUNTOS_NO_PRESENTADO
                Set rs = db.OpenRecordset("SELECT nombre_hombre, nombre_mujer FROM parejas WHERE codigo =" & aTabla(iCParejasNP, C_COD_PAREJA), dbOpenSnapshot)
                If Not rs.EOF Then
                    aTablaResumen(icParejasTabla, 4) = rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer
                End If
                If icParejasTabla >= G_MAX_FILAS_POR_PAG Then
                    ImprimirCabeceraResumen
                    Printer.FontSize = 9
                    Printer.Print mml_FRASE0702 & iParejasPart & mml_FRASE0472 & iParejas & mml_FRASE0703 & iPagina
                    SaltoLinea Printer, 4
                    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icParejasTabla, 5, aTablaResumen(), aDefCelda(), C_ALTO_CELDA_TABLAA
                    Printer.NewPage
                    iPagina = iPagina + 1
                    icParejasTabla = 1
                Else
                    icParejasTabla = icParejasTabla + 1
                End If
                rs.Close
                db.Execute ("INSERT INTO resumenfinales VALUES (" & tbCodCateg.Text & ",'" & sDorsal & "','NO PRESENTADO','POR DETERMINAR')")
            End If
        Next iCParejasNP
        
        'Comprobamos si estamos emitiendo el listado de posiciones de los dorsales que no posaron
        'En este caso eliminamos a los primeros que pasaron de fase
        Dim iDorsalesSigFase As Integer
        iDorsalesSigFase = 0
        If iFase > 0 Then
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE fase > " & iFase, dbOpenSnapshot)
            iDorsalesSigFase = rs.Fields(0)
            rs.Close
        End If
        
    ImprimirCabeceraResumen iDorsalesSigFase
    
    Printer.FontSize = 9
    Printer.Print mml_FRASE0702 & iParejasPart & mml_FRASE0472 & iParejas & mml_FRASE0703 & iPagina
    SaltoLinea Printer, 4
    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icParejasTabla, 5, aTablaResumen(), aDefCelda(), C_ALTO_CELDA_TABLAA
    Printer.Print
    Printer.FontBold = True
    Printer.Print mml_FRASE0706;
    Printer.FontBold = False
    Printer.Print sBailes
    If frmImprimirInternet.gl_bGenerarInet Then
        Printer.EndDoc
    Else
        Printer.NewPage
    End If
    
    Next iCCopias
    Exit Sub
error:
    ProcesarError "ImprimirResumen"
End Sub

