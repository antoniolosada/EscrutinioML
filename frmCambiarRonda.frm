VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCambiarRonda 
   Caption         =   "mml_FRASE1098"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCombinar 
      Caption         =   "mml_FRASE1042"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5730
      TabIndex        =   25
      Top             =   9210
      Width           =   2085
   End
   Begin VB.CommandButton cmdImpTandas 
      Caption         =   "mml_FRASE0600"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3570
      TabIndex        =   24
      Top             =   9210
      Width           =   2085
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE1098"
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   10845
      Begin VB.TextBox tbCombinarTandas 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   22
         Top             =   1740
         Width           =   375
      End
      Begin VB.CommandButton cmdACtualizar 
         Caption         =   "mml_FASE0295"
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
         Left            =   7440
         TabIndex        =   20
         Top             =   1740
         Width           =   2085
      End
      Begin VB.ComboBox cbBaile 
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
         Height          =   390
         Left            =   4110
         TabIndex        =   18
         Top             =   1740
         Width           =   3255
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
         Left            =   3540
         TabIndex        =   14
         Top             =   1290
         Width           =   5175
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
         Left            =   2700
         TabIndex        =   13
         Top             =   1290
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
         Left            =   3540
         TabIndex        =   12
         Top             =   810
         Width           =   6615
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
         Left            =   2700
         TabIndex        =   11
         Top             =   810
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
         Left            =   3540
         TabIndex        =   10
         Top             =   330
         Width           =   6615
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
         Left            =   2700
         TabIndex        =   9
         Top             =   330
         Width           =   855
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
         Left            =   8820
         TabIndex        =   8
         Top             =   1290
         Width           =   1335
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
         Left            =   2265
         Picture         =   "frmCambiarRonda.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   330
         Width           =   450
      End
      Begin VB.CommandButton CommandButton1 
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
         Left            =   2265
         Picture         =   "frmCambiarRonda.frx":046A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   825
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
         Left            =   2265
         Picture         =   "frmCambiarRonda.frx":08D4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1275
         Width           =   450
      End
      Begin VB.Label lblOrden 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   10140
         TabIndex        =   26
         Top             =   2190
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "mml_FRASE0600"
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
         Left            =   1140
         TabIndex        =   23
         Top             =   1770
         Width           =   1245
      End
      Begin VB.Label lblCambioDorsal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   180
         TabIndex        =   21
         Top             =   2190
         Width           =   9945
      End
      Begin VB.Label lblBaile 
         Caption         =   "mml_FRASE0436"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3090
         TabIndex        =   19
         Top             =   1830
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
         Left            =   540
         TabIndex        =   17
         Top             =   1290
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
         Left            =   540
         TabIndex        =   16
         Top             =   810
         Width           =   1575
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
         Left            =   540
         TabIndex        =   15
         Top             =   330
         Width           =   1575
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
      Height          =   525
      Left            =   8310
      TabIndex        =   3
      Top             =   9210
      Width           =   2385
   End
   Begin VB.CommandButton cmdIntercambiarRondas 
      Caption         =   "mml_FRASE1098"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   210
      TabIndex        =   2
      Top             =   9210
      Width           =   3285
   End
   Begin VB.Frame frmCambiarRonda 
      Caption         =   "mml_FRASE1098"
      Height          =   6435
      Left            =   0
      TabIndex        =   0
      Top             =   2700
      Width           =   10845
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridRondas 
         Height          =   6075
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Visible         =   0   'False
         Width           =   10665
         _ExtentX        =   18812
         _ExtentY        =   10716
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmCambiarRonda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub RecargarBailes()
Dim rs As Recordset
   
    If Val(tbCodComp.Text) = 0 Or Val(tbCodCat.Text) = 0 Or Val(tbCodFase.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
   
    cbBaile.Clear
    'Recuperamos los bailes
    Set rs = db.OpenRecordset("SELECT b.codigo, b.nombre FROM bailes b, bailes_categ bc WHERE b.codigo = bc.cod_baile AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase = " & IIf(Val(tbCodFase.Text) = 1, 1, 2), dbOpenSnapshot)
    While Not rs.EOF
        cbBaile.AddItem rs.Fields("codigo") & " - " & rs.Fields("nombre")
        rs.MoveNext
    Wend
    rs.Close
    Me.Refresh

End Sub

Sub CargarDorsales()
Dim rs As Recordset

    If Not C_DEBUG Then On Local Error GoTo error

    If Val(tbCodComp.Text) = 0 Or Val(tbCodCat.Text) = 0 Or Val(tbCodFase.Text) = 0 Or Val(cbBaile.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
    
    lblCambioDorsal.Caption = ""
    lblOrden.Caption = ""
    
    gridRondas.Clear
    gridRondas.Rows = 0
    gridRondas.Cols = 6
    gridRondas.ColWidth(0) = 1000
    gridRondas.ColWidth(1) = 1000
    gridRondas.ColWidth(2) = 1000
    gridRondas.ColWidth(3) = 1000
    gridRondas.ColWidth(4) = 3000
    gridRondas.ColWidth(5) = 3000
    gridRondas.AddItem "Cod." & vbTab & mml_FRASE0617 & vbTab & mml_FRASE0300 & vbTab & mml_FRASE0436 & vbTab & mml_FRASE0030 & vbTab & mml_FRASE0030
    
    Set rs = db.OpenRecordset("SELECT dc.num_dorsal, orden, dc.codigo, b.nombre, p.nombre_hombre, p.nombre_mujer FROM dorsales d, parejas p, dorsalescombinados dc, bailes b WHERE p.codigo = d.cod_pareja AND d.num_dorsal = dc.num_dorsal AND d.cod_categoria = dc.cod_categoria AND b.codigo = dc.cod_baile AND d.fase = dc.fase AND d.repesca = dc.repesca AND dc.cod_categoria = " & tbCodCat.Text & " AND dc.fase = " & tbCodFase.Text & " AND dc.repesca = " & chkRep.Value & " AND dc.cod_baile = " & Val(cbBaile.Text) & " ORDER BY dc.cod_baile, dc.orden, dc.num_dorsal", dbOpenSnapshot)
    If Not rs.EOF Then
        gridRondas.Visible = True
    Else
        gridRondas.Visible = False
    End If
    
    While Not rs.EOF
        gridRondas.AddItem rs.Fields("codigo") & vbTab & rs.Fields("orden") & vbTab & rs.Fields("num_dorsal") & vbTab & rs.Fields("nombre") & vbTab & rs.Fields("nombre_hombre") & vbTab & rs.Fields("nombre_mujer")
        rs.MoveNext
    Wend
    rs.Close
    
    If gridRondas.Rows > 1 Then
        gridRondas.FixedRows = 1
    End If
    
    Exit Sub
error:
    ProcesarError "CargarDorsales"
End Sub

Private Sub cbBaile_Click()
    CargarDorsales
End Sub

Private Sub cbBaile_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub chkRep_Click()
    RecargarBailes
End Sub

Private Sub cmdActualizar_Click()
    CargarDorsales
End Sub

Private Sub cmdCombinar_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    
    If tbCodCat.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    If Val(tbCodFase.Text) < 2 Then
        MsgBox mml_FRASE0651, vbOKOnly Or vbExclamation, mml_FRASE0096
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE1043, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        'Genera la información para la recombinación de dorsales
        CombinarDorsales tbCodCat.Text, tbCodFase.Text, chkRep.Value, Val(tbCombinarTandas.Text), True
        If MsgBox(mml_FRASE1044, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
            cmdImpTandas_Click
        End If
    End If
    Exit Sub
error:
    ProcesarError "cmdCombinar_Click"
End Sub

Private Sub cmdImpTandas_Click()
    If MsgBox(mml_FRASE0670, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        With frmImpHojasPuntuaciones
            .tbCodComp.Text = tbCodComp.Text
            .tbDescComp.Text = tbDescComp.Text
            .tbCodCat.Text = tbCodCat.Text
            .tbDescCat.Text = tbDescCat.Text
            .tbCodFase.Text = tbCodFase.Text
            .tbDescFase.Text = tbDescFase.Text
            .chkRep.Value = chkRep.Value
            
            .cmdTandas_Click
        End With
    End If
End Sub

Private Sub cmdIntercambiarRondas_Click()
Dim lCodigo As Long
Dim sCambio As String

    If Not C_DEBUG Then On Local Error GoTo error
    
    If Val(tbCodCat.Text) = 0 Or Val(tbCodFase.Text) = 0 Or Not gridRondas.Visible Then
        CamposSinCubrir
        Exit Sub
    End If
    
    With gridRondas
        If lblCambioDorsal.Caption = "" Then
            If .RowSel >= 1 Then
                If lblCambioDorsal.Caption = "" Then
                    lblCambioDorsal.Caption = .TextMatrix(.RowSel, 0) & " - " & .TextMatrix(.RowSel, 1) & ", (" & .TextMatrix(.RowSel, 2) & ") " & .TextMatrix(.RowSel, 3) & ", " & .TextMatrix(.RowSel, 4)
                    lblOrden.Caption = .TextMatrix(.RowSel, 1)
                End If
                MsgBox mml_FRASE1099, vbOKOnly Or vbInformation, G_MSG_AVISO
            End If
        Else
            If .RowSel >= 1 Then
                sCambio = .TextMatrix(.RowSel, 0) & " - " & .TextMatrix(.RowSel, 1) & ", (" & .TextMatrix(.RowSel, 2) & ") " & .TextMatrix(.RowSel, 3) & ", " & .TextMatrix(.RowSel, 4)
                If MsgBox(mml_FRASE1107 & ": " & vbCrLf & lblCambioDorsal.Caption & vbCrLf & sCambio, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
                    lCodigo = Val(lblCambioDorsal.Caption)
                    db.Execute "UPDATE dorsalescombinados SET orden = " & lblOrden.Caption & " WHERE codigo = " & .TextMatrix(.RowSel, 0)
                    db.Execute "UPDATE dorsalescombinados SET orden = " & .TextMatrix(.RowSel, 1) & " WHERE codigo = " & lCodigo
                    
                    cmdActualizar_Click
                End If
            End If
        End If
    End With
    Exit Sub
error:
    ProcesarError "cmdIntercambiarRondas_Click"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelComp_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCat.Text = ""
    tbDescCat.Text = ""
    tbCodFase.Text = ""
    tbDescFase.Text = ""
End Sub

Private Sub CommandButton1_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCat.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCat.Text = sResultado(2)
    tbCodFase.Text = ""
    tbDescFase.Text = ""


End Sub

Private Sub CommandButton2_Click()
Dim rs As Recordset
    If tbCodCat.Text = "" Then
        MsgBox mml_FRASE0504, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & tbCodCat.Text)
    If tbCodFase.Text <> "" Then
        tbDescFase.Text = sDescFase(Val(tbCodFase.Text))
    End If
    
    CubrirInfoTandas
    RecargarBailes
End Sub

Sub CubrirInfoTandas()
Dim iMaxTandas As Integer
Dim iTandasConMasDorsales As Integer
Dim iTotalDorsales As Integer

    iMaxTandas = 0
    iDorsalesTanda = CalcularDorsalesPorTandaCatExt(Val(tbCodCat.Text), Val(tbCodFase.Text), chkRep.Value, 1, iMaxTandas, iTandasConMasDorsales, iTotalDorsales)
    tbCombinarTandas.Text = iMaxTandas

End Sub

Private Sub Form_Load()
    TraducirCadenas Me
End Sub
