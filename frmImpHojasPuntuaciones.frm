VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImpHojasPuntuaciones 
   Caption         =   "mml_FRASE0596"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0041"
      Height          =   4350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.CommandButton cmdImpRondas 
         Caption         =   "mml_FRASE1098"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7890
         TabIndex        =   43
         Top             =   3210
         Width           =   1905
      End
      Begin VB.CommandButton cmdCombinar 
         Caption         =   "mml_FRASE1042"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7890
         TabIndex        =   42
         Top             =   3990
         Width           =   1125
      End
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
         Left            =   1440
         TabIndex        =   41
         Top             =   2250
         Width           =   375
      End
      Begin VB.CheckBox chkRecombinar 
         Caption         =   "Recombinar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6420
         TabIndex        =   40
         Top             =   3510
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "mml_FRASE0065"
         Default         =   -1  'True
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
         Left            =   7875
         TabIndex        =   38
         Top             =   2730
         Width           =   1920
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
         Height          =   690
         Left            =   9030
         TabIndex        =   39
         Top             =   3570
         Width           =   750
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
         Picture         =   "frmImpHojasPuntuaciones.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1305
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
         Left            =   1845
         Picture         =   "frmImpHojasPuntuaciones.frx":046A
         Style           =   1  'Graphical
         TabIndex        =   36
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
         Picture         =   "frmImpHojasPuntuaciones.frx":08D4
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         Width           =   450
      End
      Begin VB.Frame Frame3 
         Caption         =   "mml_FRASE0437"
         Height          =   735
         Left            =   6420
         TabIndex        =   31
         Top             =   2760
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
            ItemData        =   "frmImpHojasPuntuaciones.frx":0D3E
            Left            =   150
            List            =   "frmImpHojasPuntuaciones.frx":0D60
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.CheckBox chkImprimirCatVacias 
         Caption         =   "mml_FRASE0597"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         TabIndex        =   30
         Top             =   3870
         Width           =   2685
      End
      Begin VB.CheckBox chkHojasHorarioPorJuez 
         Caption         =   "mml_FRASE0598"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         TabIndex        =   29
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CheckBox chkImpIniCat 
         Caption         =   "mml_FRASE0599"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         TabIndex        =   28
         Top             =   3150
         Width           =   2415
      End
      Begin VB.CommandButton cmdTandas 
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
         Height          =   435
         Left            =   7890
         TabIndex        =   27
         Top             =   3570
         Width           =   1125
      End
      Begin VB.CheckBox chkInfoTandas 
         Caption         =   "mml_FRASE0601"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8760
         TabIndex        =   26
         Top             =   1770
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkSigFases 
         Caption         =   "mml_FRASE0602"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3825
         TabIndex        =   25
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CheckBox chkFinDoc 
         Caption         =   "mml_FRASE0013"
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
         Left            =   3825
         TabIndex        =   24
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2475
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
         Left            =   8400
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   9240
         Top             =   2220
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Caption         =   "mml_FRASE0604"
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   3615
         Begin VB.ComboBox cbJuez 
            BackColor       =   &H00FFFFC0&
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
            ItemData        =   "frmImpHojasPuntuaciones.frx":0D8A
            Left            =   780
            List            =   "frmImpHojasPuntuaciones.frx":0DDC
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox tbTanda 
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
            Left            =   2640
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "mml_FRASE0421"
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
            Left            =   60
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "mml_FRASE0605"
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
            Left            =   1800
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox tbTotalDorsales 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   8400
         TabIndex        =   16
         Top             =   2250
         Width           =   735
      End
      Begin VB.TextBox tbDorsalesTanda 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   5880
         TabIndex        =   14
         Top             =   2250
         Width           =   705
      End
      Begin VB.TextBox tbTandas 
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
         Left            =   1800
         TabIndex        =   12
         Top             =   2250
         Width           =   435
      End
      Begin VB.TextBox tbComen 
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
         Left            =   1800
         MaxLength       =   55
         TabIndex        =   10
         Top             =   1800
         Width           =   7005
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   360
         Width           =   6615
      End
      Begin VB.TextBox tbCodCat 
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Line Line1 
         X1              =   3810
         X2              =   6060
         Y1              =   3510
         Y2              =   3510
      End
      Begin VB.Label lblTandasMenosDorsales 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   3105
         TabIndex        =   34
         Top             =   2250
         Width           =   810
      End
      Begin VB.Label lblTandasMasDorsales 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
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
         Left            =   2220
         TabIndex        =   33
         Top             =   2250
         Width           =   810
      End
      Begin VB.Label Label6 
         Caption         =   "mml_FRASE0012"
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
         Left            =   6600
         TabIndex        =   17
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "mml_FRASE0256"
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
         Left            =   3930
         TabIndex        =   15
         Top             =   2280
         Width           =   1935
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
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0607"
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
         Top             =   1800
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
         Left            =   120
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmImpHojasPuntuaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const C_CORRECCION_TABLA_POR_X = 3000
Const C_MIN_X = 50

Dim iBaileActivo As Integer

Private Type sCategHorario
    sDescCateg As String
    iCodCateg As Integer
    iFase As Integer
    iRepesca As Integer
End Type
Dim aiOrdenCateg(20) As sCategHorario
Dim iNumCatGrupo As Integer


Public Sub ImprimirHojas(sCodComp As String, sCodCat As String, sCodFase As String, sDescComp As String, sDescCat As String)
    tbCodComp.Text = sCodComp
    tbCodCat.Text = sCodCat
    tbCodFase.Text = sCodFase
    tbDescComp.Text = sDescComp
    tbDescCat.Text = sDescCat
    tbDescComp.Text = Buscar("competiciones", "descripcion", tbCodComp.Text)
    tbDescCat.Text = Buscar("categorias", "descripcion", tbCodCat.Text)
    tbDescFase.Text = LiteralFase(sCodFase)
    chkInfoTandas.Value = 1
    CubrirInfoTandas
    Me.Show 1
End Sub
Public Sub ImprimirHojasCompleta(sCodComp As String, sCodCat As String, sCodFase As String, sDescComp As String, sDescCat As String)
    tbCodComp.Text = sCodComp
    tbCodCat.Text = sCodCat
    tbCodFase.Text = sCodFase
    tbDescComp.Text = sDescComp
    tbDescCat.Text = sDescCat
    tbDescComp.Text = Buscar("competiciones", "descripcion", tbCodComp.Text)
    tbDescCat.Text = Buscar("categorias", "descripcion", tbCodCat.Text)
    tbDescFase.Text = LiteralFase(sCodFase)
    chkInfoTandas.Value = 1
    CubrirInfoTandas
    
    ImpresionDirecta
End Sub

Public Sub ImprimirHojaPuntuaciones(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer, Optional iCodBaile = 0)
Dim lOrden As Long
    lOrden = iBuscarOrden(sCodCat, sCodFase, iRepesca)
    
    If G_HOJA_EXTENDIDA Or bSelecAutomatica(sCodCat, sCodFase, iMaxTandas, iMaxDorsales) Then
        ImprimirHojaPuntuacionesExt sCodComp, sCodCat, sCodFase, sIdJuez, iTanda, iNumDorsal, iMaxDorsalesTanda, iMaxTandas, iMaxDorsales, iRepesca, iHoja, iCodBaile, lOrden
    Else
        ImprimirHojaPuntuacionesNormal sCodComp, sCodCat, sCodFase, sIdJuez, iTanda, iNumDorsal, iMaxDorsalesTanda, iMaxTandas, iMaxDorsales, iRepesca, iHoja, iCodBaile, lOrden
    End If
End Sub
Function bSelecAutomatica(iCodCat As Integer, iFase As Integer, iMaxTandas As Integer, iMaxDorsales As Integer) As Boolean
Dim rs As Recordset, iNumJueces As Integer
    bSelecAutomatica = False
    If G_SELEC_HOJA_EXT_AUTO Then
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE cod_categoria = " & iCodCat, dbOpenSnapshot)
        iNumJueces = rs.Fields(0)
        rs.Close
        'Si los posibles dorsales o juces no caben
        If (((iMaxDorsales \ iMaxTandas) + IIf(iMaxDorsales Mod iMaxTandas > 0, 1, 0) > C_MAX_DORSALES_HOJA_OPTICA) And iFase > 2) Or _
            iNumJueces > C_MAX_JUECES_HOJA_OPTICA Then
            bSelecAutomatica = True
        Else
            bSelecAutomatica = False
        End If
    End If
End Function
Public Sub ImprimirHojaPuntuacionesExt(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer, Optional iCodBaile = 0, Optional lOrden As Long = -1)
Dim rs As Recordset
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim X As Integer, Y As Integer
Dim i As Integer, j As Integer, k As Integer
Dim iAncho As Integer, iAlto As Integer
Dim iCCateg As Integer
Dim iCBailes As Integer
Dim iCJueces As Integer
Dim iCDorsales As Integer
Dim bJuezPasos As Boolean
Dim sInfo As String
Dim iMaxDorsalesPorTanda As Integer
Dim iCMarcas As Integer
Dim iTamFuente As Integer
Const MAX_COL = 28
Const MAX_FILA = 80
Const INIC_CAT = 5
Const INIC_FASE = 12
Const INIC_TANDA = 20
Const INIC_BAILES = 24
Const INIC_JUEZ = 16
Const C_POS_X_MARCA_BAILE = 25
Dim iAnchoCelda As Integer
Dim bTeamMatch As Integer

Dim ALTO_BAILE As Integer
Dim ALTO_BAILE_FINAL As Integer

Dim borradorX As Integer, borradorY As Integer

    bTeamMatch = ComprobarSiTeamMatch(sCodCat)
    bJuezPasos = False
    Printer.CurrentY = MARGEN_SUPERIOR
    Y = MARGEN_SUPERIOR
    Printer.FillColor = &HFFFFFF
    Printer.FillStyle = 0
    iAncho = VarCfg("ancho_marca")
    iAlto = VarCfg("alto_marca")
    iAnchoCelda = (iAncho + C_ANCHO_ESPACIO)
    
    ALTO_BAILE = 2 * C_ANCHO_MARCAS_BAILE * iAlto
    ALTO_BAILE_FINAL = 16 * iAlto
    
    'Imprimir marca de control de calidad de escaneo
    Y = MARGEN_SUPERIOR + (2 * iAlto) * 6
    X = C_POS_CONTROL_CALIDAD_EXT * iAnchoCelda
    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    Printer.Line (X, Y + iAlto / 2)-(X + iAncho, Y + iAlto), , BF
    
    Y = MARGEN_SUPERIOR
    X = 0
    'Imprimir marcas
    iCMarcas = 1
    While iCMarcas <= C_MAX_MARCAS_X_EXT - 1
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        X = X + iAnchoCelda
        Inc iCMarcas
    Wend
    iCMarcas = 1
    While iCMarcas <= C_MAX_MARCAS_Y
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Y = Y + 2 * iAlto
        Inc iCMarcas
    Wend
    'Imprimir marca de control inferior
    Printer.Line (0, Y - 2 * iAlto)-(iAncho, Y - 2 * iAlto + iAlto), , BF
    
    If sCodFase > 1 And chkRep.Value = 0 Then
        sInfo = mml_FRASE0608 & (sCodFase / 2) * 6
        If iMaxTandas > 1 Then
            sInfo = sInfo & ") - ("
            If lblTandasMenosDorsales.Caption = "" Then
                sInfo = sInfo & mml_FRASE0609 & tbDorsalesTanda.Text
            Else
                sInfo = sInfo & "Primeras rondas " & lblTandasMasDorsales.Caption & " dorsales, Siguientes " & lblTandasMenosDorsales.Caption & " dorsales"
            End If
            Dim iMediaPorRonda As Integer
            iMediaPorRonda = ((sCodFase / 2) * 6) \ iMaxTandas
            sInfo = sInfo & " ,Media/Ronda " & iMediaPorRonda
            If ((sCodFase / 2) * 6) Mod iMaxTandas > 0 Then
                sInfo = sInfo & " faltan " & iMaxDorsales - iMaxTandas * iMediaPorRonda & " marcas"
            End If
        End If
        Printer.FontBold = False
        sInfo = sInfo & ") "
    End If
    
    Printer.FontName = "Arial"
    Printer.CurrentY = MARGEN_SUPERIOR + 2 * iAlto - iAlto / 2
    Printer.CurrentX = 0
    Printer.FontSize = 14
    Printer.Print tbDescCat.Text;
    Printer.FontSize = 12
    Printer.Print " - " & sDescFase(Val(sCodFase)) & " - " & sIdJuez;
    If iMaxTandas > 1 Then
        Printer.Print mml_FRASE0613 & iTanda & mml_FRASE0472 & iMaxTandas;
    End If
    Printer.Print "  - (" & tbDescComp.Text & ")";
    Printer.FontSize = 10
    If lOrden > 0 Then Printer.Print " (" & lOrden & ")";
    Printer.Print
    If sIdJuez <> "Control" Then
        If chkInfoTandas.Value = 1 Then
            Printer.Print sInfo & "    " & tbComen.Text
        Else
            Printer.Print tbComen.Text
        End If
    End If
    
    ' {·M} Imprimimos el hueco para la firma
    Printer.CurrentX = iAnchoCelda * C_POS_FIRMA_EXT
    Printer.CurrentY = MARGEN_SUPERIOR + iAlto * 12
    Printer.Print mml_FRASE0614
    Printer.Line (iAnchoCelda * C_POS_CUADRO_FIRMA_EXT, MARGEN_SUPERIOR + iAlto * 14)-Step((iAncho + C_ANCHO_ESPACIO) * 4.5, iAlto * 6), 0, B
    
    ' Imprimimos las categorías
    Set rs = db.OpenRecordset("SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_CAT * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y + iAlto
    Printer.Print mml_FRASE0615
    X = 2 * iAnchoCelda
    iCCateg = 0
    iTamFuente = Printer.FontSize
    Printer.FontSize = 7
    While Not rs.EOF
        If iCCateg <> 0 And iCCateg Mod C_MAX_CATEG_POR_LINEA_EXT = 0 Then
            X = 2 * iAnchoCelda
        End If
        
        Y = MARGEN_SUPERIOR + INIC_CAT * iAlto + 2 * iAlto * Int(iCCateg / C_MAX_CATEG_POR_LINEA_EXT)
        Printer.CurrentY = Y
        
        Printer.CurrentX = X
        Printer.Print rs!codigo
        Y = Y + iAlto
        If rs!codigo = sCodCat Then
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Else
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
        End If
        X = X + iAnchoCelda
        rs.MoveNext
        Inc iCCateg
    Wend
    rs.Close
    Printer.FontSize = iTamFuente
    
    
    ' Imprimimos la fase
    X = 0
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0299
    
    Printer.FontSize = 7
    Printer.FontBold = True
    
    X = 2 * iAnchoCelda
    i = 1
    While i <= 256
        Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
        Printer.CurrentX = X
        ' Si ocupamos todo el cuadro de categorias
        If iCCateg > C_MAX_CATEG_POR_LINEA_EXT * 3 Then
            Printer.CurrentY = Y + iAlto / 1.8
        Else
            Printer.CurrentY = Y + iAlto / 1.8
        End If
        Printer.Print sDescFase(i)
        Y = Y + 2 * iAlto
        If i = sCodFase Then
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Else
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
        End If
        X = X + iAnchoCelda
        i = i * 2
    Wend
    Printer.FontSize = iTamFuente
    Printer.FontBold = False
    
    'Imprimir marca de indicación de problemas
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    X = C_POS_FALLO_EXT * iAnchoCelda
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print "Err"
    Y = Y + 2 * iAlto
    Printer.ForeColor = Val(G_COLOR_MARCAS)
    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    Printer.ForeColor = C_COLOR_NEGRO
    
    ' Imprimimos la marca de identificación de hoja extendida
    X = C_POS_HOJA_EXT_EXT * (iAncho + C_ANCHO_ESPACIO)
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print "Ext"
    Y = Y + 2 * iAlto
    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
    'Imprimimos los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos FROM juez_categ WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_JUEZ_EXT * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y + iAlto
    Printer.Print mml_FRASE0421
    Printer.FontSize = 7
    Printer.FontBold = True
    X = 2 * iAnchoCelda
    iCJueces = 1
    While Not rs.EOF
        Y = MARGEN_SUPERIOR + INIC_JUEZ_EXT * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y + iAlto
        If rs!pasos = 1 Then
            Printer.Print rs!id_juez & "p"
            If rs!id_juez = sIdJuez Then
                bJuezPasos = True
            End If
        Else
            Printer.Print rs!id_juez
        End If
        Y = Y + 2 * iAlto
        If rs!id_juez = sIdJuez Then
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Else
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
        End If
        If iCJueces = C_MAX_JUECES_EXT / 2 Then
            X = 2 * iAnchoCelda
            Y = Y + iAlto
        Else
            X = X + iAnchoCelda
        End If
        Inc iCJueces
        rs.MoveNext
    Wend
    rs.Close
    Printer.FontBold = False
    Printer.FontSize = iTamFuente
    
    If bJuezPasos Then
        ALTO_BAILE = 2 * C_ANCHO_MARCAS_BAILE_JUEZ_PASOS * iAlto
    End If
    
    ' Imprimimos la marca de control
    X = C_POS_CONTROL_EXT / 2 * (iAncho + C_ANCHO_ESPACIO)
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print "Ctrl"
    Y = Y + 2 * iAlto
    If sIdJuez = "Control" Then
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
    Else
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    End If
    
    ' Imprimimos la marca de repesca
    X = C_POS_REPESCA_EXT / 2 * (iAncho + C_ANCHO_ESPACIO)
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0418
    Y = Y + 2 * iAlto
    If iRepesca = 1 Then
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
    Else
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    End If
    
    ' Imprimimos la marca de hoja2
    X = C_POS_HOJA2_EXT / 2 * iAnchoCelda
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0616
    Y = Y + 2 * iAlto
    If iHoja = 2 Then
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
    Else
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    End If
    
    ' Imprimimos la tanda
    X = 0
    Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0617
    X = 2 * iAnchoCelda
    i = 1
    While i <= C_MAX_TANDAS_EXT
        Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Print Trim$(Str$(i))
        Y = Y + 2 * iAlto
        If i = iTanda Or i = iMaxTandas Then
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Else
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
        End If
        X = X + iAnchoCelda
        i = i + 1
    Wend
    
    If sIdJuez = "Control" Then
        X = 0
        Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Line (X, Y)-(X + Printer.Width / 1.2, Y)
        Printer.CurrentX = X
        Printer.CurrentY = Y + iAlto + iAlto / 2
        Printer.Print mml_FRASE0618
        X = 0
        Y = Y + 3 * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y + iAlto / 2
        Printer.Print mml_FRASE0619
        X = 0
        Y = Y + 2 * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y + iAlto / 2
        Printer.Print mml_FRASE0620
        
        Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        X = 2 * iAnchoCelda
        iCDorsales = 1
        While Not rsDorsales.EOF Or (G_IMPRIMIR_TODOS_LOS_CUADROS And (iCDorsales - iNumDorsal) < C_MAX_DORSALES_HOJA_OPTICA)
            If iCDorsales >= iNumDorsal And (iCDorsales < iNumDorsal + iMaxDorsalesTanda Or G_IMPRIMIR_TODOS_LOS_CUADROS) Then
                Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                Printer.CurrentX = X
                Printer.CurrentY = Y
                Printer.FontBold = True
                If (iCDorsales < iNumDorsal + iMaxDorsalesTanda) Then
                    Printer.Print Trim$(rsDorsales!num_dorsal)
                Else
                    Printer.Print "___"
                End If
                Printer.FontBold = False
                Y = Y + 2 * iAlto
                ' puntuaciones
                If Not bJuezPasos Then
                    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                End If
                Y = Y + 2 * iAlto
                ' anulaciones
                If Not bJuezPasos Then
                    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                End If
                Y = Y + 2 * iAlto
                ' presente
                If Not bJuezPasos Then
                    Printer.Circle (X + 120, Y + 60), 60
                End If
                Y = Y + 2 * iAlto
                X = X + iAnchoCelda
            End If
            Inc iCDorsales
            If Not rsDorsales.EOF Then rsDorsales.MoveNext
        Wend
        rsDorsales.Close
    Else
        ' Imprimimos los bailes y dorsales
        If sCodFase = "1" Then
            ' Imprimimos una FINAL
            Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase= 1 ORDER BY posicion", dbOpenSnapshot)
            iCBailes = 0
            If iHoja = 2 Then
                While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                    Inc iCBailes
                    rsBailes.MoveNext
                Wend
            End If
            iCBailes = 0
            X = 0
            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
            While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                If iCodBaile = 0 Or iCodBaile = rsBailes!codigo Then
                    X = 0
                    Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    'Marca del baile
                    If G_NO_MARCAR_BAILES Then
                        Printer.Line (C_POS_X_MARCA_BAILE * iAnchoCelda, Y)-Step(iAncho, iAlto), , B
                    Else
                        Printer.Line (C_POS_X_MARCA_BAILE * iAnchoCelda, Y)-Step(iAncho, iAlto), , BF
                    End If
                                
                    Printer.Line (X, Y)-(X + Printer.Width / 1.2, Y)
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    Printer.Print rsBailes!Nombre
                    
                    Y = Y + 2 * iAlto
                    X = 2 * iAnchoCelda
                                
                                
                    ' Imprimimos los puestos
                    If Not bJuezPasos Then
                    Dim iFin As Integer, iMaxDor As Integer
                        iFin = IIf(bTeamMatch, 5, IIf(iMaxDorsalesTanda > C_MAX_DORSALES_HOJA_OPTICA_FINAL, C_MAX_DORSALES_HOJA_OPTICA_FINAL, iMaxDorsalesTanda))
                    
                        Printer.FontBold = True
                        iMaxDor = iFin
                        If G_IMPRIMIR_TODOS_LOS_CUADROS Then iFin = C_MAX_DORSALES_HOJA_OPTICA_FINAL
                        
                        For i = 1 To iFin
                            Printer.CurrentX = X
                            Printer.CurrentY = Y - iAlto / 3
                            If G_LINEAS_DIVISION_FINAL Then
                                Printer.Line Step(0, 0)-Step(8000, 0)
                                Printer.CurrentX = X
                            End If
                            If bTeamMatch Then
                                Printer.CurrentX = Printer.CurrentX - 60
                                Printer.FontSize = 10
                                Printer.Print Ptos2CadTeamMatch(CDbl(i + 1) / 2)
                            Else
                                Printer.FontSize = 11
                                If i <= iMaxDor Then Printer.Print i
                            End If
                            Printer.FontSize = 10
                            Y = Y + 2 * iAlto
                        Next
                        Printer.FontBold = False
                    End If
                    
                    Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
                    X = 3 * iAnchoCelda
                    If bJuezPasos Then
                        Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
                        Printer.Line (X - 2 * iAncho, Y + 2 * iAlto)-(X + iAncho * 33, Y + iAlto * 14), 0, B
                    End If
                    If rsDorsales.EOF Then
                    ' Imprimimos hojas sin conocer los dorsales
                        For iCDorsales = 1 To C_MAX_DORSAL_FINAL
                            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
                            Printer.CurrentX = X
                            Printer.CurrentY = Y
                            Printer.FontBold = True
                            Printer.Print "___"
                            ' Marca de descalificado
                            Printer.CurrentX = X + iAnchoCelda - iAncho / 2.5
                            Printer.CurrentY = Y - iAlto / 4
                            Printer.FontBold = False
                            Printer.FontSize = 8
                            If Not bTeamMatch Then
                                Printer.Print "d"
                            End If
                            Printer.FontSize = iTamFuente
                            ' Marca de anular
                            'If Not bJuezPasos Then
                            '    Printer.CurrentX = X + iAnchoCelda - iAncho / 2
                            '    Printer.CurrentY = Y + 2 * iAlto
                            '    Printer.FontBold = False
                            '    Printer.Print "a"
                            'End If
                            'Cuadro de descalificado
                            Printer.DrawStyle = 0
                            Printer.DrawMode = 13
                            If Not bTeamMatch Then
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Line (X + (iAncho + C_ANCHO_ESPACIO), Y)-(X + (iAncho + C_ANCHO_ESPACIO) + iAncho, Y + iAlto), , B
                            End If
                            Printer.ForeColor = C_COLOR_NEGRO
                            'Cuadros de los puestos
                            If Not bJuezPasos Then
                                iFin = IIf(bTeamMatch, 5, C_MAX_DORSAL_FINAL)
                                For i = 1 To iFin
                                    If Not bJuezPasos Then
                                        Printer.CurrentY = Y - iAlto / 4
                                        Printer.FontBold = False
                                        If bTeamMatch Then
                                            Printer.CurrentX = X - iAncho / 1.5 - 80
                                            Printer.FontSize = 7
                                            Printer.Print Ptos2CadTeamMatch(CDbl(i + 1) / 2)
                                        Else
                                            Printer.CurrentX = X - iAncho / 1.5
                                            Printer.FontSize = 9
                                            Printer.Print i
                                        End If
                                        Printer.FontSize = 10
                                        
                                        Printer.DrawStyle = vbDot
                                        Printer.DrawMode = 13
                                        Printer.Circle ((X + 2 * iAncho - 180), Y + 80), 60
                                        Printer.DrawStyle = 3
                                        Printer.DrawMode = 13
                                        Printer.Line ((X + 2 * iAncho), Y)-((X + 2 * iAncho) + iAncho, Y + iAlto), , B
                                        Printer.DrawStyle = 0
                                        Printer.DrawMode = 13
                                        Printer.ForeColor = Val(G_COLOR_MARCAS)
                                        If G_PUNTEO_ANULACION Then Printer.Circle ((X + 2 * iAncho - 140), Y + 80), 60
                                        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                                        Printer.ForeColor = C_COLOR_NEGRO
                                    End If
                                Next
                            End If
                            X = X + 2 * iAnchoCelda
                        Next
                    End If
                    'Imprimir hojas conociendo los dorsales
                    iCDorsales = 1
                    borradorY = 0
                    While Not rsDorsales.EOF Or (G_IMPRIMIR_TODOS_LOS_CUADROS And (iCDorsales - iNumDorsal) < C_MAX_DORSALES_HOJA_OPTICA_FINAL)
                        If iCDorsales >= iNumDorsal And (iCDorsales < iNumDorsal + iMaxDorsalesTanda Or G_IMPRIMIR_TODOS_LOS_CUADROS) Then
                            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
                            Printer.CurrentX = X
                            Printer.CurrentY = Y
                            Printer.FontBold = True
                            If rsDorsales.EOF Then
                                Printer.Print "___"
                            Else
                                Printer.Print Trim$(rsDorsales!num_dorsal)
                            End If
                            'Imprimimos el dorsal para la zona de borrador
                            If iMaxDorsalesTanda <= C_MAX_DORSALES_HOJA_OPTICA_FINAL Then
                                If Not bJuezPasos Then
                                    borradorY = IIf(borradorY = 0, Y + iAlto / 2, borradorY)
                                    borradorX = (C_POS_BORRADOR_EXT) * iAnchoCelda
                                    Printer.CurrentX = borradorX
                                    Printer.CurrentY = borradorY + iAlto / 2
                                    If rsDorsales.EOF Then
                                        Printer.Print "___"
                                    Else
                                        Printer.Print Trim$(rsDorsales!num_dorsal)
                                    End If
                                    borradorX = borradorX + iAnchoCelda / 4
                                    Printer.Line (borradorX + iAncho, borradorY)-(borradorX + 2.5 * iAncho, borradorY + 2 * iAlto), , B
                                    borradorX = borradorX + 1.5 * iAnchoCelda / 2
                                    Printer.Line (borradorX + iAncho, borradorY)-(borradorX + 2.5 * iAncho, borradorY + 2 * iAlto), , B
                                    borradorY = borradorY + 2 * iAlto
                                End If
                            End If
                            ' Marca de descalificado
                            Printer.CurrentX = X + iAnchoCelda - iAncho / 2.5
                            Printer.CurrentY = Y - iAlto / 4
                            Printer.FontBold = False
                            Printer.FontSize = 8
                            If Not bTeamMatch Then
                                Printer.Print "d"
                            End If
                            Printer.FontSize = iTamFuente
                            ' Marca de anular
                            'If Not bJuezPasos Then
                            '    Printer.CurrentX = X + iAnchoCelda - iAncho / 2
                            '    Printer.CurrentY = Y + 2 * iAlto
                            '    Printer.FontBold = False
                            '    Printer.Print "a"
                            'End If
                            'Cuadro de descalificado
                            Printer.DrawStyle = 0
                            Printer.DrawMode = 13
                            If Not bTeamMatch Then
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Line (X + (iAncho + C_ANCHO_ESPACIO), Y)-(X + (iAncho + C_ANCHO_ESPACIO) + iAncho, Y + iAlto), , B
                            End If
                            Printer.ForeColor = C_COLOR_NEGRO
                            'Cuadros de los puestos
                            If Not bJuezPasos Then
                                iFin = IIf(bTeamMatch, 5, IIf(iMaxDorsalesTanda > C_MAX_DORSALES_HOJA_OPTICA_FINAL, C_MAX_DORSALES_HOJA_OPTICA_FINAL, iMaxDorsalesTanda))
                                
                                iMaxDor = iFin
                                If G_IMPRIMIR_TODOS_LOS_CUADROS Then iFin = C_MAX_DORSALES_HOJA_OPTICA_FINAL
                                For i = 1 To iFin
                                    Y = Y + 2 * iAlto
                                    If Not bJuezPasos Then
                                        Printer.CurrentY = Y - iAlto / 4
                                        Printer.FontBold = False
                                        If bTeamMatch Then
                                            Printer.CurrentX = X - iAncho / 1.5
                                            Printer.FontSize = 5
                                            Printer.Print Ptos2CadTeamMatch(CDbl(i + 1) / 2)
                                        Else
                                            Printer.CurrentX = X - iAncho / 1.5
                                            Printer.FontSize = 9
                                            Printer.Print i
                                        End If
                                        Printer.FontSize = 10
                                        
                                        Printer.DrawStyle = vbDot
                                        Printer.DrawMode = 13
                                        Printer.Circle ((X + 2 * iAncho - 180), Y + 80), 60
                                        Printer.DrawStyle = 3
                                        Printer.DrawMode = 13
                                        Printer.Line ((X + (iAncho + C_ANCHO_ESPACIO)), Y)-((X + (iAncho + C_ANCHO_ESPACIO)) + iAncho, Y + iAlto), , B
                                        Printer.DrawStyle = 0
                                        Printer.DrawMode = 13
                                        Printer.ForeColor = Val(G_COLOR_MARCAS)
                                        If G_PUNTEO_ANULACION Then Printer.Circle ((X - 200), Y + 80), 60
                                        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                                        Printer.ForeColor = C_COLOR_NEGRO
                                    End If
                                Next
                            End If
                            X = X + 2 * iAnchoCelda
                        End If
                        Inc iCDorsales
                        rsDorsales.MoveNext
                    Wend
                    rsDorsales.Close
                End If
                Inc iCBailes
                rsBailes.MoveNext
            Wend
        Else
            ' Imprimimos los bailes y dorsales de NO FINAL
            Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase = 2 ORDER BY posicion", dbOpenSnapshot)
            iCBailes = 0
            If iHoja = 2 Then
                While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                    Inc iCBailes
                    rsBailes.MoveNext
                Wend
            End If
            iCBailes = 0
            X = 0
            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
            While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                If iCodBaile = 0 Or iCodBaile = rsBailes!codigo Then
                    X = 0
                    Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    
                    'Marca del baile
                    If G_NO_MARCAR_BAILES Then
                        Printer.Line (C_POS_X_MARCA_BAILE * iAnchoCelda, Y + 8 * iAlto)-Step(iAncho, iAlto), , B
                    Else
                        Printer.Line (C_POS_X_MARCA_BAILE * iAnchoCelda, Y + 8 * iAlto)-Step(iAncho, iAlto), , BF
                    End If
                    
                    Printer.Line (X, Y)-(X + Printer.Width / 1.2, Y)
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    Printer.FontBold = True
                    If Len(rsBailes!Nombre) > 7 Then
                        Printer.FontSize = 8
                    End If
                    Printer.Print rsBailes!Nombre
                    Printer.FontSize = iTamFuente
                    Printer.FontBold = False
                    X = 0
                    Y = Y + 1 * iAlto
                    If Not bJuezPasos Then
                        Printer.CurrentX = X
                        Printer.CurrentY = Y + iAlto / 2
                        Printer.Print mml_FRASE0621
                    End If
                    X = 0
                    Y = Y + 2 * iAlto
                    If Not bJuezPasos Then
                        Printer.CurrentX = X
                        Printer.CurrentY = Y + iAlto / 2
                        Printer.Print mml_FRASE0619
                    End If
                    X = 0
                    Y = Y + 2 * iAlto
                    If bJuezPasos Then
                        Y = Y - 4 * iAlto
                    End If
                    Printer.CurrentX = X
                    Printer.CurrentY = Y + iAlto / 2
                    Printer.Print mml_FRASE0552
                    
                    If Not bJuezPasos Then
                        X = 0
                        Y = Y + 2 * iAlto
                        Printer.CurrentX = X
                        Printer.CurrentY = Y + iAlto / 2
                        Printer.Print mml_FRASE0622
                    End If
                    
                    If iMaxTandas > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
                        Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & iRepesca & " AND d.fase =" & tbCodFase.Text & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
                    Else
                        Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
                    End If
                    X = 2 * iAnchoCelda
                    If bJuezPasos Then
                        Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                        'Printer.Line (x - iAncho, y + 2 * iAlto)-(x + iAncho * 35, y + iAlto * 11 - C_MARGEN_DESC), 0, B
                        Printer.Line (X - iAncho, Y + 4 * iAlto)-(X + iAncho * 35, Y + iAlto * 10 - C_MARGEN_DESC), 0, B
                    End If
                    iCDorsales = 1
                    If rsDorsales.EOF Then
                        For i = 1 To C_MAX_DORSALES_HOJA_OPTICA_EXT
                            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                            Printer.CurrentX = X
                            Printer.CurrentY = Y
                            Printer.Print "___"
                            Y = Y + 2 * iAlto
                            ' puntuaciones
                            If Not bJuezPasos Then
                                'Linea sólida
                                Printer.DrawStyle = 0
                                Printer.DrawMode = 13
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                                Printer.ForeColor = C_COLOR_NEGRO
                            End If
                            Y = Y + 2 * iAlto
                            ' anulaciones
                            If Not bJuezPasos Then
                                'Linea punteada
                                Printer.DrawStyle = 3
                                Printer.DrawMode = 13
                                Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                            End If
                            Y = Y + 2 * iAlto
                            If bJuezPasos Then
                                Y = Y - 4 * iAlto
                            End If
                            ' descalificaciones
                            'Linea sólida
                            Printer.DrawStyle = 0
                            Printer.DrawMode = 13
                            Printer.ForeColor = Val(G_COLOR_MARCAS)
                            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                            Printer.ForeColor = C_COLOR_NEGRO
                            
                            'Preselección
                            If Not bJuezPasos Then
                                Y = Y + 2 * iAlto
                                'Linea sólida
                                Printer.DrawStyle = vbSolid
                                Printer.DrawMode = 13
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Circle (X + 50, Y + 60), 60
                                Printer.DrawStyle = vbDot
                                Printer.Circle (X + 210, Y + 60), 60
                                Printer.ForeColor = C_COLOR_NEGRO
                                Printer.DrawStyle = vbSolid
                            End If
                            
                            X = X + iAnchoCelda
                        Next
                    End If
                    While Not rsDorsales.EOF Or (G_IMPRIMIR_TODOS_LOS_CUADROS And (iCDorsales - iNumDorsal) < C_MAX_DORSALES_HOJA_OPTICA_EXT)
                        If iCDorsales >= iNumDorsal And (iCDorsales < iNumDorsal + iMaxDorsalesTanda Or G_IMPRIMIR_TODOS_LOS_CUADROS) Then
                            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                            Printer.CurrentX = X
                            Printer.CurrentY = Y
                            Printer.FontBold = True
                            If (iCDorsales < iNumDorsal + iMaxDorsalesTanda) Then
                                Printer.Print Trim$(rsDorsales!num_dorsal)
                            Else
                                Printer.Print "___"
                            End If
                            Printer.FontBold = False
                            Y = Y + 2 * iAlto
                            ' puntuaciones
                            If Not bJuezPasos Then
                                'Linea sólida
                                Printer.DrawStyle = 0
                                Printer.DrawMode = 13
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                                Printer.ForeColor = C_COLOR_NEGRO
                            End If
                            Y = Y + 2 * iAlto
                            ' anulaciones
                            If Not bJuezPasos Then
                                'Linea punteada
                                Printer.DrawStyle = 3
                                Printer.DrawMode = 13
                                Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                            End If
                            Y = Y + 2 * iAlto
                            If bJuezPasos Then
                                Y = Y - 4 * iAlto
                            End If
                            ' descalificaciones
                            'Linea punteada
                            Printer.DrawStyle = 0
                            Printer.DrawMode = 13
                            Printer.ForeColor = Val(G_COLOR_MARCAS)
                            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                            Printer.ForeColor = C_COLOR_NEGRO
                            
                            'Preselección
                            If Not bJuezPasos Then
                                Y = Y + 2 * iAlto
                                'Linea sólida
                                Printer.DrawStyle = vbSolid
                                Printer.DrawMode = 13
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Circle (X + 50, Y + 60), 60
                                Printer.DrawStyle = vbDot
                                Printer.Circle (X + 210, Y + 60), 60
                                Printer.ForeColor = C_COLOR_NEGRO
                                Printer.DrawStyle = vbSolid
                            End If
                            
                            X = X + iAnchoCelda
                        End If
                        Inc iCDorsales
                        If Not rsDorsales.EOF Then rsDorsales.MoveNext
                    Wend
                    rsDorsales.Close
                End If
                Inc iCBailes
                rsBailes.MoveNext
            Wend
        End If
    End If
End Sub

Public Sub ImprimirHojaPuntuacionesNormal(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer, Optional iCodBaile = 0, Optional lOrden As Long = -1)
Dim rs As Recordset
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim X As Integer, Y As Integer
Dim i As Integer, j As Integer, k As Integer
Dim iAncho As Integer, iAlto As Integer
Dim iCCateg As Integer
Dim iCBailes As Integer
Dim iCDorsales As Integer
Dim bJuezPasos As Boolean
Dim sInfo As String
Dim iMaxDorsalesPorTanda As Integer
Dim iCMarcas As Integer
Dim iTamFuente As Integer
Dim bTeamMatch As Boolean

Const MAX_COL = 28
Const MAX_FILA = 80
Const INIC_CAT = 5
Const INIC_FASE = 12
Const INIC_JUEZ = 16
Const INIC_TANDA = 20
Const INIC_BAILES = 24
Const POS_X_MARCA_BAILE = 38
Dim ALTO_BAILE As Integer
Dim ALTO_BAILE_FINAL As Integer

Dim borradorX As Integer, borradorY As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    
    bTeamMatch = ComprobarSiTeamMatch(sCodCat)
    
    bJuezPasos = False
    Printer.CurrentY = MARGEN_SUPERIOR
    Y = MARGEN_SUPERIOR
    Printer.FillColor = &HFFFFFF
    Printer.FillStyle = 0
    iAncho = VarCfg("ancho_marca")
    iAlto = VarCfg("alto_marca")
    
    ALTO_BAILE = 2 * C_ANCHO_MARCAS_BAILE * iAlto
    ALTO_BAILE_FINAL = 16 * iAlto
    
    'Imprimir marca de control de calidad de escaneo
    Y = MARGEN_SUPERIOR + (2 * iAlto) * 6
    X = 19 * 2 * iAncho
    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    Printer.Line (X, Y + iAlto / 2)-(X + iAncho, Y + iAlto), , BF
    
    Y = MARGEN_SUPERIOR
    X = 0
    'Imprimir marcas
    iCMarcas = 1
    While iCMarcas <= C_MAX_MARCAS_X_NORMAL - 1
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        X = X + 2 * iAncho
        Inc iCMarcas
    Wend
    iCMarcas = 1
    While iCMarcas <= C_MAX_MARCAS_Y
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Y = Y + 2 * iAlto
        Inc iCMarcas
    Wend
    'Imprimir marca de control inferior
    Printer.Line (0, Y - 2 * iAlto)-(iAncho, Y - 2 * iAlto + iAlto), , BF
    
    If sCodFase > 1 And chkRep.Value = 0 Then
        sInfo = mml_FRASE0608 & (sCodFase / 2) * 6
        If iMaxTandas > 1 Then
            sInfo = sInfo & ") - ("
            If lblTandasMenosDorsales.Caption = "" Then
                sInfo = sInfo & mml_FRASE0609 & tbDorsalesTanda.Text
            Else
                sInfo = sInfo & "Primeras rondas " & lblTandasMasDorsales.Caption & " dorsales, Siguientes " & lblTandasMenosDorsales.Caption & " dorsales"
            End If
            Dim iMediaPorRonda As Integer
            iMediaPorRonda = ((sCodFase / 2) * 6) \ iMaxTandas
            sInfo = sInfo & " ,Media/Ronda " & iMediaPorRonda
            If ((sCodFase / 2) * 6) Mod iMaxTandas > 0 Then
                sInfo = sInfo & " faltan " & iMaxDorsales - iMaxTandas * iMediaPorRonda & " marcas"
            End If
        End If
        Printer.FontBold = False
        sInfo = sInfo & ") "
    End If
    
    Printer.FontName = "Arial"
    Printer.CurrentY = MARGEN_SUPERIOR + 2 * iAlto - iAlto / 2
    Printer.CurrentX = 0
    Printer.FontSize = 14
    Printer.Print tbDescCat.Text;
    Printer.FontSize = 12
    Printer.Print " - " & sDescFase(Val(sCodFase)) & " - " & sIdJuez;
    If iMaxTandas > 1 Then
        Printer.Print mml_FRASE0613 & iTanda & mml_FRASE0472 & iMaxTandas;
    End If
    Printer.Print "  - (" & tbDescComp.Text & ")";
    Printer.FontSize = 10
    If lOrden > 0 Then Printer.Print " (" & lOrden & ")";
    Printer.Print
    If sIdJuez <> "Control" Then
        If chkInfoTandas.Value = 1 Then
            Printer.Print sInfo & "    " & tbComen.Text
        Else
            Printer.Print tbComen.Text
        End If
    End If
    
    ' {·M} Imprimimos el hueco para la firma
    Printer.CurrentX = iAncho * 30
    Printer.CurrentY = MARGEN_SUPERIOR + iAlto * 12
    Printer.Print mml_FRASE0614
    Printer.Line (iAncho * 30, MARGEN_SUPERIOR + iAlto * 14)-Step(iAncho * 9, iAlto * 6), 0, B
    
    ' Imprimimos las categorías
    Set rs = db.OpenRecordset("SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_CAT * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0623
    X = 4 * iAncho
    iCCateg = 0
    iTamFuente = Printer.FontSize
    Printer.FontSize = 7
    While Not rs.EOF
        If iCCateg <> 0 And iCCateg Mod 18 = 0 Then
            X = 4 * iAncho
        End If
        
        Y = MARGEN_SUPERIOR + INIC_CAT * iAlto + 2 * iAlto * Int(iCCateg / 18)
        Printer.CurrentY = Y
        
        Printer.CurrentX = X
        Printer.Print rs!codigo
        Y = Y + iAlto
        If rs!codigo = sCodCat Then
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Else
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
        End If
        X = X + 2 * iAncho
        rs.MoveNext
        Inc iCCateg
    Wend
    rs.Close
    Printer.FontSize = iTamFuente
    
    
    ' Imprimimos la fase
    X = 0
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0299
    X = 4 * iAncho
    i = 1
    While i <= 256
        Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
        Printer.CurrentX = X
        'Las categorías pueden ocupar mucho
        If iCCateg > 54 Then
            Printer.CurrentY = Y + iAlto / 1.8
        Else
            Printer.CurrentY = Y
        End If
        Printer.Print sDescFase(i)
        Y = Y + 2 * iAlto
        If i = sCodFase Then
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Else
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
        End If
        X = X + 2 * iAncho
        i = i * 2
    Wend
    
    
    'Imprimir marca de indicación de problemas
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    X = 12 * 2 * iAncho
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print "Err"
    Y = Y + 2 * iAlto
    Printer.ForeColor = Val(G_COLOR_MARCAS)
    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    Printer.ForeColor = C_COLOR_NEGRO
    
    ' Imprimimos la marca de identificación de hoja extendida
    X = C_POS_HOJA_EXT * iAncho
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print "Ext"
    Y = Y + 2 * iAlto
    'solo marcamos la hoja como extendida si utiliza caracteristicas extendidas
    If BailesParciales(sCodCat) Then
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
    Else
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    End If
    
    'Imprimimos los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos FROM juez_categ WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0421
    X = 4 * iAncho
    While Not rs.EOF
        Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y
        If rs!pasos = 1 Then
            Printer.Print rs!id_juez & "p"
            If rs!id_juez = sIdJuez Then
                bJuezPasos = True
            End If
        Else
            Printer.Print rs!id_juez
        End If
        Y = Y + 2 * iAlto
        If rs!id_juez = sIdJuez Then
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Else
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
        End If
        X = X + 2 * iAncho
        rs.MoveNext
    Wend
    rs.Close
    
    If bJuezPasos Then
        ALTO_BAILE = 2 * C_ANCHO_MARCAS_BAILE_JUEZ_PASOS * iAlto
    End If
    
    ' Imprimimos la marca de control
    X = POS_CONTROL * iAncho
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print "Ctrl"
    Y = Y + 2 * iAlto
    If sIdJuez = "Control" Then
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
    Else
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    End If
    
    ' Imprimimos la marca de repesca
    X = POS_REPESCA * iAncho
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0418
    Y = Y + 2 * iAlto
    If iRepesca = 1 Then
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
    Else
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    End If
    
    ' Imprimimos la marca de hoja2
    X = POS_HOJA2 * iAncho
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0624
    Y = Y + 2 * iAlto
    If iHoja = 2 Then
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
    Else
        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
    End If
    
    ' Imprimimos la tanda
    X = 0
    Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0617
    X = 4 * iAncho
    i = 1
    While i <= 18
        Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Print Trim$(Str$(i))
        Y = Y + 2 * iAlto
        If i = iTanda Or i = iMaxTandas Then
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , BF
        Else
            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
        End If
        X = X + 2 * iAncho
        i = i + 1
    Wend
    
    If sIdJuez = "Control" Then
        X = 0
        Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Line (X, Y)-(X + Printer.Width / 1.2, Y)
        Printer.CurrentX = X
        Printer.CurrentY = Y + iAlto + iAlto / 2
        Printer.Print mml_FRASE0618
        X = 0
        Y = Y + 3 * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y + iAlto / 2
        Printer.Print mml_FRASE0619
        X = 0
        Y = Y + 2 * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y + iAlto / 2
        Printer.Print mml_FRASE0620
        
        Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        X = 4 * iAncho
        iCDorsales = 1
        While Not rsDorsales.EOF Or (G_IMPRIMIR_TODOS_LOS_CUADROS And (iCDorsales - iNumDorsal) < C_MAX_DORSALES_HOJA_OPTICA)
            If iCDorsales >= iNumDorsal And (iCDorsales < iNumDorsal + iMaxDorsalesTanda Or G_IMPRIMIR_TODOS_LOS_CUADROS) Then
                Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                Printer.CurrentX = X
                Printer.CurrentY = Y
                Printer.FontBold = True
                If (iCDorsales < iNumDorsal + iMaxDorsalesTanda) Then
                    Printer.Print Trim$(rsDorsales!num_dorsal)
                Else
                    Printer.Print "___"
                End If
                Printer.FontBold = False
                Y = Y + 2 * iAlto
                ' ausente
                If Not bJuezPasos Then
                    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                End If
                Y = Y + 2 * iAlto
                ' anulaciones
                If Not bJuezPasos Then
                    Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                End If
                Y = Y + 2 * iAlto
                ' presente
                If Not bJuezPasos Then
                    Printer.Circle (X + 120, Y + 60), 60
                End If
                Y = Y + 2 * iAlto
                X = X + 2 * iAncho
            End If
            Inc iCDorsales
            If Not rsDorsales.EOF Then rsDorsales.MoveNext
        Wend
        rsDorsales.Close
    Else
        ' Imprimimos los bailes y dorsales
        If sCodFase = "1" Then
' Imprimimos una FINAL ******************************************************************
            Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase= 1 ORDER BY posicion", dbOpenSnapshot)
            iCBailes = 0
            If iHoja = 2 Then
                While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                    Inc iCBailes
                    rsBailes.MoveNext
                Wend
            End If
            iCBailes = 0
            X = 0
            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
            While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                If iCodBaile = 0 Or iCodBaile = rsBailes!codigo Then
                    X = 0
                    Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    'Marca del baile
                    If Not G_NO_MARCAR_BAILES And (iCodBaile = 0 Or iCodBaile = rsBailes!codigo) Then
                        Printer.Line (iAncho * POS_X_MARCA_BAILE, Y)-Step(iAncho, iAlto), , BF
                    Else
                        Printer.Line (iAncho * POS_X_MARCA_BAILE, Y)-Step(iAncho, iAlto), , B
                    End If
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    Printer.Line (X, Y)-(X + Printer.Width / 1.2, Y)
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    Printer.Print rsBailes!Nombre
                    Y = Y + 2 * iAlto
                    X = 4 * iAncho
                                
                    ' Imprimimos los puestos
                    If Not bJuezPasos Then
                    Dim iFin As Integer, iMaxDor As Integer
                        iFin = IIf(bTeamMatch, 5, IIf(iMaxDorsalesTanda > C_MAX_DORSALES_HOJA_OPTICA_FINAL, C_MAX_DORSALES_HOJA_OPTICA_FINAL, iMaxDorsalesTanda))
                        Printer.FontBold = True
                        
                        iMaxDor = iFin
                        If G_IMPRIMIR_TODOS_LOS_CUADROS Then iFin = C_MAX_DORSALES_HOJA_OPTICA_FINAL
                        
                        For i = 1 To iFin
                            Printer.CurrentX = X
                            Printer.CurrentY = Y - iAlto / 3
                            If G_LINEAS_DIVISION_FINAL Then
                                Printer.Line Step(0, 0)-Step(7800, 0)
                                Printer.CurrentX = X
                            End If
                            If bTeamMatch Then
                                Printer.CurrentX = Printer.CurrentX - 60
                                Printer.FontSize = 10
                                Printer.Print Ptos2CadTeamMatch(CDbl(i + 1) / 2)
                            Else
                                Printer.FontSize = 11
                                If i <= iMaxDor Then Printer.Print i
                            End If
                            Printer.FontSize = 10
                            Y = Y + 2 * iAlto
                        Next
                        Printer.FontBold = False
                    End If
                    
                    Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
                    X = 6 * iAncho
                    If bJuezPasos Then
                        Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
                        Printer.Line (X - 2 * iAncho, Y + 2 * iAlto)-(X + iAncho * 33, Y + iAlto * 14), 0, B
                    End If
                    If rsDorsales.EOF Then
                    ' Imprimimos hojas sin conocer los dorsales
                        For iCDorsales = 1 To C_MAX_DORSAL_FINAL
                            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
                            Printer.CurrentX = X
                            Printer.CurrentY = Y
                            Printer.FontBold = True
                            Printer.Print "___"
                            ' Marca de descalificado
                            If Not bTeamMatch Then
                                Printer.CurrentX = X + 2 * iAncho - iAncho / 2
                                Printer.CurrentY = Y - iAlto / 3
                                Printer.FontBold = False
                                Printer.Print "d"
                            End If
                            ' Marca de anular
                            'If Not bJuezPasos Then
                            '    Printer.CurrentX = X + 2 * iAncho - iAncho / 2
                            '    Printer.CurrentY = Y + 2 * iAlto
                            '    Printer.FontBold = False
                            '    Printer.Print "a"
                            'End If
                            'Cuadro de descalificado
                            Printer.DrawStyle = 0
                            Printer.DrawMode = 13
                            If Not bTeamMatch Then
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Line (X + 2 * iAncho, Y)-(X + 3 * iAncho, Y + iAlto), , B
                            End If
                            Printer.ForeColor = C_COLOR_NEGRO
                            'Cuadros de los puestos
                            If Not bJuezPasos Then
                                iFin = IIf(bTeamMatch, 5, C_MAX_DORSAL_FINAL)
                                For i = 1 To iFin
                                    Y = Y + 2 * iAlto
                                    If Not bJuezPasos Then
                                        Printer.CurrentY = Y - iAlto / 4
                                        Printer.FontBold = False
                                        If bTeamMatch Then
                                            Printer.CurrentX = X - iAncho / 1.5 - 80
                                            Printer.FontSize = 7
                                            Printer.Print Ptos2CadTeamMatch(CDbl(i + 1) / 2)
                                        Else
                                            Printer.CurrentX = X - iAncho / 1.5
                                            Printer.FontSize = 9
                                            Printer.Print i
                                        End If
                                        Printer.FontSize = 10
                                        
                                        Printer.DrawStyle = vbDot
                                        Printer.DrawMode = 13
                                        Printer.Circle ((X + 2 * iAncho - 140), Y + 80), 60
                                        Printer.DrawStyle = 3
                                        Printer.DrawMode = 13
                                        Printer.Line ((X + 2 * iAncho), Y)-((X + 2 * iAncho) + iAncho, Y + iAlto), , B
                                        Printer.DrawStyle = 0
                                        Printer.DrawMode = 13
                                        Printer.ForeColor = Val(G_COLOR_MARCAS)
                                        If G_PUNTEO_ANULACION Then Printer.Circle ((X - 180), Y + 80), 60
                                        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                                        Printer.ForeColor = C_COLOR_NEGRO
                                    End If
                                Next
                            End If
                            X = X + 4 * iAncho
                        Next
                    End If
                    'Imprimir hojas conociendo los dorsales
                    iCDorsales = 1
                    borradorY = 0
                    While Not rsDorsales.EOF Or (G_IMPRIMIR_TODOS_LOS_CUADROS And (iCDorsales - iNumDorsal) < C_MAX_DORSALES_HOJA_OPTICA_FINAL)
                        If iCDorsales >= iNumDorsal And (iCDorsales < iNumDorsal + iMaxDorsalesTanda Or G_IMPRIMIR_TODOS_LOS_CUADROS) Then
                            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE_FINAL * iCBailes
                            Printer.CurrentX = X
                            Printer.CurrentY = Y
                            Printer.FontBold = True
                            If rsDorsales.EOF Then
                                Printer.Print "___"
                            Else
                                Printer.Print Trim$(rsDorsales!num_dorsal)
                            End If
                            'Imprimimos el dorsal para la zona de borrador
                            If iMaxDorsalesTanda <= C_MAX_DORSALES_HOJA_OPTICA_FINAL Then
                                If Not bJuezPasos Then
                                    borradorY = IIf(borradorY = 0, Y + iAlto / 2, borradorY)
                                    borradorX = 2 * iAncho * C_POS_BORRADOR
                                    Printer.CurrentX = borradorX + 50
                                    Printer.CurrentY = borradorY + iAlto / 2
                                    If rsDorsales.EOF Then
                                        Printer.Print "___"
                                    Else
                                        Printer.Print Trim$(rsDorsales!num_dorsal)
                                    End If
                                    borradorX = borradorX + iAncho / 2
                                    Printer.Line (borradorX + iAncho, borradorY)-(borradorX + 2.5 * iAncho, borradorY + 2 * iAlto), , B
                                    borradorX = borradorX + 1.5 * iAncho
                                    Printer.Line (borradorX + iAncho, borradorY)-(borradorX + 2.5 * iAncho, borradorY + 2 * iAlto), , B
                                    borradorY = borradorY + 2 * iAlto
                                End If
                            End If
                            ' Marca de descalificado
                            Printer.CurrentX = X + 2 * iAncho - iAncho / 2
                            Printer.CurrentY = Y - iAlto / 3
                            If Not bTeamMatch Then
                                Printer.FontBold = False
                                Printer.Print "d"
                            End If
                            ' Marca de anular
                            'If Not bJuezPasos Then
                            '    Printer.CurrentX = X + 2 * iAncho - iAncho / 2
                            '    Printer.CurrentY = Y + 2 * iAlto
                            '    Printer.FontBold = False
                            '    Printer.Print "a"
                            'End If
                            'Cuadro de descalificado
                            Printer.DrawStyle = 0
                            Printer.DrawMode = 13
                            If Not bTeamMatch Then
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Line (X + 2 * iAncho, Y)-(X + 3 * iAncho, Y + iAlto), , B
                            End If
                            Printer.ForeColor = C_COLOR_NEGRO
                            'Cuadros de los puestos
                            If Not bJuezPasos Then
                                iFin = IIf(bTeamMatch, 5, IIf(iMaxDorsalesTanda > C_MAX_DORSALES_HOJA_OPTICA_FINAL, C_MAX_DORSALES_HOJA_OPTICA_FINAL, iMaxDorsalesTanda))
                                
                                iMaxDor = iFin
                                If G_IMPRIMIR_TODOS_LOS_CUADROS Then iFin = C_MAX_DORSALES_HOJA_OPTICA_FINAL
                                For i = 1 To iFin
                                    Y = Y + 2 * iAlto
                                    If Not bJuezPasos Then
                                        Printer.CurrentY = Y - iAlto / 4
                                        Printer.FontBold = False
                                        If bTeamMatch Then
                                            Printer.CurrentX = X - iAncho / 1.5 - 80
                                            Printer.FontSize = 7
                                            Printer.Print Ptos2CadTeamMatch(CDbl(i + 1) / 2)
                                        Else
                                            Printer.CurrentX = X - iAncho / 1.5
                                            Printer.FontSize = 9
                                            Printer.Print i
                                        End If
                                        Printer.FontSize = 10
                                        
                                        Printer.DrawStyle = vbDot
                                        Printer.DrawMode = 13
                                        Printer.Circle ((X + 2 * iAncho - 140), Y + 80), 60
                                        Printer.DrawStyle = 3
                                        Printer.DrawMode = 13
                                        Printer.Line ((X + 2 * iAncho), Y)-((X + 2 * iAncho) + iAncho, Y + iAlto), , B
                                        Printer.DrawStyle = vbSolid
                                        Printer.DrawMode = 13
                                        Printer.ForeColor = Val(G_COLOR_MARCAS)
                                        If G_PUNTEO_ANULACION Then Printer.Circle ((X - 180), Y + 80), 60
                                        Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                                        Printer.ForeColor = C_COLOR_NEGRO
                                    End If
                                Next
                            End If
                            X = X + 4 * iAncho
                        End If
                        Inc iCDorsales
                        If Not rsDorsales.EOF Then rsDorsales.MoveNext
                    Wend
                    rsDorsales.Close
                End If
                Inc iCBailes
                rsBailes.MoveNext
            Wend
        Else
' Imprimimos los bailes y dorsales de NO FINAL *******************************************
            Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase = 2 ORDER BY posicion", dbOpenSnapshot)
            iCBailes = 0
            If iHoja = 2 Then
                While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                    Inc iCBailes
                    rsBailes.MoveNext
                Wend
            End If
            iCBailes = 0
            X = 0
            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
            While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                If iCodBaile = 0 Or iCodBaile = rsBailes!codigo Then
                    X = 0
                    Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    'Marca del baile
                    If Not G_NO_MARCAR_BAILES And (iCodBaile = 0 Or iCodBaile = rsBailes!codigo) Then
                        Printer.Line (POS_X_MARCA_BAILE * iAncho, Y + 8 * iAlto)-Step(iAncho, iAlto), , BF
                    Else
                        Printer.Line (POS_X_MARCA_BAILE * iAncho, Y + 8 * iAlto)-Step(iAncho, iAlto), , B
                    End If
                    
                    Printer.Line (X, Y)-(X + Printer.Width / 1.2, Y)
                    Printer.CurrentX = X
                    Printer.CurrentY = Y
                    Printer.FontBold = True
                    Printer.Print rsBailes!Nombre
                    Printer.FontBold = False
                    X = 0
                    Y = Y + 1 * iAlto
                    If Not bJuezPasos Then
                        Printer.CurrentX = X
                        Printer.CurrentY = Y + iAlto / 2
                        Printer.Print mml_FRASE0625
                    End If
                    X = 0
                    Y = Y + 2 * iAlto
                    If Not bJuezPasos Then
                        Printer.CurrentX = X
                        Printer.CurrentY = Y + iAlto / 2
                        Printer.Print mml_FRASE0619
                    End If
                    X = 0
                    Y = Y + 2 * iAlto
                    If bJuezPasos Then
                        Y = Y - 4 * iAlto
                    End If
                    Printer.CurrentX = X
                    Printer.CurrentY = Y + iAlto / 2
                    Printer.Print mml_FRASE0626
                    
                    If Not bJuezPasos Then
                        X = 0
                        Y = Y + 2 * iAlto
                        Printer.CurrentX = X
                        Printer.CurrentY = Y + iAlto / 2
                        Printer.Print mml_FRASE0627
                    End If
                    
                    If iMaxTandas > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
                        Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & iRepesca & " AND d.fase =" & tbCodFase.Text & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
                    Else
                        Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
                    End If
                    X = 4 * iAncho
                    If bJuezPasos Then
                        Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                        'Printer.Line (x - iAncho, y + 2 * iAlto)-(x + iAncho * 35, y + iAlto * 11 - C_MARGEN_DESC), 0, B
                        Printer.Line (X - iAncho, Y + 4 * iAlto)-(X + iAncho * 34, Y + iAlto * 10 - C_MARGEN_DESC), 0, B
                    End If
                    iCDorsales = 1
                    'Sin conocer los dorsales
                    If rsDorsales.EOF Then
                        For i = 1 To C_MAX_DORSALES_HOJA_OPTICA
                            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                            Printer.CurrentX = X
                            Printer.CurrentY = Y
                            Printer.Print "___"
                            Y = Y + 2 * iAlto
                            ' puntuaciones
                            If Not bJuezPasos Then
                                'Linea sólida
                                Printer.DrawStyle = 0
                                Printer.DrawMode = 13
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                                Printer.ForeColor = C_COLOR_NEGRO
                            End If
                            Y = Y + 2 * iAlto
                            ' anulaciones
                            If Not bJuezPasos Then
                                'Linea punteada
                                Printer.DrawStyle = 3
                                Printer.DrawMode = 13
                                Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                            End If
                            Y = Y + 2 * iAlto
                            If bJuezPasos Then
                                Y = Y - 4 * iAlto
                            End If
                            ' descalificaciones
                            'Linea sólida
                            Printer.DrawStyle = 0
                            Printer.DrawMode = 13
                            Printer.ForeColor = Val(G_COLOR_MARCAS)
                            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                            Printer.ForeColor = C_COLOR_NEGRO
                            
                            'Preselección
                            If Not bJuezPasos Then
                                Y = Y + 2 * iAlto
                                'Linea sólida
                                Printer.DrawStyle = vbSolid
                                Printer.DrawMode = 13
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Circle (X + 50, Y + 60), 60
                                Printer.DrawStyle = vbDot
                                Printer.Circle (X + 210, Y + 60), 60
                                Printer.ForeColor = C_COLOR_NEGRO
                                Printer.DrawStyle = vbSolid
                            End If
                            
                            X = X + 2 * iAncho
                        Next
                    End If
                    While Not rsDorsales.EOF Or (G_IMPRIMIR_TODOS_LOS_CUADROS And (iCDorsales - iNumDorsal) < C_MAX_DORSALES_HOJA_OPTICA)
                        If iCDorsales >= iNumDorsal And (iCDorsales < iNumDorsal + iMaxDorsalesTanda Or G_IMPRIMIR_TODOS_LOS_CUADROS) Then
                            Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                            Printer.CurrentX = X
                            Printer.CurrentY = Y
                            Printer.FontBold = True
                            If (iCDorsales < iNumDorsal + iMaxDorsalesTanda) Then
                                Printer.Print Trim$(rsDorsales!num_dorsal)
                            Else
                                Printer.Print "___"
                            End If
                            Printer.FontBold = False
                            Y = Y + 2 * iAlto
                            ' puntuaciones
                            If Not bJuezPasos Then
                                'Linea sólida
                                Printer.DrawStyle = 0
                                Printer.DrawMode = 13
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                                Printer.ForeColor = C_COLOR_NEGRO
                            End If
                            Y = Y + 2 * iAlto
                            ' anulaciones
                            If Not bJuezPasos Then
                                'Linea punteada
                                Printer.DrawStyle = 3
                                Printer.DrawMode = 13
                                Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                            End If
                            Y = Y + 2 * iAlto
                            If bJuezPasos Then
                                Y = Y - 4 * iAlto
                            End If
                            ' descalificaciones
                            'Linea punteada
                            Printer.DrawStyle = 0
                            Printer.DrawMode = 13
                            Printer.ForeColor = Val(G_COLOR_MARCAS)
                            Printer.Line (X, Y)-(X + iAncho, Y + iAlto), , B
                            Printer.ForeColor = C_COLOR_NEGRO
                            
                            'Preselección
                            If Not bJuezPasos And (iCDorsales - iNumDorsal + 1) < C_MAX_DORSALES_HOJA_OPTICA Then
                                Y = Y + 2 * iAlto
                                'Linea sólida
                                Printer.DrawStyle = vbSolid
                                Printer.DrawMode = 13
                                Printer.ForeColor = Val(G_COLOR_MARCAS)
                                Printer.Circle (X + 50, Y + 60), 60
                                Printer.DrawStyle = vbDot
                                Printer.Circle (X + 210, Y + 60), 60
                                Printer.ForeColor = C_COLOR_NEGRO
                                Printer.DrawStyle = vbSolid
                            End If

                            X = X + 2 * iAncho
                        End If
                        Inc iCDorsales
                        If Not rsDorsales.EOF Then rsDorsales.MoveNext
                    Wend
                    rsDorsales.Close
                End If
                Inc iCBailes
                rsBailes.MoveNext
            Wend
        End If
    End If
error: ProcesarError
    
End Sub


Public Sub ImprimirHojaPuntuacionesJuez(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer)
Dim rs As Recordset
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim X As Integer, Y As Integer
Dim i As Integer, j As Integer, k As Integer
Dim iAncho As Integer, iAlto As Integer
Dim iCCateg As Integer
Dim iCBailes As Integer
Dim iCDorsales As Integer
Dim bJuezPasos As Boolean
Dim sInfo As String
Dim iMaxDorsalesPorTanda As Integer
Dim iTamFuente As Integer
Dim rsCateg As Recordset

Const MAX_COL = 28
Const MAX_FILA = 80
Const INIC_CAT = 6
Const INIC_FASE = 8
Const INIC_JUEZ = 10
Const INIC_TANDA = 12
Const INIC_BAILES = 16
Dim ALTO_BAILE As Integer
Dim ALTO_BAILE_FINAL As Integer

    bJuezPasos = False
    Printer.CurrentY = MARGEN_SUPERIOR
    Y = MARGEN_SUPERIOR
    Printer.FillColor = &HFFFFFF
    Printer.FillStyle = 0
    iAncho = VarCfg("ancho_marca")
    iAlto = VarCfg("alto_marca")
    
    ALTO_BAILE = 10 * iAlto
    ALTO_BAILE_FINAL = 16 * iAlto
    
    If sCodFase > 1 Then
        sInfo = mml_FRASE0608 & (sCodFase / 2) * 6
        If iMaxTandas > 1 Then
            sInfo = sInfo & ") - ("
            If lblTandasMenosDorsales.Caption = "" Then
                sInfo = sInfo & mml_FRASE0609 & tbDorsalesTanda.Text
            Else
                sInfo = sInfo & "Primeras rondas " & lblTandasMasDorsales.Caption & " dorsales, Siguientes " & lblTandasMenosDorsales.Caption & " dorsales"
            End If
            sInfo = sInfo & " ,Media/Ronda " & ((sCodFase / 2) * 6) \ iMaxTandas
            If ((sCodFase / 2) * 6) Mod iMaxTandas > 0 Then
                sInfo = sInfo & " faltando " & ((sCodFase / 2) * 6) \ iMaxTandas & " marcas"
            End If
        End If
        sInfo = sInfo & ") "
    End If
    
    Printer.FontName = "Arial"
    Printer.FontSize = 12
    Printer.CurrentY = MARGEN_SUPERIOR + 2 * iAlto
    Printer.CurrentX = 0
    Printer.Print tbDescComp.Text & " - " & tbDescCat.Text
    Printer.FontSize = 10
    Printer.Print sInfo & "    " & tbComen.Text
    
    ' Imprimimos la categoría
    X = 0
    Y = MARGEN_SUPERIOR + INIC_CAT * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0631 & sCodCat
    
    ' Imprimimos la fase
    X = 0
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0299 & sCodFase;
    If iRepesca = 1 Then
        Printer.Print mml_FRASE0146;
    End If
    Printer.Print mml_FRASE0632 & iHoja
    
    'Imprimimos los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos FROM juez_categ WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0421 & sIdJuez
    X = 4 * iAncho
    rs.Close
    
    ' Imprimimos la tanda
    X = 0
    Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0605 & iTanda & mml_FRASE0472 & iMaxTandas
    
    ' Imprimimos los bailes y dorsales
    Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase = " & IIf(sCodFase > 1, "2", "1") & " ORDER BY posicion", dbOpenSnapshot)
    iCBailes = 0
    iCBailes = 0
    If iHoja = 2 Then
        While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
            Inc iCBailes
            rsBailes.MoveNext
        Wend
    End If
    iCBailes = 0
    X = 0
    Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
    While Not rsBailes.EOF
        X = 0
        Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Line (X, Y)-(X + Printer.Width / 1.2, Y)
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Print rsBailes!Nombre
        X = 0
        Y = Y + 3 * iAlto
        
        If iMaxTandas > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
            Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & iRepesca & " AND d.fase =" & tbCodFase.Text & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
        Else
            Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        End If
        
        X = 4 * iAncho
        iCDorsales = 1
        While Not rsDorsales.EOF
            If iCDorsales >= iNumDorsal And iCDorsales < iNumDorsal + iMaxDorsalesTanda Then
                Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto + ALTO_BAILE * iCBailes
                Printer.CurrentX = X
                Printer.CurrentY = Y
                Printer.Print rsDorsales!num_dorsal
                Y = Y + 2 * iAlto
                ' puntuaciones
                Printer.Line (X, Y)-(X + iAncho, Y + iAlto * 2), , B
                X = X + iAncho * 2
            End If
            Inc iCDorsales
            rsDorsales.MoveNext
        Wend
        rsDorsales.Close
        
        Inc iCBailes
        rsBailes.MoveNext
    Wend
    ' {·M} Imprimimos el hueco para la firma
    Printer.CurrentX = 0
    Printer.CurrentY = MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 4) * 2
    iTamFuente = Printer.FontSize
    Printer.FontSize = 12
    Printer.FontSize = iTamFuente
    Printer.Print mml_FRASE0614
    Printer.Line (0, MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 3) * 2)-Step(iAncho * 6, iAlto * 5), 0, B
End Sub

'Se imprimen todos los bailes en una sola hoja o en múltiples hojas
Public Function ImprimirHojaPuntuacionesBaile(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer) As Integer
Dim rsBailes As Recordset
Dim iCBailes As Integer
Dim lOrden As Long
Dim iGrupo As Integer

   If sIdJuez = "Control" Or iHoja = 2 Or sCodFase = C_FASE_GENERAL_LOOK Then Exit Function
    
   iBuscarOrden sCodCat, sCodFase, iRepesca, lOrden, iGrupo
    
   ImprimirHojaPuntuacionesBaile = G_HOJA_IMPRESA
   iCBailes = 0
   If VarCfg("bailes_por_hoja_unica") = "S" And iMaxDorsales <= MAX_DORSALES_HOJA_UNICA Then
        ImprimirHojaPuntuacionesBaile = ImprimirHojaBailes_o_UnicaParaVariosGrupos(sCodComp, sCodCat, sCodFase, sIdJuez, iTanda, iNumDorsal, iMaxDorsalesTanda, iMaxTandas, iMaxDorsales, iRepesca, iHoja, lOrden, iGrupo)
   Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase = " & IIf(sCodFase > 1, "2", "1") & " ORDER BY posicion", dbOpenSnapshot)
        'If iHoja = 2 Then
        '    While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
        '        Inc iCBailes
        '        rsBailes.MoveNext
        '    Wend
        'End If
        iCBailes = 0
        While Not rsBailes.EOF
            ImprimirHojaBaile sCodComp, sCodCat, sCodFase, sIdJuez, iTanda, iNumDorsal, iMaxDorsalesTanda, iMaxTandas, iMaxDorsales, iRepesca, iHoja, rsBailes!codigo, lOrden, iGrupo
            Inc iCBailes
            rsBailes.MoveNext
            If Not rsBailes.EOF Then Printer.NewPage
        Wend
   End If
End Function
'{*M Country}
'Imprime en una sola hoja un solo baile para todos los grupos que bailen juntos
Public Function ImprimirHojaBailes_o_UnicaParaVariosGrupos(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer, Optional lOrden As Long = -1, Optional iGrupo As Integer = -1) As Integer
Dim rsCateg As Recordset
Dim i As Integer
Dim sSQL As String

    ImprimirHojaBailes_o_UnicaParaVariosGrupos = G_HOJA_IMPRESA
    'Identificamos si tenemos activo el modo de country para impresión de múltiples grupos en la misma hoja
    If C_COUNTRY And G_COUNTRY_COMB_GRUP_HOJAS Then
        'Solo debemos imprimir las categorías que son cabecera de salida a pista, ya que con cada una de ellas se imprimen las hojas de las demás que salen a pista conjuntamente
        'Comprobamos si la categoría es cabecera de salida a pista buscando los grupos que salen a bailar juntos
        'como mucho se estima que saldrán 20 grupos juntos a pista
        sSQL = "SELECT TOP 20 h.* FROM horario h WHERE h.cod_competicion = " & sCodComp & " AND h.orden >= " & _
                "(SELECT h2.orden FROM horario h2 WHERE h2.cod_competicion = " & sCodComp & " AND h2.cod_categoria = " & sCodCat & _
                " AND h2.numfase = " & sCodFase & " AND h2.repesca = " & iRepesca & " AND h2.orden = " & lOrden & " ) ORDER BY h.orden"
        Debug.Print sSQL
        Set rsCateg = db.OpenRecordset(sSQL, dbOpenSnapshot)
        If Not rsCateg.EOF Then
            'Comprobamos si hay que imprimir la categoría porque es cabecera de grupo
            If rsCateg.Fields("inicio_grupo") = 1 Then
                'Contamos las categorias que salen juntas
                iNumCatGrupo = 1
                aiOrdenCateg(iNumCatGrupo - 1).iFase = rsCateg.Fields("numfase")
                aiOrdenCateg(iNumCatGrupo - 1).iCodCateg = rsCateg.Fields("cod_categoria")
                aiOrdenCateg(iNumCatGrupo - 1).iRepesca = rsCateg.Fields("repesca")
                aiOrdenCateg(iNumCatGrupo - 1).sDescCateg = rsCateg.Fields("grupo")
                rsCateg.MoveNext
                Do While Not rsCateg.EOF
                    If rsCateg.Fields("inicio_grupo") = 0 Then
                        Inc iNumCatGrupo
                        aiOrdenCateg(iNumCatGrupo - 1).iFase = rsCateg.Fields("numfase")
                        aiOrdenCateg(iNumCatGrupo - 1).iCodCateg = rsCateg.Fields("cod_categoria")
                        aiOrdenCateg(iNumCatGrupo - 1).iRepesca = rsCateg.Fields("repesca")
                        aiOrdenCateg(iNumCatGrupo - 1).sDescCateg = rsCateg.Fields("grupo")
                    Else
                        Exit Do
                    End If
                    rsCateg.MoveNext
                Loop
                If iNumCatGrupo = 1 Then
                    'En este grupo solo hay una categoría por lo que pasamos a modo normal de todos los bailes en una sola hoja
                    ImprimirHojaTodosBailesUnaHoja sCodComp, sCodCat, sCodFase, sIdJuez, iTanda, iNumDorsal, iMaxDorsalesTanda, iMaxTandas, iMaxDorsales, iRepesca, iHoja, lOrden, iGrupo
                Else
                    Dim rsBailes As Recordset
                    
                    If G_COUNTRY_IMP_TODOS_BAILES_JUNTOS Then
                        ImprimirHojasGrupoCategsTodosLosBailes sCodComp, sCodFase, sIdJuez, iTanda, iNumDorsal, iMaxDorsalesTanda, iMaxTandas, iMaxDorsales, iRepesca, iHoja, lOrden, iGrupo
                    Else
                        Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion " & _
                                            "FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & _
                                            tbCodCat.Text & " AND bc.fase = " & IIf(sCodFase > 1, "2", "1") & " ORDER BY posicion", dbOpenSnapshot)
                        While Not rsBailes.EOF
                            'Imprimimos todos los grupos por cada baile
                            ImprimirHojasGrupoCategs sCodComp, sCodFase, sIdJuez, iTanda, iNumDorsal, iMaxDorsalesTanda, iMaxTandas, iMaxDorsales, iRepesca, iHoja, rsBailes!codigo, lOrden, iGrupo
                            rsBailes.MoveNext
                            If Not rsBailes.EOF Then Printer.NewPage
                        Wend
                    End If
                End If
            Else
                'No imprimimos la categoría porque no es cabecera de grupo
                ImprimirHojaBailes_o_UnicaParaVariosGrupos = G_HOJA_NO_IMPRESA
                Exit Function
            End If
        End If
        
    Else ' Modo normal de todos los bailes en la misma hoja por cada juez
        ImprimirHojaTodosBailesUnaHoja sCodComp, sCodCat, sCodFase, sIdJuez, iTanda, iNumDorsal, iMaxDorsalesTanda, iMaxTandas, iMaxDorsales, iRepesca, iHoja, lOrden, iGrupo
    End If

End Function
'Imprime todas las categorias del grupo con todos los bailes
Public Sub ImprimirHojasGrupoCategsTodosLosBailes(sCodComp As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer, Optional lOrden As Long = -1, Optional iGrupo As Integer = -1)
Dim rs As Recordset
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim X As Integer, Y As Integer
Dim i As Integer, j As Integer, k As Integer
Dim iAncho As Integer, iAlto As Integer
Dim iCCateg As Integer
Dim iCBailes As Integer
Dim iCDorsales As Integer
Dim bJuezPasos As Boolean
Dim sInfo As String
Dim iMaxDorsalesPorTanda As Integer
Dim iTamFuente As Integer
Dim iPos As Integer
Dim iColTabla As Integer
Dim iMaxCol As Integer
Dim iMaxFila As Integer
Dim iCodBaile As Integer
Dim sCodCat As Integer

Const MAX_COL = 28
Const MAX_FILA = 80
Const INIC_CAT = 7
Const INIC_FASE = 8
Const INIC_JUEZ = 5
Const INIC_TANDA = 12
Const INIC_BAILES = 14
Dim ALTO_BAILE As Integer
Dim ALTO_BAILE_FINAL As Integer

    
    bJuezPasos = False
    Printer.CurrentY = MARGEN_SUPERIOR
    Y = MARGEN_SUPERIOR
    Printer.FillColor = &HFFFFFF
    Printer.FillStyle = 0
    iAncho = VarCfg("ancho_marca")
    iAlto = VarCfg("alto_marca")
    
    ALTO_BAILE = 10 * iAlto
    ALTO_BAILE_FINAL = 16 * iAlto
    
    If sCodFase > 1 Then
        sInfo = mml_FRASE0608 & (sCodFase / 2) * 6
        If iMaxTandas > 1 Then
            sInfo = sInfo & ") - ("
            If lblTandasMenosDorsales.Caption = "" Then
                sInfo = sInfo & mml_FRASE0609 & tbDorsalesTanda.Text
            Else
                sInfo = sInfo & "Primeras rondas " & lblTandasMasDorsales.Caption & " dorsales, Siguientes " & lblTandasMenosDorsales.Caption & " dorsales"
            End If
            sInfo = sInfo & " ,Media/Ronda " & ((sCodFase / 2) * 6) \ iMaxTandas
            If ((sCodFase / 2) * 6) Mod iMaxTandas > 0 Then
                sInfo = sInfo & " faltando " & ((sCodFase / 2) * 6) \ iMaxTandas & " marcas"
            End If
        End If
        sInfo = sInfo & ") "
    End If
    
    Printer.FontName = "Arial"
    Printer.FontSize = 14
    Printer.FontBold = True
    'Printer.Print tbDescCat.Text;
    Printer.Print tbDescComp.Text
    Printer.CurrentY = MARGEN_SUPERIOR + 2 * iAlto
    Printer.FontBold = False
    Printer.FontSize = 12
    Printer.CurrentX = 0
    Printer.FontSize = 10
    Printer.Print sInfo & "    " & tbComen.Text;
    Printer.Print
      
    'Imprimimos los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos, nombre FROM juez_categ jc, jueces j WHERE jc.cod_juez = j.codigo AND cod_categoria = " & tbCodCat.Text & " AND id_juez = '" & sIdJuez & "' ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0421 & " " & sIdJuez & "               " & rs!Nombre;
    X = 4 * iAncho
    rs.Close
    
    ' Imprimimos la fase
    Printer.FontBold = True
    Printer.Print "     (" & mml_FRASE0299 & sDescFase(aiOrdenCateg(i).iFase);
    If iRepesca = 1 Then
        Printer.Print mml_FRASE0146;
    End If
    Printer.Print IIf(iGrupo > 0, " - Group " & iGrupo & ")", ")")
    
    ' Imprimimos la tanda
    If iMaxTandas > 1 Then
        MsgBox "No es posible utilizar la impresión de múltiples categorias por hoja con múltiples tandas", vbOKOnly Or vbCritical, "ERROR"
        X = 0
        Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Print mml_FRASE0617 & iTanda & mml_FRASE0472 & iMaxTandas
    End If
    
    ' {·M} Imprimimos el hueco para la firma
    Dim iAnchoFirma As Integer
    iAnchoFirma = iAncho * 7
    Printer.CurrentX = Printer.Width - iAnchoFirma
    Printer.CurrentY = MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 4) * 2
    iTamFuente = Printer.FontSize
    Printer.FontSize = 12
    Printer.Print mml_FRASE0614
    Printer.Line (Printer.Width - iAnchoFirma, MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 3) * 2)-Step(iAnchoFirma, iAlto * 7), 0, B
    
    ' Iniciamos la posición
    X = 0
    Y = MARGEN_SUPERIOR + INIC_CAT * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    
    Printer.FontSize = 10
    Printer.FontBold = True
    
    Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion " & _
                        "FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & _
                        tbCodCat.Text & " AND bc.fase = " & IIf(sCodFase > 1, "2", "1") & " ORDER BY posicion", dbOpenSnapshot)
        
    While Not rsBailes.EOF
    ' Imprimimos el baile
    iCodBaile = rsBailes!codigo
    X = 0
    Printer.FontSize = 10
    Printer.Print "     -----------------------------------------------------------  " & mml_FRASE0635 & rsBailes!Nombre & "  -----------------------------------------------------------"
    Printer.FontBold = True
    
        
        For i = 0 To iNumCatGrupo - 1
            sCodCat = aiOrdenCateg(i).iCodCateg
            Printer.FontBold = True
            Printer.Print "(" & sCodCat & ") -" & aiOrdenCateg(i).sDescCateg
            'Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto
            'Printer.CurrentY = Y
            If Not rsBailes.EOF Then
                ' Imprimimos el baile
                iCodBaile = rsBailes!codigo
                
                If iMaxTandas > 1 Then
                    MsgBox "ERROR: No se pueden imprimir hojas combinadas con varias tandas", vbOKOnly Or vbCritical
                Else
                    Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & aiOrdenCateg(i).iCodCateg & " AND repesca=" & aiOrdenCateg(i).iRepesca & " AND fase =" & aiOrdenCateg(i).iFase & " ORDER BY 1", dbOpenSnapshot)
                End If
                
                Dim iNumDorsales As Integer
                
                If Not rsDorsales.EOF Then
                    rsDorsales.MoveLast
                    rsDorsales.MoveFirst
                    'En caso de imprimir múltiples grupos no puede haber múltiples tandas
                    iMaxDorsalesTanda = rsDorsales.RecordCount
                End If
                ReDim aTabla(20, 40)
                X = 4 * iAncho
                
                Dim iFilaEspaciado As Integer
                Dim iFilaTabla As Integer
                
                iCDorsales = 0
                iPos = 0
                iFilaTabla = 0
                
                iFilaEspaciado = 2
                
                If iMaxDorsalesTanda > C_MAX_DORSALES_ESPACIADOS Then
                    iFilaEspaciado = 1
                End If
                
                While Not rsDorsales.EOF
                    Inc iCDorsales
                    If iCDorsales >= iNumDorsal And iCDorsales < iNumDorsal + iMaxDorsalesTanda Then
                        aTabla(iFilaTabla, iPos) = rsDorsales!num_dorsal
                        Inc iPos
                    End If
                    If iPos > iMaxCol Then iMaxCol = iPos
                    rsDorsales.MoveNext
                    If Not rsDorsales.EOF And iPos >= C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA Then
                        iMaxCol = iPos
                        iFilaTabla = iFilaTabla + iFilaEspaciado
                        iPos = 0
                    End If
                Wend
                Inc iFilaTabla
                If iFilaEspaciado = 2 Then Inc iFilaTabla
                rsDorsales.Close
                
                ' Impresión de fases siguientes sin dorsales
                If iPos = 0 And iColTabla = 0 Then
                    If sCodFase = 1 Then
                        iMaxCol = 8
                        iFilaTabla = 2
                        Dim i1 As Integer
                        For i1 = 0 To 7
                            aTabla(0, i1) = i1 + 1 & "º"
                        Next
                    Else
                        iMaxCol = C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA
                        iFilaTabla = C_MAX_DORSALES_BAILE_POR_COL_HOJA_UNICA
                    End If
                End If
                Printer.FontSize = 10
                'X = C_MIN_X
                X = Printer.Width / 2 - Printer.TextWidth("|9999999|") * iMaxCol / 2
                DibujarTabla Printer, X, Printer.CurrentY, iFilaTabla, iMaxCol, 900, 300, False
            End If
            
            Printer.FontSize = iTamFuente
        Next
        
        rsBailes.MoveNext
    Wend
End Sub
'Imprime en una sola hoja un solo baile para todos los grupos que bailen juntos
'aiOrdenCateg() As sCategHorario es una variable global del módulo
Public Sub ImprimirHojasGrupoCategs(sCodComp As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer, iBaile As Integer, Optional lOrden As Long = -1, Optional iGrupo As Integer = -1)
Dim rs As Recordset
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim X As Integer, Y As Integer
Dim i As Integer, j As Integer, k As Integer
Dim iAncho As Integer, iAlto As Integer
Dim iCCateg As Integer
Dim iCBailes As Integer
Dim iCDorsales As Integer
Dim bJuezPasos As Boolean
Dim sInfo As String
Dim iMaxDorsalesPorTanda As Integer
Dim iTamFuente As Integer
Dim iPos As Integer
Dim iColTabla As Integer
Dim iMaxCol As Integer
Dim iMaxFila As Integer
Dim iCodBaile As Integer
Dim sCodCat As Integer

Const MAX_COL = 28
Const MAX_FILA = 80
Const INIC_CAT = 7
Const INIC_FASE = 8
Const INIC_JUEZ = 5
Const INIC_TANDA = 12
Const INIC_BAILES = 14
Dim ALTO_BAILE As Integer
Dim ALTO_BAILE_FINAL As Integer

    
    bJuezPasos = False
    Printer.CurrentY = MARGEN_SUPERIOR
    Y = MARGEN_SUPERIOR
    Printer.FillColor = &HFFFFFF
    Printer.FillStyle = 0
    iAncho = VarCfg("ancho_marca")
    iAlto = VarCfg("alto_marca")
    
    ALTO_BAILE = 10 * iAlto
    ALTO_BAILE_FINAL = 16 * iAlto
    
    If sCodFase > 1 Then
        sInfo = mml_FRASE0608 & (sCodFase / 2) * 6
        If iMaxTandas > 1 Then
            sInfo = sInfo & ") - ("
            If lblTandasMenosDorsales.Caption = "" Then
                sInfo = sInfo & mml_FRASE0609 & tbDorsalesTanda.Text
            Else
                sInfo = sInfo & "Primeras rondas " & lblTandasMasDorsales.Caption & " dorsales, Siguientes " & lblTandasMenosDorsales.Caption & " dorsales"
            End If
            sInfo = sInfo & " ,Media/Ronda " & ((sCodFase / 2) * 6) \ iMaxTandas
            If ((sCodFase / 2) * 6) Mod iMaxTandas > 0 Then
                sInfo = sInfo & " faltando " & ((sCodFase / 2) * 6) \ iMaxTandas & " marcas"
            End If
        End If
        sInfo = sInfo & ") "
    End If
    
    Printer.FontName = "Arial"
    Printer.FontSize = 14
    Printer.FontBold = True
    'Printer.Print tbDescCat.Text;
    Printer.Print tbDescComp.Text
    Printer.CurrentY = MARGEN_SUPERIOR + 2 * iAlto
    Printer.FontBold = False
    Printer.FontSize = 12
    Printer.CurrentX = 0
    Printer.FontSize = 10
    Printer.Print sInfo & "    " & tbComen.Text;
    Printer.Print
      
    'Imprimimos los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos, nombre FROM juez_categ jc, jueces j WHERE jc.cod_juez = j.codigo AND cod_categoria = " & tbCodCat.Text & " AND id_juez = '" & sIdJuez & "' ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0421 & " " & sIdJuez & "               " & rs!Nombre
    X = 4 * iAncho
    rs.Close
    
    ' Imprimimos la tanda
    If iMaxTandas > 1 Then
        MsgBox "No es posible utilizar la impresión de múltiples categorias por hoja con múltiples tandas", vbOKOnly Or vbCritical, "ERROR"
        X = 0
        Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Print mml_FRASE0617 & iTanda & mml_FRASE0472 & iMaxTandas
    End If
    
    ' {·M} Imprimimos el hueco para la firma
    Printer.CurrentX = C_MIN_X
    Printer.CurrentY = MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 4) * 2
    iTamFuente = Printer.FontSize
    Printer.FontSize = 12
    Printer.Print mml_FRASE0614
    Printer.Line (C_MIN_X, MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 3) * 2)-Step(iAncho * 13, iAlto * 7), 0, B
    
    ' Iniciamos la posición
    X = 0
    Y = MARGEN_SUPERIOR + INIC_CAT * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    
    Printer.FontSize = 10
    Printer.FontBold = True
    
    For i = 0 To iNumCatGrupo - 1
        sCodCat = aiOrdenCateg(i).iCodCateg
    
        Printer.FontBold = True
        Printer.Print aiOrdenCateg(i).sDescCateg & "  -  "
        Printer.FontBold = False
        Printer.Print mml_FRASE0631 & sCodCat & " - " & mml_FRASE0632 & iHoja & IIf(lOrden > 0, "     Orden " & lOrden, "") _
                        & IIf(iGrupo > 0, "  Grupo " & iGrupo, "")
        Printer.Print
        
        ' Imprimimos la fase
        Printer.FontBold = True
        Printer.Print mml_FRASE0299 & sDescFase(aiOrdenCateg(i).iFase);
        If iRepesca = 1 Then
            Printer.Print mml_FRASE0146;
        End If
        Printer.Print
        
        'Imprimimos la tabla del baile especificafo en el parámetro iBaile
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & _
                                                      tbCodCat.Text & " AND bc.fase = " & IIf(sCodFase > 1, "2", "1") & " AND b.codigo = " & iBaile & " ORDER BY posicion", dbOpenSnapshot)
        
        'Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto
        'Printer.CurrentY = Y
        If Not rsBailes.EOF Then
            ' Imprimimos el baile
            iCodBaile = rsBailes!codigo
            
            X = 0
            Printer.CurrentX = X
            Printer.FontSize = 10
            Printer.Print mml_FRASE0635 & rsBailes!Nombre
            Printer.FontBold = True
            
            If iMaxTandas > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
                'Si hay más de una tanda no está autorizada la impresión conjunta de categorías
                Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & iRepesca & " AND d.fase =" & tbCodFase.Text & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
            Else
                Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & aiOrdenCateg(i).iCodCateg & " AND repesca=" & aiOrdenCateg(i).iRepesca & " AND fase =" & aiOrdenCateg(i).iFase & " ORDER BY 1", dbOpenSnapshot)
            End If
            
            Dim iNumDorsales As Integer
            
            If Not rsDorsales.EOF Then
                rsDorsales.MoveLast
                rsDorsales.MoveFirst
                'En caso de imprimir múltiples grupos no puede haber múltiples tandas
                iMaxDorsalesTanda = rsDorsales.RecordCount
            End If
            ReDim aTabla(20, 40)
            X = 4 * iAncho
            
            Dim iFilaEspaciado As Integer
            Dim iFilaTabla As Integer
            
            iCDorsales = 0
            iPos = 0
            iFilaTabla = 0
            
            iFilaEspaciado = 2
            
            If iMaxDorsalesTanda > C_MAX_DORSALES_ESPACIADOS Then
                iFilaEspaciado = 1
            End If
            
            While Not rsDorsales.EOF
                Inc iCDorsales
                If iCDorsales >= iNumDorsal And iCDorsales < iNumDorsal + iMaxDorsalesTanda Then
                    aTabla(iFilaTabla, iPos) = rsDorsales!num_dorsal
                    Inc iPos
                End If
                If iPos > iMaxCol Then iMaxCol = iPos
                rsDorsales.MoveNext
                If Not rsDorsales.EOF And iPos >= C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA Then
                    iMaxCol = iPos
                    iFilaTabla = iFilaTabla + iFilaEspaciado
                    iPos = 0
                End If
            Wend
            Inc iFilaTabla
            If iFilaEspaciado = 2 Then Inc iFilaTabla
            rsDorsales.Close
            
            ' Impresión de fases siguientes sin dorsales
            If iPos = 0 And iColTabla = 0 Then
                If sCodFase = 1 Then
                    iMaxCol = 8
                    iFilaTabla = 2
                    Dim i1 As Integer
                    For i1 = 0 To 7
                        aTabla(0, i1) = i1 + 1 & "º"
                    Next
                Else
                    iMaxCol = C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA
                    iFilaTabla = C_MAX_DORSALES_BAILE_POR_COL_HOJA_UNICA
                End If
            End If
            Printer.Print
            Printer.FontSize = 10
            'X = C_MIN_X
            X = Printer.Width / 2 - Printer.TextWidth("|9999999|") * iMaxCol / 2
            DibujarTabla Printer, X, Printer.CurrentY, iFilaTabla, iMaxCol, 900, 300, False
            Printer.Print
            rsBailes.MoveNext
        End If
        
        Printer.FontSize = iTamFuente
    Next
End Sub

Public Sub ImprimirHojaTodosBailesUnaHoja(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer, Optional lOrden As Long = -1, Optional iGrupo As Integer = -1)
Dim rs As Recordset
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim X As Integer, Y As Integer
Dim i As Integer, j As Integer, k As Integer
Dim iAncho As Integer, iAlto As Integer
Dim iCCateg As Integer
Dim iCBailes As Integer
Dim iCDorsales As Integer
Dim bJuezPasos As Boolean
Dim sInfo As String
Dim iMaxDorsalesPorTanda As Integer
Dim iTamFuente As Integer
Dim iPos As Integer
Dim iColTabla As Integer
Dim iMaxCol As Integer
Dim iMaxFila As Integer
Dim iCodBaile As Integer

Const MAX_COL = 28
Const MAX_FILA = 80
Const INIC_CAT = 6
Const INIC_FASE = 8
Const INIC_JUEZ = 10
Const INIC_TANDA = 12
Const INIC_BAILES = 14
Dim ALTO_BAILE As Integer
Dim ALTO_BAILE_FINAL As Integer

    
    bJuezPasos = False
    Printer.CurrentY = MARGEN_SUPERIOR
    Y = MARGEN_SUPERIOR
    Printer.FillColor = &HFFFFFF
    Printer.FillStyle = 0
    iAncho = VarCfg("ancho_marca")
    iAlto = VarCfg("alto_marca")
    
    ALTO_BAILE = 10 * iAlto
    ALTO_BAILE_FINAL = 16 * iAlto
    
    If sCodFase > 1 Then
        sInfo = mml_FRASE0608 & (sCodFase / 2) * 6
        If iMaxTandas > 1 Then
            sInfo = sInfo & ") - ("
            If lblTandasMenosDorsales.Caption = "" Then
                sInfo = sInfo & mml_FRASE0609 & tbDorsalesTanda.Text
            Else
                sInfo = sInfo & "Primeras rondas " & lblTandasMasDorsales.Caption & " dorsales, Siguientes " & lblTandasMenosDorsales.Caption & " dorsales"
            End If
            sInfo = sInfo & " ,Media/Ronda " & ((sCodFase / 2) * 6) \ iMaxTandas
            If ((sCodFase / 2) * 6) Mod iMaxTandas > 0 Then
                sInfo = sInfo & " faltando " & ((sCodFase / 2) * 6) \ iMaxTandas & " marcas"
            End If
        End If
        sInfo = sInfo & ") "
    End If
    
    Printer.FontName = "Arial"
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.Print tbDescCat.Text;
    Printer.FontBold = False
    Printer.FontSize = 12
    Printer.Print " - " & tbDescComp.Text
    Printer.CurrentY = MARGEN_SUPERIOR + 2 * iAlto
    Printer.CurrentX = 0
    Printer.FontSize = 10
    Printer.Print sInfo & "    " & tbComen.Text;
    Printer.Print
    
    ' Imprimimos las categorías
    X = 0
    Y = MARGEN_SUPERIOR + INIC_CAT * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0631 & sCodCat & mml_FRASE0632 & iHoja & IIf(lOrden > 0, "     Orden " & lOrden, "") _
                    & IIf(iGrupo > 0, "  Grupo " & iGrupo, "")
    
    ' Imprimimos la fase
    X = 0
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    Printer.FontBold = True
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0299 & sDescFase(sCodFase);
    If iRepesca = 1 Then
        Printer.Print mml_FRASE0146;
    End If
    'Printer.Print
    
    'Imprimimos los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos, nombre FROM juez_categ jc, jueces j WHERE jc.cod_juez = j.codigo AND cod_categoria = " & tbCodCat.Text & " AND id_juez = '" & sIdJuez & "' ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0421 & " " & sIdJuez & "               " & rs!Nombre
    X = 4 * iAncho
    rs.Close
    
    ' Imprimimos la tanda
    If iMaxTandas > 1 Then
        X = 0
        Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Print mml_FRASE0617 & iTanda & mml_FRASE0472 & iMaxTandas
    End If
    
    ' {·M} Imprimimos el hueco para la firma
    Printer.CurrentX = C_MIN_X
    Printer.CurrentY = MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 4) * 2
    iTamFuente = Printer.FontSize
    Printer.FontSize = 12
    Printer.Print mml_FRASE0614
    Printer.Line (C_MIN_X, MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 3) * 2)-Step(iAncho * 13, iAlto * 7), 0, B
    
    'Imprimimos las tablas
    Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase = " & IIf(sCodFase > 1, "2", "1") & " ORDER BY posicion", dbOpenSnapshot)
    
    Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto
    Printer.CurrentY = Y
    While Not rsBailes.EOF
        ' Imprimimos el baile
        iCodBaile = rsBailes!codigo
        
        X = 0
        Printer.CurrentX = X
        Printer.FontSize = 10
        Printer.Print mml_FRASE0635 & rsBailes!Nombre
        Printer.FontBold = True
        
        If iMaxTandas > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
            Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & iRepesca & " AND d.fase =" & tbCodFase.Text & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
        Else
            Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        End If
        
        Dim iNumDorsales As Integer
        
        If Not rsDorsales.EOF Then
            rsDorsales.MoveLast
            rsDorsales.MoveFirst
        End If
        ReDim aTabla(20, 40)
        X = 4 * iAncho
        
        Dim iFilaEspaciado As Integer
        Dim iFilaTabla As Integer
        
        iCDorsales = 0
        iPos = 0
        iFilaTabla = 0
        
        iFilaEspaciado = 2
        
        If iMaxDorsalesTanda > C_MAX_DORSALES_ESPACIADOS Then
            iFilaEspaciado = 1
        End If
        
        While Not rsDorsales.EOF
            Inc iCDorsales
            If iCDorsales >= iNumDorsal And iCDorsales < iNumDorsal + iMaxDorsalesTanda Then
                aTabla(iFilaTabla, iPos) = rsDorsales!num_dorsal
                Inc iPos
            End If
            If iPos > iMaxCol Then iMaxCol = iPos
            rsDorsales.MoveNext
            If Not rsDorsales.EOF And iPos >= C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA Then
                iMaxCol = iPos
                iFilaTabla = iFilaTabla + iFilaEspaciado
                iPos = 0
            End If
        Wend
        Inc iFilaTabla
        If iFilaEspaciado = 2 Then Inc iFilaTabla
        rsDorsales.Close
        
        ' Impresión de fases siguientes sin dorsales
        If iPos = 0 And iColTabla = 0 Then
            If sCodFase = 1 Then
                iMaxCol = 8
                iFilaTabla = 2
                For i = 0 To 7
                    aTabla(0, i) = i + 1 & "º"
                Next
            Else
                iMaxCol = C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA
                iFilaTabla = C_MAX_DORSALES_BAILE_POR_COL_HOJA_UNICA
            End If
        End If
        Printer.Print
        Printer.FontSize = 10
        'X = C_MIN_X
        X = Printer.Width / 2 - Printer.TextWidth("|9999999|") * iMaxCol / 2
        DibujarTabla Printer, X, Printer.CurrentY, iFilaTabla, iMaxCol, 900, 300, False
        Printer.Print
        rsBailes.MoveNext
    Wend
    
    Printer.FontSize = iTamFuente
End Sub
Public Sub ImprimirHojaBaile(sCodComp As Integer, sCodCat As Integer, sCodFase As Integer, sIdJuez As String, iTanda As Integer, iNumDorsal As Integer, iMaxDorsalesTanda As Integer, iMaxTandas As Integer, iMaxDorsales As Integer, iRepesca As Integer, iHoja As Integer, iCodBaile As Integer, Optional lOrden As Long = -1, Optional iGrupo As Integer = -1)
Dim rs As Recordset
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim X As Integer, Y As Integer
Dim i As Integer, j As Integer, k As Integer
Dim iAncho As Integer, iAlto As Integer
Dim iCCateg As Integer
Dim iCBailes As Integer
Dim iCDorsales As Integer
Dim bJuezPasos As Boolean
Dim sInfo As String
Dim iMaxDorsalesPorTanda As Integer
Dim iTamFuente As Integer
Dim iPos As Integer
Dim iColTabla As Integer
Dim iMaxCol As Integer
Dim iIncCol As Integer

Const MAX_COL = 28
Const MAX_FILA = 80
Const INIC_CAT = 6
Const INIC_FASE = 8
Const INIC_JUEZ = 10
Const INIC_TANDA = 12
Const INIC_BAILES = 16
Dim ALTO_BAILE As Integer
Dim ALTO_BAILE_FINAL As Integer

    
    iIncCol = 2
    bJuezPasos = False
    Printer.CurrentY = MARGEN_SUPERIOR
    Y = MARGEN_SUPERIOR
    Printer.FillColor = &HFFFFFF
    Printer.FillStyle = 0
    iAncho = VarCfg("ancho_marca")
    iAlto = VarCfg("alto_marca")
    
    ALTO_BAILE = 10 * iAlto
    ALTO_BAILE_FINAL = 16 * iAlto
    
    If sCodFase > 1 Then
        sInfo = mml_FRASE0608 & (sCodFase / 2) * 6
        If iMaxTandas > 1 Then
            sInfo = sInfo & ") - ("
            If lblTandasMenosDorsales.Caption = "" Then
                sInfo = sInfo & mml_FRASE0609 & tbDorsalesTanda.Text
            Else
                sInfo = sInfo & "Primeras rondas " & lblTandasMasDorsales.Caption & " dorsales, Siguientes " & lblTandasMenosDorsales.Caption & " dorsales"
            End If
            sInfo = sInfo & " ,Media/Ronda " & ((sCodFase / 2) * 6) \ iMaxTandas
            If ((sCodFase / 2) * 6) Mod iMaxTandas > 0 Then
                sInfo = sInfo & " faltando " & ((sCodFase / 2) * 6) \ iMaxTandas & " marcas"
            End If
        End If
        sInfo = sInfo & ") "
    End If
    
    Printer.FontName = "Arial"
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.Print tbDescCat.Text;
    Printer.FontBold = False
    Printer.FontSize = 12
    Printer.Print " - " & tbDescComp.Text
    Printer.CurrentY = MARGEN_SUPERIOR + 2 * iAlto
    Printer.CurrentX = 0
    Printer.FontSize = 10
    Printer.Print sInfo & "    " & tbComen.Text;
    Printer.Print
    
    ' Imprimimos las categorías
    X = 0
    Y = MARGEN_SUPERIOR + INIC_CAT * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0631 & sCodCat & mml_FRASE0632 & iHoja & IIf(lOrden > 0, "     Orden " & lOrden, "") _
                    & IIf(iGrupo > 0, "  Grupo " & iGrupo, "")
    
    ' Imprimimos la fase
    X = 0
    Y = MARGEN_SUPERIOR + INIC_FASE * iAlto
    Printer.FontBold = True
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0299 & sDescFase(sCodFase);
    If iRepesca = 1 Then
        Printer.Print mml_FRASE0146;
    End If
    Printer.Print
    
    'Imprimimos los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos, nombre FROM juez_categ jc, jueces j WHERE jc.cod_juez = j.codigo AND cod_categoria = " & tbCodCat.Text & " AND id_juez = '" & sIdJuez & "' ORDER BY 1", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_JUEZ * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0421 & sIdJuez & "               " & rs!Nombre
    X = 4 * iAncho
    rs.Close
    
    ' Imprimimos la tanda
    If iMaxTandas > 1 Then
        X = 0
        Y = MARGEN_SUPERIOR + INIC_TANDA * iAlto
        Printer.CurrentX = X
        Printer.CurrentY = Y
        Printer.Print mml_FRASE0617 & iTanda & mml_FRASE0472 & iMaxTandas
    End If
    
    ' Imprimimos el baile
    Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo and b.codigo = " & iCodBaile & " ORDER BY posicion", dbOpenSnapshot)
    X = 0
    Y = MARGEN_SUPERIOR + INIC_BAILES * iAlto
    Printer.CurrentX = X
    Printer.CurrentY = Y
    Printer.Print mml_FRASE0635 & rsBailes!Nombre & mml_FRASE0636 & sDescFase(sCodFase)
    Printer.FontBold = True
    
    If iMaxTandas > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
        Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & iRepesca & " AND d.fase =" & tbCodFase.Text & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
    Else
        Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & iRepesca & " AND fase =" & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
    End If
    
    If Not rsDorsales.EOF Then
        rsDorsales.MoveLast
        If rsDorsales.RecordCount > MAX_DORSLES_HOJA_PUNT_COL_DOBLE Then
            iIncCol = 1
        End If
        rsDorsales.MoveFirst
    End If
    ReDim aTabla(100, 32)
    X = 4 * iAncho
    iCDorsales = 0
    iPos = 0
    iColTabla = 0
    While Not rsDorsales.EOF
        Inc iCDorsales
        If iCDorsales >= iNumDorsal And iCDorsales < iNumDorsal + iMaxDorsalesTanda Then
            aTabla(iPos, iColTabla) = rsDorsales!num_dorsal
            Inc iPos
        End If
        If iPos > iMaxCol Then iMaxCol = iPos
        rsDorsales.MoveNext
        If Not rsDorsales.EOF And iPos > C_MAX_DORSALES_BAILE_POR_COL Then
            iMaxCol = iPos
            iColTabla = iColTabla + iIncCol
            iPos = 0
        End If
    Wend
    rsDorsales.Close
    
    ' Impresión de fases siguientes sin dorsales
    If iPos = 0 And iColTabla = 0 Then
        If sCodFase = 1 Then
            iMaxCol = 8
            iColTabla = 0
            For i = 0 To 7
                aTabla(i, 0) = i + 1 & "º"
            Next
        Else
            iMaxCol = C_MAX_DORSALES_BAILE_POR_COL
            iColTabla = C_MAX_COLS_HOJA_POR_BAILE - 2
        End If
    End If
    Printer.Print
    Printer.FontSize = 24
    X = Printer.Width / 2 - Printer.TextWidth("|999|") * (iColTabla + 2) / 2 '- C_CORRECCION_TABLA_POR_X
    DibujarTabla Printer, X, Printer.CurrentY, iMaxCol, iColTabla + 2, 1000, 600, False
    
    ' {·M} Imprimimos el hueco para la firma
    Printer.CurrentX = 0
    Printer.CurrentY = MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 4) * 2
    iTamFuente = Printer.FontSize
    Printer.FontSize = 12
    Printer.Print mml_FRASE0614
    Printer.Line (0, MARGEN_SUPERIOR + iAlto * (C_MAX_MARCAS_Y - 3) * 2)-Step(iAncho * 13, iAlto * 7), 0, B
    Printer.FontSize = iTamFuente
End Sub


Private Sub chkFinDoc_Click()
    If chkFinDoc.Value = 0 Then
        MsgBox "Si imprime en PDF debe cerrar la aplicación para poder abrirlo.", vbOKOnly Or vbInformation, mml_FRASE0084
    End If
End Sub

Private Sub chkRep_Click()
    tbTandas.Text = ""
    CubrirInfoTandas
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
            cmdTandas_Click
        End If
    End If
    Exit Sub
error:
    ProcesarError "cmdCombinar_Click"
End Sub

Private Sub cmdImprimir_Click()
Dim rsCat As Recordset
Dim rsDorsales As Recordset
Dim sFaltaCateg As String
Dim i As Integer

    
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0637, vbOKOnly Or vbExclamation, mml_FRASE0096
        Exit Sub
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    Me.Tag = mml_FRASE0464
    On Local Error GoTo Pcancelar1
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir1
Pcancelar1:
    Exit Sub
Pseguir1:
    
    ImpresionDirecta

End Sub
Public Sub ImpresionDirecta()
Dim rsCat As Recordset
Dim rsDorsales As Recordset
Dim sFaltaCateg As String
Dim i As Integer

    iBaileActivo = 0
    '******************************************************************************************
    If chkHojasHorarioPorJuez.Value = 1 Then
        'Imprimimos todas las hojas por juez y ordenadas por el horario
            sFaltaCateg = ""
            If cbJuez.Text <> "" Then
                If cbJuez.Text = "Control" Then
                    Set rsCat = db.OpenRecordset("SELECT h.orden, h.hora, h.cod_categoria, h.numfase, h.grupo, 'Control' as id_juez, h.cod_baile FROM horario h, juez_categ j WHERE j.id_juez = 'A' AND h.cod_categoria = j.cod_categoria AND h.cod_competicion = " & tbCodComp.Text & " ORDER BY id_juez,orden,hora", dbOpenSnapshot)
                Else
                    Set rsCat = db.OpenRecordset("SELECT h.orden, h.hora, h.cod_categoria, h.numfase, h.grupo, j.id_juez, h.cod_baile  FROM horario h, juez_categ j WHERE h.cod_categoria = j.cod_categoria AND h.cod_competicion = " & tbCodComp.Text & " AND id_juez = '" & cbJuez.Text & "' ORDER BY id_juez,orden,hora", dbOpenSnapshot)
                End If
            Else ' TODOS LOS JUECES
                If GenerarControl(Val(tbCodComp.Text)) And VarCfg("tipo_hoja_puntuaciones") <> "hoja_rec_por_baile" Then
                    Set rsCat = db.OpenRecordset("SELECT h.orden, h.hora, h.cod_categoria, h.numfase, h.grupo, j.id_juez, h.cod_baile  FROM horario h, juez_categ j WHERE h.cod_categoria = j.cod_categoria AND h.cod_competicion = " & tbCodComp.Text & " UNION " & _
                                        "SELECT h.orden, h.hora, h.cod_categoria, h.numfase, h.grupo, 'Control', h.cod_baile  FROM horario h, juez_categ j WHERE j.id_juez = 'A' AND h.cod_categoria = j.cod_categoria AND h.cod_competicion = " & tbCodComp.Text & " ORDER BY id_juez,orden,hora", dbOpenSnapshot)
                Else
                    Set rsCat = db.OpenRecordset("SELECT h.orden, h.hora, h.cod_categoria, h.numfase, h.grupo, j.id_juez, h.cod_baile  FROM horario h, juez_categ j WHERE h.cod_categoria = j.cod_categoria AND h.cod_competicion = " & tbCodComp.Text & " ORDER BY id_juez,orden,hora", dbOpenSnapshot)
                End If
            End If
            If rsCat.EOF Then
                MsgBox mml_FRASE0638, vbOKOnly Or vbExclamation, mml_FRASE0096
                rsCat.Close
                Exit Sub
            End If
            'Recorremos todas las categorías seleccionadas (Si hemos seleccionado por juez, con todos los jueces)
            While Not rsCat.EOF
                If InStr(rsCat!grupo, cbPista.List(cbPista.ListIndex)) > 0 Then
                    'Contamos los dorsales de la categoria
                    Set rsDorsales = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria =" & rsCat!cod_categoria & " AND fase = " & rsCat!numfase, dbOpenSnapshot)
                    If rsDorsales.Fields(0) = 0 And (chkImprimirCatVacias.Value = 0 Or rsCat!numfase = C_FASE_GENERAL_LOOK) Then
                        'La categoría es una eliminatoria y falta
                        If VarCfg("avisar_falta_categ") = "S" Then
                            MsgBox mml_FRASE0640 & rsCat!grupo & mml_FRASE0641 & sDescFase(rsCat!numfase) & mml_FRASE0642, vbOKOnly Or vbInformation, mml_FRASE0086
                        End If
                        If rsCat!numfase = C_FASE_GENERAL_LOOK And G_IMPRIMIR_AVISO_GENERAL_LOOK Then
                            sFaltaCateg = sFaltaCateg & mml_FRASE0640 & rsCat!grupo & mml_FRASE0641 & sDescFase(rsCat!numfase) & mml_FRASE0643 & Chr(13) & Chr$(10)
                        Else
                            sFaltaCateg = sFaltaCateg & mml_FRASE0640 & rsCat!grupo & mml_FRASE0641 & UCase(sDescFase(rsCat!numfase)) & mml_FRASE0644 & Chr(13) & Chr$(10)
                        End If
                        GoTo continuar
                    Else
                        If sFaltaCateg <> "" Then
                            If InStr(sFaltaCateg, mml_FRASE0645) = 0 Then Printer.FontBold = True
                            Printer.FontSize = 12
                            Printer.Print sFaltaCateg
                            Printer.FontBold = False
                            Printer.NewPage
                        End If
                        sFaltaCateg = ""
                    End If
                    rsDorsales.Close
                    tbCodCat.Text = rsCat!cod_categoria
                    tbCodCat.Refresh
                    tbDescCat.Text = rsCat!grupo
                    tbDescCat.Refresh
                    tbCodFase.Text = rsCat!numfase
                    tbCodFase.Refresh
                    tbTandas.Text = ""
                    CubrirInfoTandas
                    cbJuez.Text = rsCat!id_juez
                    cbJuez.Refresh
                    For i = 1 To Val(tbTandas.Text)
                    Dim iBaile As Integer
                        Me.Tag = mml_FRASE0464
                        tbTanda.Text = i
                        tbTanda.Refresh
                        ProcesarEventos
                        iBaile = IIf(rsCat!cod_baile >= 0, rsCat!cod_baile, 0)
                        ImprimirHojasPuntuaciones True, iBaile
                    Next
                End If ' Pista
continuar:
                    rsCat.MoveNext
            Wend
            rsCat.Close
            On Local Error Resume Next
            Printer.EndDoc
            MsgBox mml_FRASE0646, vbOKOnly Or vbInformation, mml_FRASE0086
    
    '******************************************************************************************
    ElseIf chkImpIniCat.Value = 1 Then
        Set rsCat = db.OpenRecordset("SELECT DISTINCT cod_categoria, fase, c.descripcion FROM categorias c, dorsales d WHERE c.codigo = cod_categoria AND cod_categoria IN ( SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & ") ORDER BY 1,2", dbOpenSnapshot)
        If rsCat.EOF Then
            MsgBox mml_FRASE0647, vbOKOnly Or vbExclamation, mml_FRASE0096
            rsCat.Close
            Exit Sub
        End If
        While Not rsCat.EOF
            If InStr(rsCat!grupo, cbPista.List(cbPista.ListIndex)) > 0 Then
                If rsCat!fase <> C_FASE_GENERAL_LOOK Then
                    tbCodCat.Text = rsCat!cod_categoria
                    tbDescCat.Text = rsCat!DESCRIPCION
                    tbCodFase.Text = rsCat!fase
                    Me.Tag = mml_FRASE0464
                    CubrirInfoTandas
                    ImprimirHojasPuntuaciones
                End If
            End If
            rsCat.MoveNext
        Wend
        rsCat.Close
        MsgBox mml_FRASE0648, vbOKOnly Or vbInformation, mml_FRASE0086
    '******************************************************************************************
    Else
        Dim rs As Recordset
        Dim iOrden As Integer
        ' Si estamos imprimiendo solo esta categoría y tenemos activadas las opciones de combinación
        ' iris generará una hoja por cada grupo de categorías y solo las imprimirá si tenemos seleccionada
        ' la categoría cabecera, con lo que debemos cambiar la categoría por la categoría cabecera del grupo
        If C_COUNTRY And G_COUNTRY_COMB_GRUP_HOJAS Then
            'Identificamos el grupo en el que está incluída la categoría
            Set rs = db.OpenRecordset("SELECT orden FROM horario WHERE cod_competicion = " & tbCodComp.Text & _
                " AND cod_categoria = " & tbCodCat.Text & " AND numfase = " & tbCodFase.Text, dbOpenSnapshot)
                If Not rs.EOF Then
                    iOrden = rs.Fields("orden")
                Else
                    Exit Sub
                End If
            rs.Close
            Set rs = db.OpenRecordset("SELECT * FROM horario WHERE cod_competicion = " & tbCodComp.Text & _
                " AND orden = (SELECT MAX(orden) FROM horario WHERE cod_competicion = " & tbCodComp.Text & _
                " AND orden <= " & iOrden & " AND inicio_grupo = 1)", dbOpenSnapshot)
            If Not rs.EOF Then
                If Val(tbCodCat.Text) <> rs.Fields("cod_categoria") Or _
                    Val(tbCodFase.Text) <> rs.Fields("numfase") Or _
                    chkRep.Value <> rs.Fields("repesca") Then
                    MsgBox mml_FRASE1264, vbOKOnly Or vbInformation, ""
                    tbCodCat.Text = rs.Fields("cod_categoria")
                    tbDescCat.Text = sDescCategoria(Val(tbCodCat.Text))
                    tbCodFase.Text = rs.Fields("numfase")
                    tbDescFase.Text = sDescFase(Val(tbCodFase.Text))
                    chkRep.Value = rs.Fields("repesca")
                End If
            End If
            rs.Close
        End If
        ImprimirHojasPuntuaciones False
    End If

End Sub

Private Sub ImprimirHojasPuntuaciones(Optional bCalcularTandas As Boolean = True, Optional iCodBaile As Integer = 0)
Dim rs As Recordset
Dim rsBailes As Recordset
Dim rsJueces As Recordset
Dim rsDorsales As Recordset
Dim iCTanda As Integer
Dim iCDorsales As Integer
Dim iDorsalesPorTanda As Integer
Dim sTipoHojas As String
Dim iMaxHojas As Integer
Dim i As Integer
Dim iPistas As Integer
Dim iCCopias As Integer
Dim rsBailesJuez As Recordset
Dim aPistas(10) As String
Dim iDorsalInicialTanda As Integer
Dim iCBailes As Integer
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    'Comprobación del tipo de hoja según la pista
    If Len(G_PISTAS_HOJAS_OPTICAS) > 0 Then
        iPistas = DividirCampo(G_PISTAS_HOJAS_OPTICAS, aPistas, ",")
        For i = 0 To iPistas - 1
            If InStr(tbDescCat.Text, aPistas(i)) > 0 Then
                AsignarParametro "tipo_hoja_puntuaciones", "hojas_rec_optico"
            Else
                AsignarParametro "tipo_hoja_puntuaciones", "hoja_rec_por_baile"
            End If
        Next
    End If
    
    
    If G_IMP_HOJAS_BAILE_EN_FINALES Then
        'Las finales deben ser hojas normales y las eliminatorias ópticas
        If Val(tbCodFase.Text) > 1 Then
            AsignarParametro "tipo_hoja_puntuaciones", "hojas_rec_optico"
        Else
            AsignarParametro "tipo_hoja_puntuaciones", "hoja_rec_por_baile"
        End If
    End If

    If Val(tbTandas.Text) = 0 Or tbTandas.Text = "" Or tbDorsalesTanda.Text = "" Or tbTotalDorsales.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If VarCfg("tipo_hoja_puntuaciones") = "hoja_rec_optico" And Val(tbDorsalesTanda.Text) > C_MAX_DORSALES_HOJA_OPTICA_EXT Then
        MsgBox mml_FRASE0649 & C_MAX_DORSALES_HOJA_OPTICA_EXT, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If

   If Not Me.Tag = mml_FRASE0464 Then ' Si se generan automáticamente no se pregunta por el cuadro de diálogo de impresión
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
    Else
        If cbJuez.Text = "" And tbTanda.Text = "" And bCalcularTandas Then
            CubrirInfoTandas
        End If
    End If
    
    'Genera la información para la recombinación de dorsales
    CombinarDorsales tbCodCat.Text, tbCodFase.Text, chkRep.Value, Val(tbCombinarTandas.Text), chkRecombinar.Value
    
    For iCCopias = 1 To CDialog.Copies
        ' Comprobamos si los bailes caben en una hoja
        Set rsBailes = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCat.Text & " AND fase= " & IIf(tbCodFase.Text = "1", 1, 2) & " ORDER BY 1", dbOpenSnapshot)
        If rsBailes.Fields(0) > C_BAILES_POR_HOJA Then
            iMaxHojas = 2
        Else
            iMaxHojas = 1
        End If
        rsBailes.Close
        sTipoHojas = VarCfg("tipo_hoja_puntuaciones")
'*************************************************************************************
'Imprimimos el juez y la tanda indicada
        If cbJuez.Text <> "" And tbTanda.Text <> "" Then
            If Val(tbTanda.Text) > Val(tbTandas.Text) Then
                MsgBox mml_FRASE0650, vbOKOnly Or vbInformation, mml_FRASE0096
                Exit Sub
            End If
'Imprimimos una sola hoja ***********************************************************
            iDorsalInicialTanda = CalcularDorsalInicialTandaCat(Val(tbCodCat.Text), Val(tbCodFase.Text), chkRep.Value, Val(tbTanda.Text), Val(tbTandas.Text), iDorsalesPorTanda)
            
            'Si hay reconocimiento parcial de bailes tenemos que imprimir varias hojas por juez
            If BailesParciales(Val(tbCodCat.Text)) Then
                If iCodBaile > 0 Then
                    Set rsBailesJuez = db.OpenRecordset("SELECT codigo FROM bailes b, bailes_categ bc WHERE bc.cod_baile = " & iCodBaile & " AND b.codigo = bc.cod_baile AND bc.cod_categoria = " & tbCodCat.Text & " AND fase = " & IIf(Val(tbCodFase.Text) = 1, 1, 2) & " ORDER by posicion", dbOpenSnapshot)
                Else
                    Set rsBailesJuez = db.OpenRecordset("SELECT codigo FROM bailes b, bailes_categ bc WHERE b.codigo = bc.cod_baile AND bc.cod_categoria = " & tbCodCat.Text & " AND fase = " & IIf(Val(tbCodFase.Text) = 1, 1, 2) & " ORDER by posicion", dbOpenSnapshot)
                End If
                If Not rsBailesJuez.EOF Then
                    rsBailesJuez.MoveLast
                    rsBailesJuez.MoveFirst
                End If
                iCBailes = 0
                While Not rsBailesJuez.EOF
                    iCBailes = iCBailes + 1
                    If sTipoHojas = "hoja_rec_optico" Then
                        i = IIf(iCBailes > C_BAILES_POR_HOJA, 2, 1)
                        ImprimirHojaPuntuaciones tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                                    cbJuez.Text, Val(tbTanda.Text), iDorsalInicialTanda, iDorsalesPorTanda, _
                                    Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i, rsBailesJuez!codigo
                    End If
                    rsBailesJuez.MoveNext
                    If Not rsBailesJuez.EOF Then Printer.NewPage
                Wend
                rsBailesJuez.Close
            Else ' Se imprimen todos los bailes
                If sTipoHojas = "hoja_rec_optico" Then
                    For i = 1 To iMaxHojas
                        ImprimirHojaPuntuaciones tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                                    cbJuez.Text, Val(tbTanda.Text), iDorsalInicialTanda, iDorsalesPorTanda, _
                                    Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i
                            'ImprimirHojaPuntuaciones tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                            '    cbJuez.Text, Val(tbTanda.Text), (Val(tbTanda.Text) - 1) * Val(tbDorsalesTanda) + 1, iDorsalesPorTanda, Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i
                        If i < iMaxHojas Then Printer.NewPage
                    Next
                ElseIf sTipoHojas = "hoja_rec_por_juez" Then
                    For i = 1 To iMaxHojas
                            ImprimirHojaPuntuacionesJuez tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                                cbJuez.Text, Val(tbTanda.Text), iDorsalInicialTanda, iDorsalesPorTanda, _
                                Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i
                            'ImprimirHojaPuntuacionesJuez tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                            '    cbJuez.Text, Val(tbTanda.Text), (Val(tbTanda.Text) - 1) * Val(tbDorsalesTanda) + 1, iDorsalesPorTanda, Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i
                        If i < iMaxHojas Then Printer.NewPage
                    Next
                ElseIf sTipoHojas = "hoja_rec_por_baile" Then ' Hojas normales (todos los bailes en una hoja, o en country con opción de uno con varios grupos)
                    Dim iHojaImpresa  As Integer
                    For i = 1 To iMaxHojas
                        iHojaImpresa = ImprimirHojaPuntuacionesBaile(tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                            cbJuez.Text, Val(tbTanda.Text), (Val(tbTanda.Text) - 1) * Val(tbDorsalesTanda) + 1, iDorsalesPorTanda, Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i)
                    Next
                End If
            End If
            If chkFinDoc.Value = 1 Then
                Printer.EndDoc
            Else
                If iHojaImpresa = G_HOJA_IMPRESA Then
                    Printer.NewPage
                End If
            End If
            
        '*************************************************************************************
        Else
            '*************************************************************************************
            'Imprimimos las hojas de todos los jueces de esta categoría
            
            ' Por cada juez debemos imprimir todas las hojas de las tandas ***********************************************************
            Set rsJueces = db.OpenRecordset("SELECT DISTINCT id_juez, pasos FROM juez_categ WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1", dbOpenSnapshot)
            While Not rsJueces.EOF
                iCDorsales = 1
                For iCTanda = 1 To Val(tbTandas.Text)
                
                    'Recuperamos los dorsales por tanda y el dorsal incial de la tanda
                    iDorsalInicialTanda = CalcularDorsalInicialTandaCat(Val(tbCodCat.Text), Val(tbCodFase.Text), chkRep.Value, iCTanda, Val(tbTandas.Text), iDorsalesPorTanda)
                    
                    If BailesParciales(Val(tbCodCat.Text)) And sTipoHojas = "hoja_rec_optico" Then
                    
                        If iCodBaile > 0 Then
                            Set rsBailesJuez = db.OpenRecordset("SELECT codigo FROM bailes b, bailes_categ bc WHERE bc.cod_baile = " & iCodBaile & " AND b.codigo = bc.cod_baile AND bc.cod_categoria = " & tbCodCat.Text & " AND fase = " & IIf(Val(tbCodFase.Text) = 1, 1, 2) & " ORDER by posicion", dbOpenSnapshot)
                        Else
                            Set rsBailesJuez = db.OpenRecordset("SELECT codigo FROM bailes b, bailes_categ bc WHERE b.codigo = bc.cod_baile AND bc.cod_categoria = " & tbCodCat.Text & " AND fase = " & IIf(Val(tbCodFase.Text) = 1, 1, 2) & " ORDER by posicion", dbOpenSnapshot)
                        End If
                        
                        If Not rsBailesJuez.EOF Then
                            rsBailesJuez.MoveLast
                            rsBailesJuez.MoveFirst
                        End If
                        iCBailes = 0
                        While Not rsBailesJuez.EOF
                            iCBailes = iCBailes + 1
                                i = IIf(iCBailes > C_BAILES_POR_HOJA, 2, 1)
                                ImprimirHojaPuntuaciones tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                                            rsJueces!id_juez, iCTanda, iCDorsales, iDorsalesPorTanda, _
                                            Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i, rsBailesJuez!codigo
                                Printer.NewPage
                            rsBailesJuez.MoveNext
                        Wend
                        rsBailesJuez.Close
                    Else ' Se imprimen todos los bailes
                        If sTipoHojas = "hoja_rec_optico" Then
                            For i = 1 To iMaxHojas
                                ImprimirHojaPuntuaciones tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                                            rsJueces!id_juez, iCTanda, iCDorsales, iDorsalesPorTanda, Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i
                                Printer.NewPage
                            Next
                        ElseIf sTipoHojas = "hoja_rec_por_juez" Then
                            For i = 1 To iMaxHojas
                                ImprimirHojaPuntuacionesJuez tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                                            rsJueces!id_juez, iCTanda, iCDorsales, iDorsalesPorTanda, Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i
                                Printer.NewPage
                            Next
                        ElseIf sTipoHojas = "hoja_rec_por_baile" Then
                            For i = 1 To iMaxHojas
                                ImprimirHojaPuntuacionesBaile tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                                            rsJueces!id_juez, iCTanda, iCDorsales, iDorsalesPorTanda, Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, i
                                Printer.NewPage
                            Next
                        End If
                    End If
                    
                    iCDorsales = iCDorsales + iDorsalesPorTanda
                Next
                rsJueces.MoveNext
            Wend
            rsJueces.Close
            ' {·M15} Imprimimos la hoja de control
            If sTipoHojas = "hoja_rec_optico" And GenerarControl(Val(tbCodComp.Text)) Then
                iCDorsales = 1
                For iCTanda = 1 To Val(tbTandas.Text)
                    ' Si la tanda es la última imprimimos hasta el final
                    iDorsalInicialTanda = CalcularDorsalInicialTandaCat(Val(tbCodCat.Text), Val(tbCodFase.Text), chkRep.Value, iCTanda, Val(tbTandas.Text), iDorsalesPorTanda)
                    
                    ImprimirHojaPuntuaciones tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, _
                                    "Control", iCTanda, iCDorsales, iDorsalesPorTanda, Val(tbTandas.Text), Val(tbTotalDorsales), chkRep.Value, 1
                    Printer.NewPage
                    iCDorsales = iCDorsales + iDorsalesPorTanda
                Next
            End If
        End If
    Next iCCopias
    If chkSigFases.Value = 1 Then
        If Val(tbCodFase.Text) > 1 Then
            tbCodFase.Text = Val(tbCodFase.Text) / 2
            Me.Tag = mml_FRASE0464
            tbDescFase.Text = ""
            tbTandas.Text = ""
            CubrirInfoTandas
            cmdImprimir.Enabled = False
            ImprimirHojasPuntuaciones
        End If
    End If
    If chkFinDoc.Value = 1 Then
        Printer.EndDoc
    End If
    Me.Tag = "Run"
    Exit Sub
error:
    ProcesarError "ImprimirHojasPuntuaciones"
End Sub

Private Sub cmdImpRondas_Click()

    If Not C_DEBUG Then On Local Error GoTo error
    With frmCambiarRonda
        .tbCodComp.Text = tbCodComp.Text
        .tbDescComp.Text = tbDescComp.Text
        .tbCodCat.Text = tbCodCat.Text
        .tbDescCat.Text = tbDescCat.Text
        .tbCodFase.Text = tbCodFase.Text
        .tbDescFase.Text = tbDescFase.Text
        .chkRep.Value = chkRep.Value
        .RecargarBailes
        .Show vbModal
    End With
    
    Exit Sub
error:
    ProcesarError "cmdImpRondas_Click"
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

Sub cmdTandas_Click()
Dim iCCopias As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    
    If tbCodCat.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    If Val(tbCodFase.Text) < 2 Then
        MsgBox mml_FRASE0651, vbOKOnly Or vbExclamation, mml_FRASE0096
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
        ImprimirTandas
    Next iCCopias
    Exit Sub
error:
    ProcesarError "cmdTandas_Click"
End Sub
Public Sub ImprimirTandas()
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rs As Recordset
Dim iLineas As Integer
Dim iEscala As Integer
Dim iMaxLineasPorPag As Integer
Dim iPag As Integer
Dim icFilasTabla As Integer
Dim aTabla() As String
Dim aDefCelda(3) As TCelda
Dim sCateg As String
Dim rsBailes As Recordset


        
        Select Case tbCodFase.Text
            Case 1:
                tbDescFase.Text = mml_FRASE0329
            Case 2:
                tbDescFase.Text = "SEMI-FINAL"
            Case Else
                tbDescFase.Text = "1/" & tbCodFase.Text & " FINAL"
        End Select
        
        iPag = 1
        icFilasTabla = 1
        ReDim aTabla(G_MAX_FILAS_POR_PAG + 4, 4)
        
        aTabla(0, 0) = mml_FRASE0654
        aTabla(0, 1) = mml_FRASE0655
        aTabla(0, 2) = mml_FRASE0656
        aTabla(0, 3) = mml_FRASE0657
        
        aDefCelda(0).Ancho = 1200
        aDefCelda(0).Justificado = eccentro
        aDefCelda(1).Ancho = 4200
        aDefCelda(1).Justificado = ecizquierda
        aDefCelda(2).Ancho = 4200
        aDefCelda(2).Justificado = ecizquierda
        aDefCelda(3).Ancho = 1000
        aDefCelda(3).Justificado = eccentro
                
        Set rs = db.OpenRecordset("SELECT descripcion FROM categorias WHERE codigo = " & tbCodCat.Text, dbOpenSnapshot)
        sCateg = rs!DESCRIPCION & " (" & tbDescFase.Text & ")"
        If chkRep.Value = 1 Then
            sCateg = sCateg & mml_FRASE0658
        End If
        rs.Close
        
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase = 2 ORDER BY posicion", dbOpenSnapshot)
        While Not rsBailes.EOF
            aTabla(icFilasTabla, 0) = ""
            aTabla(icFilasTabla, 1) = ".14b" & rsBailes!Nombre
            aTabla(icFilasTabla, 2) = ""
            aTabla(icFilasTabla, 3) = ""
            Inc icFilasTabla
            ' Ahora recuperamos a los participantes de cada ronda
            iLineas = 0
            Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden, p.nombre_hombre, p.nombre_mujer, dc.codigo FROM dorsales d,dorsalescombinados dc, parejas p WHERE p.codigo = d.cod_pareja AND d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & chkRep.Value & " AND d.fase =" & tbCodFase.Text & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
            While Not rsDorsales.EOF
                aTabla(icFilasTabla, 0) = ".12n" & rsDorsales!num_dorsal
                aTabla(icFilasTabla, 1) = ".12n" & rsDorsales!nombre_hombre
                aTabla(icFilasTabla, 2) = ".12n" & rsDorsales!nombre_mujer
                aTabla(icFilasTabla, 3) = ".12n" & rsDorsales!orden
                rsDorsales.MoveNext
                
                Inc icFilasTabla
                If icFilasTabla >= G_MAX_FILAS_POR_PAG - 3 Then
                    Printer.FontBold = False
                    Printer.FontSize = 10
                    Printer.Print mml_FRASE0659 & iPag
                    ImprimirCabecera sCateg
                    DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
                    Printer.NewPage
                    Inc iPag
                    icFilasTabla = 1
                End If
            Wend
            rsBailes.MoveNext
        Wend
        Printer.FontBold = False
        Printer.FontSize = 10
        Printer.Print mml_FRASE0659 & iPag
        ImprimirCabecera sCateg
        DibujarTablaA Printer, Printer.CurrentX + C_MARGEN_IZQ_TABLAA, Printer.CurrentY, icFilasTabla, 4, aTabla(), aDefCelda(), C_ALTO_CELDA_TABLAA
        Printer.EndDoc

End Sub
Sub ImprimirCabecera(sCateg As String)
Dim rs As Recordset
Dim X As Integer, Y As Integer

    Printer.PaintPicture frmMenu.picEPA.Picture, Printer.Width - frmMenu.picEPA.Width - G_MARGEN_EPA, G_MARGEN_EPA_Y
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
    Printer.FontBold = True
    Printer.FontSize = 13
    Centrado Printer, sCateg, Printer.Width
    SaltoLinea Printer, 4
    Printer.Line Step(Printer.Width / 20, 0)-Step(Printer.Width - (Printer.Width / 20) * 3, 0)
    Printer.DrawWidth = 1
    SaltoLinea Printer, 10
End Sub

Private Sub CommandButton1_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCat.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text, "descripcion", " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCat.Text = sResultado(2)
    tbCodFase.Text = ""
    tbDescFase.Text = ""

End Sub

Private Sub CommandButton2_Click()
    If tbCodCat.Text = "" Then
        MsgBox mml_FRASE0504, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & tbCodCat.Text)
    If tbCodFase.Text <> "" Then
        tbDescFase.Text = sDescFase(Val(tbCodFase.Text))
        cbJuez.Refresh
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        tbTandas.Text = ""
        CubrirInfoTandas
    End If
    cmdImprimir.Enabled = True
End Sub

Public Sub CubrirInfoTandas()
Dim rs As Recordset
Dim iTotalDorsales As Integer
Dim iMaxDorsalesPorTanda As Integer
Dim iDorsalesTanda As Integer
Dim iDorsalesTandaSol As Integer
Dim iTandasConMasDorsales As Integer
Dim iMaxTandas As Integer

    'Si pedimos combinar se genera la combinación para las tandas adecuadas
    If G_DORSALES_COMBINADOS And CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
        'Primero generamos las tandas para la recombinación
        iMaxTandas = Val(tbTandas.Text)
        iDorsalesTanda = CalcularDorsalesPorTandaCatExt(Val(tbCodCat.Text), Val(tbCodFase.Text), chkRep.Value, 1, iMaxTandas, iTandasConMasDorsales, iTotalDorsales)
        tbCombinarTandas.Text = iMaxTandas
    Else
        tbCombinarTandas.Text = ""
    End If
    If ComprobarSiTandaUnica Then
        tbTandas.Text = "1"
        If tbCombinarTandas.Text = "" Then tbCombinarTandas.Text = tbTandas.Text
    Else
        'tbTandas.Text = "0"
    End If
    
    iMaxTandas = Val(tbTandas.Text)
    iDorsalesTanda = CalcularDorsalesPorTandaCatExt(Val(tbCodCat.Text), Val(tbCodFase.Text), chkRep.Value, 1, iMaxTandas, iTandasConMasDorsales, iTotalDorsales)
    If Not ComprobarSiTandaUnica Then
        tbTandas.Text = iMaxTandas
        tbTandas.Refresh
    
        If tbCombinarTandas.Text = "" Then tbCombinarTandas.Text = tbTandas.Text
    End If
    tbTotalDorsales.Text = iTotalDorsales
    
    If iTandasConMasDorsales > 0 Then
        tbDorsalesTanda.Text = iDorsalesTanda
        lblTandasMasDorsales.Caption = iTandasConMasDorsales & " x " & iDorsalesTanda
        lblTandasMenosDorsales.Caption = Val(tbTandas.Text) - iTandasConMasDorsales & " x " & iDorsalesTanda - 1
    Else
        tbDorsalesTanda.Text = iDorsalesTanda
        lblTandasMasDorsales.Caption = Val(tbTandas.Text) & " x " & iDorsalesTanda
        lblTandasMenosDorsales.Caption = ""
    End If
    
    cbJuez.Clear
    Set rs = db.OpenRecordset("SELECT DISTINCT id_juez FROM juez_categ WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1", dbOpenSnapshot)
    While Not rs.EOF
        cbJuez.AddItem rs!id_juez
        rs.MoveNext
    Wend
    If GenerarControl(Val(tbCodComp.Text)) Then
        cbJuez.AddItem "Control"
    End If
    cbJuez.Refresh
    rs.Close
    
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    chkRecombinar.Value = 1
    
    Me.Tag = "Run"
    CargarPistas cbPista
End Sub

Private Sub tbCodCat_GotFocus()
    tbCodCat.SelStart = 0
    tbCodCat.SelLength = Len(tbCodCat.Text)
End Sub

Private Sub tbCodCat_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbCodCat_LostFocus()
    ComprobarCategyFase tbCodCat, tbDescCat, tbCodFase, tbDescFase

End Sub
Private Sub tbCodCat_Change()
    If CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
        chkRecombinar.Enabled = True
    Else
        chkRecombinar.Enabled = False
    End If
    

End Sub

Private Sub tbTandas_LostFocus()
    If tbTotalDorsales.Text <> "" Then
        CubrirInfoTandas
    End If
End Sub


Function ComprobarSiTandaUnica() As Boolean
Dim rs As Recordset

    If Not C_DEBUG Then On Local Error GoTo error
    
    ComprobarSiTandaUnica = False
    ' Si se utilizan hojas ópticas no es posible la tanda única
    If VarCfg("tipo_hoja_puntuaciones") = "hoja_rec_por_baile" Then
        ComprobarSiTandaUnica = ImpTandaUnicaCateg(Val(tbCodCat.Text))
    End If

    If ComprobarSiTandaUnica Then
        tbTandas.Text = "1"
    '        tbTandas_LostFocus
        DoEvents
    End If
    Exit Function
error:
    ProcesarError "ComprobarSiTandaUnica"
End Function


