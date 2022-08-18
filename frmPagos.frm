VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1023"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2670
      Picture         =   "frmPagos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   465
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
      Left            =   3960
      TabIndex        =   5
      Top             =   60
      Width           =   4935
   End
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
      Left            =   3120
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "mml_FRASE0886"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   8220
      Width           =   2355
   End
   Begin VB.Frame mrcMarco 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   10275
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "mml_FRASE0251"
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
         Left            =   4560
         TabIndex        =   11
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "mml_FRASE1028"
         Default         =   -1  'True
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
         Left            =   3240
         TabIndex        =   10
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox tbTexto 
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
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   3195
      End
      Begin VB.ComboBox cbOrden 
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
         Left            =   7620
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   2535
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgParejas 
         Height          =   6855
         Left            =   60
         TabIndex        =   3
         Top             =   600
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12091
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label25 
         Caption         =   "mml_FRASE0190"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Top             =   180
         Width           =   1245
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
      Left            =   960
      TabIndex        =   7
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBorrar_Click()
    tbTexto.Text = ""
    tbTexto.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    CargarDatos
    tbTexto.SetFocus
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    Exit Sub
error:
    ProcesarError "cmdSelCateg_Click"
End Sub

Private Sub dgParejas_DblClick()
Dim iPagado As Integer
    If Not C_DEBUG Then On Local Error GoTo error
    dgParejas.Col = 0
    If dgParejas.CellBackColor = vbGreen Then
        iPagado = 0
        dgParejas.CellBackColor = vbRed
    Else
        iPagado = 1
        dgParejas.CellBackColor = vbGreen
    End If
    db.Execute "UPDATE parejas SET pagado = " & iPagado & " WHERE codigo = " & dgParejas.Text
    Exit Sub
error:
    ProcesarError "dgParejas_DblClick"
End Sub

Private Sub Form_Load()
    If Not C_DEBUG Then On Local Error GoTo error
    TraducirCadenas Me
    dgParejas.Cols = 8
    dgParejas.ColWidth(0) = 800
    dgParejas.ColWidth(1) = 1000
    dgParejas.ColWidth(2) = 2200
    dgParejas.ColWidth(3) = 2200
    dgParejas.ColWidth(4) = 1200
    dgParejas.ColWidth(5) = 600
    dgParejas.ColWidth(6) = 1500
    dgParejas.ColWidth(7) = 1000
    
    cbOrden.AddItem mml_FRASE1024
    cbOrden.AddItem mml_FRASE1025
    cbOrden.AddItem mml_FRASE1026
    cbOrden.AddItem mml_FRASE1027
    cbOrden.ListIndex = 1
    
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    Exit Sub
error:
    ProcesarError "Form_Load"
End Sub

Sub CargarDatos()
Dim rs As Recordset
    If Not C_DEBUG Then On Local Error GoTo error
    If Val(tbCodComp.Text) = 0 Then Exit Sub
    
    Set rs = db.OpenRecordset("SELECT * FROM parejas WHERE (nif_hombre LIKE '*" & tbTexto.Text & "*' OR nombre_hombre LIKE '*" & tbTexto.Text & "*') AND cod_competicion = " & tbCodComp.Text & " ORDER BY " & Val(cbOrden.List(cbOrden.ListIndex)))
    dgParejas.Rows = 0
    mrcMarco.Visible = False
    While Not rs.EOF
        dgParejas.AddItem rs!codigo & vbTab & rs!nif_hombre & vbTab & rs!nombre_hombre & vbTab & rs!nombre_mujer & vbTab & sDescModalidad(rs!cod_modalidad) & vbTab & rs!categoria & vbTab & rs!escuelas & vbTab & rs!telefonos
        dgParejas.Row = dgParejas.Rows - 1
        If rs!pagado = 1 Then
            dgParejas.CellBackColor = vbGreen
        Else
            dgParejas.CellBackColor = vbRed
        End If
        rs.MoveNext
    Wend
    rs.Close
    mrcMarco.Visible = True
    mrcMarco.Refresh
error:
    ProcesarError "CargarDatos"
End Sub

Private Sub tbCodComp_Change()
    CargarDatos
End Sub
