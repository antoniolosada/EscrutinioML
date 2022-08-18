VERSION 5.00
Begin VB.Form frmBuscarPart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1028"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   5265
      Left            =   30
      TabIndex        =   7
      Top             =   1650
      Width           =   11715
      Begin VB.ListBox lstBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4860
         Left            =   150
         TabIndex        =   8
         Top             =   270
         Width           =   11415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1665
      Left            =   0
      TabIndex        =   3
      Top             =   -30
      Width           =   11745
      Begin VB.CommandButton cmdSalir 
         Caption         =   "mml_FRASE0029"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9420
         TabIndex        =   10
         Top             =   1110
         Width           =   1905
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "mml_FRASE1028"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9420
         TabIndex        =   9
         Top             =   300
         Width           =   1905
      End
      Begin VB.TextBox tbMNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2100
         MaxLength       =   100
         TabIndex        =   1
         Top             =   900
         Width           =   4515
      End
      Begin VB.TextBox tbHNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2100
         MaxLength       =   100
         TabIndex        =   0
         Top             =   420
         Width           =   4515
      End
      Begin VB.TextBox tbDorsal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8160
         TabIndex        =   2
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "mml_FRASE0300"
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
         Left            =   7050
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "mml_FRASE0384"
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
         Left            =   810
         TabIndex        =   5
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "mml_FRASE0389"
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
         Left            =   810
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmBuscarPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_iCodComp As Integer
Sub BuscarParticipantes(iCodComp As Integer)
    m_iCodComp = iCodComp
    
    Me.Show vbNomodal
End Sub

Private Sub cmdBuscar_Click()
Dim rs As Recordset
Dim sSelecDorsal As String
Dim sSQL As String

    If Val(tbDorsal.Text) > 0 Then
        sSelecDorsal = " AND d.num_dorsal = " & tbDorsal.Text
    Else
        sSelecDorsal = ""
    End If
    If Not C_DEBUG Then On Local Error GoTo error
        lstBuscar.Clear
        sSQL = "SELECT DISTINCT c.codigo, d.num_dorsal, nombre_hombre, nombre_mujer, c.descripcion FROM dorsales d, parejas p, categorias c WHERE c.codigo = d.cod_categoria AND d.cod_pareja = p.codigo AND p.nombre_hombre LIKE '*" & tbHNombre.Text & "*' AND p.nombre_mujer LIKE '*" & tbMNombre.Text & "*' AND c.cod_competicion = " & m_iCodComp & " " & sSelecDorsal & " ORDER BY c.descripcion"
        Debug.Print sSQL
        Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
        While Not rs.EOF
            lstBuscar.AddItem rs.Fields("codigo") & " " & mml_FRASE0300 & ": " & rs.Fields("num_dorsal") & " - " & rs.Fields("descripcion") & ", " & rs.Fields("nombre_hombre") & " - " & rs.Fields("nombre_mujer")
            rs.MoveNext
        Wend
        rs.Close
    Exit Sub
error:
    ProcesarError "cmdBuscar_Click"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
End Sub

Private Sub lstBuscar_Click()
Dim lCodCateg As Long
    If lstBuscar.ListIndex >= 0 Then
        lCodCateg = Val(lstBuscar.List(lstBuscar.ListIndex))
        If lCodCateg > 0 Then
            frmADorsales.tbCodCateg.Text = lCodCateg
            frmADorsales.tbDescCateg.Text = sDescCategoria(lCodCateg)
            On Local Error Resume Next
            EstablecerFase frmADorsales.cbFase, MinFaseCateg(frmADorsales.tbCodCateg.Text)
            frmADorsales.cmdActualizar_Click

            frmADorsales.Show vbNomodal
            Unload Me
        End If
    End If
End Sub

Private Sub tbDorsal_GotFocus()
    tbDorsal.SelStart = 0
    tbDorsal.SelLength = Len(tbDorsal.Text)

End Sub

Private Sub tbDorsal_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbHNombre_GotFocus()
    tbHNombre.SelStart = 0
    tbHNombre.SelLength = Len(tbHNombre.Text)
End Sub
Private Sub tbmNombre_GotFocus()
    tbMNombre.SelStart = 0
    tbMNombre.SelLength = Len(tbMNombre.Text)
End Sub

