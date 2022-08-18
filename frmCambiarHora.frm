VERSION 5.00
Begin VB.Form frmCambiarHora 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mml_FRASE0060"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "mml_FRASE0060"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   585
      Width           =   2070
   End
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
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   45
      Width           =   2070
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0060"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.TextBox tbDifHora 
         Alignment       =   2  'Center
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
         Left            =   1260
         MaxLength       =   5
         TabIndex        =   1
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "mml_FRASE0534"
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
         Left            =   180
         TabIndex        =   2
         Top             =   375
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCambiarHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalcular_Click()
Dim rs As Recordset
Dim sHora As String
Dim sComentarios As String

    db.Execute ("UPDATE cfg SET valor = '" & tbDifHora.Text & "' WHERE variable = 'dif_hora'")
    
    Set rs = db.OpenRecordset("SELECT * FROM publicar", dbOpenDynaset)
    While Not rs.EOF
        If Not IsNull(rs!comentarios) Then
            ' Localizada la hora, procedemos a su cambio
            sHora = Mid$(rs!comentarios, 1, 5)
            If IsDate(sHora) Then
                sComentarios = rs!comentarios
                rs.Edit
                rs!comentarios = Format$(DateAdd("n", Val(tbDifHora.Text), CDate(sHora)), "hh:mm") & Mid$(sComentarios, 6)
                rs.Update
            End If
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    ' No se cambia el horario, el programa HORARIO lo calcula dinámicamente
    'Set rs = db.OpenRecordset("SELECT * FROM horario", dbOpenDynaset)
    'While Not rs.EOF
    '    rs.Edit
    '    rs!hora = DateAdd("n", Val(tbDifHora.Text), rs!hora)
    '    rs.Update
    '    rs.MoveNext
    'Wend
    'rs.Close
    
    MsgBox mml_FRASE0535, vbOKOnly Or vbInformation, mml_FRASE0084
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbDifHora.Text = Val(VarCfg("dif_hora"))
End Sub
