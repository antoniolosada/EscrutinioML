VERSION 5.00
Begin VB.Form frmInsertarSeparadorFases 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnadirFases 
      Caption         =   "mml_FRASE0625"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.ListBox lstFasesHorario 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3195
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   240
         Width           =   9615
      End
   End
End
Attribute VB_Name = "frmInsertarSeparadorFases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cancelar As Integer

Private Sub cmdAnadirFases_Click()
    Cancelar = 0
    Hide
End Sub

Private Sub Form_Load()
Dim rs As Recordset
Dim i As Integer

    TraducirCadenas Me
    
    Cancelar = 1
    i = 0
    lstFasesHorario.Clear
    db.Execute ("DELETE FROM faseshorario WHERE inicio_sesion IS NULL")
    Set rs = db.OpenRecordset("SELECT * FROM faseshorario", dbOpenSnapshot)
    While Not rs.EOF
        lstFasesHorario.AddItem rs.Fields("des_fase") & "     [" & rs.Fields("duracion") & "min]"
        lstFasesHorario.ItemData(i) = rs.Fields("inicio_sesion")
        rs.MoveNext
    Wend
    rs.Close
End Sub

Public Function SeleccionarFasesHorario() As Integer
    Me.Show vbModal
    SeleccionarFasesHorario = Cancelar
End Function
