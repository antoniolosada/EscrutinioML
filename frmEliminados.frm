VERSION 5.00
Begin VB.Form frmEliminados 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "mml_FRASE0584"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11460
      TabIndex        =   3
      Top             =   8220
      Width           =   3315
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
      Height          =   525
      Left            =   5760
      TabIndex        =   1
      Top             =   8220
      Width           =   3435
   End
   Begin VB.Frame Frame1 
      Height          =   8145
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   14775
      Begin VB.TextBox tbEliminados 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7875
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   180
         Width           =   14595
      End
   End
End
Attribute VB_Name = "frmEliminados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub cmdBorrar_Click()
    If Not C_DEBUG Then On Local Error GoTo error
    If MsgBox(mml_FRASE1184, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbYes Then
        Kill G_FICHERO_ELIMINADOS
        MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    End If
    Exit Sub
error:
    ProcesarError "cmdBorrar_Click"
End Sub

Private Sub cmdSalir_Click()
    Unload Me

End Sub

Private Sub Form_Load()
Dim iFile

    If Not C_DEBUG Then On Local Error GoTo error
    TraducirCadenas Me
    iFile = FreeFile
    Open G_FICHERO_ELIMINADOS For Input As #iFile
    While Not EOF(iFile)
        Line Input #iFile, sCad
        tbEliminados.Text = tbEliminados.Text & sCad & vbCrLf
    Wend
    Close #iFile
    Exit Sub
error:
    ProcesarError "Form_Load"
End Sub
