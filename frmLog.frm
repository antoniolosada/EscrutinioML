VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0885"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   12795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "mml_FRASE0886"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4545
      TabIndex        =   1
      Top             =   5490
      Width           =   2850
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0887"
      Height          =   5370
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   12705
      Begin VB.TextBox tbLog 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Text            =   "frmLog.frx":0000
         Top             =   225
         Width           =   12570
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub VisualizarLog(sLog As String)
    tbLog.Text = sLog
    Me.Show
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub
