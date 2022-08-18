VERSION 5.00
Begin VB.Form frmPresentacion 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "frmPresentacion"
   Picture         =   "frmPresentacion.frx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3975
      Top             =   90
   End
End
Attribute VB_Name = "frmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
' Creamos la presentación
'    InitializeSurfaceCapture Me
'    CreateSurfacefromMask Me, 0
'    ReleaseSurfaceCapture Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    HookSurfaceHwnd Me
End Sub

Private Sub Timer1_Timer()
    AbrirBaseDeDatos
    
    Me.Hide
    Timer1.Enabled = False
    frmMenu.Show
End Sub
