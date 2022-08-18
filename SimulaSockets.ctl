VERSION 5.00
Begin VB.UserControl SimulaSockets 
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   975
   ScaleHeight     =   1230
   ScaleWidth      =   975
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   780
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   360
      Top             =   480
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "SimulaSockets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim Mensaje As String

Event DataArrival(ByVal n As Long)
Event Connect()
Event SendComplete()
Event Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Sub GetData(a As String)
   a = "200 " & Mensaje & vbCrLf
End Sub
Public Sub SendData(a As String)
   If Left(a, 1) = "." Then
      Timer1.Enabled = True
      Mensaje = "mensaje aceptado y enviado"
   End If
   If Left(a, 4) = "RCPT" Then
      Timer1.Enabled = True
      Mensaje = "Destinatario aceptado"
   End If
   If Left(a, 4) = "HELO" Then
      Timer1.Enabled = True
      Mensaje = "Bienvenido"
   End If
   If Left(a, 4) = "MAIL" Then
      Timer1.Enabled = True
      Mensaje = "Emisor aceptado"
   End If
   If Left(a, 4) = "DATA" Then
      Timer1.Enabled = True
      Mensaje = "Para finalizar pon un punto"
   End If
   Timer2.Enabled = True
End Sub

Public Sub Connect()
   Mensaje = "Conectado"
   Timer1.Enabled = True
   Timer3.Enabled = True
End Sub
Public Property Let Protocol(a As Integer)
End Property
Public Property Let RemotePort(a As Integer)
End Property
Public Property Let Remotehost(a As String)
End Property

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   RaiseEvent DataArrival(4 + Len(Mensaje))
End Sub
Private Sub Timer2_Timer()
   Timer2.Enabled = False
   RaiseEvent SendComplete
End Sub
Private Sub Timer3_Timer()
   Timer3.Enabled = False
   RaiseEvent Connect
End Sub

