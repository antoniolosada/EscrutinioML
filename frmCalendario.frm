VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0530"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2580
   ClipControls    =   0   'False
   Icon            =   "frmCalendario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2580
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView mvCal 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   60948482
      CurrentDate     =   37613
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function Mostrar()
    TraducirCadenas Me
    mvCal.Value = Now
    Me.Show 1
    Mostrar = mvCal.Value
End Function

Private Sub mvCal_DblClick()
    Hide
End Sub
