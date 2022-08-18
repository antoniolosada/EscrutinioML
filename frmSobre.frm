VERSION 5.00
Begin VB.Form frmSobre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mml_FRASE0837"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   660
      Top             =   5790
   End
   Begin VB.CommandButton cmdInfoVersion 
      Caption         =   "Info Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9090
      TabIndex        =   10
      Top             =   5790
      Width           =   2655
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
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   5685
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   11685
      Begin VB.Label lblLic 
         Alignment       =   2  'Center
         Caption         =   "Licencia/License"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   4800
         Width           =   5295
      End
      Begin VB.Label lblLicencia 
         Alignment       =   2  'Center
         Caption         =   "mml_FRASE0020"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   5160
         Width           =   5235
      End
      Begin VB.Label lblFinLicencia 
         Alignment       =   2  'Center
         Caption         =   "mml_FRASE0020"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   3330
         TabIndex        =   7
         Top             =   4350
         Width           =   5235
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   $"frmSobre.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1425
         Left            =   270
         TabIndex        =   6
         Top             =   2190
         Width           =   10905
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "mml_FRASE0838"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3060
         TabIndex        =   5
         Top             =   1350
         Width           =   5295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "mml_FRASE0839"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3060
         TabIndex        =   4
         Top             =   990
         Width           =   5295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "mml_FRASE0840"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3060
         TabIndex        =   3
         Top             =   630
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "mml_FRASE0841"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3060
         TabIndex        =   2
         Top             =   270
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdInfoVersion_Click()
    On Local Error Resume Next
    frm1InfoVersion.Show vbModal
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim iBDLon As Integer
Dim sFecha As String * 255
    
    TraducirCadenas Me
    iBDLon = GetPrivateProfileString(mml_FRASE0033, "FinLicencia", "", sFecha, 255, G_PATH_ESCRUTINIO & "Licencia.ini")
    If iBDLon > 0 Then
        lblFinLicencia.Caption = mml_FRASE0965 & sFecha
    Else
        MsgBox mml_FRASE0966 & G_PATH_ESCRUTINIO, vbOKOnly Or vbInformation, mml_FRASE0084
    End If
    lblLicencia.Caption = Licencia
    
End Sub
Private Sub Lic()
Dim sLic As String

    sLic = frmMenu.ComprobarBase
    If IsDate(sLic) Then
        If CDate(sLic) > Now Then
            lblLic.ForeColor = vbBlack
        Else
            lblLic.ForeColor = vbBlue
        End If
    End If
End Sub

Public Function Licencia() As String
Dim lValor As Long, lLong As Long, sValor As String * 100
Dim sNumero As String, sNum1 As String, sNum2 As String
Dim ComputerInfo As cComputerInfo
Dim sCodigo As String
  
  Set ComputerInfo = New cComputerInfo

    Licencia = ""
    RegOpenKey HKEY_CURRENT_USER, "Applications\GenNumber" & Chr$(0), lValor
    If lValor = 0 Then Exit Function
    lLong = 80
    RegQueryValue lValor, "" & Chr$(0), sValor, lLong
    RegCloseKey lValor
    
    sNumero = Mid$(sValor, 1, InStr(sValor, Chr$(0)) - 1)
    
    sNum1 = Mid$(sNumero, 1, InStr(sNumero, "-") - 1)
    sNum2 = Mid$(sNumero, InStr(sNumero, "-") + 1)
    
    sNum1 = Val("&H" & sNum1) Xor 69
    sNum2 = Val("&H" & sNum2) Xor 69
    
    
    sCodigo = Trim$(Str$(Val(sNum1) Xor ComputerInfo.ProcessorRevision))
    
    lblFinLicencia.Caption = CDate("01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2))
    
    Licencia = Hex$(Val(sNum1)) & "-" & Hex$(Val(sNum2))
End Function

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Lic
End Sub
