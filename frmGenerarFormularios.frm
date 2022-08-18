VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGenerarFormularios 
   Caption         =   "Generar Formularios"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox tbProceso 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   180
      TabIndex        =   1
      Top             =   390
      Width           =   4815
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar Formularios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   3765
   End
End
Attribute VB_Name = "frmGenerarFormularios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerar_Click()
Dim sFichero As String
Dim sPath As String
Dim i As Integer
    CD.FileName = "D:\Documentos\Proy\Escrutinio\EsrutinioML\frmEnlacePPC.frm"
    CD.InitDir = "D:\Documentos\Proy\Escrutinio\EsrutinioML\"
    
    CD.ShowSave
    sFichero = CD.FileName
    
    If sFichero <> "" Then
        sPath = sExtraerPath(sFichero)
        For i = 1 To 3
            tbProceso.Text = "Copiando Fichero frmPuntuacionesBaile" & i
            tbProceso.Refresh
            FileCopy sPath & "\frmAPuntuacionesBaile.frm", sPath & "\frmAPuntuacionesBaile" & Trim$(Str$(i)) & ".frm.BAK"
            FileCopy sPath & "\frmAPuntuacionesBaile.frx", sPath & "\frmAPuntuacionesBaile" & Trim$(Str$(i)) & ".frx"
            Sustituciones sPath & "\frmAPuntuacionesBaile" & Trim$(Str$(i)) & ".frm", "frmAPuntuacionesBaile", "frmAPuntuacionesBaile" & Trim$(Str$(i))
        Next
        For i = 1 To 3
            tbProceso.Text = "Copiando Fichero frmEnlacePPC" & i
            tbProceso.Refresh
            FileCopy sPath & "\frmEnlacePPC.frm", sPath & "\frmEnlacePPC" & Trim$(Str$(i)) & ".frm.BAK"
            FileCopy sPath & "\frmEnlacePPC.frx", sPath & "\frmEnlacePPC" & Trim$(Str$(i)) & ".frx"
            Sustituciones sPath & "\frmEnlacePPC" & Trim$(Str$(i)) & ".frm", "frmEnlacePPC", "frmEnlacePPC" & Trim$(Str$(i))
        Next
        For i = 1 To 3
            tbProceso.Text = "Copiando Fichero frmEnlacePPC_HTML" & i
            tbProceso.Refresh
            FileCopy sPath & "\frmEnlacePPC_HTML.frm", sPath & "\frmEnlacePPC_HTML" & Trim$(Str$(i)) & ".frm.BAK"
            FileCopy sPath & "\frmEnlacePPC_HTML.frx", sPath & "\frmEnlacePPC_HTML" & Trim$(Str$(i)) & ".frx"
            Sustituciones sPath & "\frmEnlacePPC_HTML" & Trim$(Str$(i)) & ".frm", "frmEnlacePPC_HTML", "frmEnlacePPC_HTML" & Trim$(Str$(i))
        Next
    End If
    tbProceso.Text = "Operación realizada"
    MsgBox "Operación realizada, Ahora debe abandonar el IDE de Visual Basic", vbOKOnly Or vbInformation, "MENSAJE"
    End
End Sub

Sub Sustituciones(sFichero As String, sCad1 As String, sCad2 As String)
Dim iFile As Integer, iFileS As Integer
Dim sCad As String
Dim i As Integer
    iFile = FreeFile
    Open sFichero & ".BAK" For Input As #iFile
    iFileS = FreeFile
    Open sFichero For Output As #iFileS
    While Not EOF(iFile)
        Line Input #iFile, sCad
        i = InStr(sCad, sCad1)
        While i > 0
            sCad = Mid$(sCad, 1, i - 1) & sCad2 & Mid$(sCad, i + Len(sCad1))
            i = InStr(i + Len(sCad1), sCad, sCad1)
        Wend
        Print #iFileS, sCad
    Wend
    Close #iFile
    Close #iFileS
End Sub
