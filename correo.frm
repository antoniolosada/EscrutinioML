VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form correo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0213"
   ClientHeight    =   7035
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSoloMovil 
      Caption         =   "mml_FRASE0214"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   3000
      Width           =   2565
   End
   Begin VB.FileListBox fileInet 
      Height          =   2235
      Left            =   3720
      TabIndex        =   21
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   480
      TabIndex        =   16
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdSelCateg 
         Appearance      =   0  'Flat
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
         Left            =   1920
         Picture         =   "correo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox tbCodComp 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox tbDescComp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox tbInfo 
         BackColor       =   &H00C0E0FF&
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
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   7935
      End
      Begin ComctlLib.ProgressBar pbarProgreso 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0215"
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdProcEnvios 
      Appearance      =   0  'Flat
      Caption         =   "mml_FRASE0216"
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
      Left            =   4650
      TabIndex        =   15
      Top             =   2640
      Width           =   2445
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
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
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Smtp 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   12
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox De 
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox Para 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   8
      Top             =   2220
      Width           =   3855
   End
   Begin MSWinsockLib.Winsock Sck 
      Left            =   450
      Top             =   6540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   25
   End
   Begin ComctlLib.ProgressBar Progreso 
      Height          =   615
      Left            =   4650
      TabIndex        =   7
      Top             =   1920
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   1085
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Retardo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   6360
   End
   Begin VB.TextBox MensajeEntrante 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Text            =   "correo.frx":046A
      Top             =   5520
      Width           =   8055
   End
   Begin VB.CommandButton Enviar 
      Appearance      =   0  'Flat
      Caption         =   "mml_FRASE0217"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7170
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox MensajeSaliente 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3660
      Width           =   8055
   End
   Begin VB.TextBox Asunto 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   3360
      Width           =   8055
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   6120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "mml_FRASE0218"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   13
      Top             =   2700
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "mml_FRASE0219"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   1980
      Width           =   300
   End
   Begin VB.Label dds 
      AutoSize        =   -1  'True
      Caption         =   "mml_FRASE0220"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label6 
      Caption         =   "mml_FRASE0221"
      Height          =   555
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "mml_FRASE0086"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "mml_FRASE0222"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3420
      Width           =   735
   End
End
Attribute VB_Name = "correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RetardoDesconexion As Boolean
Dim FicheroAnexo As String
Dim NFich As String
Dim bSalir As Boolean


'Private Sub Base64_Click()
'   Dim Caracter As String * 1
'   Dim Trio(3) As Integer
'   Dim Cont As Integer
'   Dim ContLinea As Integer
'   Dim Cuatro(4) As Integer
'   Dim Base64 As String
'
'   Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
'
'   ContLinea = 0
'   MensajeSaliente = ""
'   MensajeEntrante = ""
'   If FicheroAnexo <> "" Then
'      Open NFich For Binary As #3 Len = 3
'      Cont = 0
'      ContTotal = 0
'      Progreso.Max = FileLen(NFich)
'      While Not ContTotal = LOF(3)
'         ContTotal = ContTotal + 1
'         Caracter = Input(1, 3)
'         Cont = Cont + 1
'         Trio(Cont) = Asc(Caracter)
'         'MensajeSaliente = MensajeSaliente + Caracter
'         If Cont = 3 Then
'            Cuatro(1) = Int(Trio(1) / 4)
'            Cuatro(2) = (Trio(1) - Int(Trio(1) / 4) * 4) * 16 + Int(Trio(2) / 16)
'            Cuatro(3) = (Trio(2) - (Int(Trio(2) / 16) * 16)) * 4 + Int(Trio(3) / 64)
'            Cuatro(4) = Trio(3) - Int(Trio(3) / 64) * 64
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(1) + 1, 1)
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(2) + 1, 1)
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(3) + 1, 1)
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(4) + 1, 1)
'            Cont = 0
'            ContLinea = ContLinea + 4
'            If ContLinea = 76 Then
'               MensajeEntrante = MensajeEntrante + vbCrLf
'               ContLinea = 0
'            End If
'         End If
'         DoEvents
'      Wend
'      Select Case Cont
'         Case 1
'            Cuatro(1) = Int(Trio(1) / 4)
'            Cuatro(2) = (Trio(1) - Int(Trio(1) / 4) * 4) * 16
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(1) + 1, 1)
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(2) + 1, 1) + "=="
'         Case 2
'            Cuatro(1) = Int(Trio(1) / 4)
'            Cuatro(2) = (Trio(1) - Int(Trio(1) / 4) * 4) * 16 + Int(Trio(2) / 16)
'            Cuatro(3) = (Trio(2) - (Int(Trio(2) / 16) * 16)) * 4
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(1) + 1, 1)
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(2) + 1, 1)
'            MensajeEntrante = MensajeEntrante + Mid(Base64, Cuatro(3) + 1, 1) + "="
'         End Select
'      Close #3
'   End If
'End Sub


Private Sub cmdProcEnvios_Click()
Dim rs As Recordset
Dim rsDorsales As Recordset
Dim rsPos As Recordset
Dim sDirFichas As String
Dim i As Integer, j As Integer
Dim iCont As Integer
Dim aeMail(10) As String
Dim iConteMail As Integer

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0223, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    bSalir = False
    sDirFichas = VarCfg("dir_fichas")
    ' Seleccionamos todas las parejas que solicitaron el envío de la información por correo.
    Set rs = db.OpenRecordset("SELECT codigo, email, emailmovil, email_selec, emailmovil_selec, nombre_hombre,nombre_mujer FROM parejas WHERE (email_selec <> 0 OR emailmovil_selec <> 0) AND cod_competicion=" & tbCodComp.Text, dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        pbarProgreso.Max = rs.RecordCount
        rs.MoveFirst
    End If
    iCont = 1
    While Not rs.EOF
        pbarProgreso.Value = iCont
        iCont = iCont + 1
        ' Por cada pareja debemos localizar los dorsales de la pareja
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal, cod_categoria, descripcion FROM dorsales d, categorias c WHERE c.codigo = d.cod_categoria AND cod_pareja=" & rs!codigo, dbOpenSnapshot)
        While Not rsDorsales.EOF
            ' Por cada dorsal, comprobamos el grupo al que pertenece recopilamos la información de los ficheros
            ' generados para el grupo, y de toda la información del grupo
            ' Debemos enviar un correo por cada uno de los archivos
            If rs!email_selec <> 0 And Not chkSoloMovil.Value Then
                iConteMail = DividirCampo(rs!email, aeMail, ",")
                For j = 1 To iConteMail
                    fileInet.Path = sDirFichas & "\Inet\"
                    fileInet.Pattern = rsDorsales!DESCRIPCION & "*.*"
                    fileInet.Refresh
                    For i = 0 To fileInet.ListCount - 1
                        tbInfo.Text = mml_FRASE0224 & rsDorsales!num_dorsal & " - " & rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer
                        Para.Text = aeMail(j - 1)
                        Asunto.Text = mml_FRASE0225 & rsDorsales!num_dorsal & " - " & mml_FRASE0226 & fileInet.List(i)
                        MensajeSaliente.Text = mml_FRASE0227 + Chr$(13) + Chr$(10) + _
                                               mml_FRASE0225 & rsDorsales!num_dorsal & " - " & rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer + Chr$(13) + Chr$(10) + _
                                               mml_FRASE0228 + Chr$(13) + Chr$(10) + _
                                               "" + Chr$(13) + Chr$(10) + _
                                               mml_FRASE0229 + Chr$(13) + Chr$(10) + _
                                               mml_FRASE0230 + Chr$(13) + Chr$(10)
                        NFich = sDirFichas & "\Inet\" & fileInet.List(i)
                        FicheroAnexo = fileInet.List(i)
                        Call Enviar_Click
                        While Not Enviar.Enabled And Not bSalir
                            DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
                            Sleep 100
                        Wend
                    Next
                Next
            End If
            If rs!emailmovil_selec <> 0 Then
                'Localizamos la posición final
                Set rsPos = db.OpenRecordset("SELECT * FROM resumenfinales WHERE cod_categoria = " & rsDorsales!cod_categoria & " AND dorsal = " & rsDorsales!num_dorsal, dbOpenSnapshot)
                If Not rsPos.EOF Then
                    tbInfo.Text = mml_FRASE0231 & rsDorsales!num_dorsal & " - " & rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer
                    Para.Text = rs!email
                    Asunto.Text = mml_FRASE0225 & rsDorsales!num_dorsal & mml_FRASE0232 & rsPos!posicion & mml_FRASE0233 & rsPos!puntos
                    MensajeSaliente.Text = mml_FRASE0234 + Chr$(13) + Chr$(10) + _
                                           mml_FRASE0225 & rsDorsales!num_dorsal & " - " & rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer + Chr$(13) + Chr$(10) + _
                                           mml_FRASE0235 & rsPos!posicion & mml_FRASE0233 & rsPos!puntos + Chr$(13) + Chr$(10) + _
                                           mml_FRASE0236 + Chr$(13) + Chr$(10) + _
                                           mml_FRASE0237 + Chr$(13) + Chr$(10) + _
                                           "" + Chr$(13) + Chr$(10) + _
                                           mml_FRASE0229 + Chr$(13) + Chr$(10) + _
                                           mml_FRASE0230 + Chr$(13) + Chr$(10)
                    NFich = sDirFichas & "\Inet\" & fileInet.List(i)
                    FicheroAnexo = fileInet.List(i)
                    Call Enviar_Click
                    While Not Enviar.Enabled And Not bSalir
                        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
                        Sleep 100
                    Wend
                End If
            End If
            
            rsDorsales.MoveNext
        Wend
        rsDorsales.Close
        
        rs.MoveNext
    Wend
    rs.Close
    MsgBox mml_FRASE0238, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)

End Sub

Private Sub Command1_Click()
    bSalir = True
    Unload Me
End Sub

Private Sub Enviar_Click()

   ' Verifico que la configuracion esta completa
   If Smtp = "" Or De = "" Or Para = "" Then
      MsgBox (mml_FRASE0239)
   Else
      ' Inicializo el mensaje
      Cabecera(1) = "HELO Atawalabachala"
      Cabecera(2) = "MAIL FROM: <" & De.Text & ">"
      Cabecera(3) = "RCPT TO: <" & Para.Text & ">"
      Cabecera(4) = "DATA"
      Cabecera(5) = "From: " & De.Text
      Cabecera(6) = "To: " & Para.Text
      Cabecera(7) = "Subject: " & Asunto.Text
      Cabecera(8) = "DATE: "
      Cabecera(9) = "MIME-Version: 1.0"
      Cabecera(10) = "Content-Type: multipart/mixed;"
      Boundary = "----=_NextPart_000_0002_01BD22EE.C1291DA0"
      Cabecera(11) = Chr(9) & "boundary=""" & Boundary & """"
      Cabecera(12) = "X-Priority: 3"
      Cabecera(13) = "X-MSMail - Priority: Normal"
      Cabecera(14) = "X-MimeOLE: Producido por Raul Gimenez V1.0"
      Cabecera(15) = ""
      Cabecera(16) = mml_FRASE0240
      Cabecera(17) = ""
      Cabecera(18) = "--" & Boundary
      Cabecera(19) = "Content-Type: text/plain;"
      Cabecera(20) = Chr(9) & "charset=""x-user-defined"""
      Cabecera(21) = "Content-Transfer-Encoding: 8bit"
      Cabecera(22) = ""
      Cabecera(23) = MensajeSaliente.Text
      Cabecera(24) = ""
      InicioBoundaryAnexo = 24
      Cabecera(25) = "--" & Boundary
      Cabecera(26) = "Content-Type: application/octet-stream;"
      Cabecera(27) = Chr(9) & "Name=""" & FicheroAnexo & """"
      Cabecera(28) = "Content-Disposition: attachment;"
      Cabecera(29) = Chr(9) & "filename=""" & FicheroAnexo & """"
      Cabecera(30) = "Content-Transfer-Encoding: base64"
      InicioAnexo = 30
      Cabecera(31) = ""
      ' Inicializo la ventana de entrada
      MensajeEntrante = ""
      ' Inicializo la barra de progreso
      Progreso.Value = 0
      If FicheroAnexo = "" Then
         Progreso.Max = 500
      Else
         Progreso.Max = FileLen(NFich) * 1.4 + 500
      End If
      ' Inicializo el asunto y el mensaje saliente
      If MensajeSaliente = "" Then MensajeSaliente = mml_FRASE0241
      If Asunto = "" Then Asunto = mml_FRASE0242
      ' Anulo el boton de enviar
      Enviar.Enabled = False
      If Not ConexionEstablecida Then
         ' Si es la primera vez que envio un mensaje ,
         ' establezco la conexión.
         ' El mensaje se enviará en el evento "Connect"
         Sck.Protocol = sckTCPProtocol
         Sck.RemotePort = 25
         Sck.Remotehost = Smtp.Text
         Sck.Connect
      Else
         ' Si no es la primera vez, comienzo el envio del mensaje
         ' El resto del mensaje se enviará en el evento "SendComplete"
         ' NOTA: Me salto el comando HELO
         Paso = 2
         EnviarDatos (Cabecera(Paso) & vbCrLf)
         Progreso.Value = 10
      End If
   End If
End Sub

Private Sub Form_Load()
   TraducirCadenas Me
   
   ConexionEstablecida = False
   NFich = VarCfg("logo")
   FicheroAnexo = mml_FRASE0243
   
   Smtp.Text = VarCfg("servidor_smtp")
   De.Text = VarCfg("origen_mensaje")
End Sub

Private Sub Form_Terminate()
   Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim Resultado As Integer
   If Enviar.Enabled = False Then
      Resultado = MsgBox(mml_FRASE0244, vbYesNo)
      If Resultado = vbNo Then
         Cancel = -1
      Else
         'Sck.Close
         End
      End If
   End If
End Sub


Private Sub Salir_Click()
   End
End Sub

Private Sub Sck_Connect()
   ConexionEstablecida = True
   Paso = 0
End Sub

Private Sub Sck_DataArrival(ByVal bytesTotal As Long)
   Dim S As String
   Dim C As String * 1
   If Enviar.Enabled = False Then
      ' leo los mensajes que me llegan desde el servidor SMTP.
      Sck.GetData S
      MensajeEntrante.Text = MensajeEntrante.Text & S
      ' Inicializaciones varias
      ' Desactivo el timer para el control del TimeOut.
      TimeOut.Enabled = False
      ' Evaluo el mensaje del servidor SMTP
      C = Left(S, 1)
      If (C = 2 Or C = 3) Then
         Select Case Paso
            Case 0 To 900
               ' Envio la siguiente parte del mensaje
               Paso = Paso + 1
               Progreso.Value = Progreso.Value + 10
               EnviarDatos (Cabecera(Paso) & vbCrLf)
            Case 998, 999
               ' Esta es la respuesta al comando "."
               ' Espero un poco antes de desconectar
               Retardo.Enabled = True
               TimeOut.Enabled = False
               Progreso.Value = Progreso.Max
         End Select
      Else
         Paso = 999
         'Progreso.Value = Progreso.Value + 10
         EnviarDatos ("RSET" & vbCrLf)
      End If
   End If
End Sub

Private Sub Sck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   TimeOut.Enabled = False
   EnviarDatos ("RSET" & vbCrLf)
   MsgBox "(" & Number & ") " & Description
   'Sck.Close
   Enviar.Enabled = True
End Sub

Private Sub Sck_SendComplete()
   Dim Caracter As String * 1
   Dim Trio(3) As Integer
   Dim Cont As Integer
   Dim ContLinea As Integer
   Dim Cuatro(4) As Integer
   Dim Pos As Long
   Dim salir As Boolean
   Dim Base64 As String
   
   Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
   ' Desactivo el timer para el control del TimeOut.
   TimeOut.Enabled = False
   ' Envio la siguiente parte del mensaje
   Select Case Paso
      Case 5 To InicioBoundaryAnexo - 1
            Paso = Paso + 1
            Progreso.Value = Progreso.Value + 10
            EnviarDatos (Cabecera(Paso) & vbCrLf)
      Case InicioBoundaryAnexo To InicioAnexo - 1
         If FicheroAnexo <> "" Then
            Paso = Paso + 1
            Progreso.Value = Progreso.Value + 10
            EnviarDatos (Cabecera(Paso) & vbCrLf)
         Else
            ' Envio el mensaje de mml_FRASE0245
            Paso = 998
            Progreso.Value = Progreso.Value + 10
            EnviarDatos ("." & vbCrLf)
         End If
      Case InicioAnexo
         If FicheroAnexo <> "" Then
            Open NFich For Binary As #3
            ContTotal = 0
            Paso = Paso + 1
            Progreso.Value = Progreso.Value + 10
            EnviarDatos (Cabecera(Paso) & vbCrLf)
         Else
            ' Envio el mensaje de mml_FRASE0245
            Paso = 998
            Progreso.Value = Progreso.Value + 10
            EnviarDatos ("." & vbCrLf)
         End If
      Case InicioAnexo + 1
         salir = False
         Cont = 0
         ContLinea = 0
         Cadena = ""
         While ContTotal <> LOF(3) And Not salir
            Caracter = Input(1, 3)
            ContTotal = ContTotal + 1
            Cont = Cont + 1
            Trio(Cont) = Asc(Caracter)
            If Cont = 3 Then
               Cuatro(1) = Int(Trio(1) / 4)
               Cuatro(2) = (Trio(1) - Int(Trio(1) / 4) * 4) * 16 + Int(Trio(2) / 16)
               Cuatro(3) = (Trio(2) - (Int(Trio(2) / 16) * 16)) * 4 + Int(Trio(3) / 64)
               Cuatro(4) = Trio(3) - Int(Trio(3) / 64) * 64
               Cont = 0
               ContLinea = ContLinea + 4
               Cadena = Cadena & Mid(Base64, Cuatro(1) + 1, 1) & Mid(Base64, Cuatro(2) + 1, 1) & Mid(Base64, Cuatro(3) + 1, 1) & Mid(Base64, Cuatro(4) + 1, 1)
               If ContLinea = 76 Then
                  salir = True
               End If
            End If
         Wend
         If ContTotal = LOF(3) Then
            Close #3
            Paso = 900
            Select Case Cont
               Case 1
                  Cuatro(1) = Int(Trio(1) / 4)
                  Cuatro(2) = (Trio(1) - Int(Trio(1) / 4) * 4) * 16
                  Cadena = Cadena & Mid(Base64, Cuatro(1) + 1, 1) & Mid(Base64, Cuatro(2) + 1, 1) & "=="
               Case 2
                  Cuatro(1) = Int(Trio(1) / 4)
                  Cuatro(2) = (Trio(1) - Int(Trio(1) / 4) * 4) * 16 + Int(Trio(2) / 16)
                  Cuatro(3) = (Trio(2) - (Int(Trio(2) / 16) * 16)) * 4
                  Cadena = Cadena & Mid(Base64, Cuatro(1) + 1, 1) & Mid(Base64, Cuatro(2) + 1, 1) & Mid(Base64, Cuatro(3) + 1, 1) & "="
            End Select
         End If
        ' Progreso.Value = Progreso.Value + 76
         EnviarDatos (Cadena & vbCrLf)
      Case 900
         Paso = 901
         ' Envio el mensaje de mml_FRASE0245
         Progreso.Value = Progreso.Value + 10
         EnviarDatos (vbCrLf)
      Case 901
         Paso = 902
         ' Envio el mensaje de mml_FRASE0245
         Progreso.Value = Progreso.Value + 10
         EnviarDatos ("--" & Boundary & "--" & vbCrLf)
      Case 902
         Paso = 997
         ' Envio el mensaje de mml_FRASE0245
         Progreso.Value = Progreso.Value + 10
         EnviarDatos (vbCrLf)
      Case 997
         Paso = 998
         ' Envio el mensaje de mml_FRASE0245
         Progreso.Value = Progreso.Value + 10
         EnviarDatos ("." & vbCrLf)
   End Select
End Sub

Private Sub Retardo_Timer()
   ' Despues de enviar el mensaje espero X segundos antes
   ' dar por terminado el envio del mensaje
   Enviar.Enabled = True
   Retardo.Enabled = False
   Progreso.Value = 0
End Sub

Private Sub TimeOut_Timer()
   Enviar.Enabled = True
   MsgBox mml_FRASE0246 & vbCrLf & mml_FRASE0247
End Sub

