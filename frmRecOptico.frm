VERSION 5.00
Begin VB.Form frmRecOptico 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "mml_FRASE0902"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "frmRecOptico.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   638
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelCateg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1260
      Picture         =   "frmRecOptico.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmdHojasProc 
      Height          =   315
      Left            =   8670
      Picture         =   "frmRecOptico.frx":08AC
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Hojas procesadas"
      Top             =   270
      Width           =   375
   End
   Begin VB.CommandButton cmdEditarFichero 
      Height          =   315
      Left            =   9030
      Picture         =   "frmRecOptico.frx":0C61
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "mml_FRASE0903"
      Top             =   270
      Width           =   375
   End
   Begin VB.PictureBox picFicha 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11250
      Picture         =   "frmRecOptico.frx":1181
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   26
      ToolTipText     =   "mml_FRASE0904"
      Top             =   0
      Width           =   315
      Begin VB.CommandButton Command11 
         Caption         =   "Command2"
         Height          =   195
         Left            =   600
         TabIndex        =   28
         Top             =   210
         Width           =   255
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command1"
         Height          =   195
         Left            =   600
         TabIndex        =   27
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.Frame mrcImg 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   13500
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   11895
      Begin VB.Image imgFicha 
         BorderStyle     =   1  'Fixed Single
         Height          =   10995
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8895
      End
   End
   Begin VB.Frame frmBarra 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   345
      Left            =   9390
      TabIndex        =   18
      Top             =   270
      Width           =   2205
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   0
         Picture         =   "frmRecOptico.frx":14C3
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "mml_FRASE0905"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   360
         Picture         =   "frmRecOptico.frx":1A05
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "mml_FRASE0028"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdCalcular 
         Height          =   315
         Left            =   720
         Picture         =   "frmRecOptico.frx":1F73
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "mml_FRASE0058"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   1080
         Picture         =   "frmRecOptico.frx":24A1
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "mml_FRASE0906"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   1440
         Picture         =   "frmRecOptico.frx":2A41
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "mml_FRASE0907"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   1800
         Picture         =   "frmRecOptico.frx":2FD7
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "mml_FRASE0895"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picMin 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11550
      Picture         =   "frmRecOptico.frx":356D
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   15
      ToolTipText     =   "mml_FRASE0904"
      Top             =   0
      Width           =   375
      Begin VB.CommandButton Command9 
         Caption         =   "Command1"
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   210
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command2"
         Height          =   195
         Left            =   600
         TabIndex        =   16
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.TextBox tbHojasProc 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Text            =   "mml_FRASE0789"
      ToolTipText     =   "mml_FRASE0790"
      Top             =   240
      Width           =   8685
   End
   Begin VB.CommandButton cmdDesplazar 
      Height          =   315
      Left            =   11580
      Picture         =   "frmRecOptico.frx":3A5F
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "mml_FRASE0968"
      Top             =   270
      Width           =   315
   End
   Begin VB.PictureBox picActivo 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   10560
      ScaleHeight     =   195
      ScaleWidth      =   615
      TabIndex        =   6
      ToolTipText     =   "mml_FRASE0904"
      Top             =   0
      Width           =   675
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   210
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdParar 
      Caption         =   "mml_FRASE0791"
      Height          =   270
      Left            =   9720
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdActivar 
      Caption         =   "mml_FRASE0792"
      Height          =   270
      Left            =   8730
      TabIndex        =   4
      Top             =   0
      Width           =   1005
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "mml_FRASE0793"
      Enabled         =   0   'False
      Height          =   270
      Left            =   7320
      TabIndex        =   3
      Top             =   0
      Width           =   1410
   End
   Begin VB.TextBox tbDescComp 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2415
      TabIndex        =   1
      Top             =   0
      Width           =   4935
   End
   Begin VB.TextBox tbCodComp 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Timer tmrLecFichas 
      Enabled         =   0   'False
      Left            =   120
      Top             =   0
   End
   Begin VB.PictureBox pbFicha 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   9615
      Left            =   0
      ScaleHeight     =   637
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   788
      TabIndex        =   8
      Top             =   495
      Width           =   11880
      Begin VB.PictureBox pbFichaOculta 
         Height          =   4815
         Left            =   3480
         ScaleHeight     =   4755
         ScaleWidth      =   6675
         TabIndex        =   11
         Top             =   2640
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.ListBox lbDir 
         Height          =   5130
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.PictureBox pbDesc 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   480
         ScaleHeight     =   133
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   663
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   10005
      End
   End
   Begin VB.Label Label5 
      Caption         =   "mml_FRASE0215"
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
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   1170
   End
End
Attribute VB_Name = "frmRecOptico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCFichas As Integer
Dim sCHoraFichas As String
Dim iRenombrar As Integer
Dim bProcesando As Boolean
Dim sNombreFicha As String
Dim Directorio As TDirectorio
Dim sFichero As String
Dim bMarcaControl As Boolean
Dim bHojaExtendida As Boolean
Dim bFalloHoja As Boolean

Dim igCodCat As Integer
Dim igFase As Integer
Dim igHojaRepesca As Integer
        


' Cálculo de pendientes
Dim dPteY As Double
Dim dPteX As Double
Dim dTermIndepX As Double, dTermIndepY As Double
Dim dEsX As Double, dEsY As Double
Dim dPteY1 As Double
Dim dPteX1 As Double
Dim dTermIndepX1 As Double, dTermIndepY1 As Double
' *********************************************
' Display the picture at the correct scale.
' *********************************************
Private Sub DrawPicture()
    Dim wid As Single
    Dim hgt As Single
    
    If ScaleFactor <> 1 Then
        wid = pbFichaOculta.ScaleWidth * ScaleFactor
        hgt = pbFichaOculta.ScaleHeight * ScaleFactor
        pbFicha.Cls
        pbFicha.Move 1, 16, wid, hgt
        pbFicha.Refresh
        pbFicha.PaintPicture pbFichaOculta.Picture, _
            0, 0, wid, hgt, 0, 0, _
            pbFichaOculta.ScaleWidth, _
            pbFichaOculta.ScaleHeight
    End If
End Sub



Private Sub cmdActivar_Click()
    If tmrLecFichas.Enabled = True Then
        tmrLecFichas.Enabled = False
        cmdActivar.Caption = mml_FRASE0794
        tbDescComp.BackColor = &HC0C0FF
    Else
        tmrLecFichas.Interval = VarCfg("refresco_lectura_fichas")
        tmrLecFichas.Enabled = True
        cmdActivar.Caption = mml_FRASE0792
        tbDescComp.BackColor = &HC0FFC0
    End If
End Sub

Private Sub cmdCalcular_Click()
    If tbCodComp.Text <> "" And igCodCat <> 0 And igFase <> 0 Then
        frmCalcular.tbCodComp.Text = tbCodComp.Text
        frmCalcular.tbDescComp.Text = tbDescComp.Text
        frmCalcular.tbCodCat.Text = igCodCat
        frmCalcular.tbDescCat.Text = sDescCategoria(igCodCat)
        frmCalcular.tbCodFase.Text = igFase
        frmCalcular.tbDescFase.Text = sDescFase(igFase)
        frmCalcular.chkRep.Value = igHojaRepesca
        tbHojasProc.Text = C_HOJAS_PROC
    End If
    frmCalcular.MostrarCalcular
End Sub

Private Sub cmdDesplazar_Click()
    If cmdDesplazar.Tag = mml_FRASE0795 Then
        DesplazarVisualizacionHoja 0
        cmdDesplazar.Tag = ""
    Else
        DesplazarVisualizacionHoja G_MAX_FILA_VIS_HOJA + 1
        cmdDesplazar.Tag = mml_FRASE0795
    End If
End Sub

Private Sub cmdEditarFichero_Click()
Dim sDirFichas As String

    sDirFichas = VarCfg("dir_fichas")
    If sFichero <> "" Then
        Shell G_APP_GRAFICA & " """ & sFichero & """", vbNormalFocus
    End If
End Sub

Private Sub cmdGrabar_Click()
    GrabarInfoHoja
End Sub

Private Sub cmdHojasProc_Click()
Dim rs As Recordset, sJuez As String, sCad As String

    Set rs = db.OpenRecordset("SELECT id_juez,tanda FROM hojas_reconocidas WHERE repesca = " & igHojaRepesca & " AND cod_categoria = " & igCodCat & " AND fase = " & igFase & " GROUP BY id_juez, tanda", dbOpenSnapshot)
    While Not rs.EOF
        If sJuez <> rs!id_juez Then
            If sCad <> "" Then
                sCad = sCad & Chr$(13) & Chr$(10)
            End If
            sJuez = rs!id_juez
            sCad = sCad & mml_FRASE0485 & rs!id_juez & ": "
        End If
        sCad = sCad & rs!tanda & " "
        rs.MoveNext
    Wend
    rs.Close
    If sCad <> "" Then
        MsgBox mml_FRASE0789 & Chr$(13) & Chr$(10) & sCad, vbOKOnly Or vbInformation, mml_FRASE0147
    End If

End Sub

Private Sub cmdParar_Click()
    If cmdParar.Caption = mml_FRASE0791 Then
        cmdParar.Caption = mml_FRASE0796
        picActivo.BackColor = C_COLOR_VERDE
    Else
        cmdParar.Caption = mml_FRASE0791
        picActivo.BackColor = C_COLOR_ROJO
    End If
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
End Sub



Private Sub Command3_Click()
    If igFase > 1 Then
        If tbCodComp.Text <> "" And igCodCat <> 0 And igFase <> 0 Then
            frmPublicar.tbCodComp.Text = tbCodComp.Text
            frmPublicar.tbDescComp.Text = tbDescComp.Text
            frmPublicar.tbCodCat.Text = igCodCat
            frmPublicar.tbDescCat.Text = sDescCategoria(igCodCat)
            frmPublicar.tbCodFase.Text = igFase - 1
            frmPublicar.tbDescFase.Text = sDescFase(igFase - 1)
            frmPublicar.chkRep.Value = igHojaRepesca
        End If
    Else
        MsgBox mml_FRASE0797, vbOKOnly Or vbInformation, mml_FRASE0096
    End If
    frmPublicar.Show 1

End Sub

Private Sub Command4_Click()
    If tbCodComp.Text <> "" And igCodCat <> 0 And igFase <> 0 Then
        frmImpHojasPuntuaciones.tbCodComp.Text = tbCodComp.Text
        frmImpHojasPuntuaciones.tbDescComp.Text = tbDescComp.Text
        frmImpHojasPuntuaciones.tbCodCat.Text = igCodCat
        frmImpHojasPuntuaciones.tbDescCat.Text = sDescCategoria(igCodCat)
        frmImpHojasPuntuaciones.tbCodFase.Text = igFase
        frmImpHojasPuntuaciones.tbDescFase.Text = sDescFase(igFase)
        frmImpHojasPuntuaciones.chkRep.Value = igHojaRepesca
    End If
     frmImpHojasPuntuaciones.Show 1

End Sub

Private Sub Command5_Click()
    If tbCodComp.Text <> "" And igCodCat <> 0 And igFase <> 0 Then
        frmImprimirFinal.tbCodComp.Text = tbCodComp.Text
        frmImprimirFinal.tbDescComp.Text = tbDescComp.Text
        frmImprimirFinal.tbCodCateg.Text = igCodCat
        frmImprimirFinal.tbDescCateg.Text = sDescCategoria(igCodCat)
        frmImprimirFinal.tbCodFase.Text = igFase
        frmImprimirFinal.tbDescFase.Text = sDescFase(igFase)
        frmImprimirFinal.chkRep.Value = igHojaRepesca
    End If
    frmImprimirFinal.Show 1

End Sub

Private Sub Command7_Click()
    If tbCodComp.Text <> "" And igCodCat <> 0 And igFase <> 0 Then
        frmADorsales.tbCodComp.Text = tbCodComp.Text
        frmADorsales.tbDescComp.Text = tbDescComp.Text
        frmADorsales.tbCodCateg.Text = igCodCat
        frmADorsales.tbDescCateg.Text = sDescCategoria(igCodCat)
        frmADorsales.cbRepesca.Value = igHojaRepesca
    End If
    frmADorsales.Show 0

End Sub

Private Sub Command8_Click()
    If tbCodComp.Text <> "" And igCodCat <> 0 And igFase <> 0 Then
        frmDescalificados.tbCodComp.Text = tbCodComp.Text
        frmDescalificados.tbDescComp.Text = tbDescComp.Text
        frmDescalificados.tbCodCateg.Text = igCodCat
        frmDescalificados.tbDescCateg.Text = sDescCategoria(igCodCat)
        frmDescalificados.chkRep.Value = igHojaRepesca
    End If
    frmDescalificados.Show 0

End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    tbHojasProc.Text = C_HOJAS_PROC
    tmrLecFichas.Enabled = False
    cmdActivar.Caption = mml_FRASE0794
    tbDescComp.BackColor = &HC0C0FF
    tmrLecFichas.Interval = VarCfg("refresco_lectura_fichas")
    bProcesando = False
    iRenombrar = VarCfg(mml_FRASE0798)
    sCHoraFichas = Format(Time(), "hhmmss")
    
    If C_REC_OPTICO_RAPIDO Then
        pbFicha.Visible = False
        mrcImg.Visible = True
    Else
        pbFicha.Visible = True
        mrcImg.Visible = False
    End If
End Sub


Private Sub Form_Resize()
On Local Error Resume Next
    If frmRecOptico.Width < C_ANCHURA_LECT_OPTICA Then
        MsgBox mml_FRASE0799, vbOKOnly Or vbInformation, mml_FRASE0084
        frmRecOptico.Width = C_ANCHURA_LECT_OPTICA
    End If
On Local Error GoTo 0
End Sub

Private Sub Form_Terminate()
    Me.Tag = mml_FRASE0029
End Sub

Private Sub pbFicha_DblClick()
    Call tmrLecFichas_Timer
End Sub

Private Sub pbFicha_Resize()
    Me.Refresh
End Sub

Private Sub picActivo_DblClick()
    If frmRecOptico.Height < Screen.Height Then
        Maximizar
    Else
        frmRecOptico.Top = 0
        frmRecOptico.Left = 0
        frmRecOptico.Height = C_ALTURA_LECT_OPTICA
    End If
    On Local Error Resume Next
    VerMenu True
    G_MARCA_MAYOR = IIf(VarCfg("marca_mayor") = "S", True, False)
End Sub

Private Sub picFicha_Click()
    If mrcImg.Visible Then
        C_REC_OPTICO_RAPIDO = False
        mrcImg.Visible = False
        pbFicha.Visible = True
    Else
        C_REC_OPTICO_RAPIDO = True
        mrcImg.Visible = True
        pbFicha.Visible = False
    End If
End Sub

Private Sub picMin_Click()
    Me.Height = 550
    frmRecOptico.pbFicha.Top = G_VIS_HOJA_POS_INIC
End Sub

Private Sub tbHojasProc_DblClick()
    tbHojasProc.Text = C_HOJAS_PROC
End Sub

Private Sub tbHojasProc_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub tmrLecFichas_Timer()
Dim sDirFichas As String
Dim i As Integer
Dim Msj As String
Dim iCFicheros As Integer
Dim sFichSalida As String
Dim sSigFich As String

    If Not bProcesando Then
    
        On Local Error GoTo error
        bProcesando = True
        sDirFichas = VarCfg("dir_fichas")
        
        lbDir.Clear
        iCFicheros = 0
        sSigFich = Dir$(sDirFichas & "\*." & C_EXTENSION_FICHEROS)
        Do While sSigFich <> ""
            lbDir.AddItem sSigFich
            sSigFich = Dir
            Inc iCFicheros
        Loop
        
        picActivo.FontName = "Arial"
        picActivo.FontSize = 7
        picActivo.Cls
        picActivo.Print iCFicheros
        If cmdParar.Caption <> mml_FRASE0791 Then
            If iCFicheros = 0 Then
                picActivo.BackColor = C_COLOR_VERDE_MEDIO
            Else
                picActivo.BackColor = C_COLOR_AMARILLO_MEDIO
            End If
        End If
        For i = 0 To lbDir.ListCount - 1
            If C_REC_OPTICO_RAPIDO Then
                mrcImg.Visible = True
                pbFicha.Visible = False
            End If
            
            
            bFalloHoja = False
            picActivo.Cls
            picActivo.Print Str$(i + 1) & " d " & iCFicheros
            sFichero = lbDir.List(i)
            If iRenombrar = 1 Then
                On Local Error GoTo repetir
            End If
            Inc iCFichas
repetir:
            If Not C_DEBUG Then On Local Error GoTo error
            
            If cmdParar.Caption = mml_FRASE0791 Then Exit For
            
            If ScaleFactor = 1 Then
                pbFicha.Picture = LoadPicture(sDirFichas & "\" & sFichero)
            Else
                pbFichaOculta.Picture = LoadPicture(sDirFichas & "\" & sFichero)
            End If
            DrawPicture
            If C_REC_OPTICO_RAPIDO Then
                imgFicha.Picture = LoadPicture(sDirFichas & "\" & sFichero)
            End If
            
            frmRecOptico.Caption = sFichero
            sNombreFicha = ""
            ProcesarFicha
            If bFalloHoja And G_PARAR_REC_SI_FALLO Then
                cmdParar_Click
                DoEvents
            End If
            'Lo copiamos y borramos al acabar de reconocer
            If iRenombrar = 1 Then
                'Name sDirFichas & "\" & sFichero As sDirFichas & "\" & sFichero & ".PROC_sCHoraFichas_" & iCFichas
                'FileCopy sDirFichas & "\" & sFichero, sDirFichas & "\TMP\P" & sCHoraFichas & "_" & iCFichas & "_" & sFichero
                If sNombreFicha <> "" Then
                    If Directorio.bNuevo Then
                        On Local Error Resume Next
                        MkDir sDirFichas & "\TMP\" & Directorio.sDirectorio
                        On Local Error GoTo error
                    End If
                    sFichSalida = sDirFichas & "\TMP\" & Directorio.sDirectorio & "\" & sNombreFicha & "." & C_EXTENSION_FICHEROS
                    FileCopy sDirFichas & "\" & sFichero, sFichSalida
                Else
                    sFichSalida = sDirFichas & "\TMP\ERRORES\Ficha" & Trim$(Str$(iCFichas)) & "." & C_EXTENSION_FICHEROS
                    FileCopy sDirFichas & "\" & sFichero, sFichSalida
                End If
                If bFalloHoja = True And C_PREGUNTAR_EDIC_HOJA Then
                    If MsgBox(mml_FRASE0800, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
                        Shell G_APP_GRAFICA & " """ & sDirFichas & "\" & sFichero & """", vbNormalFocus
                        If MsgBox(mml_FRASE0801, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
                            GoTo NoBorrar
                        End If
                    End If
                End If
                Kill sDirFichas & "\" & sFichero
NoBorrar:
                sFichero = sFichSalida
                Inc iCFichas
            End If
        Next
        bProcesando = False
    End If
error:
If Err.Number <> 0 Then
   Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
         & Err.Source & Chr(13) & Err.Description
   MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
   bProcesando = False
End If

End Sub
Sub ProcesarFicha()
Dim X As Integer, Y As Integer
Dim iMinX As Integer, iMinY As Integer
Dim iAnchoMarca As Integer
Dim iAltoMarca As Integer

    'Comenzamos con tamaño pequeño
    C_MAX_MARCAS_X = C_MAX_MARCAS_X_NORMAL
    
    DesplazarVisualizacionHoja 0
    frmRecOptico.ZOrder
    iMinX = VarCfg("margen_lectura_optica_X")
    iMinY = VarCfg("margen_lectura_optica_Y")
    X = 0
    While X <= C_MAX_MARCAS_X - 1
        If X = 0 Then
            C_MARGEN_BUSQUEDA = C_MARGEN_BUSQUEDA_INIC
        Else
            C_MARGEN_BUSQUEDA = C_MARGEN_BUSQUEDA_NORMAL
        End If
        If Not BuscarMarca(X, 0, iMinX, iMinY) Then
            MsgBox mml_FRASE0892 & X & mml_FRASE0802, vbOKOnly Or vbCritical, mml_FRASE0096
            bFalloHoja = True
            Exit Sub
        End If
        iAnchoMarca = aMarcas(X, 0).iXf - aMarcas(X, 0).iXi
        iMinX = aMarcas(X, 0).iXf + (C_NUM_PUNTOS_BLANCOS + C_MARGEN)
        iMinY = aMarcas(X, 0).iYi - (C_NUM_PUNTOS_BLANCOS + C_MARGEN)
        If iMinY < 0 Then iMinY = 0
        'MsgBox x
        If X = C_MAX_MARCAS_X_NORMAL - 1 Then
            If iMinX > C_LIM_HOJA_EXT Then
                bHojaExtendida = False
            Else
                bHojaExtendida = True
                C_MAX_MARCAS_X = C_MAX_MARCAS_X_EXT
            End If
        End If
        Inc X
    Wend
    
    
    iAltoMarca = aMarcas(X - 1, 0).iYf - aMarcas(X - 1, 0).iYi
    iMinX = aMarcas(X - 1, 0).iXi - (C_NUM_PUNTOS_BLANCOS + C_MARGEN)
    iMinY = aMarcas(X - 1, 0).iYi + iAltoMarca + (C_NUM_PUNTOS_BLANCOS + C_MARGEN)
    If iMinY < 0 Then iMinY = 0
    
    For Y = 1 To C_MAX_MARCAS_Y - 1
        If Not BuscarMarca(C_MAX_MARCAS_X, Y, iMinX, iMinY) Then
            MsgBox mml_FRASE0892 & C_MAX_MARCAS_X & "," & Y & mml_FRASE0803, vbOKOnly Or vbCritical, mml_FRASE0096
            bFalloHoja = True
            Exit Sub
        End If
        iAltoMarca = aMarcas(C_MAX_MARCAS_X, Y).iYf - aMarcas(C_MAX_MARCAS_X, Y).iYi
        iMinX = aMarcas(C_MAX_MARCAS_X, Y).iXi - (C_NUM_PUNTOS_BLANCOS + C_MARGEN)
        iMinY = aMarcas(C_MAX_MARCAS_X, Y).iYi + iAltoMarca + (C_NUM_PUNTOS_BLANCOS + C_MARGEN)
        If iMinY < 0 Then iMinY = 0
        'MsgBox y
    Next Y
    
    'Localizamos la marca de control inferior
    iAltoMarca = aMarcas(0, 0).iYf - aMarcas(0, 0).iYi
    iMinX = G_MARGEN_X_MARCA_CONTROL
    iMinY = aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 2).iYi
    If BuscarMarca(0, C_MAX_MARCAS_Y - 1, iMinX, iMinY) Then
        bMarcaControl = True
    Else
        If G_MARCA_CONTROL Then
            If MsgBox(mml_FRASE0804, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                mrcImg.Visible = False
                pbFicha.Visible = True
                Maximizar
            End If
        End If
        bMarcaControl = False
        If G_DESPLAZAR_SI_CONTROL_NO_LOCALIZADO Then
            DesplazarVisualizacionHoja G_MAX_FILA_VIS_HOJA + 1
        End If
    End If
    
    
    Me.Caption = mml_FRASE0805
    'Comprobamos si estamos dentro de los límites
    Debug.Print aMarcas(0, 0).iYi
    Debug.Print aMarcas(C_MAX_MARCAS_X - 1, 0).iYi
    Debug.Print aMarcas(C_MAX_MARCAS_X, 1).iXi
    Debug.Print aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iXi

    'Comprobamos la tolerancia
    If Abs(aMarcas(0, 0).iYi - aMarcas(C_MAX_MARCAS_X - 1, 0).iYi) > C_TOLERANCIA_Y Or _
        Abs(aMarcas(C_MAX_MARCAS_X, 1).iXi - aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iXi) > C_TOLERANCIA_X Then
        MsgBox mml_FRASE0806, vbOKOnly Or vbCritical, mml_FRASE0096
    Else
        ProcesarMarcas
    End If
    
End Sub

Sub ProcesarMarcas()
Dim ix As Integer, iy As Integer
Dim X As Integer, Y As Integer
Dim cx As Integer, cy As Integer
Dim iCPuntosNegros As Long
Dim iCPuntosBlancos As Long
Dim iSalto As Integer
Dim sMarca As String
    
    'Borramos la variable para que se calcule de nuevo solo una vez por hoja
    'dentro de ColorPunto()
    dPteY = 0
    pbFicha.FontBold = True
    
    'Procesamos la marca de control inferior
    'iCPuntosNegros = ProcesarMarca(0, C_MAX_MARCAS_Y - 2)
    'If iCPuntosNegros >= C_MARCA Then
    '    aMarcas(0, C_MAX_MARCAS_Y).iXi = 1
    'Else
    '    aMarcas(0, C_MAX_MARCAS_Y).iXi = 0
    'End If
    'ix = aMarcas(0, 0).iXi + 6
    'iy = aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iYi
    'pbFicha.CurrentX = (ix + (iy - aMarcas(C_MAX_MARCAS_X - 1, 0).iYi) * dPteX) * dEsX
    'pbFicha.CurrentY = (iy + (ix - aMarcas(C_MAX_MARCAS_X - 1, 0).iXi) * dPteY) * dEsY
    'pbFicha.ForeColor = C_COLOR_X_MARCA
    'pbFicha.Print iCPuntosNegros
    
    For X = 2 To C_MAX_MARCAS_X - 2
        DoEvents: DoEvents:
        For Y = 2 To C_MAX_MARCAS_Y - 1
            'Me.Caption = mml_FRASE0807 & X & " , " & Y & ")"
            iCPuntosNegros = ProcesarMarca(X, Y)
            
            If iCPuntosNegros >= C_MARCA Then
                ix = aMarcas(X, 0).iXi + 6
                iy = aMarcas(C_MAX_MARCAS_X, Y + 1).iYi
                
                pbFicha.CurrentX = (ix + (iy - aMarcas(C_MAX_MARCAS_X - 1, 0).iYi) * dPteX) * dEsX
                pbFicha.CurrentY = (iy + (ix - aMarcas(C_MAX_MARCAS_X - 1, 0).iXi) * dPteY) * dEsY
                
                pbFicha.ForeColor = C_COLOR_X_MARCA
                pbFicha.Print "X"
                aMarcas(X, Y + 1).iXi = 1
            ElseIf iCPuntosNegros < C_MARCA And iCPuntosNegros > C_BLANCO Then
                'Indeterminado
                pbFicha.ForeColor = &HFF
                ix = aMarcas(X, 0).iXi
                iy = aMarcas(C_MAX_MARCAS_X, Y + 1).iYi
                
                If iCPuntosNegros >= C_BLANCO + (C_MARCA - C_BLANCO) / 2 Then
                    aMarcas(X, Y + 1).iXi = 1
                Else
                    aMarcas(X, Y + 1).iXi = 0
                End If
                
                pbFicha.CurrentY = (iy + (ix - aMarcas(C_MAX_MARCAS_X - 1, 0).iXi) * dPteY) * dEsY
                If G_MOSTRAR_NUM_PUNTOS Then
                    pbFicha.CurrentX = (ix + (iy - aMarcas(C_MAX_MARCAS_X - 1, 0).iYi) * dPteX) * dEsX
                    pbFicha.Print iCPuntosNegros
                Else
                    pbFicha.CurrentX = (ix + (iy - aMarcas(C_MAX_MARCAS_X - 1, 0).iYi) * dPteX) * dEsX - 10
                    pbFicha.Print "?"
                End If
                
                aMarcas(X, Y + 1).iXi = -1
            Else
                aMarcas(X, Y + 1).iXi = 0
            End If
            aMarcas(X, Y + 1).iXf = iCPuntosNegros
        Next Y
    Next X
    
    If bHojaExtendida Then
        GrabarInfoHojaExt
    Else
        GrabarInfoHoja
    End If
End Sub
Function ProcesarMarca(X As Integer, Y As Integer) As Integer
Dim iCPuntosNegros As Integer
Dim iCPuntosBlancos As Integer
Dim iSalto As Integer
Dim cy As Integer, cx As Integer
    
    iCPuntosNegros = 0
    iCPuntosBlancos = 0
    iSalto = C_SALTO_PUNTOS_MARCA
    For cy = aMarcas(C_MAX_MARCAS_X, Y + 1).iYi - G_MARGEN_MARCA_Y To aMarcas(C_MAX_MARCAS_X, Y + 1).iYf + G_MARGEN_MARCA_Y
        cx = aMarcas(X, 0).iXi
        While cx <= aMarcas(X, 0).iXf
            If ValorColor(ColorPunto(cx, cy)) < C_UMBRAL Then
                If iSalto = C_SALTO_PUNTOS_MARCA Then
                    iSalto = 1
                    cx = cx - (C_SALTO_PUNTOS_MARCA)
                Else
                    Inc iCPuntosNegros
                End If
            Else
                If iCPuntosBlancos > C_NUM_PUNTOS_BLANCOS Then
                    iSalto = C_SALTO_PUNTOS_MARCA
                    iCPuntosBlancos = 0
                Else
                    Inc iCPuntosBlancos
                End If
            End If
            cx = cx + iSalto
        Wend
    Next
    ProcesarMarca = iCPuntosNegros

End Function
Function ColorPunto(ix As Integer, iy As Integer) As Long
Dim i As Integer
Dim X As Integer, Y As Integer, X1 As Integer, Y1 As Integer

    If dPteY = 0 Then
        dPteY = (aMarcas(C_MAX_MARCAS_X - 1, 0).iYi - aMarcas(0, 0).iYi) / (aMarcas(C_MAX_MARCAS_X - 1, 0).iXi - aMarcas(0, 0).iXi)
        dTermIndepY = aMarcas(C_MAX_MARCAS_X - 1, 0).iYi - aMarcas(0, 0).iYi
        dPteX = (aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iXi - aMarcas(C_MAX_MARCAS_X, 1).iXi) / (aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iYi - aMarcas(C_MAX_MARCAS_X, 1).iYi)
        dTermIndepX = aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iXi - aMarcas(C_MAX_MARCAS_X, 1).iXi
        
        If bMarcaControl Then
            dPteY1 = (aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iYi - aMarcas(0, C_MAX_MARCAS_Y - 1).iYi) / (aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iXi - aMarcas(0, C_MAX_MARCAS_Y - 1).iXi)
            dTermIndepY1 = aMarcas(C_MAX_MARCAS_X, C_MAX_MARCAS_Y - 1).iYi - aMarcas(0, C_MAX_MARCAS_Y - 1).iYi
            dPteX1 = (aMarcas(0, C_MAX_MARCAS_Y - 1).iXi - aMarcas(0, 0).iXi) / (aMarcas(0, C_MAX_MARCAS_Y - 1).iYi - aMarcas(0, 0).iYi)
            dTermIndepX1 = aMarcas(0, C_MAX_MARCAS_Y - 1).iXi - aMarcas(0, 0).iXi
        End If
        
        dEsX = (pbFicha.ScaleWidth / pbFicha.Width)
        dEsY = (pbFicha.ScaleHeight / pbFicha.Height)
    End If
    
    If bMarcaControl Then
        X = (ix + (iy - aMarcas(C_MAX_MARCAS_X, 1).iYi) * dPteX)
        Y = (iy - dTermIndepY + (ix - aMarcas(0, 0).iXi) * dPteY)
        
        X1 = (ix + (iy - aMarcas(0, 0).iYi) * dPteX1)
        Y1 = (iy - dTermIndepY1 + (ix - aMarcas(0, C_MAX_MARCAS_Y - 1).iXi) * dPteY1)
        
        If iy > G_MARGEN_CONTROL_PTE Then
            X = X1
            Y = Y1
        End If
        ColorPunto = pbFicha.Point(X * dEsX, Y * dEsY)
        If G_MARCAR_PUNTOS Then
            If ValorColor(ColorPunto) < C_UMBRAL Then
                pbFicha.PSet (X * dEsX, Y * dEsY), C_COLOR_MARCA
            End If
        End If
    Else
        ColorPunto = pbFicha.Point((ix + (iy - aMarcas(C_MAX_MARCAS_X - 1, 0).iYi) * dPteX) * dEsX, (iy - dTermIndepY + (ix - aMarcas(0, 0).iXi) * dPteY) * dEsY)
        If G_MARCAR_PUNTOS Then
            If ValorColor(ColorPunto) < C_UMBRAL Then
                pbFicha.PSet ((ix + (iy - aMarcas(C_MAX_MARCAS_X - 1, 0).iYi) * dPteX) * dEsX, (iy - dTermIndepY + (ix - aMarcas(0, 0).iXi) * dPteY) * dEsY), C_COLOR_MARCA
            End If
        End If
    End If
End Function

Function BuscarMarca(ix As Integer, iy As Integer, iMinX As Integer, iMinY As Integer) As Boolean
Dim iCPuntosBlancosX As Integer
Dim iCPuntosNegrosX As Integer
Dim iCPuntosBlancosY As Integer
Dim iCPuntosX As Integer
Dim X As Integer, Y As Integer
Dim iPosXi As Integer
Dim iMarcaXi As Integer, iMarcaXf As Integer
Dim iMarcaXi1 As Integer, iMarcaXf1 As Integer
Dim iMarcaYi As Integer, iMarcaYf As Integer
Dim dEsX As Double, dEsY As Double
Dim iAnchoMarca As Integer
Dim iMaxX As Integer, iMaxY As Integer
Dim iPesoX As Integer
Const C_PESO_MINIMO = 4

    BuscarMarca = False
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    dEsX = (pbFicha.ScaleWidth / pbFicha.Width)
    dEsY = (pbFicha.ScaleHeight / pbFicha.Height)
    'Primero debemos localizar las marcas
    iMaxX = IIf(pbFicha.Width < (iMinX + C_TAM_MARCA_X * C_MARGEN_BUSQUEDA - C_MARGEN), pbFicha.Width, (iMinX + C_TAM_MARCA_X * C_MARGEN_BUSQUEDA))
    iMaxY = IIf(pbFicha.Height < (iMinY + C_TAM_MARCA_Y * C_MARGEN_BUSQUEDA - C_MARGEN), pbFicha.Height, (iMinY + C_TAM_MARCA_Y * C_MARGEN_BUSQUEDA))
    
    iPesoX = 0
    iMarcaXi = iMinX
    iMarcaXf = iMaxX
    iMarcaXi1 = iMinX
    iMarcaXf1 = iMaxX
    For Y = iMinY To iMaxY
        iCPuntosBlancosX = 0
        For X = iMinX To iMaxX
            If ValorColor(pbFicha.Point(X * dEsX, Y * dEsY)) < C_UMBRAL Then
                'Reconoce una marca negra
                iPosXi = IIf(iPosXi = 0, X, iPosXi)
                iCPuntosBlancosX = 0
                Inc iCPuntosNegrosX
                Inc iCPuntosX
                If G_MARCAR_BUS_MARCA Then
                    pbFicha.PSet (X * (pbFicha.ScaleWidth / pbFicha.Width), Y * (pbFicha.ScaleHeight / pbFicha.Height)), &HFF
                    If X Mod 8 = 0 Then pbFicha.Refresh
                End If
                'DoEvents
                'MsgBox x & "," & y & "-" & iCPuntosNegrosX & "-" & iMarcaXi & "," & iMarcaXf & "." & iMarcaYi & "," & iMarcaYf & "|" & iPosXi & "|" & iCPuntosNegrosX
                'Sleep 5
            Else
                If G_MARCAR_BUS_MARCA Then
                    pbFicha.PSet (X * (pbFicha.ScaleWidth / pbFicha.Width), Y * (pbFicha.ScaleHeight / pbFicha.Height)), &HC0FFFF
                End If
                Inc iCPuntosBlancosX
                Inc iCPuntosX
                'Comprobamos si finalizo la zona negra
                If iCPuntosBlancosX >= C_NUM_PUNTOS_BLANCOS Then
                    ' Si la marca no tiene el tamaño mínimo es una falsa alarma
                    ' Solo entramos si encontramos algún punto negro
                    If (iCPuntosNegrosX < C_LIM_MIN_PTOS_NEGROS) Then
                        iCPuntosNegrosX = 0
                        iCPuntosX = 0
                        iPosXi = 0
                    Else
                        ' Si se han leido un número suficiente de puntos negros
                        If iCPuntosNegrosX >= C_TAM_MIN_MARCA_X And iCPuntosX - C_NUM_PUNTOS_BLANCOS <= C_TAM_MAX_MARCA_X Then
                            'Comprobamos siempre la primera posición del punto negro
                            iMarcaXi1 = iPosXi
                            iMarcaXf1 = X - iCPuntosBlancosX + 1
                            
                            'Primero comprobamos que la linea leída es suficientemente ancha y no es demasiado ancha
                            If (iMarcaXf1 - iMarcaXi1) < C_TAM_MAX_MARCA_X And (iMarcaXf1 - iMarcaXi1) > C_TAM_MIN_MARCA_X Then
                                'Comprobamos cuantas líneas tienen la misma posición en X y anchura
                                If Abs(iMarcaXi - iMarcaXi1) < 3 And Abs(iMarcaXf - iMarcaXf1) < 3 Then
                                    Inc iPesoX
                                Else
                                    iPesoX = 0
                                End If
                                
                                ' Si no se reconoció ninguna linea válida para marca todavía o
                                ' Si se reconoce una marca completa y es anterior a la que se reconoció primero
                                ' o si está próxima a la ya reconocida pero se ajusta mejoral tamaño de una
                                ' marca
                                
                        Dim iDifMarca1 As Integer, iDifMarca As Integer
                                iDifMarca = iMarcaXf - iMarcaXi - C_TAM_MARCA_X + 1
                                iDifMarca1 = iMarcaXf1 - iMarcaXi1 - C_TAM_MARCA_X + 1
                        If Not G_MARCA_MAYOR Then
                                '0. Si el la primera línea buena reconocida
                                '   o si se reconocio una marca distinta anterior
                                '1. Primera línea válida por anchura
                                '1. Primera línea válida por anchura
                                '2. La línea está próxima a la reconocida ( es la misma marca)
                                '3. y está más próxima a la anchura de una marca perfecta
                                '4. o está igual de cerca pero con tamaño superior
                                If iMarcaXi = iMinX Or _
                                   (iMarcaXi1 < iMarcaXi - C_TAM_MARCA_X And iPesoX < C_PESO_MINIMO) Or _
                                   (Abs(iMarcaXi1 - iMarcaXi) < C_TAM_MARCA_X And _
                                     (Abs(iDifMarca1) < Abs(iDifMarca) Or _
                                      (Abs(iDifMarca1) = Abs(iDifMarca) And iDifMarca1 > iDifMarca)) _
                                   ) Then
                                    iMarcaXf = iMarcaXf1
                                    iMarcaXi = iMarcaXi1
                                End If
                        Else
                                '0. Si el la primera línea buena reconocida
                                '   o si se reconocio una marca distinta anterior
                                '1. Primera línea válida por anchura
                                '1. Primera línea válida por anchura
                                '2. La línea está próxima a la reconocida ( es la misma marca)
                                '3. y tiene una tamaño adecuado y es mas grande que la anterior
                                If iMarcaXi = iMinX Or _
                                   (iMarcaXi1 < iMarcaXi - C_TAM_MARCA_X And iPesoX < C_PESO_MINIMO) Or _
                                   (Abs(iMarcaXi1 - iMarcaXi) < C_TAM_MARCA_X And _
                                     (iDifMarca1 > iDifMarca)) Then
                                    iMarcaXf = iMarcaXf1
                                    iMarcaXi = iMarcaXi1
                                End If
                        End If
                                iMarcaXf1 = iMarcaXf
                                iMarcaXi1 = iMarcaXi
                            End If
                            If Y - iMarcaYi >= C_TAM_MIN_MARCA_Y Then
                                iMarcaYf = Y
                            End If
                            If iMarcaYi = 0 Then
                                'Comienzo en y posible asignado
                                iMarcaYi = Y
                            End If
                            iPosXi = 0
                            GoTo continuar
                        Else
                            ' Llegamos aquí con Tamaño pequeño en X o si no encontramos negros
                            ' Si en Y el tamaño es correcto
                            ' ya tenemos un posible comienzo este es el final
                            If iMarcaYi > 0 Then
                                ' Si no es suficiente ancho en Y empezamos de nuevo
                                If iMarcaYf - iMarcaYi < C_TAM_MIN_MARCA_Y Then
                                    iMarcaXi = iMaxX
                                    iMarcaYi = 0
                                    iMarcaYf = 0
                                    iMarcaXf = 0
                                Else
                                    BuscarMarca = True
                                    GoTo salir
                                End If
                            End If
                            iPosXi = 0
                        End If
                    End If
                End If
            End If
        Next X
continuar:
    If iMarcaYi > 0 And X = iMaxX + 1 Then
        If Y - iMarcaYi >= C_TAM_MIN_MARCA_Y And iMarcaYf - iMarcaYi >= C_TAM_MIN_MARCA_Y Then
            BuscarMarca = True
            GoTo salir
        End If
    End If
    'DoEvents
    iCPuntosBlancosX = 0
    iCPuntosNegrosX = 0
    iCPuntosX = 0
    Next Y
    'Marcamos
    Debug.Print mml_FRASE0808
salir:
    iAnchoMarca = iMarcaXf - iMarcaXi
    'Solo si encuentra la marca
    If BuscarMarca Then
        ' Tamaño perfecto
        pbFicha.Line (iMarcaXi * dEsX, iMarcaYi * dEsY)-((iMarcaXi + C_TAM_MARCA_X) * dEsX, (iMarcaYi + C_TAM_MARCA_Y) * dEsY), 0, BF
        
        ' Tamaño encontrado
        pbFicha.Line (iMarcaXi * dEsX, iMarcaYi * dEsY)-(iMarcaXf * dEsX, iMarcaYf * dEsY), &HFFFF80, BF
        pbFicha.CurrentX = iMarcaXi * dEsX
        pbFicha.CurrentY = iMarcaYi * dEsY
        
        pbFicha.ForeColor = 0
        If iy = 0 Then
            pbFicha.Print ix
        Else
            pbFicha.Print iy
        End If
    End If
    
    aMarcas(ix, iy).iXi = iMarcaXi
    aMarcas(ix, iy).iXf = iMarcaXf
    aMarcas(ix, iy).iYi = iMarcaYi
    aMarcas(ix, iy).iYf = iMarcaYf
End Function

Function BuscarMarca1(ix As Integer, iy As Integer, iMinX As Integer, iMinY As Integer) As Boolean
Dim iCPuntosBlancosX As Integer
Dim iCPuntosNegrosX As Integer
Dim iCPuntosBlancosY As Integer
Dim iCPuntosX As Integer
Dim X As Integer, Y As Integer
Dim iPosXi As Integer
Dim iMarcaXi As Integer, iMarcaXf As Integer
Dim iMarcaXi1 As Integer, iMarcaXf1 As Integer
Dim iMarcaYi As Integer, iMarcaYf As Integer
Dim dEsX As Double, dEsY As Double
Dim iAnchoMarca As Integer
Dim iMaxX As Integer, iMaxY As Integer

    BuscarMarca1 = False
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    dEsX = (pbFicha.ScaleWidth / pbFicha.Width)
    dEsY = (pbFicha.ScaleHeight / pbFicha.Height)
    'Primero debemos localizar las marcas
    iMaxX = IIf(pbFicha.Width < (iMinX + C_TAM_MARCA_X * C_MARGEN_BUSQUEDA - C_MARGEN), pbFicha.Width, (iMinX + C_TAM_MARCA_X * C_MARGEN_BUSQUEDA))
    iMaxY = IIf(pbFicha.Height < (iMinY + C_TAM_MARCA_Y * C_MARGEN_BUSQUEDA - C_MARGEN), pbFicha.Height, (iMinY + C_TAM_MARCA_Y * C_MARGEN_BUSQUEDA))
    
    iMarcaXi = iMinX
    iMarcaXf = iMaxX
    iMarcaXi1 = iMinX
    iMarcaXf1 = iMaxX
    For Y = iMinY To iMaxY
        For X = iMinX To iMaxX
            If ValorColor(pbFicha.Point(X * dEsX, Y * dEsY)) < C_UMBRAL Then
                'Reconoce una marca negra
                iPosXi = IIf(iPosXi = 0, X, iPosXi)
                iCPuntosBlancosX = 0
                Inc iCPuntosNegrosX
                Inc iCPuntosX
                If G_MARCAR_BUS_MARCA Then
                    pbFicha.PSet (X * (pbFicha.ScaleWidth / pbFicha.Width), Y * (pbFicha.ScaleHeight / pbFicha.Height)), &HE04000
                End If
                'DoEvents
                'MsgBox x & "," & y & "-" & iCPuntosNegrosX & "-" & iMarcaXi & "," & iMarcaXf & "." & iMarcaYi & "," & iMarcaYf & "|" & iPosXi & "|" & iCPuntosNegrosX
                'Sleep 10
            Else
                If G_MARCAR_BUS_MARCA Then
                    pbFicha.PSet (X * (pbFicha.ScaleWidth / pbFicha.Width), Y * (pbFicha.ScaleHeight / pbFicha.Height)), &HC0FFFF
                End If
                Inc iCPuntosBlancosX
                Inc iCPuntosX
                'Comprobamos si finalizo la zona negra
                If iCPuntosBlancosX >= C_NUM_PUNTOS_BLANCOS Then
                    ' Si la marca no tiene el tamaño mínimo es una falsa alarma
                    ' Solo entramos si encontramos algún punto negro
                    If (iCPuntosNegrosX < C_LIM_MIN_PTOS_NEGROS) Then
                        iCPuntosNegrosX = 0
                        iCPuntosX = 0
                        iPosXi = 0
                    Else
                        ' Si la anchura es suficiente
                        If iCPuntosNegrosX >= C_TAM_MIN_MARCA_X And iCPuntosX - C_NUM_PUNTOS_BLANCOS <= C_TAM_MAX_MARCA_X Then
                            ' El tamaña en X es correcto
                            
                            'Comprobamos siempre la primera posición del punto negro en donde
                            'tenemos una línea suficientemente ancha
                            iMarcaXi1 = iPosXi
                            
                            'Comprobamos que el final pueda ser adecuado por tamaño y que se aproxime
                            ' más al tamaño de una marca que el final anterior
                            If ((X - iCPuntosBlancosX + 1) - iMarcaXi1 + 1 < C_TAM_MAX_MARCA_X And _
                                 (X - iCPuntosBlancosX + 1) - iMarcaXi1 + 1 > C_TAM_MIN_MARCA_X) Then
                                iMarcaXf1 = X - iCPuntosBlancosX + 1
                                Debug.Print iMarcaXf1
                            End If
                            
                            ' Si se reconoce una marca completa y es anterior a la que se reconoció primero
                            ' Tiene preferencia
                            If (iMarcaXf1 - iMarcaXi1) < C_TAM_MAX_MARCA_X And _
                               (iMarcaXf1 - iMarcaXi1) > C_TAM_MIN_MARCA_X And _
                                iMarcaXi1 < iMarcaXi - C_TAM_MARCA And iMarcaXf1 < iMarcaXf - C_TAM_MARCA Then
                                iMarcaXf = iMarcaXf1
                                iMarcaXi = iMarcaXi1
                            Else
                                ' Ahora comprobamos que los nuevos valores están más cerca del tamaño de una
                                ' marca
                                If Abs(C_TAM_MARCA + c_MARGEN_MARCA - (iMarcaXf1 - iMarcaXi1)) < Abs(C_TAM_MARCA + c_MARGEN_MARCA - (iMarcaXf - iMarcaXi)) Then
                                    iMarcaXf = iMarcaXf1
                                    iMarcaXi = iMarcaXi1
                                End If
                            End If
                            iMarcaXf1 = iMarcaXf
                            iMarcaXi1 = iMarcaXi
                            
                            
                            If Y - iMarcaYi >= C_TAM_MIN_MARCA_Y Then
                                iMarcaYf = Y
                            End If
                            If iMarcaYi = 0 Then
                                'Comienzo en y posible asignado
                                iMarcaYi = Y
                            End If
                            iPosXi = 0
                            GoTo continuar
                        Else
                            ' Llegamos aquí con Tamaño pequeño en X o si no encontramos negros
                            ' Si en Y el tamaño es correcto
                            ' ya tenemos un posible comienzo este es el final
                            If iMarcaYi > 0 Then
                                ' Si no es suficiente ancho en Y empezamos de nuevo
                                If iMarcaYf - iMarcaYi < C_TAM_MIN_MARCA_Y Then
                                    iMarcaXi = iMaxX
                                    iMarcaYi = 0
                                    iMarcaYf = 0
                                    iMarcaXf = 0
                                Else
                                    BuscarMarca1 = True
                                    GoTo salir
                                End If
                            End If
                            iPosXi = 0
                        End If
                    End If
                End If
            End If
        Next X
continuar:
    If iMarcaYi > 0 And X = iMaxX + 1 Then
        If Y - iMarcaYi >= C_TAM_MIN_MARCA_Y And iMarcaYf - iMarcaYi >= C_TAM_MIN_MARCA_Y Then
            BuscarMarca1 = True
            GoTo salir
        End If
    End If
    'DoEvents
    iCPuntosBlancosX = 0
    iCPuntosNegrosX = 0
    iCPuntosX = 0
    Next Y
    'Marcamos
    Debug.Print mml_FRASE0808
salir:
    iAnchoMarca = iMarcaXf - iMarcaXi
    'Solo si encuentra la marca
    If BuscarMarca1 Then
        ' Tamaño perfecto
        pbFicha.Line (iMarcaXi * dEsX, iMarcaYi * dEsY)-((iMarcaXi + C_TAM_MARCA_X) * dEsX, (iMarcaYi + C_TAM_MARCA_Y) * dEsY), 0, BF
        
        ' Tamaño encontrado
        pbFicha.Line (iMarcaXi * dEsX, iMarcaYi * dEsY)-(iMarcaXf * dEsX, iMarcaYf * dEsY), &HFFFF80, BF
        pbFicha.CurrentX = iMarcaXi * dEsX
        pbFicha.CurrentY = iMarcaYi * dEsY
        
        pbFicha.ForeColor = 0
        If iy = 0 Then
            pbFicha.Print ix
        Else
            pbFicha.Print iy
        End If
    End If
    
    aMarcas(ix, iy).iXi = iMarcaXi
    aMarcas(ix, iy).iXf = iMarcaXf
    aMarcas(ix, iy).iYi = iMarcaYi
    aMarcas(ix, iy).iYf = iMarcaYf
End Function

Sub GrabarInfoHoja()
Dim rs As Recordset
Dim ix As Integer
Dim iy As Integer
Dim iCodCat As Integer
Dim iFase As Integer
Dim sIdJuez As String
Dim iTanda As Integer
Dim iMaxTandas As Integer
Dim iMaxDorsales As Integer
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim iCBailes As Integer
Dim iCDorsales As Integer
Dim iMarca As Integer
Dim iPuesto As Integer
Dim iCPuestos As Integer
Dim aPuestos(7) As Integer
Dim i As Integer, j As Integer
Dim iDorsales As Integer, iJueces As Integer, iBailes As Integer, iPuestos As Integer
Dim iHojas As Integer
Dim bJuezPasos As Boolean
Dim ijuecespasos As Integer
Dim iAlto As Integer, iAncho As Integer
Dim iXini As Integer, iYini As Integer, iAltoZona As Integer, iAnchoZona As Integer
Dim lColor As Long, dEsX As Double, dEsY As Double
Dim bHojaControl As Boolean
Dim iHojaRepesca As Integer
Dim iHoja2 As Integer
Dim iHojaExt As Integer
Dim iNumHojas As Integer
Dim iMaxHojasRec As Integer
Dim iContMarcas As Integer
Dim iCBailesHoja As Integer
Dim iAlturaRecOptico As Integer
Dim iDorsalesTanda As Integer
Dim iDorsalInicialTanda As Integer
Dim bTeamMatch As Boolean
Dim sAntHojasProc As String


    If Not C_DEBUG Then On Local Error GoTo error

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        cmdGrabar.Enabled = True
        sNombreFicha = ""
        Exit Sub
    End If
    
    Me.Caption = mml_FRASE0809 & sFichero
    VerMenu False
    
    
    'Comprobar referencia de ajuste
    ComprobarReferencia 19, 6
    
    'Primero debemos comprobar que todas las marcas contienen marcas o espacios
    ' con total seguridad
    'Comprobamos la cabecera
    ComprobarMarcas 2, G_POS_INIC_CATEG, 19, 5
    ComprobarMarcas 2, 7, 10, 7
    ComprobarMarcas 2, 9, 14, 9
    ComprobarMarcas 2, 11, 19, 11
    
    'Fallos
    ComprobarMarcas 12, 7, 12, 7
    'Control
    ComprobarMarcas 12, 9, 12, 9
    'Repesca
    ComprobarMarcas 13, 9, 13, 9
    'Hoja2
    ComprobarMarcas 14, 9, 14, 9
    VerMenu True
    
    cmdGrabar.Enabled = False
    
    'Recuperamos la información de la categoría
    ' categorías
    Set rs = db.OpenRecordset("SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY 1", dbOpenSnapshot)
    ix = 2
    iy = G_POS_INIC_CATEG
    Do While Not rs.EOF
        If ix Mod 20 = 0 Then
            Inc iy
            ix = 2
        End If
        If aMarcas(ix, iy).iXi = 1 Then
            iCodCat = rs!codigo
            Exit Do
        End If
        rs.MoveNext
        Inc ix
    Loop
    rs.Close
    
    bTeamMatch = ComprobarSiTeamMatch(iCodCat)
    
    If iCodCat = 0 Then
        GrabarPosicion Me
        Maximizar
        MsgBox mml_FRASE0810, vbOKOnly Or vbCritical, mml_FRASE0096
        RestaurarPosicion Me
        Exit Sub
    End If
    
    'Localizamos la descripción de la categoria para crear el directorio
    Set rs = db.OpenRecordset("SELECT codigo, descripcion FROM categorias WHERE codigo = " & iCodCat, dbOpenSnapshot)
    If Not rs.EOF Then
        If Directorio.sDirectorio = rs!DESCRIPCION & "_" & rs!codigo Then
            Directorio.bNuevo = False
        Else
            Directorio.sDirectorio = rs!DESCRIPCION & "_" & rs!codigo
            Directorio.bNuevo = True
        End If
    End If
    rs.Close
    
    ' fase
    ix = 2
    iy = 7
    
    Do While ix <= 10
        If aMarcas(ix, iy).iXi = 1 Then
            iFase = 2 ^ (ix - 2)
            Exit Do
        End If
        Inc ix
    Loop
    
    'jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos FROM juez_categ WHERE cod_categoria = " & iCodCat & " ORDER BY 1", dbOpenSnapshot)
    ix = 2
    iy = 9
    While Not rs.EOF
        If aMarcas(ix, iy).iXi = 1 Then
            sIdJuez = rs!id_juez
            If rs!pasos = 1 Then
                bJuezPasos = True
            Else
                bJuezPasos = False
            End If
        End If
        Inc ix
        rs.MoveNext
    Wend
    rs.Close
    
    ix = POS_CONTROL / 2
    iy = 9
    'Hoja de control
    If aMarcas(ix, iy).iXi = 1 Then
        bHojaControl = True
        sIdJuez = "Control"
    Else
        bHojaControl = False
    End If
    
    ix = POS_REPESCA / 2
    iy = 9
    'Hoja de repesca
    If aMarcas(ix, iy).iXi = 1 Then
        iHojaRepesca = 1
    Else
        iHojaRepesca = 0
    End If
    
    ix = POS_HOJA2 / 2
    iy = 9
    'Hoja2
    If aMarcas(ix, iy).iXi = 1 Then
        iHoja2 = 1
    Else
        iHoja2 = 0
    End If
    
    ix = C_POS_HOJA_EXT / 2
    iy = 7
    'HojaExt
    If aMarcas(ix, iy).iXi = 1 Then
        iHojaExt = 1
    Else
        iHojaExt = 0
    End If
    
    ' tanda
    ix = 2
    iy = 11
    
    Do While ix <= 19
        If aMarcas(ix, iy).iXi = 1 Then
            iTanda = IIf(iTanda = 0, (ix - 1), iTanda)
            iMaxTandas = IIf(iMaxTandas < (ix - 1), (ix - 1), iMaxTandas)
        End If
        Inc ix
    Loop
    
    Me.Caption = mml_FRASE0811 & iCodCat & mml_FRASE0641 & iFase & mml_FRASE0485 & sIdJuez & IIf(bJuezPasos, mml_FRASE0812, "") & mml_FRASE0813 & iTanda & mml_FRASE0814 & iMaxTandas & IIf(bHojaControl, mml_FRASE0815, "") & mml_FRASE0816 & iHojaRepesca & mml_FRASE0817 & iHoja2 & " - " & sFichero
    sNombreFicha = "C" & iCodCat & "F" & iFase & "J" & sIdJuez & "T" & iTanda & IIf(bHojaControl, "CT", "NC") & "R" & iHojaRepesca & "H2" & iHoja2
    
    'Recuperamos la información de si ha habido fallos
    If aMarcas(12, 7).iXi = 1 Then
        GrabarPosicion Me
        Maximizar
        If MsgBox(mml_FRASE0818, vbYesNo Or vbCritical, mml_FRASE0084) = vbNo Then
            RestaurarPosicion Me
            bFalloHoja = True
            Exit Sub
        End If
        RestaurarPosicion Me
    End If
    
    ' Comprobamos el número de bailes
    Set rsBailes = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ bc WHERE bc.cod_categoria = " & iCodCat & " AND bc.fase = " & IIf(iFase = 1, 1, 2), dbOpenSnapshot)
    iCBailes = rsBailes.Fields(0)
    rsBailes.Close
    
    If iCBailes > C_BAILES_POR_HOJA Then
        iNumHojas = 2
    Else
        iNumHojas = 1
    End If
    
    ' Comprobamos el número de dorsales
    Set rsDorsales = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase =" & iFase & " ORDER BY 1", dbOpenSnapshot)
        iMaxDorsales = rsDorsales.Fields(0)
    rsDorsales.Close
    
    If iMaxDorsales = 0 Then
        MsgBox mml_FRASE0819, vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    End If
    
    If iCBailes > C_BAILES_POR_HOJA Then
        If iHoja2 = 1 Then
            iCBailesHoja = iCBailes - C_BAILES_POR_HOJA
        Else
            iCBailesHoja = C_BAILES_POR_HOJA
        End If
    Else
        iCBailesHoja = iCBailes
    End If
    'Comprobamos los bailes
    If bHojaControl Then
        'Comprobar marcas de ausente y anulación
        ComprobarMarcas 2, 13, 19, 14
    ElseIf iFase = 1 Then ' FINAL
        ' Comprobamos marcas de descalificación
        For i = 4 To 4 + (iMaxDorsales - 1) * 2 Step 2
            For j = 0 To iCBailesHoja - 1
                ComprobarMarcas i, 12 + j * C_ANCHO_MARCAS_BAILE_FINAL, i, 12 + j * C_ANCHO_MARCAS_BAILE_FINAL
            Next
        Next
        If Not bJuezPasos Then
            ' Comprobamos marcas de puntuación y anulación
            For i = 0 To iCBailesHoja - 1
                ComprobarMarcas 3, 13 + i * C_ANCHO_MARCAS_BAILE_FINAL, 3 + (iMaxDorsales - 1) * 2, 19 + i * C_ANCHO_MARCAS_BAILE_FINAL
            Next
        End If
        'Comprobar las marcas de la presencia de bailes
        For j = 0 To iCBailesHoja - 1
            ComprobarMarcas C_REC_POS_X_MARCA_BAILE, 12 + j * C_ANCHO_MARCAS_BAILE_FINAL, C_REC_POS_X_MARCA_BAILE, 12 + j * C_ANCHO_MARCAS_BAILE_FINAL
        Next
    Else ' FASES ELIMINATORIAS
        If Not bJuezPasos Then
            For i = 0 To iCBailesHoja - 1
                ComprobarMarcas 2, 13 + i * C_ANCHO_MARCAS_BAILE, 19, 15 + i * C_ANCHO_MARCAS_BAILE
            Next
        Else
            For i = 0 To iCBailesHoja - 1
                ComprobarMarcas 2, 15 + i * C_ANCHO_MARCAS_BAILE, 19, 15 + i * C_ANCHO_MARCAS_BAILE
            Next
        End If
        'Comprobar las marcas de la presencia de bailes
        For j = 0 To iCBailesHoja - 1
            ComprobarMarcas C_REC_POS_X_MARCA_BAILE, 16 + j * C_ANCHO_MARCAS_BAILE, C_REC_POS_X_MARCA_BAILE, 16 + j * C_ANCHO_MARCAS_BAILE
        Next
    End If
    
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM hojas_reconocidas WHERE cod_categoria = " & iCodCat & " AND id_juez='" & sIdJuez & IIf(iHoja2 = 1, "2", "") & "' AND fase = " & iFase & " AND tanda = " & iTanda & " AND repesca = " & iHojaRepesca, dbOpenSnapshot)
    If Not rs.EOF Then
        If Not BailesParciales(iCodCat) And rs.Fields(0) > 0 Then
            GrabarPosicion Me
            Maximizar
            If MsgBox(mml_FRASE0820, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                RestaurarPosicion Me
                Exit Sub
            End If
            RestaurarPosicion Me
        End If
    End If
    rs.Close
        
    sAntHojasProc = tbHojasProc.Text
    tbHojasProc.Text = tbHojasProc.Text & RTrim$(Left$(sIdJuez & " ", 2)) & Trim$(Str$(iTanda)) & " "
    'db.Execute ("INSERT INTO hojas_reconocidas VALUES(" & iCodCat & ",'" & sIdJuez & IIf(iHoja2 = 1, "2", "") & "'," & iFase & "," & iTanda & "," & iHojaRepesca & ")")
    
    If bHojaControl Then
        iCDorsales = 0
        ix = 2
        iy = 13
        Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase =" & iFase & " ORDER BY 1", dbOpenSnapshot)
        iDorsalInicialTanda = CalcularDorsalInicialTandaCat(iCodCat, iFase, iHojaRepesca, iTanda, iMaxTandas, iDorsalesTanda)
        While iCDorsales < iDorsalInicialTanda - 1 And Not rsDorsales.EOF
            Inc iCDorsales
            rsDorsales.MoveNext
        Wend
        iCDorsales = 0
        While iCDorsales < iDorsalesTanda And Not rsDorsales.EOF
            ' Comprobamos la casilla de anulación
            iMarca = IIf(2 * aMarcas(ix, iy).iXi + aMarcas(ix, iy + 1).iXi = 2, 1, 0)
            If iMarca = 1 Then
                Debug.Print "UPDATE dorsales SET no_presentado = " & iMarca & "  WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND cod_categoria =" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase
                db.Execute ("UPDATE dorsales SET no_presente = " & iMarca & " WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND cod_categoria =" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase)
            End If
            rsDorsales.MoveNext
            Inc iCDorsales
            Inc ix
        Wend
    Else
    ' Hojas de puntuación -----------------------------------------------------------
        'Reconocer los bailes
        iCBailes = 0
        iCDorsales = 0
        If iFase > 1 Then 'Eliminatorias ********************************************************
            iy = 13
            Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & iCodCat & " AND bc.fase = 2 ORDER BY posicion", dbOpenSnapshot)
            If iHoja2 = 1 Then
                While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                    Inc iCBailes
                    rsBailes.MoveNext
                Wend
                iCBailes = 0
            End If
            While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                iContMarcas = 0
                iCDorsales = 0
                ix = 2
                
                If G_REC_HOJA_EXT And iHojaExt = 1 And BailesParciales(iCodCat) And aMarcas(C_REC_POS_X_MARCA_BAILE, iy + 3).iXi = 0 Then
                    GoTo Cont_Semi
                End If
                If BailesParciales(iCodCat) Then sNombreFicha = "B" & rsBailes!codigo & sNombreFicha
                
                If iMaxTandas > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(iCodCat) Then
                    Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & iCodCat & " AND d.repesca=" & iHojaRepesca & " AND d.fase =" & iFase & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
                Else
                    Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase =" & iFase & " ORDER BY 1", dbOpenSnapshot)
                End If
                
               iDorsalInicialTanda = CalcularDorsalInicialTandaCat(iCodCat, iFase, iHojaRepesca, iTanda, iMaxTandas, iDorsalesTanda)
               While iCDorsales < iDorsalInicialTanda - 1 And Not rsDorsales.EOF
                    Inc iCDorsales
                    rsDorsales.MoveNext
                Wend
                iCDorsales = 0
                While iCDorsales < iDorsalesTanda And Not rsDorsales.EOF
                    ' Comprobamos la casilla de anulación
                    iMarca = IIf(2 * aMarcas(ix, iy).iXi + aMarcas(ix, iy + 1).iXi = 2, 1, 0)
                    If iMarca = 1 Then Inc iContMarcas
                    Debug.Print "INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "'," & iMarca & "," & iFase & "," & iHojaRepesca & ")"
                    If Not bJuezPasos Then
                        'Borramos cualquier marca igual
                        db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND cod_categoria=" & iCodCat & " AND cod_baile=" & rsBailes!codigo & " AND cod_juez='" & sIdJuez & "' AND fase=" & iFase & " AND repesca=" & iHojaRepesca)
                        ' Si el dorsal está descalificado no le asignamos puntuación aunque la tenga
                        If (Not bJuezPasos And aMarcas(ix, iy + 2).iXi = 1) Or (bJuezPasos And aMarcas(ix, iy).iXi = 1) Then
                            db.Execute ("INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "',0," & iFase & "," & iHojaRepesca & ")")
                        Else
                            db.Execute ("INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "'," & iMarca & "," & iFase & "," & iHojaRepesca & ")")
                        End If
                    End If
                    
                    'Descalificación
                    db.Execute ("DELETE FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND num_dorsal=" & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes!codigo)
                    If (Not bJuezPasos And aMarcas(ix, iy + 2).iXi = 1) Or (bJuezPasos And aMarcas(ix, iy).iXi = 1) Then
                        If bJuezPasos Then
                            'Agrandamos la pantalla
                            iAlturaRecOptico = 0
                            If frmRecOptico.Height < C_ALTURA_LECT_OPTICA Then
                                iAlturaRecOptico = frmRecOptico.Height
                                frmRecOptico.Height = C_ALTURA_LECT_OPTICA
                                ProcesarEventos
                            End If
                            
                            iAncho = aMarcas(2, 0).iXf - aMarcas(2, 0).iXi + 1
                            iAlto = aMarcas(2, 0).iYf - aMarcas(2, 0).iYi + 1
                            iXini = aMarcas(1, 0).iXf
                            iYini = aMarcas(C_MAX_MARCAS_X, iy + 1).iYi
                            iAltoZona = iAlto * 5.5
                            iAnchoZona = iAncho * 36
                            dEsX = (pbDesc.ScaleWidth / pbDesc.Width)
                            dEsY = (pbDesc.ScaleHeight / pbDesc.Height)
                            pbDesc.Visible = True
                            pbDesc.Cls
                            pbDesc.Refresh
                            tbHojasProc.Visible = False
                            VerMenu False
                            frmBarra.Visible = False
                            DoEvents: DoEvents: DoEvents
                            For i = iXini To iXini + iAnchoZona
                                For j = iYini To iYini + iAltoZona
                                    lColor = pbFicha.Point(i, j)
                                    If ValorColor(lColor) < C_UMBRAL Or lColor = C_COLOR_MARCA Then
                                        lColor = 0
                                    Else
                                        lColor = &HFFFFFF
                                    End If
                                    pbDesc.PSet ((i - iXini) * dEsX, (j - iYini) * dEsY), lColor
                                Next j
                                If i Mod 50 = 0 Then DoEvents
                            Next i
                            SavePicture CaptureWindow(pbDesc.hwnd, True, 0, 0, (pbDesc.Width) - 4, (pbDesc.Height) - 4), "desc.bmp"
                            pbDesc.Picture = LoadPicture("Desc.bmp")
                            frmBarra.Visible = True
                            If iAlturaRecOptico > 0 Then
                                frmRecOptico.Height = iAlturaRecOptico
                                ProcesarEventos
                            End If
                        End If
                        Set rs = db.OpenRecordset(mml_FRASE0043, dbOpenTable)
                            rs.AddNew
                                rs!codigo = MaxCod("descalificaciones")
                                rs!cod_categoria = iCodCat
                                rs!fase = iFase
                                rs!id_juez = sIdJuez
                                rs!cod_Baile = rsBailes!codigo
                                rs!num_dorsal = rsDorsales!num_dorsal
                                rs!repesca = iHojaRepesca
                                If bJuezPasos Then
                                    GuardarBinary rs!anotacion, pbDesc
                                End If
                            rs.Update
                        rs.Close
                        pbDesc.Visible = False
                        tbHojasProc.Visible = True
                        VerMenu True
                        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
                    End If
                    rsDorsales.MoveNext
                    Inc iCDorsales
                    Inc ix
                Wend
                If G_AVISO_NUM_MARCAS Then
                ' Si hay varias tandas permitimos marcar una menos o más en cada tanda
                Dim iMargen As Integer
                    If iMaxTandas > 1 Then
                        iMargen = 1
                    Else
                        iMargen = 0
                    End If
                    If iHojaRepesca = 0 And Not bJuezPasos Then
                        If Abs(iContMarcas - (iFase / 2 * 6) / iMaxTandas) > iMargen Then
                            GrabarPosicion Me
                            Maximizar
                            DesplazarVisualizacionHoja iy
                            MsgBox mml_FRASE0821 & rsBailes!Nombre & mml_FRASE0822 & (iFase / 2 * 6) / iMaxTandas & mml_FRASE0823 & iContMarcas, vbOKOnly Or vbInformation, mml_FRASE0084
                            DesplazarVisualizacionHoja 0
                            RestaurarPosicion Me
                        End If
                    End If
                End If
                rsDorsales.Close
Cont_Semi:
                iy = iy + C_ANCHO_MARCAS_BAILE
                rsBailes.MoveNext
                Inc iCBailes
            Wend
            rsBailes.Close
        Else ' FINAL ********************************************************************
            iy = 13
            Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & iCodCat & " AND bc.fase = 1 ORDER BY posicion", dbOpenSnapshot)
            iCBailes = 0
            If iHoja2 = 1 Then
                While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                    Inc iCBailes
                    rsBailes.MoveNext
                Wend
                iCBailes = 0
            End If
            While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                ix = 3
                
                If G_REC_HOJA_EXT And iHojaExt = 1 And BailesParciales(iCodCat) And aMarcas(C_REC_POS_X_MARCA_BAILE, iy - 1).iXi = 0 Then
                    GoTo Cont_Final
                End If
                If BailesParciales(iCodCat) Then sNombreFicha = "B" & rsBailes!codigo & sNombreFicha
                
                Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase =" & iFase & " ORDER BY 1", dbOpenSnapshot)
                iCPuestos = 0
                While Not rsDorsales.EOF
                    If Not bJuezPasos Then
                        'Comprobamos el puesto
                        iPuesto = 0
                        For i = 0 To IIf(bTeamMatch, C_MAX_POS_TEAMMATCH, iMaxDorsales - 1)
                            ' Comprobamos la casilla de anulación
                            iMarca = IIf(2 * aMarcas(ix, iy + i).iXi + aMarcas(ix + 1, iy + i).iXi = 2, 1, 0)
                            If iMarca = 1 Then
                                ' Comprobamos que un dorsal no tenga varios puestos
                                If iPuesto <> 0 Then
                                    GrabarPosicion Me
                                    Maximizar
                                    DesplazarVisualizacionHoja iy
                                    MsgBox mml_FRASE0824 & rsDorsales!num_dorsal & mml_FRASE0489 & rsBailes!Nombre, vbOKOnly Or vbCritical, mml_FRASE0096
                                    DesplazarVisualizacionHoja 0
                                    RestaurarPosicion Me
                                    bFalloHoja = True
                                Else
                                    iPuesto = i + 1
                                    ' Comprobamos que no se repitan puestos
                                    If Not bTeamMatch Then
                                        For j = 0 To iCPuestos - 1
                                            If aPuestos(j) = iPuesto Then
                                                GrabarPosicion Me
                                                Maximizar
                                                DesplazarVisualizacionHoja iy
                                                MsgBox mml_FRASE0825 & iPuesto & mml_FRASE0826 & rsDorsales!num_dorsal & mml_FRASE0489 & rsBailes!Nombre, vbOKOnly Or vbCritical, mml_FRASE0096
                                                DesplazarVisualizacionHoja 0
                                                RestaurarPosicion Me
                                                bFalloHoja = True
                                            End If
                                        Next
                                        aPuestos(iCPuestos) = iPuesto
                                        Inc iCPuestos
                                    End If
                                End If
                            End If
                        Next
                        ' Si se ha descalificado el dorsal le asignamos el último puesto
                        ' Comprobamos que se haya marcado un puesto
                        If iPuesto = 0 Or aMarcas(ix + 1, iy - 1).iXi = 1 Then
                            'Comprobamos que el dorsal está marcado como persente
                            Dim rsDesc As Recordset, iNoPresente As Integer
                            Set rsDesc = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND fase = " & iFase & " AND repesca = " & iHojaRepesca & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND no_presente > 0", dbOpenSnapshot)
                            iNoPresente = rsDesc.Fields(0)
                            rsDesc.Close
                            ' Si no hay puesto y dosales < 8 y no se ha descalificado
                            ' y está presente se avisa de la situación
                            If iNoPresente = 0 And iPuesto = 0 And iMaxDorsales < 8 And aMarcas(ix + 1, iy - 1).iXi = 0 Then
                                GrabarPosicion Me
                                Maximizar
                                DesplazarVisualizacionHoja iy
                                MsgBox mml_FRASE0827 & rsDorsales!num_dorsal & mml_FRASE0489 & rsBailes!Nombre & mml_FRASE0828, vbOKOnly Or vbCritical, mml_FRASE0096
                                DesplazarVisualizacionHoja 0
                                RestaurarPosicion Me
                                bFalloHoja = True
                            End If
                            'Si la marca de descalificación se emplea para enviar al final
                            ', pero ordenados. se deja el puesto
                            ' para conservar el puesto tiene que existir, estár descalificado y el juez ordena las descalificaciones
                            If Not (iPuesto > 0 And aMarcas(ix + 1, iy - 1).iXi = 1 And InStr(G_ORDENAR_DESCALIFICADOS_FINAL, sIdJuez) > 0) Then
                                'Si está descalificado -> detrás del puesto mas alto posible
                                If aMarcas(ix + 1, iy - 1).iXi = 1 Then
                                    iPuesto = C_ULTIMO_PUESTO
                                Else
                                    iPuesto = iMaxDorsales
                                End If
                            End If
                        End If
                        Debug.Print "INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "'," & iPuesto & "," & iFase & "," & iHojaRepesca & ")"
                        'Borramos cualquier marca igual
                        db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND cod_categoria=" & iCodCat & " AND cod_baile=" & rsBailes!codigo & " AND cod_juez='" & sIdJuez & "' AND fase=" & iFase & " AND repesca=" & iHojaRepesca)
                        If bTeamMatch Then
                            db.Execute ("INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "','" & CDbl(iPuesto + 1) / 2 & "'," & iFase & "," & iHojaRepesca & ")")
                        Else
                            db.Execute ("INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "'," & iPuesto & "," & iFase & "," & iHojaRepesca & ")")
                        End If
                    End If
                    'Descalificación
                    db.Execute ("DELETE FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND num_dorsal=" & rsDorsales!num_dorsal & " AND cod_baile=" & rsBailes!codigo)
                    If aMarcas(ix + 1, iy - 1).iXi = 1 Then
                        If bJuezPasos Then
                            'Agrandamos la pantalla
                            iAlturaRecOptico = 0
                            If frmRecOptico.Height < C_ALTURA_LECT_OPTICA Then
                                iAlturaRecOptico = frmRecOptico.Height
                                frmRecOptico.Height = C_ALTURA_LECT_OPTICA
                                ProcesarEventos
                            End If

                            iAncho = aMarcas(2, 0).iXf - aMarcas(2, 0).iXi + 1
                            iAlto = aMarcas(2, 0).iYf - aMarcas(2, 0).iYi + 1
                            iXini = aMarcas(2, 0).iXi
                            iYini = aMarcas(C_MAX_MARCAS_X, iy).iYi
                            iAltoZona = iAlto * 11
                            iAnchoZona = iAncho * 34
                            dEsX = (pbDesc.ScaleWidth / pbDesc.Width)
                            dEsY = (pbDesc.ScaleHeight / pbDesc.Height)
                            pbDesc.Cls
                            tbHojasProc.Visible = False
                            pbDesc.Visible = True
                            VerMenu False
                            frmBarra.Visible = False
                            pbDesc.Refresh
                            DoEvents: DoEvents: DoEvents
                            For i = iXini To iXini + iAnchoZona
                                For j = iYini To iYini + iAltoZona
                                    lColor = pbFicha.Point(i, j)
                                    If ValorColor(lColor) < C_UMBRAL Or lColor = C_COLOR_MARCA Then
                                        lColor = 0
                                    Else
                                        lColor = &HFFFFFF
                                    End If
                                    'If G_MARCAR_PUNTOS Then pbDesc.PSet ((i - iXini) * dEsX, (j - iYini) * dEsY), lColor
                                    pbDesc.PSet ((i - iXini) * dEsX, (j - iYini) * dEsY), lColor
                                Next j
                                If i Mod 50 = 0 Then DoEvents
                            Next i
                            SavePicture CaptureWindow(pbDesc.hwnd, True, 0, 0, (pbDesc.Width) - 4, (pbDesc.Height) - 4), "desc.bmp"
                            pbDesc.Picture = LoadPicture("Desc.bmp")
                            frmBarra.Visible = True
                            If iAlturaRecOptico > 0 Then
                                frmRecOptico.Height = iAlturaRecOptico
                                ProcesarEventos
                            End If
                        End If
                        Set rs = db.OpenRecordset(mml_FRASE0043, dbOpenTable)
                            rs.AddNew
                                rs!codigo = MaxCod("descalificaciones")
                                rs!cod_categoria = iCodCat
                                rs!fase = iFase
                                rs!id_juez = sIdJuez
                                rs!cod_Baile = rsBailes!codigo
                                rs!num_dorsal = rsDorsales!num_dorsal
                                rs!repesca = iHojaRepesca
                                If bJuezPasos Then
                                    GuardarBinary rs!anotacion, pbDesc
                                End If
                            rs.Update
                        rs.Close
                        pbDesc.Visible = False
                        VerMenu True
                        tbHojasProc.Visible = True
                        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
                    End If
                    rsDorsales.MoveNext
                    ix = ix + 2
                Wend
                rsDorsales.Close
                
                If Not bJuezPasos Then
                    'Comprobamos el avance de puestos por descalificaciones
                    Dim iCDescalif As Integer
                    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo, dbOpenSnapshot)
                    iCDescalif = rs.Fields(0)
                    rs.Close
                    If iCDescalif > 0 And G_DEC_POSICIONES_POR_DESCALIFICACION Then
                        Dim iCPuesto As Integer
                        
                        'Repasamos todos los dorsales y Buscamos los puestos vacios y desplazamos los siguientes
                        Set rs = db.OpenRecordset("SELECT num_dorsal, puesto FROM puntuaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND cod_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & " AND num_dorsal NOT IN (SELECT num_dorsal FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & ") ORDER BY puesto", dbOpenSnapshot)
                        iCPuesto = 1
                        While Not rs.EOF
                            db.Execute ("UPDATE puntuaciones SET puesto = " & iCPuesto & " WHERE cod_categoria=" & iCodCat & " AND cod_baile=" & rsBailes!codigo & " AND cod_juez='" & sIdJuez & "' AND fase=" & iFase & " AND repesca=" & iHojaRepesca & " AND num_dorsal = " & rs!num_dorsal)
                            Inc iCPuesto
                            rs.MoveNext
                        Wend
                        rs.Close
                        
                        'Ahora repasamos y ordenados los descalificados
                        If InStr(G_ORDENAR_DESCALIFICADOS_FINAL, sIdJuez) > 0 Then
                            If C_DESC_SIN_PUESTO Then
                                sSQL = "SELECT num_dorsal, puesto FROM puntuaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND cod_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & " AND puesto < " & C_ULTIMO_PUESTO & " AND num_dorsal IN (SELECT num_dorsal FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & ") ORDER BY puesto"
                            Else
                                sSQL = "SELECT num_dorsal, puesto FROM puntuaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND cod_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & " AND num_dorsal IN (SELECT num_dorsal FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & ") ORDER BY puesto"
                            End If
                            Debug.Print sSQL
                            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
                            While Not rs.EOF
                                db.Execute ("UPDATE puntuaciones SET puesto = " & iCPuesto & " WHERE cod_categoria=" & iCodCat & " AND cod_baile=" & rsBailes!codigo & " AND cod_juez='" & sIdJuez & "' AND fase=" & iFase & " AND repesca=" & iHojaRepesca & " AND num_dorsal = " & rs!num_dorsal)
                                ' Si se ordenan las descalificaciones no deben figurar
                                db.Execute ("DELETE FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND num_dorsal=" & rs!num_dorsal & " AND cod_baile=" & rsBailes!codigo)
                                Inc iCPuesto
                                rs.MoveNext
                            Wend
                            rs.Close
                        End If
                        
                    End If
                End If
Cont_Final:
                iy = iy + C_ANCHO_MARCAS_BAILE_FINAL
                Inc iCBailes
                rsBailes.MoveNext
            Wend
            rsBailes.Close
        End If
        Me.Caption = Me.Caption & mml_FRASE0908
        
        If bFalloHoja And G_NO_PROC_HOJAS_ERROR Then
            tbHojasProc.Text = sAntHojasProc
        Else
            db.Execute ("INSERT INTO hojas_reconocidas VALUES(" & iCodCat & ",'" & sIdJuez & IIf(iHoja2 = 1, "2", "") & "'," & iFase & "," & iTanda & "," & iHojaRepesca & ")")
        End If
        ' Comprobamos si hemos recopilado la información de todos los dorsales
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase = " & iFase, dbOpenSnapshot)
        iPuestos = rs.Fields(0)
        rs.Close
        ' Comprobamos el número de jueces
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & iCodCat, dbOpenSnapshot)
        iJueces = rs.Fields(0)
        rs.Close
        ' Comprobamos el número de jueces de pasos y figuras
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 1 AND cod_categoria = " & iCodCat, dbOpenSnapshot)
        ijuecespasos = rs.Fields(0)
        rs.Close
        ' Comprobamos el número de bailes
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & iCodCat & " AND fase = " & IIf(iFase > 1, 2, 1), dbOpenSnapshot)
        iBailes = rs.Fields(0)
        rs.Close
        ' Comprobamos el número de dorsales
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase = " & iFase, dbOpenSnapshot)
        iDorsales = rs.Fields(0)
        rs.Close
        
        'Localizamos el número de hojas de esta fase
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM hojas_reconocidas WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase, dbOpenSnapshot)
        iHojas = rs.Fields(0)
        rs.Close
        
        If GenerarControl(Val(tbCodComp.Text)) Then
            ' Cada juez tiene x tandas x hojas por tanda(Si bailes > 5) + el número de tandas de la hoja de control
            iMaxHojasRec = (iMaxTandas * iNumHojas * (iJueces + ijuecespasos) + iMaxTandas)
        Else
            ' Cada juez tiene x tandas x hojas por tanda(Si bailes > 5)
            iMaxHojasRec = iMaxTandas * iNumHojas * (iJueces + ijuecespasos)
        End If
        
        Me.Caption = mml_FRASE0909 & iPuestos & mml_FRASE0472 & iDorsales * iBailes * iJueces & mml_FRASE0910 & iHojas & mml_FRASE0472 & iMaxHojasRec ' Mas 1 de control
        If bJuezPasos Then
            Me.Caption = Me.Caption & mml_FRASE0911
        End If
        If iHojaRepesca Then
            Me.Caption = Me.Caption & mml_FRASE0912
        End If
        
        igCodCat = iCodCat
        igFase = iFase
        igHojaRepesca = iHojaRepesca
        
        If (Not BailesParciales(iCodCat) And (iMaxHojasRec = iHojas Or G_NO_CONTAR_HOJAS) And iPuestos = iDorsales * iBailes * iJueces) Or _
           (BailesParciales(iCodCat) And ((iMaxHojasRec Mod iHojas = 0 And iHojas > 1) Or G_NO_CONTAR_HOJAS) And (iPuestos Mod (iDorsales * iJueces)) = 0) Then
            If MsgBox(mml_FRASE0913 & iCodCat & mml_FRASE0641 & iFase & mml_FRASE0914, vbYesNo Or vbQuestion, mml_FRASE0086) = vbYes Then
                frmCalcular.tbCodComp.Text = tbCodComp.Text
                frmCalcular.tbCodCat.Text = iCodCat
                frmCalcular.tbCodFase.Text = iFase
                frmCalcular.chkRep.Value = iHojaRepesca
                tbHojasProc.Text = C_HOJAS_PROC
                frmCalcular.MostrarCalcular
            End If
        End If
    End If
error:
Dim Msj
If Err.Number <> 0 Then
   Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
         & Err.Source & Chr(13) & Err.Description
   MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
End If
End Sub

Function ComprobarReferencia(iXi As Integer, iYi As Integer) As Integer

pbFicha.CurrentX = aMarcas(iXi - 1, 0).iXi
pbFicha.CurrentY = aMarcas(C_MAX_MARCAS_X, iYi).iYi
pbFicha.Print aMarcas(iXi, iYi).iXf
ComprobarReferencia = aMarcas(iXi, iYi).iXf

End Function

Function ComprobarMarcas(iXi As Integer, iYi As Integer, iXf As Integer, iYf As Integer) As Boolean
Dim X As Integer, Y As Integer
Dim sAsume As String
    ComprobarMarcas = True
    For X = iXi To iXf
        For Y = iYi To iYf
            If aMarcas(X, Y).iXi = -1 Then
                If aMarcas(X, Y).iXf >= C_MEDIO Then
                    aMarcas(X, Y).iXi = 1
                    sAsume = mml_FRASE0892
                Else
                    aMarcas(X, Y).iXi = 0
                    sAsume = mml_FRASE0915
                End If
                GrabarPosicion Me
                Maximizar
                aMarcas(X, Y).iXi = frmMarca.Visualizar(X, Y, aMarcas(X, Y).iXi, aMarcas(X, Y).iXf, bFalloHoja)
                RestaurarPosicion Me
                ComprobarMarcas = False
            End If
        Next
    Next
End Function

Sub Maximizar()
    frmRecOptico.Top = 0
    frmRecOptico.Left = 0
    frmRecOptico.Height = Screen.Height
End Sub

Sub VerMenu(bVisible As Boolean)
Static iAncho As Integer
    On Local Error Resume Next
    If frmMenu.Visible Then AppActivate C_TITULO_VENTANA_PRINCIPAL
    frmRecOptico.Show vbNomodal, Me
    ProcesarEventos
    If Not bVisible Then
        iAncho = frmMenu.Width
        frmMenu.Width = 0
    ElseIf iAncho > 0 Then
        frmMenu.Width = iAncho
    End If
    'frmMenu.Visible = bVisible
    On Local Error GoTo 0
End Sub


Sub GrabarInfoHojaExt()
Dim rs As Recordset
Dim ix As Integer
Dim iy As Integer
Dim iCodCat As Integer
Dim iFase As Integer
Dim sIdJuez As String
Dim iTanda As Integer
Dim iMaxTandas As Integer
Dim iMaxDorsales As Integer
Dim rsBailes As Recordset
Dim rsDorsales As Recordset
Dim iCBailes As Integer
Dim iCDorsales As Integer
Dim iMarca As Integer
Dim iPuesto As Integer
Dim iCPuestos As Integer
Dim aPuestos(7) As Integer
Dim i As Integer, j As Integer
Dim iDorsales As Integer, iJueces As Integer, iBailes As Integer, iPuestos As Integer
Dim iHojas As Integer
Dim bJuezPasos As Boolean
Dim ijuecespasos As Integer
Dim iAlto As Integer, iAncho As Integer
Dim iXini As Integer, iYini As Integer, iAltoZona As Integer, iAnchoZona As Integer
Dim lColor As Long, dEsX As Double, dEsY As Double
Dim bHojaControl As Boolean
Dim iHojaRepesca As Integer
Dim iHoja2 As Integer
Dim iHojaExt As Integer
Dim iNumHojas As Integer
Dim iMaxHojasRec As Integer
Dim iContMarcas As Integer
Dim iCBailesHoja As Integer
Dim iAlturaRecOptico As Integer
Dim iDorsalInicialTanda As Integer
Dim iDorsalesTanda As Integer
Dim bTeamMatch As Boolean
Dim sAntHojasProc As String

On Local Error GoTo error

    If tbCodComp.Text = "" Then
        CamposSinCubrir
        cmdGrabar.Enabled = True
        sNombreFicha = ""
        Exit Sub
    End If
    
    Me.Caption = mml_FRASE0916 & sFichero
    VerMenu False
    
    
    'Comprobar referencia de ajuste
    ComprobarReferencia C_MAX_CATEG_POR_LINEA_EXT + 1, 6
    
    'Primero debemos comprobar que todas las marcas contienen marcas o espacios
    ' con total seguridad
    'Comprobamos la cabecera
    ComprobarMarcas 2, G_POS_INIC_CATEG, C_MAX_CATEG_POR_LINEA_EXT + 1, 5
    'Fases
    ComprobarMarcas 2, 7, 10, 7
    'Jueces
    ComprobarMarcas 2, 8, C_MAX_JUECES_EXT / 2, 9
    'Tandas
    ComprobarMarcas 2, 11, C_MAX_CATEG_POR_LINEA_EXT + 1, 11
    
    'Fallos
    ComprobarMarcas C_POS_FALLO_EXT, 7, C_POS_FALLO_EXT, 7
    'Hoja extendida
    ComprobarMarcas C_POS_HOJA_EXT_EXT, 7, C_POS_HOJA_EXT_EXT, 7
    'Control
    ComprobarMarcas C_POS_CONTROL_EXT / 2, 9, C_POS_CONTROL_EXT / 2, 9
    'Repesca
    ComprobarMarcas C_POS_REPESCA_EXT / 2, 9, C_POS_REPESCA_EXT / 2, 9
    'Hoja2
    ComprobarMarcas C_POS_HOJA2_EXT / 2, 9, C_POS_HOJA2_EXT / 2, 9
    VerMenu True
    
    cmdGrabar.Enabled = False
    
    'Recuperamos la información de la categoría
    ' categorías
    Set rs = db.OpenRecordset("SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY 1", dbOpenSnapshot)
    ix = 2
    iy = G_POS_INIC_CATEG
    Do While Not rs.EOF
        If ix Mod (C_MAX_CATEG_POR_LINEA_EXT + 2) = 0 Then
            Inc iy
            ix = 2
        End If
        If aMarcas(ix, iy).iXi = 1 Then
            iCodCat = rs!codigo
            Exit Do
        End If
        rs.MoveNext
        Inc ix
    Loop
    rs.Close
    
    bTeamMatch = ComprobarSiTeamMatch(iCodCat)
    
    If iCodCat = 0 Then
        GrabarPosicion Me
        Maximizar
        MsgBox mml_FRASE0810, vbOKOnly Or vbCritical, mml_FRASE0096
        RestaurarPosicion Me
        Exit Sub
    End If
    
    'Localizamos la descripción de la categoria para crear el directorio
    Set rs = db.OpenRecordset("SELECT codigo, descripcion FROM categorias WHERE codigo = " & iCodCat, dbOpenSnapshot)
    If Not rs.EOF Then
        If Directorio.sDirectorio = rs!DESCRIPCION & "_" & rs!codigo Then
            Directorio.bNuevo = False
        Else
            Directorio.sDirectorio = rs!DESCRIPCION & "_" & rs!codigo
            Directorio.bNuevo = True
        End If
    End If
    rs.Close
    
    ' fase
    ix = 2
    iy = 7
    
    Do While ix <= 10
        If aMarcas(ix, iy).iXi = 1 Then
            iFase = 2 ^ (ix - 2)
            Exit Do
        End If
        Inc ix
    Loop
    
    'jueces
    Set rs = db.OpenRecordset("SELECT id_juez, pasos FROM juez_categ WHERE cod_categoria = " & iCodCat & " ORDER BY 1", dbOpenSnapshot)
    ix = 2
    iy = 8
    While Not rs.EOF
        If ix Mod (C_MAX_JUECES_EXT / 2 + 2) = 0 Then
            Inc iy
            ix = 2
        End If
        If aMarcas(ix, iy).iXi = 1 Then
            sIdJuez = rs!id_juez
            If rs!pasos = 1 Then
                bJuezPasos = True
            Else
                bJuezPasos = False
            End If
        End If
        Inc ix
        rs.MoveNext
    Wend
    rs.Close
    
    ix = C_POS_CONTROL_EXT / 2
    iy = 9
    'Hoja de control
    If aMarcas(ix, iy).iXi = 1 Then
        bHojaControl = True
        sIdJuez = "Control"
    Else
        bHojaControl = False
    End If
    
    ix = C_POS_REPESCA_EXT / 2
    iy = 9
    'Hoja de repesca
    If aMarcas(ix, iy).iXi = 1 Then
        iHojaRepesca = 1
    Else
        iHojaRepesca = 0
    End If
    
    ix = C_POS_HOJA2_EXT / 2
    iy = 9
    'Hoja2
    If aMarcas(ix, iy).iXi = 1 Then
        iHoja2 = 1
    Else
        iHoja2 = 0
    End If
    
    ix = C_POS_HOJA_EXT_EXT
    iy = 7
    'Hoja Extendida
    If aMarcas(ix, iy).iXi = 1 Then
        iHojaExt = 1
    Else
        iHojaExt = 0
    End If
    
    ' tanda
    ix = 2
    iy = 11
    
    Do While ix <= C_MAX_MARCAS_X_EXT - 1
        If aMarcas(ix, iy).iXi = 1 Then
            iTanda = IIf(iTanda = 0, (ix - 1), iTanda)
            iMaxTandas = IIf(iMaxTandas < (ix - 1), (ix - 1), iMaxTandas)
        End If
        Inc ix
    Loop
    
    Me.Caption = mml_FRASE0811 & iCodCat & mml_FRASE0641 & iFase & mml_FRASE0485 & sIdJuez & IIf(bJuezPasos, mml_FRASE0812, "") & mml_FRASE0813 & iTanda & mml_FRASE0814 & iMaxTandas & IIf(bHojaControl, mml_FRASE0815, "") & mml_FRASE0816 & iHojaRepesca & mml_FRASE0817 & iHoja2 & " - " & sFichero & "Ext " & iHojaExt
    sNombreFicha = "C" & iCodCat & "F" & iFase & "J" & sIdJuez & "T" & iTanda & IIf(bHojaControl, "CT", "NC") & "R" & iHojaRepesca & "H2" & iHoja2
    
    'Recuperamos la información de si ha habido fallos
    If aMarcas(C_POS_FALLO_EXT, 7).iXi = 1 Then
        GrabarPosicion Me
        Maximizar
        If MsgBox(mml_FRASE0818, vbYesNo Or vbCritical, mml_FRASE0084) = vbNo Then
            RestaurarPosicion Me
            bFalloHoja = True
            Exit Sub
        End If
        RestaurarPosicion Me
    End If
    
    ' Comprobamos el número de bailes
    Set rsBailes = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ bc WHERE bc.cod_categoria = " & iCodCat & " AND bc.fase = " & IIf(iFase = 1, 1, 2), dbOpenSnapshot)
    iCBailes = rsBailes.Fields(0)
    rsBailes.Close
    
    If iCBailes > C_BAILES_POR_HOJA Then
        iNumHojas = 2
    Else
        iNumHojas = 1
    End If
    
    ' Comprobamos el número de dorsales
    Set rsDorsales = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase =" & iFase & " ORDER BY 1", dbOpenSnapshot)
        iMaxDorsales = rsDorsales.Fields(0)
    rsDorsales.Close
    
    If iCBailes > C_BAILES_POR_HOJA Then
        If iHoja2 = 1 Then
            iCBailesHoja = iCBailes - C_BAILES_POR_HOJA
        Else
            iCBailesHoja = C_BAILES_POR_HOJA
        End If
    Else
        iCBailesHoja = iCBailes
    End If
    
    'Comprobamos los bailes
    If bHojaControl Then
        'Comprobar marcas de ausente y anulación
        ComprobarMarcas 2, 13, C_MAX_MARCAS_X_EXT - 2, 14
    ElseIf iFase = 1 Then ' FINAL
        ' Comprobamos marcas de descalificación
        For i = 4 To 4 + (iMaxDorsales - 1) * 2 Step 2
            For j = 0 To iCBailesHoja - 1
                ComprobarMarcas i, 12 + j * C_ANCHO_MARCAS_BAILE_FINAL, i, 12 + j * C_ANCHO_MARCAS_BAILE_FINAL
            Next
        Next
        If Not bJuezPasos Then
            ' Comprobamos marcas de puntuación y anulación
            For i = 0 To iCBailesHoja - 1
                ComprobarMarcas 3, 13 + i * C_ANCHO_MARCAS_BAILE_FINAL, 3 + (iMaxDorsales - 1) * 2, 19 + i * C_ANCHO_MARCAS_BAILE_FINAL
            Next
        End If
        'Comprobar las marcas de la presencia de bailes
        For j = 0 To iCBailesHoja - 1
            ComprobarMarcas C_REC_POS_X_MARCA_BAILE_EXT, 12 + j * C_ANCHO_MARCAS_BAILE_FINAL, C_REC_POS_X_MARCA_BAILE_EXT, 12 + j * C_ANCHO_MARCAS_BAILE_FINAL
        Next
    Else ' FASES ELIMINATORIAS
        If Not bJuezPasos Then
            For i = 0 To iCBailesHoja - 1
                ComprobarMarcas 2, 13 + i * C_ANCHO_MARCAS_BAILE, C_MAX_MARCAS_X_EXT - 2, 15 + i * C_ANCHO_MARCAS_BAILE
            Next
        Else
            For i = 0 To iCBailesHoja - 1
                ComprobarMarcas 2, 15 + i * C_ANCHO_MARCAS_BAILE, C_MAX_MARCAS_X_EXT - 2, 15 + i * C_ANCHO_MARCAS_BAILE
            Next
        End If
        'Comprobar las marcas de la presencia de bailes
        For j = 0 To iCBailesHoja - 1
            ComprobarMarcas C_REC_POS_X_MARCA_BAILE_EXT, 16 + j * C_ANCHO_MARCAS_BAILE, C_REC_POS_X_MARCA_BAILE_EXT, 16 + j * C_ANCHO_MARCAS_BAILE
        Next
    End If
    
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM hojas_reconocidas WHERE cod_categoria = " & iCodCat & " AND id_juez='" & sIdJuez & IIf(iHoja2 = 1, "2", "") & "' AND fase = " & iFase & " AND tanda = " & iTanda & " AND repesca = " & iHojaRepesca, dbOpenSnapshot)
    If Not rs.EOF Then
        If Not BailesParciales(iCodCat) And rs.Fields(0) > 0 Then
            GrabarPosicion Me
            Maximizar
            If MsgBox(mml_FRASE0820, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                RestaurarPosicion Me
                Exit Sub
            End If
            RestaurarPosicion Me
        End If
    End If
    rs.Close
        
    sAntHojasProc = tbHojasProc.Text
    tbHojasProc.Text = tbHojasProc.Text & RTrim$(Left$(sIdJuez & " ", 2)) & Trim$(Str$(iTanda)) & " "
    'db.Execute ("INSERT INTO hojas_reconocidas VALUES(" & iCodCat & ",'" & sIdJuez & IIf(iHoja2 = 1, "2", "") & "'," & iFase & "," & iTanda & "," & iHojaRepesca & ")")
    
    If bHojaControl Then
        iCDorsales = 0
        ix = 2
        iy = 13
        Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase =" & iFase & " ORDER BY 1", dbOpenSnapshot)
        iDorsalInicialTanda = CalcularDorsalInicialTandaCat(iCodCat, iFase, iHojaRepesca, iTanda, iMaxTandas, iDorsalesTanda)
        While iCDorsales < iDorsalInicialTanda - 1 And Not rsDorsales.EOF
            Inc iCDorsales
            rsDorsales.MoveNext
        Wend
        iCDorsales = 0
        While iCDorsales < iDorsalesTanda And Not rsDorsales.EOF
            ' Comprobamos la casilla de anulación
            iMarca = IIf(2 * aMarcas(ix, iy).iXi + aMarcas(ix, iy + 1).iXi = 2, 1, 0)
            If iMarca = 1 Then
                Debug.Print "UPDATE dorsales SET no_presentado = " & iMarca & "  WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND cod_categoria =" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase
                db.Execute ("UPDATE dorsales SET no_presente = " & iMarca & " WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND cod_categoria =" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase)
            End If
            rsDorsales.MoveNext
            Inc iCDorsales
            Inc ix
        Wend
    Else
    ' Hojas de puntuación -----------------------------------------------------------
        'Reconocer los bailes
        iCBailes = 0
        iCDorsales = 0
        If iFase > 1 Then 'Eliminatorias ********************************************************
            iy = 13
            Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & iCodCat & " AND bc.fase = 2 ORDER BY posicion", dbOpenSnapshot)
            If iHoja2 = 1 Then
                While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                    Inc iCBailes
                    rsBailes.MoveNext
                Wend
                iCBailes = 0
            End If
            While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                iContMarcas = 0
                iCDorsales = 0
                ix = 2
                
                If G_REC_HOJA_EXT And iHojaExt = 1 And BailesParciales(iCodCat) And aMarcas(C_REC_POS_X_MARCA_BAILE_EXT, iy + 3).iXi = 0 Then
                    GoTo Cont_Semi
                End If
                If BailesParciales(iCodCat) Then sNombreFicha = "B" & rsBailes!codigo & sNombreFicha
                
                If iMaxTandas > 1 And G_DORSALES_COMBINADOS And CombinarDorsalesCateg(iCodCat) Then
                    Set rsDorsales = db.OpenRecordset("SELECT d.num_dorsal, dc.orden FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & iCodCat & " AND d.repesca=" & iHojaRepesca & " AND d.fase =" & iFase & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
                Else
                    Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase =" & iFase & " ORDER BY 1", dbOpenSnapshot)
                End If
                
                iDorsalInicialTanda = CalcularDorsalInicialTandaCat(iCodCat, iFase, iHojaRepesca, iTanda, iMaxTandas, iDorsalesTanda)
                While iCDorsales < iDorsalInicialTanda - 1 And Not rsDorsales.EOF
                    Inc iCDorsales
                    rsDorsales.MoveNext
                Wend
                iCDorsales = 0
                While iCDorsales < iDorsalesTanda And Not rsDorsales.EOF
                    ' Comprobamos la casilla de anulación
                    iMarca = IIf(2 * aMarcas(ix, iy).iXi + aMarcas(ix, iy + 1).iXi = 2, 1, 0)
                    If iMarca = 1 Then Inc iContMarcas
                    Debug.Print "INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "'," & iMarca & "," & iFase & "," & iHojaRepesca & ")"
                    If Not bJuezPasos Then
                        'Borramos cualquier marca igual
                        db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND cod_categoria=" & iCodCat & " AND cod_baile=" & rsBailes!codigo & " AND cod_juez='" & sIdJuez & "' AND fase=" & iFase & " AND repesca=" & iHojaRepesca)
                        ' Si el dorsal está descalificado no le asignamos puntuación aunque la tenga
                        If (Not bJuezPasos And aMarcas(ix, iy + 2).iXi = 1) Or (bJuezPasos And aMarcas(ix, iy).iXi = 1) Then
                            db.Execute ("INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "',0," & iFase & "," & iHojaRepesca & ")")
                        Else
                            db.Execute ("INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "'," & iMarca & "," & iFase & "," & iHojaRepesca & ")")
                        End If
                    End If
                    
                    'Descalificación
                    db.Execute ("DELETE FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND num_dorsal=" & rsDorsales!num_dorsal & " AND cod_baile = " & rsBailes!codigo)
                    If (Not bJuezPasos And aMarcas(ix, iy + 2).iXi = 1) Or (bJuezPasos And aMarcas(ix, iy).iXi = 1) Then
                        If bJuezPasos Then
                            'Agrandamos la pantalla
                            iAlturaRecOptico = 0
                            If frmRecOptico.Height < C_ALTURA_LECT_OPTICA Then
                                iAlturaRecOptico = frmRecOptico.Height
                                frmRecOptico.Height = C_ALTURA_LECT_OPTICA
                                ProcesarEventos
                            End If
                            
                            iAncho = aMarcas(2, 0).iXf - aMarcas(2, 0).iXi + 1
                            iAlto = aMarcas(2, 0).iYf - aMarcas(2, 0).iYi + 1
                            iXini = aMarcas(1, 0).iXf
                            iYini = aMarcas(C_MAX_MARCAS_X, iy + 1).iYi
                            iAltoZona = iAlto * 5.5
                            iAnchoZona = iAncho * 36
                            dEsX = (pbDesc.ScaleWidth / pbDesc.Width)
                            dEsY = (pbDesc.ScaleHeight / pbDesc.Height)
                            pbDesc.Visible = True
                            pbDesc.Cls
                            pbDesc.Refresh
                            tbHojasProc.Visible = False
                            VerMenu False
                            frmBarra.Visible = False
                            DoEvents: DoEvents: DoEvents
                            For i = iXini To iXini + iAnchoZona
                                For j = iYini To iYini + iAltoZona
                                    lColor = pbFicha.Point(i, j)
                                    If ValorColor(lColor) < C_UMBRAL Or lColor = C_COLOR_MARCA Then
                                        lColor = 0
                                    Else
                                        lColor = &HFFFFFF
                                    End If
                                    pbDesc.PSet ((i - iXini) * dEsX, (j - iYini) * dEsY), lColor
                                Next j
                                If i Mod 50 = 0 Then DoEvents
                            Next i
                            SavePicture CaptureWindow(pbDesc.hwnd, True, 0, 0, (pbDesc.Width) - 4, (pbDesc.Height) - 4), "desc.bmp"
                            pbDesc.Picture = LoadPicture("Desc.bmp")
                            frmBarra.Visible = True
                            If iAlturaRecOptico > 0 Then
                                frmRecOptico.Height = iAlturaRecOptico
                                ProcesarEventos
                            End If
                        End If
                        Set rs = db.OpenRecordset(mml_FRASE0043, dbOpenTable)
                            rs.AddNew
                                rs!codigo = MaxCod("descalificaciones")
                                rs!cod_categoria = iCodCat
                                rs!fase = iFase
                                rs!id_juez = sIdJuez
                                rs!cod_Baile = rsBailes!codigo
                                rs!num_dorsal = rsDorsales!num_dorsal
                                rs!repesca = iHojaRepesca
                                If bJuezPasos Then
                                    GuardarBinary rs!anotacion, pbDesc
                                End If
                            rs.Update
                        rs.Close
                        pbDesc.Visible = False
                        tbHojasProc.Visible = True
                        VerMenu True
                        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
                    End If
                    rsDorsales.MoveNext
                    Inc iCDorsales
                    Inc ix
                Wend
                If G_AVISO_NUM_MARCAS Then
                ' Si hay varias tandas permitimos marcar una menos o más en cada tanda
                Dim iMargen As Integer
                    If iMaxTandas > 1 Then
                        iMargen = 1
                    Else
                        iMargen = 0
                    End If
                    If iHojaRepesca = 0 And Not bJuezPasos Then
                        If Abs(iContMarcas - (iFase / 2 * 6) / iMaxTandas) > iMargen Then
                            GrabarPosicion Me
                            Maximizar
                            DesplazarVisualizacionHoja iy
                            MsgBox mml_FRASE0821 & rsBailes!Nombre & mml_FRASE0822 & (iFase / 2 * 6) / iMaxTandas & mml_FRASE0823 & iContMarcas, vbOKOnly Or vbInformation, mml_FRASE0084
                            DesplazarVisualizacionHoja 0
                            RestaurarPosicion Me
                        End If
                    End If
                End If
                rsDorsales.Close
Cont_Semi:
                iy = iy + C_ANCHO_MARCAS_BAILE
                rsBailes.MoveNext
                Inc iCBailes
            Wend
            rsBailes.Close
        Else ' FINAL ********************************************************************
            iy = 13
            Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & iCodCat & " AND bc.fase = 1 ORDER BY posicion", dbOpenSnapshot)
            iCBailes = 0
            If iHoja2 = 1 Then
                While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                    Inc iCBailes
                    rsBailes.MoveNext
                Wend
                iCBailes = 0
            End If
            While Not rsBailes.EOF And iCBailes < C_BAILES_POR_HOJA
                ix = 3
                
                If G_REC_HOJA_EXT And iHojaExt = 1 And BailesParciales(iCodCat) And aMarcas(C_REC_POS_X_MARCA_BAILE_EXT, iy - 1).iXi = 0 Then
                    'Reconocimiento parcial y baile no localizado
                    GoTo Cont_Final
                End If
                If BailesParciales(iCodCat) Then sNombreFicha = "B" & rsBailes!codigo & sNombreFicha
                
                Set rsDorsales = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase =" & iFase & " ORDER BY 1", dbOpenSnapshot)
                iCPuestos = 0
                While Not rsDorsales.EOF
                    If Not bJuezPasos Then
                        'Comprobamos el puesto
                        iPuesto = 0
                        For i = 0 To IIf(bTeamMatch, C_MAX_POS_TEAMMATCH, iMaxDorsales - 1)
                            ' Comprobamos la casilla de anulación
                            iMarca = IIf(2 * aMarcas(ix, iy + i).iXi + aMarcas(ix + 1, iy + i).iXi = 2, 1, 0)
                            If iMarca = 1 Then
                                ' Comprobamos que un dorsal no tenga varios puestos
                                If iPuesto <> 0 Then
                                    GrabarPosicion Me
                                    Maximizar
                                    DesplazarVisualizacionHoja iy
                                    MsgBox mml_FRASE0824 & rsDorsales!num_dorsal & mml_FRASE0489 & rsBailes!Nombre, vbOKOnly Or vbCritical, mml_FRASE0096
                                    DesplazarVisualizacionHoja 0
                                    RestaurarPosicion Me
                                    bFalloHoja = True
                                Else
                                    iPuesto = i + 1
                                    ' Comprobamos que no se repitan puestos
                                    If Not bTeamMatch Then
                                        For j = 0 To iCPuestos - 1
                                            If aPuestos(j) = iPuesto Then
                                                GrabarPosicion Me
                                                Maximizar
                                                DesplazarVisualizacionHoja iy
                                                MsgBox mml_FRASE0825 & iPuesto & mml_FRASE0826 & rsDorsales!num_dorsal & mml_FRASE0489 & rsBailes!Nombre, vbOKOnly Or vbCritical, mml_FRASE0096
                                                DesplazarVisualizacionHoja 0
                                                RestaurarPosicion Me
                                                bFalloHoja = True
                                            End If
                                        Next
                                        aPuestos(iCPuestos) = iPuesto
                                        Inc iCPuestos
                                    End If
                                End If
                            End If
                        Next
                        ' Si se ha descalificado el dorsal le asignamos el último puesto
                        ' Comprobamos que se haya marcado un puesto
                        If iPuesto = 0 Or aMarcas(ix + 1, iy - 1).iXi = 1 Then
                            'Comprobamos que el dorsal está marcado como persente
                            Dim rsDesc As Recordset, iNoPresente As Integer
                            Set rsDesc = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND fase = " & iFase & " AND repesca = " & iHojaRepesca & " AND num_dorsal = " & rsDorsales!num_dorsal & " AND no_presente > 0", dbOpenSnapshot)
                            iNoPresente = rsDesc.Fields(0)
                            rsDesc.Close
                            ' Si no hay puesto y dosales < 8 y no se ha descalificado
                            ' y está presente se avisa de la situación
                            If iNoPresente = 0 And iPuesto = 0 And iMaxDorsales < 8 And aMarcas(ix + 1, iy - 1).iXi = 0 Then
                                GrabarPosicion Me
                                Maximizar
                                DesplazarVisualizacionHoja iy
                                MsgBox mml_FRASE0827 & rsDorsales!num_dorsal & mml_FRASE0489 & rsBailes!Nombre & mml_FRASE0828, vbOKOnly Or vbCritical, mml_FRASE0096
                                DesplazarVisualizacionHoja 0
                                RestaurarPosicion Me
                                bFalloHoja = True
                            End If
                            'Si la marca de descalificación se emplea para enviar al final
                            ', pero ordenados. se deja el puesto
                            ' para conservar el puesto tiene que existir, estár descalificado y el juez ordena las descalificaciones
                            If Not (iPuesto > 0 And aMarcas(ix + 1, iy - 1).iXi = 1 And InStr(G_ORDENAR_DESCALIFICADOS_FINAL, sIdJuez) > 0) Then
                                'Si está descalificado -> detrás del puesto mas alto posible
                                If aMarcas(ix + 1, iy - 1).iXi = 1 Then
                                    iPuesto = C_ULTIMO_PUESTO
                                Else
                                    iPuesto = iMaxDorsales
                                End If
                            End If
                        End If
                        Debug.Print "INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "'," & iPuesto & "," & iFase & "," & iHojaRepesca & ")"
                        'Borramos cualquier marca igual
                        db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal=" & rsDorsales!num_dorsal & " AND cod_categoria=" & iCodCat & " AND cod_baile=" & rsBailes!codigo & " AND cod_juez='" & sIdJuez & "' AND fase=" & iFase & " AND repesca=" & iHojaRepesca)
                        If bTeamMatch Then
                            db.Execute ("INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "','" & CDbl(iPuesto + 1) / 2 & "'," & iFase & "," & iHojaRepesca & ")")
                        Else
                            db.Execute ("INSERT INTO puntuaciones VALUES(" & rsDorsales!num_dorsal & "," & iCodCat & "," & rsBailes!codigo & ",'" & sIdJuez & "'," & iPuesto & "," & iFase & "," & iHojaRepesca & ")")
                        End If
                    End If
                    'Descalificación
                    db.Execute ("DELETE FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND num_dorsal=" & rsDorsales!num_dorsal & " AND cod_baile=" & rsBailes!codigo)
                    If aMarcas(ix + 1, iy - 1).iXi = 1 Then
                        If bJuezPasos Then
                            'Agrandamos la pantalla
                            iAlturaRecOptico = 0
                            If frmRecOptico.Height < C_ALTURA_LECT_OPTICA Then
                                iAlturaRecOptico = frmRecOptico.Height
                                frmRecOptico.Height = C_ALTURA_LECT_OPTICA
                                ProcesarEventos
                            End If
                            
                            iAncho = aMarcas(2, 0).iXf - aMarcas(2, 0).iXi + 1
                            iAlto = aMarcas(2, 0).iYf - aMarcas(2, 0).iYi + 1
                            iXini = aMarcas(2, 0).iXi
                            iYini = aMarcas(C_MAX_MARCAS_X, iy).iYi
                            iAltoZona = iAlto * 11
                            iAnchoZona = iAncho * 34
                            dEsX = (pbDesc.ScaleWidth / pbDesc.Width)
                            dEsY = (pbDesc.ScaleHeight / pbDesc.Height)
                            pbDesc.Cls
                            tbHojasProc.Visible = False
                            pbDesc.Visible = True
                            VerMenu False
                            frmBarra.Visible = False
                            pbDesc.Refresh
                            DoEvents: DoEvents: DoEvents
                            For i = iXini To iXini + iAnchoZona
                                For j = iYini To iYini + iAltoZona
                                    lColor = pbFicha.Point(i, j)
                                    If ValorColor(lColor) < C_UMBRAL Or lColor = C_COLOR_MARCA Then
                                        lColor = 0
                                    Else
                                        lColor = &HFFFFFF
                                    End If
                                    'If G_MARCAR_PUNTOS Then pbDesc.PSet ((i - iXini) * dEsX, (j - iYini) * dEsY), lColor
                                    pbDesc.PSet ((i - iXini) * dEsX, (j - iYini) * dEsY), lColor
                                Next j
                                If i Mod 50 = 0 Then DoEvents
                            Next i
                            SavePicture CaptureWindow(pbDesc.hwnd, True, 0, 0, (pbDesc.Width) - 4, (pbDesc.Height) - 4), "desc.bmp"
                            pbDesc.Picture = LoadPicture("Desc.bmp")
                            frmBarra.Visible = True
                            If iAlturaRecOptico > 0 Then
                                frmRecOptico.Height = iAlturaRecOptico
                                ProcesarEventos
                            End If
                        End If
                        Set rs = db.OpenRecordset(mml_FRASE0043, dbOpenTable)
                            rs.AddNew
                                rs!codigo = MaxCod("descalificaciones")
                                rs!cod_categoria = iCodCat
                                rs!fase = iFase
                                rs!id_juez = sIdJuez
                                rs!cod_Baile = rsBailes!codigo
                                rs!num_dorsal = rsDorsales!num_dorsal
                                rs!repesca = iHojaRepesca
                                If bJuezPasos Then
                                    GuardarBinary rs!anotacion, pbDesc
                                End If
                            rs.Update
                        rs.Close
                        pbDesc.Visible = False
                        VerMenu True
                        tbHojasProc.Visible = True
                        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
                    End If
                    rsDorsales.MoveNext
                    ix = ix + 2
                Wend
                rsDorsales.Close

                If Not bJuezPasos Then
                    'Comprobamos el avance de puestos por descalificaciones
                    Dim iCDescalif As Integer
                    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo, dbOpenSnapshot)
                    iCDescalif = rs.Fields(0)
                    rs.Close
                    If iCDescalif > 0 And G_DEC_POSICIONES_POR_DESCALIFICACION Then
                        Dim iCPuesto As Integer
                        
                        'Repasamos todos los dorsales y Buscamos los puestos vacios y desplazamos los siguientes
                        Set rs = db.OpenRecordset("SELECT num_dorsal, puesto FROM puntuaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND cod_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & " AND num_dorsal NOT IN (SELECT num_dorsal FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & ") ORDER BY puesto", dbOpenSnapshot)
                        iCPuesto = 1
                        While Not rs.EOF
                            db.Execute ("UPDATE puntuaciones SET puesto = " & iCPuesto & " WHERE cod_categoria=" & iCodCat & " AND cod_baile=" & rsBailes!codigo & " AND cod_juez='" & sIdJuez & "' AND fase=" & iFase & " AND repesca=" & iHojaRepesca & " AND num_dorsal = " & rs!num_dorsal)
                            Inc iCPuesto
                            rs.MoveNext
                        Wend
                        rs.Close
                        
                        'Ahora repasamos y ordenados los descalificados
                        If InStr(G_ORDENAR_DESCALIFICADOS_FINAL, sIdJuez) > 0 Then
                            If C_DESC_SIN_PUESTO Then
                                sSQL = "SELECT num_dorsal, puesto FROM puntuaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND cod_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & " AND puesto < " & C_ULTIMO_PUESTO & " AND num_dorsal IN (SELECT num_dorsal FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & ") ORDER BY puesto"
                            Else
                                sSQL = "SELECT num_dorsal, puesto FROM puntuaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND cod_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & " AND num_dorsal IN (SELECT num_dorsal FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND cod_baile=" & rsBailes!codigo & ") ORDER BY puesto"
                            End If
                            Debug.Print sSQL
                            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
                            While Not rs.EOF
                                db.Execute ("UPDATE puntuaciones SET puesto = " & iCPuesto & " WHERE cod_categoria=" & iCodCat & " AND cod_baile=" & rsBailes!codigo & " AND cod_juez='" & sIdJuez & "' AND fase=" & iFase & " AND repesca=" & iHojaRepesca & " AND num_dorsal = " & rs!num_dorsal)
                                ' Si se ordenan las descalificaciones no deben figurar
                                db.Execute ("DELETE FROM descalificaciones WHERE cod_categoria=" & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase & " AND id_juez='" & sIdJuez & "' AND num_dorsal=" & rs!num_dorsal & " AND cod_baile=" & rsBailes!codigo)
                                Inc iCPuesto
                                rs.MoveNext
                            Wend
                            rs.Close
                        End If
                        
                    End If
                End If
Cont_Final:
                iy = iy + C_ANCHO_MARCAS_BAILE_FINAL
                Inc iCBailes
                rsBailes.MoveNext
            Wend
            rsBailes.Close
        End If
        Me.Caption = Me.Caption & mml_FRASE0908
        
        If bFalloHoja And G_NO_PROC_HOJAS_ERROR Then
            tbHojasProc.Text = sAntHojasProc
        Else
            db.Execute ("INSERT INTO hojas_reconocidas VALUES(" & iCodCat & ",'" & sIdJuez & IIf(iHoja2 = 1, "2", "") & "'," & iFase & "," & iTanda & "," & iHojaRepesca & ")")
        End If
        ' Comprobamos si hemos recopilado la información de todos los dorsales
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase = " & iFase, dbOpenSnapshot)
        iPuestos = rs.Fields(0)
        rs.Close
        ' Comprobamos el número de jueces
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & iCodCat, dbOpenSnapshot)
        iJueces = rs.Fields(0)
        rs.Close
        ' Comprobamos el número de jueces de pasos y figuras
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 1 AND cod_categoria = " & iCodCat, dbOpenSnapshot)
        ijuecespasos = rs.Fields(0)
        rs.Close
        ' Comprobamos el número de bailes
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & iCodCat & " AND fase = " & IIf(iFase > 1, 2, 1), dbOpenSnapshot)
        iBailes = rs.Fields(0)
        rs.Close
        ' Comprobamos el número de dorsales
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase = " & iFase, dbOpenSnapshot)
        iDorsales = rs.Fields(0)
        rs.Close
        
        'Localizamos el número de hojas de esta fase
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM hojas_reconocidas WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase=" & iFase, dbOpenSnapshot)
        iHojas = rs.Fields(0)
        rs.Close
        
        If GenerarControl(Val(tbCodComp.Text)) Then
            ' Cada juez tiene x tandas x hojas por tanda(Si bailes > 5) + el número de tandas de la hoja de control
            iMaxHojasRec = (iMaxTandas * iNumHojas * (iJueces + ijuecespasos) + iMaxTandas)
        Else
            ' Cada juez tiene x tandas x hojas por tanda(Si bailes > 5)
            iMaxHojasRec = iMaxTandas * iNumHojas * (iJueces + ijuecespasos)
        End If
        
        Me.Caption = mml_FRASE0909 & iPuestos & mml_FRASE0472 & iDorsales * iBailes * iJueces & mml_FRASE0910 & iHojas & mml_FRASE0472 & iMaxHojasRec ' Mas 1 de control
        If bJuezPasos Then
            Me.Caption = Me.Caption & mml_FRASE0911
        End If
        If iHojaRepesca Then
            Me.Caption = Me.Caption & mml_FRASE0912
        End If
        
        igCodCat = iCodCat
        igFase = iFase
        igHojaRepesca = iHojaRepesca
        
        If (Not BailesParciales(iCodCat) And (iMaxHojasRec = iHojas Or G_NO_CONTAR_HOJAS) And iPuestos = iDorsales * iBailes * iJueces) Or _
           (BailesParciales(iCodCat) And ((iMaxHojasRec Mod iHojas = 0 And iHojas > 1) Or G_NO_CONTAR_HOJAS) And (iPuestos Mod (iDorsales * iJueces)) = 0) Then
            If MsgBox(mml_FRASE0913 & iCodCat & mml_FRASE0641 & iFase & mml_FRASE0914, vbYesNo Or vbQuestion, mml_FRASE0086) = vbYes Then
                frmCalcular.tbCodComp.Text = tbCodComp.Text
                frmCalcular.tbCodCat.Text = iCodCat
                frmCalcular.tbCodFase.Text = iFase
                frmCalcular.chkRep.Value = iHojaRepesca
                tbHojasProc.Text = C_HOJAS_PROC
                frmCalcular.MostrarCalcular
            End If
        End If
    End If
error:
Dim Msj
If Err.Number <> 0 Then
   Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
         & Err.Source & Chr(13) & Err.Description
   MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
End If
End Sub

