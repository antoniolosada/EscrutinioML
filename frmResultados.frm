VERSION 5.00
Begin VB.Form frmResultados 
   Caption         =   "mml_FRASE0917"
   ClientHeight    =   8745
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmResultados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPublic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8805
      Left            =   0
      ScaleHeight     =   8745
      ScaleWidth      =   12150
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   12210
      Begin VB.ListBox lbSinOrden 
         Height          =   900
         Left            =   3150
         TabIndex        =   18
         Top             =   330
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ListBox lbConOrden 
         Height          =   1110
         Left            =   645
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   315
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.FileListBox filePublic 
         Height          =   2400
         Left            =   7980
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.PictureBox pbLibreta 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   9015
      Left            =   -15
      Picture         =   "frmResultados.frx":0BC2
      ScaleHeight     =   8955
      ScaleWidth      =   690
      TabIndex        =   12
      Tag             =   "Libreta"
      Top             =   -15
      Visible         =   0   'False
      Width           =   750
      Begin VB.PictureBox pbResultados 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4800
         Left            =   3690
         Picture         =   "frmResultados.frx":4B7C
         ScaleHeight     =   4800
         ScaleWidth      =   4680
         TabIndex        =   14
         Top             =   2025
         Visible         =   0   'False
         Width           =   4680
      End
   End
   Begin VB.Frame mrcInic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   8775
      Left            =   -30
      TabIndex        =   2
      Top             =   -60
      Width           =   12165
      Begin VB.Timer tmrPublic 
         Left            =   9570
         Top             =   210
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2880
         Left            =   7200
         Picture         =   "frmResultados.frx":8E41
         ScaleHeight     =   2880
         ScaleWidth      =   4575
         TabIndex        =   10
         Top             =   4200
         Width           =   4575
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   4800
         Picture         =   "frmResultados.frx":17983
         ScaleHeight     =   795
         ScaleWidth      =   1365
         TabIndex        =   9
         Top             =   4560
         Width           =   1365
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2490
         Left            =   135
         Picture         =   "frmResultados.frx":190D1
         ScaleHeight     =   2490
         ScaleWidth      =   1890
         TabIndex        =   3
         Top             =   1335
         Width           =   1890
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1320
         Left            =   2385
         Picture         =   "frmResultados.frx":2877B
         ScaleHeight     =   1320
         ScaleWidth      =   9405
         TabIndex        =   13
         Top             =   1740
         Width           =   9405
      End
      Begin VB.Label lblEquipo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "mml_FRASE0919"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   50.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   0
         TabIndex        =   19
         Top             =   45
         Width           =   11985
      End
      Begin VB.Label lblCab1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "mml_FRASE0920"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   42
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   105
         TabIndex        =   11
         Top             =   -180
         Width           =   11955
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "mml_FRASE0921"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   32.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   2250
         TabIndex        =   8
         Top             =   855
         Width           =   8895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "mml_FRASE0047"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   33.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1710
         TabIndex        =   7
         Top             =   3090
         Width           =   10575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "mml_FRASE0922"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   345
         TabIndex        =   6
         Top             =   4800
         Width           =   6735
      End
      Begin VB.Label lblCab 
         BackColor       =   &H00FFFFFF&
         Caption         =   "mml_FRASE0923"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   32.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   11925
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "mml_FRASE0230"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   4
         Top             =   7500
         Width           =   7575
      End
   End
   Begin VB.Timer tmrRefrescar 
      Interval        =   2000
      Left            =   1920
      Top             =   7920
   End
   Begin VB.PictureBox pbAnt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   0
      Picture         =   "frmResultados.frx":4F895
      ScaleHeight     =   8715
      ScaleWidth      =   6075
      TabIndex        =   0
      Tag             =   "Ant"
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox pbAct 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8790
      Left            =   6120
      Picture         =   "frmResultados.frx":62B7D
      ScaleHeight     =   8730
      ScaleWidth      =   5835
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
   End
End
Attribute VB_Name = "frmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_FUENTE = "Arial"
Const C_FUENTE_NEGRITA = "Arial Black"
Const C_NEGRITA = False

'Const C_FUENTE = "Comic Sans MS"
'Const C_FUENTE_NEGRITA = "Comic Sans MS"
'Const C_NEGRITA = True

Dim dHoraUltimaPublicacion As Date
Dim dHoraAntPublicacion As Date
Dim iContLibreta As Integer
Dim sPista As String

Dim m_lCodCateg As Long
Dim m_iCodFase As Integer
Dim m_iRepesca As Integer
Dim m_iCodBaile As Integer
Dim m_sDesc As String
Dim m_sComen As String
Dim gs_ImagenDerecha As String

Public Function Consulta() As Boolean
    Consulta = True
End Function

Public Sub Actualizar()
    ActualizarVariables
    dHoraUltimaPublicacion = CDate("1/1/1900 0:00:0")
    dHoraAntPublicacion = CDate("1/1/1900 0:00:0")
    Call tmrRefrescar_Timer
End Sub

Public Sub SetsPista(sCad As String)
    sPista = sCad
End Sub


Public Sub ActualizarVariables()

    CargarVariablesConfiguracion

    dHoraAntPublicacion = Now
    tmrRefrescar.Enabled = False
    tmrRefrescar.Interval = VarCfg("refresco_resultados")
    tmrRefrescar.Enabled = True
    iContLibreta = 0
    lblCab = G_CAB_RESULTADOS
    lblCab1 = G_CAB1_RESULTADOS
    If G_CAB_RESULTADOS = "" And G_CAB1_RESULTADOS = "" Then
        lblEquipo.Visible = True
    Else
        lblEquipo.Visible = False
    End If
    
    filePublic.Path = G_DIR_PUBLICIDAD
    filePublic.Refresh
    tmrPublic.Interval = VarCfg("refresco_publicidad")
        
    tmrPublic.Enabled = True
    
End Sub


Private Sub Form_Load()
    TraducirCadenas Me
    ActualizarVariables
    pbAct.Tag = "EPA_peque"
    If G_RESULTADOS_ALTO_CONTRASTE Then
        gs_ImagenDerecha = "EPA_peque1_altocontraste.gif"
    Else
        gs_ImagenDerecha = "EPA_peque1.gif"
    End If
    
    
    If G_RESULTADOS_ALTO_CONTRASTE Then
        pbAnt.Picture = Nothing
        pbAct.Picture = Nothing
        pbAct.BackColor = 0
        pbAnt.BackColor = 0
    End If
End Sub

Private Sub Form_Resize()
    pbAnt.Height = Me.Height
    pbAnt.Width = Me.Width / 2
    pbAct.Height = Me.Height
    pbAct.Width = Me.Width / 2
    pbAct.Left = pbAnt.Width + 1
    
    picPublic.Height = Me.Height
    picPublic.Width = Me.Width
End Sub

Private Sub pbAct_DblClick()
    Call tmrRefrescar_Timer
End Sub

Private Sub pbAnt_DblClick()
    dHoraUltimaPublicacion = CDate("1/1/1900 0:00:0")
    dHoraAntPublicacion = CDate("1/1/1900 0:00:0")
End Sub

Private Sub tmrPublic_Timer()
Static iImagen As Integer
Dim i As Integer
Dim sFichero As String
Dim iEspera As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    If mrcInic.Visible And G_DIR_PUBLICIDAD <> "" And VarCfg("publicidad_activa") <> "N" Then
        If iImagen < filePublic.ListCount Then
            tmrRefrescar_Timer
            If Not mrcInic.Visible Then Exit Sub
            For i = 0 To picPublic.Height Step G_SALTO_PUBLIC
                picPublic.Top = i
                Sleep 1
                DoEvents
            Next
            sFichero = filePublic.List(iImagen)
            iEspera = Val(Left$(sFichero, 2))
            If iEspera > 0 Then
                tmrPublic.Enabled = False
                tmrPublic.Interval = iEspera * 1000
                tmrPublic.Enabled = True
            Else
                tmrPublic.Interval = VarCfg("refresco_publicidad")
            End If
            If Not C_DEBUG Then On Local Error Resume Next
            picPublic.Picture = LoadPicture(filePublic.Path & "\" & sFichero)
            picPublic.Refresh
            If Not C_DEBUG Then On Local Error GoTo 0
            tmrRefrescar_Timer
            If Not mrcInic.Visible Then Exit Sub
            Espera G_ESPERA_NO_PUBLIC / 3
            tmrRefrescar_Timer
            If Not mrcInic.Visible Then Exit Sub
            Espera G_ESPERA_NO_PUBLIC / 3
            tmrRefrescar_Timer
            If Not mrcInic.Visible Then Exit Sub
            Espera G_ESPERA_NO_PUBLIC / 3
            tmrRefrescar_Timer
            If Not mrcInic.Visible Then Exit Sub
            picPublic.Visible = True
            For i = picPublic.Height To 0 Step -G_SALTO_PUBLIC
                picPublic.Top = i
                Sleep 1
                DoEvents
            Next
            tmrRefrescar_Timer
            If Not mrcInic.Visible Then Exit Sub
            iImagen = iImagen + 1
        Else
            iImagen = 0
        End If
    Else
        picPublic.Visible = False
    End If
    Exit Sub
error:
    ProcesarError "tmrPublic"
End Sub

Private Sub tmrRefrescar_Timer()
Dim rs As Recordset
Dim rsParejas As Recordset
Dim i As Long
Dim sComen As String
Dim doEscala As Double
Dim bPublicPendiente As Boolean
Static iProcActivo As Integer
Static iRetrasoResultados As Integer

    ProcesarEventos
    
    If VarCfg("publicacion_pendiente") = "N" Then
        If iRetrasoResultados < G_RETRASO_RESULTADOS Then
            Inc iRetrasoResultados
            Exit Sub
        End If
        bPublicPendiente = False
    Else
        bPublicPendiente = True
    End If
    
    If iProcActivo = 1 Then Exit Sub
    iProcActivo = 1
    
    iRetrasoResultados = 0
    If bPublicPendiente Then
        db.Execute "UPDATE cfg SET valor = 'N' WHERE variable = 'publicacion_pendiente'"
    End If
    
    If pbLibreta.Visible Then
        iContLibreta = iContLibreta + 1
        If iContLibreta = VarCfg("refresco_libreta") Then
            pbLibreta.Visible = False
            pbLibreta.Width = 585
            iContLibreta = 0
            frmResultados.Caption = mml_FRASE0925
        End If
    End If
    tmrRefrescar.Interval = VarCfg("refresco_resultados")
    If G_PUBLICAR_POSICION Then
        'Mantenemos publicadas las ?ltimas fases de cada grupo( ya que al publicar se borran las anteriores)
        Set rs = db.OpenRecordset("SELECT *, 0 as cod_baile FROM publicar WHERE descripcion LIKE '*" & sPista & "*' ORDER BY hora_publicacion DESC", dbOpenSnapshot)
    Else
        ' Solo imprimimos los dorsales de las categor?as que no han salido a bailar
        ' y por lo tanto no tienen puntuaciones O los que se publican parcialmente
        ' m?s los X siguientes grupos del horario de la hora actual
        ' que no tengan puntuaciones y que tengan dorsales y que no est?n publicados por duplicado
        If G_PUBLICAR_NUM_GRUPOS_RESULTADOS > 0 Then
            sSQL = "SELECT hora_publicacion, cod_categoria, fase, descripcion, repesca, comentarios, 0 as cod_baile  FROM publicar p WHERE  p.descripcion LIKE '*" & sPista & "*' AND " & _
                " ((SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = p.cod_categoria AND fase = p.fase AND repesca = p.repesca)=0 " & _
                "   OR (SELECT COUNT(*) FROM categorias WHERE mostrar_posicion = 1 and codigo = p.cod_categoria) = 1 ) " & _
                " UNION (SELECT top " & Val(G_PUBLICAR_NUM_GRUPOS_RESULTADOS) & " h.hora, h.cod_categoria, h.numfase, h.grupo, h.repesca, '" & C_SALIDA_PROXIMA & "', h.cod_baile  FROM horario h WHERE  h.grupo LIKE '*" & sPista & "*' AND " & _
                "   h.hora > #" & sEstimacionInversa(Now) & "# AND cod_competicion = " & VarCfg("horario_codcompeticion") & _
                "   AND (SELECT COUNT(*) FROM puntuaciones p WHERE p.cod_categoria = h.cod_categoria AND p.fase = h.numfase AND p.repesca = h.repesca)=0 " & _
                "   AND (SELECT COUNT(*) FROM dorsales d WHERE d.cod_categoria = h.cod_categoria AND d.fase = h.numfase AND d.repesca = h.repesca)>0 " & _
                "   AND (SELECT COUNT(*) FROM publicar p1 WHERE p1.cod_categoria = h.cod_categoria AND p1.fase = h.numfase AND p1.repesca = h.repesca)=0 " & _
                " ORDER BY 1) ORDER BY hora_publicacion DESC;"
            Debug.Print sSQL
            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
        Else
        ' Solo imprimimos los dorsales de las categor?as que no han salido a bailar
        ' y por lo tanto no tienen puntuaciones
            Set rs = db.OpenRecordset("SELECT * FROM publicar p WHERE p.descripcion LIKE '*" & sPista & "*' AND (SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = p.cod_categoria AND fase = p.fase AND repesca = p.repesca)=0 ORDER BY hora_publicacion DESC", dbOpenSnapshot)
        End If
    End If
    'Set rs = db.OpenRecordset("SELECT * FROM publicar p ORDER BY hora_publicacion DESC", dbOpenSnapshot)
    If VarCfg("publicidad_activa") = "N" And Not rs.EOF Then
        pbAnt.Visible = True
        pbAct.Visible = True
        mrcInic.Visible = False
        DoEvents: DoEvents: DoEvents:
        If dHoraUltimaPublicacion < rs!hora_publicacion Then
            sComen = IIf(IsNull(rs!comentarios), "", rs!comentarios)
            If InStr(sComen, C_SALIDA_PROXIMA) = 0 Then
                dHoraUltimaPublicacion = rs!hora_publicacion
                'pbAct.BackColor = &HC0C0FF&H008080FF&
                pbAct.BackColor = &HC0FFFF
                doEscala = pbLibreta.ScaleWidth / pbLibreta.Width
                pbLibreta.Width = 0
                pbLibreta.Visible = True
                pbResultados.Visible = True
                frmResultados.Caption = mml_FRASE0926
                
                If G_RESULTADOS_ALTO_CONTRASTE Then
                    pbLibreta.Picture = Nothing
                    pbLibreta.BackColor = C_COLOR_NEGRO
                Else
                    pbLibreta.BackColor = &HC0FFFF
                End If
                pbLibreta.Refresh
                If Not G_RESULTADOS_UNO_A_UNO Then
                    Publicar rs!cod_categoria, rs!fase, rs!repesca, rs!DESCRIPCION, sComen, pbLibreta, C_LONG_NOMBRE_LIBRETA, rs!cod_baile
                End If
                For i = 20 To 13000 Step Val(VarCfg("salto_libreta"))
                    pbLibreta.Width = i
                    pbLibreta.Refresh
                    Espera G_VELOCIDAD_LIBRETA
                    DoEvents: DoEvents
                Next
                If G_RESULTADOS_UNO_A_UNO Then
                    ProcesarEventos
                    Publicar rs!cod_categoria, rs!fase, rs!repesca, rs!DESCRIPCION, sComen, pbLibreta, C_LONG_NOMBRE_LIBRETA, rs!cod_baile
                End If
                pbResultados.Visible = False
                pbResultados.Refresh
                If pbAct.Tag = "EPA_peque" Then
                    pbAct.Picture = LoadPicture(G_IMAGENES_EPA_PEQUE & gs_ImagenDerecha)
                    pbAct.Tag = gs_ImagenDerecha
                End If
                Publicar rs!cod_categoria, rs!fase, rs!repesca, rs!DESCRIPCION, sComen, pbAct, C_LON_MAX_NOMBRE, rs!cod_baile
                m_lCodCateg = rs.Fields("cod_categoria")
                m_iCodFase = rs.Fields("fase")
                m_iRepesca = rs.Fields("repesca")
                m_iCodBaile = rs.Fields("cod_baile")
                m_sDesc = rs.Fields("descripcion")
                m_sComen = sComen
            End If
        ElseIf m_lCodCateg > 0 And m_iCodFase > 0 Then
            Static iEsperaActualizacion As Integer
            
            If G_ACTUALIZAR_ULTIMA_PUBLICACION > 0 Then
                If (iEsperaActualizacion + 1) Mod G_ACTUALIZAR_ULTIMA_PUBLICACION = 0 Then
                    If pbAct.Tag = "EPA_peque" Then
                        pbAct.Picture = LoadPicture(G_IMAGENES_EPA_PEQUE & gs_ImagenDerecha)
                        pbAct.Tag = gs_ImagenDerecha
                    End If
                    Publicar m_lCodCateg, m_iCodFase, m_iRepesca, m_sDesc, m_sComen, pbAct, C_LON_MAX_NOMBRE, m_iCodBaile
                End If
                
                If iEsperaActualizacion > 30000 Then
                    iEsperaActualizacion = 0
                End If
                
                Inc iEsperaActualizacion
            End If
        End If
        Do While Not rs.EOF
            If dHoraAntPublicacion <= rs!hora_publicacion Or rs!hora_publicacion = dHoraUltimaPublicacion Then
                rs.MoveNext
                If rs.EOF Then
                    rs.MoveFirst
                    'Evitamos publicar los resultados que est?n en el panel derecho
                    If G_NO_PUBLICAR_COMO_ANT_PANEL_DERECHO And rs!hora_publicacion = dHoraUltimaPublicacion Then
                        rs.MoveNext
                    End If
                    If Not rs.EOF Then
                        dHoraAntPublicacion = rs!hora_publicacion
                        sComen = IIf(IsNull(rs!comentarios), "", rs!comentarios)
                        pbAnt.Refresh
                        Publicar rs!cod_categoria, rs!fase, rs!repesca, rs!DESCRIPCION, sComen, pbAnt, C_LON_MAX_NOMBRE, rs!cod_baile
                    Else
                        dHoraAntPublicacion = CDate("1/1/1900 0:00:0")
                    End If
                    Exit Do
                End If
            Else
                dHoraAntPublicacion = rs!hora_publicacion
                sComen = IIf(IsNull(rs!comentarios), "", rs!comentarios)
                pbAnt.Refresh
                Publicar rs!cod_categoria, rs!fase, rs!repesca, rs!DESCRIPCION, sComen, pbAnt, C_LON_MAX_NOMBRE, rs!cod_baile
                Exit Do
            End If
        Loop
    Else
        pbAnt.Cls
        pbAct.Cls
        pbAct.Picture = LoadPicture(G_IMAGENES_EPA_PEQUE & "EPA_peque.gif")
        pbAct.Tag = "EPA_peque"
        pbAnt.Visible = False
        pbAct.Visible = False
        
        mrcInic.Visible = True
        DoEvents: DoEvents: DoEvents:
    End If
    rs.Close
    
    iProcActivo = 0
End Sub

Function Publicar(ByVal iCategoria As Integer, iFase As Integer, sRep As Integer, sDesc As String, ByVal sComen As String, pbCuadro As PictureBox, iLongNombre As Integer, iCodBaile As Integer) As Boolean
Dim rsParejas As Recordset
Dim iDorsalIni As Integer
Dim i As Integer
    If ComprobarSiTeamMatch(iCategoria) Then
        PublicarPantalla iCategoria, iFase, sRep, sDesc, sComen, pbCuadro, iLongNombre, iCodBaile
    Else
        Set rsParejas = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND d.cod_categoria = " & iCategoria & " AND d.fase = " & iFase & " AND d.repesca=" & sRep & " ORDER BY 1", dbOpenSnapshot)
        iDorsalIni = 0
        While iDorsalIni < rsParejas.Fields(0)
            PublicarPantalla iCategoria, iFase, sRep, sDesc, sComen, pbCuadro, iLongNombre, iCodBaile, iDorsalIni
            iDorsalIni = iDorsalIni + C_MAX_PAREJAS_PANTALLA
            For i = 0 To G_RETARDO_RESULTADOS_MULTIPLES_PANTALLAS
                Sleep 30
                DoEvents
                Sleep 30
                DoEvents
                Sleep 30
                DoEvents
            Next
        Wend
        rsParejas.Close
    End If
End Function

Function PublicarPantalla(iCategoria As Integer, iFase As Integer, sRep As Integer, sDesc As String, ByVal sComen As String, pbCuadro As PictureBox, iLongNombre As Integer, iCodBaile As Integer, Optional ByVal iPosDorsalIni As Integer = 0) As Boolean
Dim rsParejas As Recordset
Dim iCParejas As Integer
Dim iYInic As Integer
Dim sMargen As String
Dim iPosX As Integer, iPosY As Integer
Dim sCad As String
Dim sLado As String
Dim iMaxParejasCol As Integer
Dim iMaxCarLinea As Integer
Dim rsPuestos As Recordset
Dim iRestarTamFila  As Integer
Dim iTamFuente As Integer
Dim sLinea As String
Dim bTeamMatch As Boolean
Dim bSalidaAPista As Boolean
Dim iCarPorLinea As Integer

    bSalidaAPista = (InStr(sComen, C_SALIDA_PROXIMA) > 0)
    
    If bSalidaAPista Then
        bTeamMatch = False
    Else
        bTeamMatch = ComprobarSiTeamMatch(iCategoria)
    End If

    If iCodBaile > 0 Then
        sComen = sComen & " " & sNombreBaile(iCodBaile) & " "
    ElseIf iCodBaile < 0 Then
        sComen = sComen & mml_FRASE0964
    End If
    
    PublicarPantalla = True
    If pbCuadro.Tag = "Libreta" Then
        sLado = "         "
    Else
        sLado = ""
    End If
    sMargen = ""
    pbCuadro.Cls
    
    pbCuadro.FontName = C_FUENTE
    If VarCfg("reloj_activo") = "S" Then
        If pbCuadro.Tag = "Libreta" Or pbCuadro.Tag = mml_FRASE0924 Then
            If pbCuadro.Tag = "Libreta" Then
                pbCuadro.CurrentY = 0
                pbCuadro.FontSize = 40
                pbCuadro.CurrentX = frmResultados.Width - pbCuadro.TextWidth("99:9999")
                pbCuadro.FontBold = False
            ElseIf pbCuadro.Tag = mml_FRASE0924 Then
                pbCuadro.CurrentY = 6500
                pbCuadro.CurrentX = 0
                pbCuadro.FontSize = 95
                pbCuadro.FontBold = True
            End If
            pbCuadro.ForeColor = C_COLOR_VERDE_MEDIO
            pbCuadro.Print Format$(Now, "hh:mm")
        End If
    End If
    
    pbCuadro.CurrentX = 0
    pbCuadro.CurrentY = 0
    pbCuadro.FontName = C_FUENTE_NEGRITA
    pbCuadro.ForeColor = 0
    pbCuadro.FontBold = C_NEGRITA
    pbCuadro.FontSize = C_FUENTE_TITULO_RESULTADOS
    
    If G_RESULTADOS_ALTO_CONTRASTE Then
        If InStr(sComen, C_SALIDA_PROXIMA) = 0 Then
            pbCuadro.ForeColor = C_COLOR_BLANCO
        Else
            pbCuadro.ForeColor = C_COLOR_AMARILLO_CLARO
        End If
    Else
        If InStr(sComen, C_SALIDA_PROXIMA) = 0 Then
            pbCuadro.ForeColor = C_COLOR_ROJO_OSCURO
        Else
            pbCuadro.ForeColor = C_COLOR_AZUL_OSCURO
        End If
    End If
    pbCuadro.Print sLado & sDesc & "-" & sDescFase(iFase);
    pbCuadro.FontSize = 14
    pbCuadro.Print
    pbCuadro.FontSize = C_FUENTE_COMENTARIO
    
    If G_RESULTADOS_ALTO_CONTRASTE Then
        pbCuadro.ForeColor = C_COLOR_BLANCO
    Else
        pbCuadro.ForeColor = 0
    End If
    
    'Comprobamos si es una repesca
    If sComen = "" And sRep = 1 Then
        sComen = mml_FRASE0912
    End If
    
    'Comprobar si tenemos que publicar la hora estimada
    If G_PUBLICAR_HORA_ESTIMADA Then
        Dim rs As Recordset
        
        Set rs = db.OpenRecordset("SELECT hora FROM horario WHERE cod_categoria = " & iCategoria & " AND numfase = " & iFase & " AND repesca = " & sRep, dbOpenSnapshot)
        If Not rs.EOF Then
            sComen = Format$(CDate(sEstimacion(rs!hora)), "hh:mm") & " " & sComen
        End If
        rs.Close
    End If
    
    pbCuadro.Print sLado & sComen
    pbCuadro.Line -Step(pbCuadro.Width, 0)
    pbCuadro.FontBold = False
    pbCuadro.FontSize = 2
    pbCuadro.Print
    pbCuadro.FontSize = IIf(pbCuadro.Tag = "Libreta", C_FUENTE_PEQUE_LIBRETA, C_FUENTE_PEQUE)
    iYInic = pbCuadro.CurrentY
    Set rsParejas = db.OpenRecordset("SELECT num_dorsal, nombre_hombre, nombre_mujer FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND d.cod_categoria = " & iCategoria & " AND d.fase = " & iFase & " AND d.repesca=" & sRep & " ORDER BY 1", dbOpenSnapshot)
    
    If bTeamMatch Then
        Dim i As Integer, dPuntos As Double
            
        If pbCuadro.Tag = "Libreta" Then
            PrintConSombra pbCuadro, sLado & mml_FRASE0962
            
            pbCuadro.Print
            pbCuadro.FontSize = C_FUENTE_MEDIANA_LIBRETA
            Set rs = db.OpenRecordset("SELECT * FROM resultadosteammatch WHERE tipo = " & C_TEAMMATCH_TOTAL & " AND cod_categoria = " & iCategoria & " ORDER BY puntuacion", dbOpenSnapshot)
            i = 0
            dPuntos = 0
            While Not rs.EOF
                If dPuntos <> rs!puntuacion Then
                    Inc i
                    dPuntos = rs!puntuacion
                End If
                PrintConSombra pbCuadro, sLado & i & "? - " & rs!comunidad & " con " & rs!puntuacion & " Ptos."
                rs.MoveNext
            Wend
            rs.Close
        Else
            pbCuadro.FontSize = C_FUENTE_MEDIANA
            PrintConSombra pbCuadro, sLado & mml_FRASE0963
            Set rs = db.OpenRecordset("SELECT * FROM resultadosteammatch WHERE tipo = " & C_TEAMMATCH_TOTAL & " AND cod_categoria = " & iCategoria & " ORDER BY puntuacion", dbOpenSnapshot)
            i = 0
            dPuntos = 0
            While Not rs.EOF
                If dPuntos <> rs!puntuacion Then
                    Inc i
                    dPuntos = rs!puntuacion
                End If
                PrintConSombra pbCuadro, sLado & i & "? - " & rs!comunidad & " -> " & rs!puntuacion & " Ptos."
                rs.MoveNext
            Wend
            rs.Close
            pbCuadro.Print
            PrintConSombra pbCuadro, sLado & "__Ptos. de " & sDesc & "______________________"
            Set rs = db.OpenRecordset("SELECT * FROM resultadosteammatch WHERE tipo = " & C_TEAMMATCH_POR_CATEGORIA & " AND cod_categoria = " & iCategoria & " ORDER BY puntuacion", dbOpenSnapshot)
            While Not rs.EOF
                PrintConSombra pbCuadro, sLado & rs!comunidad & " -> " & rs!puntuacion & " Ptos."
                rs.MoveNext
            Wend
            rs.Close
        End If
    Else '*************************************************************************************************
        If Not rsParejas.EOF Then
            rsParejas.MoveLast
            iCParejas = rsParejas.RecordCount
            If pbCuadro.Tag = "Libreta" Then
                If rsParejas.RecordCount > C_MAX_DORSALES_PANT_COMPLETA_LIBRETA Then
                    'Linea de divisi?n
                    pbCuadro.Line (pbCuadro.Width / 2, pbCuadro.CurrentY)-(pbCuadro.Width / 2, pbCuadro.Height)
                    If rsParejas.RecordCount > C_MAX_PAREJAS_POR_COLUMNA_LIBRETA * 2 Then
                        pbCuadro.FontSize = C_FUENTE_MUY_PEQUE_LIBRETA
                    End If
                Else
                    sMargen = "     "
                    pbCuadro.FontSize = IIf(rsParejas.RecordCount > C_MAX_PAREJAS_PARA_FUENTE_GRANDE_LIBRETA, C_FUENTE_MEDIANA_LIBRETA, C_FUENTE_GRANDE_LIBRETA)
                End If
            Else
                If rsParejas.RecordCount > C_MAX_DORSALES_PANT_COMPLETA Then
                    pbCuadro.Line (pbCuadro.Width / 2, pbCuadro.CurrentY)-(pbCuadro.Width / 2, pbCuadro.Height)
                Else
                    sMargen = " " & sLado
                    pbCuadro.FontSize = IIf(rsParejas.RecordCount > C_MAX_PAREJAS_PARA_FUENTE_GRANDE, C_FUENTE_MEDIANA, C_FUENTE_GRANDE)
                End If
            End If
            rsParejas.MoveFirst
        End If
    
        pbCuadro.CurrentY = iYInic
        pbCuadro.CurrentX = 0
        If G_RESULTADOS_ALTO_CONTRASTE Then
            pbCuadro.ForeColor = C_COLOR_BLANCO
        Else
            pbCuadro.ForeColor = &HC00000
        End If
        If pbCuadro.Tag = "Libreta" Then
            If (G_PUBLICAR_POSICION Or MostrarPosicion(iCategoria)) And Not bSalidaAPista Then
                pbCuadro.CurrentY = iYInic - 40
                iTamFuente = pbCuadro.FontSize
                pbCuadro.FontSize = C_TAM_FUENTE_TITULO_LIBRETA
                If iFase = 1 Then
                    pbCuadro.Print mml_FRASE0829;
                Else
                    pbCuadro.Print mml_FRASE0830;
                End If
                pbCuadro.FontSize = 4
                pbCuadro.Print
                pbCuadro.FontSize = iTamFuente
            End If
        Else
            If (G_PUBLICAR_POSICION Or MostrarPosicion(iCategoria)) And Not bSalidaAPista Then
                If iFase = 1 Then
                    pbCuadro.Print sMargen & mml_FRASE0831
                Else
                    pbCuadro.Print sMargen & mml_FRASE0832
                End If
            Else
                pbCuadro.Print sMargen & mml_FRASE0833
            End If
        End If
        If G_RESULTADOS_ALTO_CONTRASTE Then
            pbCuadro.ForeColor = C_COLOR_BLANCO
        Else
            pbCuadro.ForeColor = 0
        End If
    
    
        If pbCuadro.Tag = "Libreta" Then
            If Round(pbCuadro.FontSize) = C_FUENTE_MUY_PEQUE_LIBRETA Then
                iMaxParejasCol = C_MAX_PAREJAS_POR_COLUMNA_LIBRETA_PEQUE
            Else
                iMaxParejasCol = C_MAX_PAREJAS_POR_COLUMNA_LIBRETA
            End If
        Else
            iMaxParejasCol = C_MAX_PAREJAS_POR_COLUMNA
        End If
        
        If pbCuadro.Tag = "Libreta" Then
            iMaxCarLinea = C_MAX_INFO_RESULTADOS_LINEA_LIBRETA
        Else
            iMaxCarLinea = C_MAX_INFO_RESULTADOS_LINEA
        End If
    Dim lbOrden As ListBox
        'Cargamos los datos en la lista
        lbConOrden.Clear
        lbSinOrden.Clear
        iCParejas = 0
        If (G_PUBLICAR_POSICION Or MostrarPosicion(iCategoria)) And Not bSalidaAPista Then
            Set lbOrden = lbConOrden
        Else
            Set lbOrden = lbSinOrden
        End If
        
        
        While Not rsParejas.EOF And iCParejas < C_MAX_PAREJAS_PANTALLA
            If iPosDorsalIni > 0 Then
                Dec iPosDorsalIni
                rsParejas.MoveNext
            Else
                sCad = sMargen
                iRestarTamFila = 0
                
                'Publicar posici?n actual
                If (G_PUBLICAR_POSICION Or MostrarPosicion(iCategoria)) And Not bSalidaAPista Then
                    iRestarTamFila = 3
                    If iFase = 1 Then ' FINAL
                        Set rsPuestos = db.OpenRecordset("SELECT puesto FROM cal_conjunto WHERE cod_categoria = " & iCategoria & " AND num_dorsal = " & rsParejas!num_dorsal & " AND regla='FIN' ORDER BY 1", dbOpenSnapshot)
                            If rsPuestos.EOF Then
                                ' Si todav?a no hay c?lculos se muestra No Calculado NC
                                sCad = sCad & "(--)"
                            Else
                                sCad = sCad & "(" & Format$(Val(rsPuestos!puesto), "#") & "?)"
                            End If
                        rsPuestos.Close
                    Else ' No final
                      Dim iSumaMarcas As Integer
                        iSumaMarcas = 0
                        Set rsPuestos = db.OpenRecordset("SELECT posiciones_mi FROM cal_baile WHERE cod_categoria = " & iCategoria & " AND repesca=" & sRep & " AND fase = " & iFase & " AND num_dorsal = " & rsParejas!num_dorsal & " AND puesto=" & BD_POS_POSICION & " ORDER BY 1", dbOpenSnapshot)
                        If rsPuestos.EOF Then
                            sCad = sCad & "(--)"
                        Else
                            While Not rsPuestos.EOF
                                iSumaMarcas = iSumaMarcas + rsPuestos!posiciones_mi
                                rsPuestos.MoveNext
                            Wend
                            sCad = sCad & "(" & Format$(iSumaMarcas, "#") & ")"
                        End If
                        rsPuestos.Close
                    End If
                End If
                
                iCarPorLinea = iMaxCarLinea
                sLinea = UCase(Left$(rsParejas!nombre_mujer, 1))
                sLinea = Left$(rsParejas!num_dorsal & "-" & Nombre(rsParejas!nombre_hombre, iLongNombre) & IIf(sLinea = "I" Or sLinea = mml_FRASE0564, " e ", mml_FRASE0035) & Nombre(rsParejas!nombre_mujer, iLongNombre), C_MAX_CAR_POR_FILA - iRestarTamFila)
                If (((pbCuadro.Tag <> "Libreta" And rsParejas.RecordCount <= C_MAX_DORSALES_PANT_COMPLETA) Or (pbCuadro.Tag = "Libreta" And rsParejas.RecordCount <= C_MAX_DORSALES_PANT_COMPLETA_LIBRETA)) And Len(sLinea) <= C_MAX_CAR_LINEA_LIBRETA) Then
                    sCad = sCad & Mid$(sLinea, 1, C_MAX_CAR_LINEA_LIBRETA)
                ElseIf Len(sLinea) > iMaxCarLinea - iRestarTamFila Then
                    sCad = sCad & Mid$(sLinea, 1, iMaxCarLinea - iRestarTamFila) & "."
                    iCarPorLinea = iMaxCarLinea - iRestarTamFila
                Else
                    sCad = sCad & Mid$(sLinea, 1, iMaxCarLinea - iRestarTamFila)
                    iCarPorLinea = iMaxCarLinea - iRestarTamFila
                End If
                'If iCParejas < iMaxParejasCol And pbCuadro.Tag = "Libreta" Then
                '    If Round(pbCuadro.FontSize) = C_FUENTE_PEQUE_LIBRETA Then
                '        sCad = "     " & sCad
                '    ElseIf Round(pbCuadro.FontSize) = C_FUENTE_MUY_PEQUE_LIBRETA Then
                '        sCad = "               " & sCad
                '    End If
                'End If
                
                lbOrden.AddItem sCad
                rsParejas.MoveNext
                Inc iCParejas
            End If
        Wend
        
        lbOrden.Refresh
        
        Dim iInicial As Integer
        Dim iFinal As Integer
        Dim iSalto As Integer
        Dim iCParejasImpresas As Integer
        
        If (G_PUBLICAR_POSICION Or MostrarPosicion(iCategoria)) And Not bSalidaAPista Then
            iInicial = lbOrden.ListCount - 1
            iFinal = 0
            iSalto = -1
        Else
            iInicial = 0
            iFinal = lbOrden.ListCount - 1
            iSalto = 1
        End If
        
        'Mostrar los dorsales en la pantalla
        Dim iContPArejas As Integer
        Dim iIniPosX As Integer, iIniPosY As Integer
        
        iIniPosX = pbCuadro.CurrentX
        iIniPosY = pbCuadro.CurrentY
        iContPArejas = 0
        For iCParejas = iInicial To iFinal Step iSalto
        
            If Not G_PUBLICAR_POSICION And Not MostrarPosicion(iCategoria) And pbCuadro.Tag = "Libreta" And G_RESULTADOS_UNO_A_UNO Then
                Espera (G_ESPERA_ENTRE_PART)
                pbResultados.Visible = False
            End If
    
            ' Para que quepan el m?ximo de parejas en una columna hay que ajustarlas
            If iCParejasImpresas = iMaxParejasCol Then
                pbCuadro.CurrentY = iYInic
                pbCuadro.CurrentX = pbCuadro.Width / 2 + 60
                If G_RESULTADOS_ALTO_CONTRASTE Then
                    pbCuadro.ForeColor = C_COLOR_BLANCO
                Else
                    pbCuadro.ForeColor = &HC00000
                End If
                If pbCuadro.Tag <> "Libreta" Then
                    If (G_PUBLICAR_POSICION Or MostrarPosicion(iCategoria)) And Not bSalidaAPista Then
                        If iFase = 1 Then
                            pbCuadro.Print sMargen & mml_FRASE0831
                        Else
                            pbCuadro.Print sMargen & mml_FRASE0832
                        End If
                    Else
                        pbCuadro.Print sMargen & mml_FRASE0833
                    End If
                End If
                If G_RESULTADOS_ALTO_CONTRASTE Then
                    pbCuadro.ForeColor = C_COLOR_BLANCO
                Else
                    pbCuadro.ForeColor = 0
                End If
            End If
            
            ' Recuperamos la pareja
            sCad = lbOrden.List(iCParejas)
            
            'Ajuste de cambio de columna
            If iCParejasImpresas >= iMaxParejasCol Then
                pbCuadro.CurrentX = pbCuadro.Width / 2 + 60
            End If
            If iCParejasImpresas < iMaxParejasCol And pbCuadro.Tag = "Libreta" Then
                If Round(pbCuadro.FontSize) = C_FUENTE_PEQUE_LIBRETA Then
                    sCad = "     " & sCad
                ElseIf Round(pbCuadro.FontSize) = C_FUENTE_MUY_PEQUE_LIBRETA Then
                    sCad = "               " & sCad
                End If
            End If
            
            iPosX = pbCuadro.CurrentX
            iPosY = pbCuadro.CurrentY
            'Se mueve un poco para imprimir la sombra
            pbCuadro.CurrentX = iPosX - 50
            pbCuadro.CurrentY = iPosY + 50
            'pbCuadro.ForeColor = &HC0C0C0
            pbCuadro.ForeColor = &HFFFFFF
            
            If Not G_RESULTADOS_ALTO_CONTRASTE Then
                'Imprime la sombra
                pbCuadro.Print sCad
            End If
            
            pbCuadro.CurrentX = iPosX
            pbCuadro.CurrentY = iPosY
            'Cambiamos a verde si por ahora entra en la siguiente fase y hay puntuaciones
            If G_RESULTADOS_ALTO_CONTRASTE Then
                If Mid$(Trim$(sCad), 1, 2) <> "--" And InStr(sComen, C_SALIDA_PROXIMA) = 0 And iFase > 1 And (G_PUBLICAR_POSICION Or MostrarPosicion(iCategoria)) And Not bSalidaAPista Then
                    If (iInicial - iCParejas) < (iFase - 1) * 6 Then
                        pbCuadro.ForeColor = C_COLOR_BLANCO
                    ElseIf (iInicial - iCParejas) = (iFase - 1) * 6 Then
                        pbCuadro.ForeColor = C_COLOR_AMARILLO_CLARO
                    Else
                        pbCuadro.ForeColor = C_COLOR_AMARILLO_CLARO
                    End If
                Else
                    pbCuadro.ForeColor = C_COLOR_BLANCO
                End If
            Else
                If Mid$(Trim$(sCad), 1, 2) <> "--" And InStr(sComen, C_SALIDA_PROXIMA) = 0 And iFase > 1 And (G_PUBLICAR_POSICION Or MostrarPosicion(iCategoria)) And Not bSalidaAPista Then
                    If (iInicial - iCParejas) < (iFase - 1) * 6 Then
                        pbCuadro.ForeColor = &H8000&
                    ElseIf (iInicial - iCParejas) = (iFase - 1) * 6 Then
                        pbCuadro.ForeColor = &H8080&
                    Else
                        pbCuadro.ForeColor = &H80&
                    End If
                Else
                    pbCuadro.ForeColor = 0
                End If
            End If
            iTamFuente = pbCuadro.FontSize
            
            'Imprime la pareja
            pbCuadro.Print sCad;
            Inc iCParejasImpresas
            If pbCuadro.Tag = "Libreta" And rsParejas.RecordCount >= C_MAX_PAREJAS_POR_COLUMNA_LIBRETA Then
                pbCuadro.FontSize = iTamFuente - 1
            End If
            pbCuadro.Print
            pbCuadro.FontSize = iTamFuente
        Next
        DoEvents: DoEvents
        rsParejas.Close
    End If
End Function


Function Nombre(sNom As String, iLongNombre) As String
Dim iPos As Integer

    iPos = InStr(sNom, ",")
    If iPos > 0 Then
        Nombre = Mid$(sNom, iPos + 1)
    Else
        iPos = InStr(sNom, " ")
        If iPos > 0 Then
            Nombre = Mid$(sNom, 1, iPos - 1)
        Else
            Nombre = sNom
        End If
    End If
    
    If Len(Nombre) > iLongNombre Then
        Nombre = Left$(Nombre, iLongNombre - 1) & "."
    End If
    Nombre = Trim$(Nombre)
End Function

Function sDescFase(iFase As Integer) As String
    Select Case iFase
        Case 1:
            sDescFase = mml_FRASE0329
        Case 2:
            sDescFase = mml_FRASE0330
        Case 4:
            sDescFase = mml_FRASE0652
        Case 8:
            sDescFase = mml_FRASE0653
        Case Else
            sDescFase = "1/" & Trim$(Str$(iFase))
    End Select
End Function

Sub PrintConSombra(pbCuadro As PictureBox, sCad As String)
Dim iPosX As Integer, iPosY As Integer
    iPosX = pbCuadro.CurrentX
    iPosY = pbCuadro.CurrentY
    
    pbCuadro.CurrentX = pbCuadro.CurrentX - 30
    pbCuadro.CurrentY = pbCuadro.CurrentY + 40
    If G_RESULTADOS_ALTO_CONTRASTE Then
        pbCuadro.ForeColor = vbBlack
    Else
        pbCuadro.ForeColor = vbWhite
    End If
    pbCuadro.Print sCad
    
    pbCuadro.CurrentX = iPosX
    pbCuadro.CurrentY = iPosY
    If G_RESULTADOS_ALTO_CONTRASTE Then
        pbCuadro.ForeColor = vbWhite
    Else
        pbCuadro.ForeColor = vbBlack
    End If
    pbCuadro.Print sCad
    
End Sub

