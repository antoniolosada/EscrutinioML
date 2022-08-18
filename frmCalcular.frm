VERSION 5.00
Begin VB.Form frmCalcular 
   Caption         =   "mml_FRASE0058"
   ClientHeight    =   6645
   ClientLeft      =   270
   ClientTop       =   555
   ClientWidth     =   10140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tbLog 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3360
      Width           =   10095
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0058"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton CommandButton2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1845
         Picture         =   "frmCalcular.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1305
         Width           =   450
      End
      Begin VB.CommandButton CommandButton1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1845
         Picture         =   "frmCalcular.frx":046A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   855
         Width           =   450
      End
      Begin VB.CommandButton cmdSelComp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1845
         Picture         =   "frmCalcular.frx":08D4
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   450
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
         Height          =   495
         Left            =   8640
         TabIndex        =   26
         Top             =   2745
         Width           =   1350
      End
      Begin VB.CommandButton cmdPublicar 
         Caption         =   "mml_FRASE0186"
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
         Left            =   7200
         TabIndex        =   25
         Top             =   2745
         Width           =   1395
      End
      Begin VB.CommandButton cmdImprimirHojas 
         Caption         =   "mml_FRASE0463"
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
         Left            =   5175
         TabIndex        =   24
         Top             =   2745
         Width           =   1980
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "mml_FRASE0065"
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
         Left            =   3600
         TabIndex        =   23
         Top             =   2745
         Width           =   1530
      End
      Begin VB.CommandButton cmdAutomatico 
         Caption         =   "mml_FRASE0462"
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
         Left            =   1755
         TabIndex        =   22
         Top             =   2745
         Width           =   1800
      End
      Begin VB.CommandButton cmdCalcular 
         Caption         =   "mml_FRASE0058"
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
         Left            =   90
         TabIndex        =   21
         Top             =   2745
         Width           =   1620
      End
      Begin VB.Frame frmSelecParejas 
         Caption         =   "mml_FRASE0457"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   8535
         Begin VB.TextBox tbEmpates 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   7440
            TabIndex        =   20
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox tbMaxParejas 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   1080
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox tbMinParejas 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            Left            =   2280
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox tbMaxSemiOFinal 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   13
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "mml_FRASE0458"
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
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "mml_FRASE0459"
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
            Left            =   1680
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "mml_FRASE0460"
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
            Left            =   2880
            TabIndex        =   17
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label7 
            Caption         =   "mml_FRASE0461"
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
            Left            =   5640
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkRep 
         Caption         =   "mml_FRASE0289"
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
         Left            =   8400
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox tbDescFase 
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
         TabIndex        =   9
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox tbCodFase 
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
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox tbDescCat 
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
         TabIndex        =   6
         Top             =   840
         Width           =   6615
      End
      Begin VB.TextBox tbCodCat 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   5
         TabIndex        =   5
         Top             =   840
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
         TabIndex        =   3
         Top             =   360
         Width           =   6615
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
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblAutoPPC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AutoPPC"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9180
         TabIndex        =   30
         Top             =   1950
         Width           =   825
      End
      Begin VB.Label lblComp 
         Caption         =   "mml_FRASE0299"
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "mml_FRASE0301"
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
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCalcular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iJueces As Integer ' Número de jueces
Dim iParejas As Integer ' número de parejas
Dim iBailes As Integer ' Número de bailes detectado en la BD

Dim sDatosReglas() As String

Dim POS_MAY_ABS  As Integer
Dim POS_TOTAL_BAILE  As Integer

' Tablas para calcular posiciones en bailes individuales
' La primera columna se reserva para los dorsales
Dim aPosBaile() As Integer ' (Parejas, Jueces, Bailes)
Dim aCalBaile() As TCal ' (Parejas, Puestos, Bailes) ' Puestos + posición de mayoría abs de bailes + pos num_dorsales
' Public aPosTotalBaile(PAREJAS_FINAL, PAREJAS_FINAL, BAILES)   As Integer ' (Parejas, Puestos, Bailes)
' Tablas para calcular posiciones en el conjunto de bailes
Dim aSumaTotalConjunto() As Double ' (Suma total , número de dorsal)
Dim aCalConjunto() As TCal ' (Parejas, puestos) en la posición 0 el número de dorsal
Dim aOrdenFinal() As Double   ' (Parejas, Dorsales + Posición)

Dim aBailes() As TCodDesc
Dim iNumBailes As Integer
Dim aDorsales() As Integer
Dim iNumDorsales As Integer
Dim aJueces() As String
Dim iNumJueces As Integer
' Al cerrar el cuadro indicará si se han impreso las hojas
Dim ImpresionHojas As Boolean

Private Sub addLog(sCad As String)
    tbLog.Text = sCad & Chr$(13) & Chr$(10) & tbLog.Text
End Sub

Private Sub cmdAutomatico_Click()
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    'Publicar resultados
    If VarCfg("publicacion") = "publicacion_auto" Then
        If Val(tbCodFase.Text) > 1 Then
            frmPublicar.PublicacionCompleta tbCodComp.Text, tbCodCat.Text, Val(tbCodFase.Text) / 2, tbDescCat.Text, 0, tbDescCat.Text, tbDescComp.Text
        End If
    End If
    'Imprimir tablas, resumen y sig fase
    frmImprimirFinal.ImpresionCompleta tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, chkRep.Value
    
    
    'Imprimir hojas de puntuaciones
    'Comprobación de si la pista es por PPC y no necesita hojas
    Dim iPistas As Integer
    Dim aPistas(10) As String
    Dim bPistaPPC As Boolean
    Dim i As Integer
    
    bPistaPPC = False
    If Len(G_PISTAS_PPC) > 0 Then
        iPistas = DividirCampo(G_PISTAS_PPC, aPistas, ",")
        For i = 0 To iPistas - 1
            If InStr(tbDescCat.Text, aPistas(i)) > 0 Then
                bPistaPPC = True
            End If
        Next
    End If
    
    If Not bPistaPPC And Val(tbCodFase.Text) > 1 And G_AUTO_IMP_HOJAS_PUNTUACION Then
        If G_IMP_HOJAS_BAILE_EN_FINALES Then
            'Las finales deben ser hojas normales y las eliminatorias ópticas
            If Val(tbCodFase.Text) / 2 > 1 Then
                AsignarParametro "tipo_hoja_puntuaciones", "hojas_rec_optico"
            Else
                AsignarParametro "tipo_hoja_puntuaciones", "hoja_rec_por_baile"
            End If
        End If
    
        frmImpHojasPuntuaciones.tbCodComp.Text = tbCodComp.Text
        frmImpHojasPuntuaciones.tbDescComp.Text = tbDescComp.Text
        frmImpHojasPuntuaciones.tbCodCat.Text = tbCodCat.Text
        frmImpHojasPuntuaciones.tbDescCat.Text = tbDescCat.Text
        frmImpHojasPuntuaciones.tbCodFase.Text = Val(tbCodFase.Text) / 2
        frmImpHojasPuntuaciones.chkRep.Value = 0
        frmImpHojasPuntuaciones.CubrirInfoTandas
        frmImpHojasPuntuaciones.Tag = mml_FRASE0464
        frmImpHojasPuntuaciones.ImpresionDirecta
        
        'Si corresponde imprimimos las hojas de combinación de dorsales
        If CombinarDorsalesCateg(Val(tbCodCat.Text)) Then
            CombinarDorsales tbCodCat.Text, tbCodFase.Text, chkRep.Value, Val(frmImpHojasPuntuaciones.tbCombinarTandas.Text), True
            For i = 1 To C_NO_COPIAS_COMBINACION
                frmImpHojasPuntuaciones.ImprimirTandas
            Next
        End If
    End If
    
    If Not G_GEN_AUTO_RESULTADOS_PPC Then MsgBox "Se han realizado las impresiones de hojas y publicaciones solicitadas.", vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Public Sub cmdCalcular_Click()
Dim rs As Recordset, bResultado As Boolean, bTeamMatch, bCountryAmateur As Boolean

    If Val(tbCodComp.Text) = 0 Or Val(tbCodCat.Text) = 0 Or Val(tbCodFase.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
    bTeamMatch = ComprobarSiTeamMatch(Val(tbCodCat.Text))
    bCountryAmateur = ComprobarSiCountryAmateur(Val(tbCodCat.Text))
    
    Set rs = db.OpenRecordset(" SELECT descripcion FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    tbDescComp.Text = rs!DESCRIPCION
    rs.Close
    Set rs = db.OpenRecordset(" SELECT descripcion FROM categorias WHERE codigo = " & tbCodCat.Text, dbOpenSnapshot)
    tbDescCat.Text = rs!DESCRIPCION
    rs.Close
    DescFase
    If Not BailesParciales(Val(tbCodCat.Text)) And VarCfg("cal_retraso_al_calcular") = "S" Then
        'Solo recalculamos el retraso del horario si es un cálculo inicial (si ya estaban generados los resultados y es una regeneración no se recalcula el retraso)
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM cal_baile WHERE cod_categoria = " & tbCodCat.Text & " AND fase = " & tbCodFase.Text & " AND repesca = " & chkRep.Value, dbOpenSnapshot)
        If rs.Fields(0) = 0 Then
            CalcularRetrasoAlCalcular Val(tbCodCat.Text), Val(tbCodFase.Text), Val(chkRep.Value)
        End If
        rs.Close
    End If
    If Val(tbCodFase.Text) = 1 Then
        If bTeamMatch Then
            CalcularTeamMatch
        ElseIf bCountryAmateur Then
            CalcularCountryAmateur
        Else
            If UCase(VarCfg("calcular_todos_los_puestos")) = "S" Then
                bResultado = CalcularFinal
            Else
                'No soporta reconocimiento parcial de bailes ni de jueces
                CalcularFinal1
            End If
        End If
    Else
        bResultado = CalcularNoFinal
    End If
    
    If Not bResultado Then Exit Sub
    
    cmdPublicar.Default = True
    If G_GEN_AUTO_RESULTADOS_PPC Then
        cmdAutomatico_Click
        ImpresionHojas = True
    ElseIf G_PREGUNTAR_IMPRESION_AUTO Then
        If MsgBox(mml_FRASE0467, vbYesNo Or vbQuestion, mml_FRASE0099) = vbYes Then
            cmdAutomatico_Click
            ImpresionHojas = True
        End If
    End If
    cmdAutomatico.Enabled = True
End Sub
Sub CalcularTeamMatch()
Dim rs As Recordset

    tbLog.Text = ""
    'Calculamos el sumatorio de todos los bailes de la categoría actual
    Set rs = db.OpenRecordset("SELECT provincia, SUM(puesto) FROM puntuaciones pu, parejas pa, dorsales d WHERE pu.cod_categoria = d.cod_categoria AND d.num_dorsal = pu.num_dorsal AND d.cod_pareja = pa.codigo AND d.fase = pu.fase AND pu.cod_categoria = " & tbCodCat.Text & " AND pu.fase = " & tbCodFase.Text & " GROUP BY provincia ORDER BY provincia", dbOpenSnapshot)
    tbLog.Text = tbLog.Text & CR & LF & "Resultados totales de la última categoría" & CR & LF
    tbLog.Text = tbLog.Text & "----------------------------------------------------" & CR & LF
    db.Execute "DELETE FROM resultadosteammatch WHERE cod_categoria = " & tbCodCat.Text
    While Not rs.EOF
        tbLog.Text = tbLog.Text & "COMUNIDAD:  " & rs!provincia & "   PTOS.TOTALES:  " & rs.Fields(1) & CR & LF
        db.Execute "INSERT INTO resultadosteammatch VALUES (" & tbCodCat.Text & "," & C_TEAMMATCH_POR_CATEGORIA & ",'" & rs!provincia & "','" & rs.Fields(1) & "')"
        rs.MoveNext
    Wend
    rs.Close
    'Calculamos el sumatorio de todos los bailes de todas las categorias del TeamMatch
    Set rs = db.OpenRecordset("SELECT provincia, SUM(puesto) FROM puntuaciones pu, parejas pa, dorsales d WHERE pu.cod_categoria = d.cod_categoria AND d.num_dorsal = pu.num_dorsal AND d.cod_pareja = pa.codigo AND d.fase = pu.fase AND pu.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & tbCodComp.Text & " AND descripcion LIKE '*" & C_TEAM_MATCH & "*') AND pu.fase = 1 GROUP BY provincia ORDER BY provincia", dbOpenSnapshot)
    tbLog.Text = tbLog.Text & CR & LF & "CLASIFICACIÓN PROVISIONAL" & CR & LF
    tbLog.Text = tbLog.Text & "----------------------------------------------------" & CR & LF
    While Not rs.EOF
        tbLog.Text = tbLog.Text & "COMUNIDAD:  " & rs!provincia & "   PTOS.TOTALES:  " & rs.Fields(1) & CR & LF
        db.Execute "INSERT INTO resultadosteammatch VALUES (" & tbCodCat.Text & "," & C_TEAMMATCH_TOTAL & ",'" & rs!provincia & "','" & rs.Fields(1) & "')"
        rs.MoveNext
    Wend
    rs.Close
End Sub
Sub CalcularCountryAmateur()
Dim rs As Recordset
Dim rsBaileDorsal As Recordset

    tbLog.Text = ""
    'Recuperamos los dorsales del grupo actual
    'Para cada dorsal sumamos todas las puntuaciones de todos los bailes
    
    
    'Calculamos el sumatorio de todos los bailes de la categoría actual por dorsal
    'En cada baile hay que establecer si el resultado fue malo-regular-bueno, para ello calculamos_ la puntuación máxima que corresponde a 3*número de juevos = pm
    'malo = 0 <= resultado < pm/3
    'regular = pm/3 < resultado < pm/3*2
    'bueno = pm/3*2 < resultado <= pm
    'Estos datos se guardan en resultadosfinales
    '
    'Se sumarán los datos de todos los bailes realizando la misma operación, pero ahora pm = 3*número de bailes
    
    'Borramos las puntuaciones calculadas
    db.Execute ("DELETE FROM cal_baile WHERE cod_categoria = " & tbCodCat.Text & " AND fase=" & tbCodFase.Text)
    db.Execute ("DELETE FROM cal_conjunto WHERE cod_categoria = " & tbCodCat.Text)
    
    Dim NumeroJueces As Integer
    'Identificamo el numero de jueces
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS num_jueces FROM categorias c, juez_categ j WHERE c.codigo = j.cod_categoria AND c.codigo = " & tbCodCat.Text, dbOpenSnapshot)
    If Not rs.EOF Then
        NumeroJueces = rs.Fields("num_jueces")
    End If
    rs.Close
    
    Dim PuntosMinimos As Integer
    PuntosMinimos = NumeroJueces
    Dim PuntosMaximos As Integer
    PuntosMaximos = NumeroJueces * 3
    Dim PuntuacionRepartir As Integer
    PuntuacionRepartir = PuntosMaximos - NumeroJueces
    
    'Localizamos el la máxima puntuación que tiene un dorsal para dividir estos puntos en tres grupos
    'Set rs = db.OpenRecordset("SELECT SUM(puesto) as suma_puestos, pu.cod_baile, pu.num_dorsal, pu.cod_categoria FROM puntuaciones pu, parejas pa, " & _
    '"dorsales d WHERE pu.cod_categoria = d.cod_categoria AND d.num_dorsal = pu.num_dorsal AND " & _
    '"d.cod_pareja = pa.codigo AND d.fase = pu.fase AND pu.cod_categoria = " & tbCodCat.Text & _
    '" AND pu.fase = " & tbCodFase.Text & " GROUP BY pu.cod_baile, pu.num_dorsal, pu.cod_categoria", dbOpenSnapshot)
    'While Not rs.EOF
    '    If PuntosMaximos < rs.Fields("suma_puestos") Then
    '        PuntosMaximos = rs.Fields("suma_puestos")
    '    End If
    '    rs.MoveNext
    'Wend
    'rs.Close
    
    Set rsBaileDorsal = db.OpenRecordset("SELECT SUM(puesto) as suma_puestos, pu.cod_baile, pu.num_dorsal, pu.cod_categoria FROM puntuaciones pu, parejas pa, " & _
    "dorsales d WHERE pu.cod_categoria = d.cod_categoria AND d.num_dorsal = pu.num_dorsal AND " & _
    "d.cod_pareja = pa.codigo AND d.fase = pu.fase AND pu.cod_categoria = " & tbCodCat.Text & _
    " AND pu.fase = " & tbCodFase.Text & " GROUP BY pu.cod_baile, pu.num_dorsal, pu.cod_categoria", dbOpenSnapshot)
    
    tbLog.Text = tbLog.Text & CR & LF & "Resultados por dorsal de la última categoría" & CR & LF
    tbLog.Text = tbLog.Text & "----------------------------------------------------" & CR & LF
    While Not rsBaileDorsal.EOF
        Dim Posicion As Integer
        Posicion = rsBaileDorsal.Fields("suma_puestos")
        If Posicion <= PuntosMinimos + PuntuacionRepartir / 3 Then
            Posicion = 1
        ElseIf Posicion <= PuntosMinimos + PuntuacionRepartir / 3 * 2 Then
            Posicion = 2
        Else
            Posicion = 3
        End If
        'posiciones_mi contiene el puesto para puesto = 0
        tbLog.Text = tbLog.Text & "Dorsal: " & rsBaileDorsal.Fields("num_dorsal") & ", Baile: " & rsBaileDorsal.Fields("cod_baile") & ", Pos: " & Posicion & CR & LF
        db.Execute "INSERT INTO cal_baile VALUES (" & rsBaileDorsal.Fields("num_dorsal") & "," & rsBaileDorsal.Fields("cod_categoria") & "," & rsBaileDorsal.Fields("cod_baile") & ",0," & Posicion & ",0,1,0)"
        rsBaileDorsal.MoveNext
    Wend
    rsBaileDorsal.Close
    
    'Localizamos la suma máxima de posiciones para estableces el tope máximo y el valor bien-regular-mal
    Dim rsPuestos As Recordset
    Set rsPuestos = db.OpenRecordset("SELECT SUM(posiciones_mi) as suma_puestos, num_dorsal, cod_categoria " & _
                            "FROM cal_baile WHERE cod_categoria = " & tbCodCat.Text & " AND fase = 1 AND repesca = 0 AND puesto = 0 " & _
                            "GROUP BY num_dorsal, cod_categoria", dbOpenSnapshot)
    Dim MinimaSumaPuestos As Integer
    MinimaSumaPuestos = 9999
    While Not rsPuestos.EOF
        If MinimaSumaPuestos > rsPuestos.Fields("suma_puestos") Then
            MinimaSumaPuestos = rsPuestos.Fields("suma_puestos")
        End If
        rsPuestos.MoveNext
    Wend
    rsPuestos.Close
    
     'Calculamos la el valor de posición
    Set rsPuestos = db.OpenRecordset("SELECT SUM(posiciones_mi) as suma_puestos, num_dorsal, cod_categoria " & _
                            "FROM cal_baile WHERE cod_categoria = " & tbCodCat.Text & " AND fase = 1 AND repesca = 0 " & _
                            "GROUP BY num_dorsal, cod_categoria", dbOpenSnapshot)
    Dim SumaPuestos As Integer
    Dim PuntosRepartir As Double
    'Maximos puntos = jueces * puntuación máxima (3)
    
    'Se modifica para que no dependa de las puntuaciones
    MinimaSumaPuestos = NumeroJueces
    PuntosRepartir = NumeroJueces * 3 - MinimaSumaPuestos
    
    tbLog.Text = tbLog.Text & "Posición final: ********************************" & CR & LF
    While Not rsPuestos.EOF
        SumaPuestos = rsPuestos.Fields("suma_puestos")
        If SumaPuestos < MinimaSumaPuestos + PuntosRepartir / 3 Then
            Posicion = 1
        ElseIf SumaPuestos < MinimaSumaPuestos + PuntosRepartir / 3 * 2 Then
            Posicion = 2
        Else
            Posicion = 3
        End If
        'puesto contiene el puesto final para regla = "FIN"
        tbLog.Text = tbLog.Text & "Dorsal: " & rsPuestos.Fields("num_dorsal") & ", PosFinal: " & Posicion & CR & LF
        db.Execute "INSERT INTO cal_conjunto VALUES (" & rsPuestos.Fields("num_dorsal") & "," & rsPuestos.Fields("cod_categoria") & "," & Posicion & ",'FIN',0)"
        rsPuestos.MoveNext
    Wend
    rsPuestos.Close
    
End Sub
Private Function CalcularFinal() As Boolean
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim iMayAbs As Integer ' Valor que marca la mayoría
Dim iPos As Integer, iSum As Integer ' Variables donde se almacena los valores acumulados sin que  pasen de la mayoría
Dim iPos1 As Integer, iSum1 As Integer ' Variables donde se almacena los valores acumulados si pasan de la mayoría
Dim iCBaile As Integer, iCJuez As Integer, iCDorsal As Integer ' Contadores para crear los arrays
Dim iPuesto As Integer ' Posición final por vaile
Dim iJuecesCateg As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    
    CalcularFinal = True
    
    If tbCodCat.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        CalcularFinal = False
        Exit Function
    End If
    
    'Borramos las puntuaciones de los no presentados
    db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal IN (SELECT DISTINCT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text & " AND no_presente>0) AND cod_categoria =" & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text)
    'Enviamos a los no presentados al final
    'db.Execute ("UPDATE puntuaciones SET puesto = " & C_MAX_PUESTO_FINAL_NO_PRESENTADO & " WHERE num_dorsal IN (SELECT DISTINCT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text & " AND no_presente>0) AND cod_categoria =" & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text)
        
    
    'Primero borramos cualquier calculo anterior
    tbLog.Text = ""
    addLog (mml_FRASE0468)
    db.Execute ("DELETE FROM cal_baile WHERE cod_categoria = " & tbCodCat.Text & " AND fase=" & tbCodFase.Text)
    db.Execute ("DELETE FROM cal_conjunto WHERE cod_categoria = " & tbCodCat.Text)
    
    ' Comprobamos el número de jueces de la categoría
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCat.Text, dbOpenSnapshot)
    iJuecesCateg = rs.Fields(0)
    addLog (mml_FRASE0469 & iJuecesCateg)
    rs.Close
    ' Comprobamos el número de jueces en las puntuaciones
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iJueces = 0
    If Not rs.EOF Then
        rs.MoveLast
        iJueces = rs.RecordCount
    End If
    iNumJueces = iJueces
    addLog (mml_FRASE0470 & iJueces)
    rs.Close
    If iJueces <> iJuecesCateg Then
        If MsgBox(mml_FRASE0471 & iJueces & mml_FRASE0472 & iJuecesCateg & ". ¿Continuar?", vbYesNo Or vbInformation, mml_FRASE0084) = vbNo Then
            CalcularFinal = False
        End If
    End If
    'comprobamos los bailes
    
    If C_CALCULOS_PARCIALES Or BailesParciales(Val(tbCodCat.Text)) Then
        ' Comprobamos el número de bailes que hay en las puntuaciones
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ")", dbOpenSnapshot)
    Else
        ' Comprobamos el número de bailes
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCat.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
    End If
    iBailes = rs.Fields(0)
    iNumBailes = iBailes
    addLog (mml_FRASE0473 & iBailes)
    rs.Close
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    addLog (mml_FRASE0474 & iParejas)
    If rs.Fields(0) Mod iBailes * iJueces > 0 Then
        If MsgBox(mml_FRASE0475, vbYesNo Or vbCritical, mml_FRASE0096) = vbNo Then
            CalcularFinal = False
            Exit Function
        End If
    End If
    rs.Close
    ' Comprobamos si están todos los dorsales
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca = " & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    If rs.Fields(0) <> iParejas Then
        If Not G_GEN_AUTO_RESULTADOS_PPC Then
            If MsgBox(mml_FRASE0476 & iParejas & mml_FRASE0472 & rs.Fields(0) & mml_FRASE0477, vbYesNo Or vbCritical, mml_FRASE0096) = vbNo Then
                CalcularFinal = False
                Exit Function
            End If
        End If
    End If
    rs.Close
    
    
    'Calculamos el valor de mayoría absoluta
    iMayAbs = Int(iJueces / 2)
    ' Primero comenzamos calculando las posiciones
    
'Inicializamos el espacio para las tablas *********************************************
POS_MAY_ABS = iParejas + 1
POS_TOTAL_BAILE = iParejas + 2
' Tablas para calcular posiciones en bailes individuales
' La primera columna se reserva para los dorsales
ReDim aPosBaile(iParejas, iJueces + 1, iBailes) As Integer ' (Parejas, Jueces, Bailes)
ReDim aCalBaile(iParejas, iParejas + 2, iBailes) As TCal ' (Parejas, Puestos, Bailes) ' Puestos + posición de mayoría abs de bailes + pos num_dorsales
' Public aPosTotalBaile(PAREJAS_FINAL, PAREJAS_FINAL, BAILES)   As Integer ' (Parejas, Puestos, Bailes)
' Tablas para calcular posiciones en el conjunto de bailes
ReDim aSumaTotalConjunto(iParejas, 1) As Double ' (Suma total , número de dorsal)
ReDim aCalConjunto(iParejas, iParejas) As TCal ' (Parejas, puestos) en la posición 0 el número de dorsal
ReDim aOrdenFinal(iParejas, 2) ' (Parejas, Dorsales + Posición)

ReDim aBailes(iBailes) As TCodDesc
ReDim aDorsales(iParejas) As Integer
ReDim aJueces(iJueces) As String
'**************************************************************************************
Dim iNumDesc As Integer
    
    If C_CALCULOS_PARCIALES Or BailesParciales(Val(tbCodCat.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase =" & BAILES_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCat.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo idependiente
    iCBaile = 0
    While Not rsBailes.EOF
        addLog (mml_FRASE0478 & rsBailes.Fields(1) & "   ----------------------")
        aBailes(iCBaile).codigo = rsBailes.Fields(0)
        aBailes(iCBaile).DESCRIPCION = rsBailes.Fields(1)
        'Para cada dorsal
        iCDorsal = 0
        ' Recuperamos los dorsales a clasificar
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            addLog (mml_FRASE0225 & rsDorsales!num_dorsal)
            
            addLog (mml_FRASE0479 & rsDorsales!num_dorsal)
            ' Por si hay descalificaciones cambiamos todas las puntuaciones > C_PUESTO_NEG
            db.Execute ("UPDATE puntuaciones SET puesto = puesto-" & C_PUESTO_NEG & " WHERE puesto >" & C_PUESTO_NEG & " AND cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
            ' Comprobamos las descalificaciones de jueces normales------------------------------------
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones d, juez_categ j WHERE j.pasos = 0 AND d.id_juez = j.id_juez AND j.cod_categoria = d.cod_categoria AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca = " & chkRep.Value & " AND d.fase = " & tbCodFase.Text & " AND d.cod_baile=" & rsBailes!cod_baile & " AND d.num_dorsal=" & rsDorsales!num_dorsal, dbOpenSnapshot)
            iNumDesc = rs.Fields(0)
            rs.Close
            ' Estamos en la final y es descalificado por más de un juez
            If iNumDesc >= G_MIN_JUECES_DESCALIFICACION Then
                If MsgBox(mml_FRASE0300 & rsDorsales!num_dorsal & mml_FRASE0480 & iNumDesc & mml_FRASE0481, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                    db.Execute ("UPDATE puntuaciones SET puesto = " & C_PUESTO_NEG & "+puesto WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
                    If Not G_GEN_AUTO_RESULTADOS_PPC Then MsgBox mml_FRASE0482, vbOKOnly Or vbInformation, mml_FRASE0086
                End If
            End If
            ' Comprobamos las descalificaciones del juez de pasos y figuras --------------------------
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones d, juez_categ j WHERE j.pasos = 1 AND d.id_juez = j.id_juez AND j.cod_categoria = d.cod_categoria AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & chkRep.Value & " AND d.fase = " & tbCodFase.Text & " AND d.cod_baile=" & rsBailes!cod_baile & " AND d.num_dorsal=" & rsDorsales!num_dorsal, dbOpenSnapshot)
            iNumDesc = rs.Fields(0)
            rs.Close
            ' Si ha sido descalificada por el juez de pasos y figuras
            If iNumDesc > 0 Then
                If MsgBox(mml_FRASE0300 & rsDorsales!num_dorsal & mml_FRASE0480 & iNumDesc & mml_FRASE0483, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                    db.Execute ("UPDATE puntuaciones SET puesto = " & C_PUESTO_NEG & "+puesto WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
                    If Not G_GEN_AUTO_RESULTADOS_PPC Then MsgBox mml_FRASE0482, vbOKOnly Or vbInformation, mml_FRASE0086
                End If
            End If
            '-----------------------------------------------------------------------------------------
            
            'Grabamos en la primera columna el número de dorsal
            aCalBaile(iCDorsal, POS_NUM_DORSAL, iCBaile).iNPos = rsDorsales!num_dorsal
            'Primero recuperamos todos los puestos (puntuaciones de cada juez) del dorsal en este baile
            Set rsPuestos = db.OpenRecordset( _
            "SELECT puesto,cod_juez FROM puntuaciones " & _
            "WHERE fase = 1 " & _
            " AND repesca=" & chkRep.Value & _
            " AND cod_categoria = " & tbCodCat.Text & _
            " AND cod_baile = " & rsBailes!cod_baile & _
            " AND num_dorsal = " & rsDorsales!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            
            ' Almacenamos los puestos ordenados por la puntuación de cada juez
            Set rsPuestosJuez = db.OpenRecordset( _
            "SELECT puesto,cod_juez FROM puntuaciones " & _
            "WHERE fase = 1 " & _
            " AND repesca=" & chkRep.Value & _
            " AND cod_categoria = " & tbCodCat.Text & _
            " AND cod_baile = " & rsBailes!cod_baile & _
            " AND num_dorsal = " & rsDorsales!num_dorsal & " ORDER BY 2", dbOpenSnapshot)
            iCJuez = 1
            While Not rsPuestosJuez.EOF
                aPosBaile(iCDorsal, iCJuez, iCBaile) = rsPuestosJuez!Puesto
                'Control de puestos con valor 0
                If aPosBaile(iCDorsal, iCJuez, iCBaile) = 0 Then
                    addLog (mml_FRASE0484 & iCDorsal & mml_FRASE0485 & iCJuez & mml_FRASE0486 & iCBaile)
                End If
                iCJuez = iCJuez + 1
                rsPuestosJuez.MoveNext
            Wend
            'Grabamos en la primera columna el número de dorsal
            aPosBaile(iCDorsal, POS_NUM_DORSAL, iCBaile) = rsDorsales!num_dorsal
            rsPuestosJuez.Close
            
            
            'Calculamos las mejores posiciones y la suma de las mismas
            iPos = 0
            iSum = 0
            iPos1 = 0
            iSum1 = 0
            iCJuez = 1 ' La posición 0 queda reservada para el número de dorsal
            For j = 1 To iParejas ' Puestos
                Do While Not rsPuestos.EOF
                    If rsPuestos!Puesto <= j Then
                        iCJuez = iCJuez + 1
                        iPos = iPos + 1 ' Numero de puestos acumulados
                        iSum = iSum + rsPuestos!Puesto
                        rsPuestos.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                'Colocamos 0 a no ser que se supere la mayoría absoluta
                If iPos > iMayAbs Then 'iPos = Suma de puestos ant 0 = al estudiado
                    'Si en alguno pasamos de la mayoría podemos ordenarlo
                    iPos1 = iPos
                    iSum1 = iSum
                    'Comprobamos si es la primera vez que pasamos la mayoría abs
                    If aCalBaile(iCDorsal, POS_MAY_ABS, iCBaile).iNPos = 0 Then
                        aCalBaile(iCDorsal, POS_MAY_ABS, iCBaile).iNPos = j
                    End If
                End If
                aCalBaile(iCDorsal, j, iCBaile).iNPos = iPos1
                aCalBaile(iCDorsal, j, iCBaile).iSum = iSum1
                db.Execute ("INSERT INTO cal_baile VALUES(" & rsDorsales!num_dorsal & "," & tbCodCat.Text & "," & rsBailes!cod_baile & "," & j & "," & iPos1 & "," & iSum1 & " ,1," & chkRep.Value & ")")
            Next j
            While Not rsPuestos.EOF
                rsPuestos.MoveNext
            Wend
            If rsPuestos.RecordCount < iJueces Then
                MsgBox mml_FRASE0487 & rsDorsales!num_dorsal & mml_FRASE0488 & rsPuestos.RecordCount & mml_FRASE0489 & rsBailes!Nombre, vbOKOnly Or vbCritical, mml_FRASE0096
                CalcularFinal = False
                Exit Function
            End If
            
            
            'Si no se ha presentado
            If aCalBaile(iCDorsal, POS_MAY_ABS, iCBaile).iNPos = 0 Then
                'Si no hemos superado la mayoría en algún puesto, le corresponde el último
                aCalBaile(iCDorsal, POS_MAY_ABS, iCBaile).iNPos = iParejas + 1
            End If
            rsPuestos.Close
            iCDorsal = iCDorsal + 1
            rsDorsales.MoveNext
        Wend ' Dorsales
        iCBaile = iCBaile + 1
        rsBailes.MoveNext
        rsDorsales.Close
    Wend 'Bailes
    rsBailes.Close
    
    ' Ahora tenemos todos los cuadros de todos los bailes y tenemos que calcular
    
    Call Ordenacion
    Exit Function
error:
    ProcesarError "CalcularFinal"
End Function

Private Function iValorCriterio(iPareja As Integer, ByVal iCriterio As Integer, iBaile As Integer) As Integer
    ' El menor => el mejor
    If iCriterio = 0 Then
        iValorCriterio = aCalBaile(iPareja, POS_MAY_ABS, iBaile).iNPos
    Else
        iValorCriterio = -(aCalBaile(iPareja, iCriterio, iBaile).iNPos * OP_CRITERIOS - aCalBaile(iPareja, iCriterio, iBaile).iSum)
    End If
    
End Function

Private Sub Ordenacion()
Dim i As Integer, j As Integer, k As Integer, X As Integer
Dim iPuestoAct As Integer, iNumPuestos As Integer, iValorPuesto As Integer, iValor As Integer
Dim iCCriterio As Integer
Dim iCPuesto As Integer
Dim iCParejas As Integer
Dim iCBaile As Integer

ReDim sDatosReglas(iParejas)

    addLog (mml_FRASE0490)

    For iCBaile = 0 To iBailes - 1
        'Inicializamos los puestos a 1
        For iCParejas = 0 To iParejas - 1
            aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos = 1
            aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iSum = 1
        Next iCParejas
        
        'Por cada puesto evaluamos todos los criterios hasta el desempate
        iCPuesto = 1
        Do While iCPuesto <= iParejas
        '·{M} Para comprobar hasta la última posición
            'For iCCriterio = 0 To iParejas - 1 ' Criterios = (Puestos-1)PosMayAbs
            For iCCriterio = 0 To iParejas  ' Criterios = (Puestos-1)PosMayAbs
                iNumPuestos = 0
                iValorPuesto = OP_CRITERIOS * 10 - 1 ' Valor más grande
                ' Localizamos el mejor valor para este criterio
                For iCParejas = 0 To iParejas - 1
                    ' Solo ordenamos las que tengan un determinado puesto
                    If iCPuesto = aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos Then
                        iValor = iValorCriterio(iCParejas, iCCriterio, iCBaile)
                        If iValorPuesto > iValor Then
                            iValorPuesto = iValor
                        End If
                    End If
                Next iCParejas
                ' Si hay más de una pareja empatada con este criterio pasamos al siguiente
                For iCParejas = 0 To iParejas - 1
                    ' Solo las que tengan un determinado puesto
                    If iCPuesto = aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos Then
                        iValor = iValorCriterio(iCParejas, iCCriterio, iCBaile)
                        ' Asignamos a todos los que tengan este valor iPuesto, a los demás si no tienen puesto ya definido el valor superior
                        If iValorPuesto = iValor Then
                            aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos = iCPuesto
                            iNumPuestos = iNumPuestos + 1
                        ElseIf aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos >= iCPuesto Then
                            aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos = iCPuesto + 1
                        End If
                    End If
                Next iCParejas
                ' Si iNumPuestos > 1 => Empate si no seguimos con el siguiente puesto
                If iNumPuestos = 1 Then
                    Exit For
                End If
            Next iCCriterio
            ' Si hemos acabado todos los criterios y persiste el empate
            If iNumPuestos > 1 Then
                addLog (mml_FRASE0491 & iNumPuestos & mml_FRASE0492 & iCPuesto & mml_FRASE0493 & aBailes(iCBaile).DESCRIPCION)
                'Avanzamos el mismo número de puestos que hay en el empate -1,
                'ya que están asignados por los empatados
                ' A todos los empatados les asignamos un valor medio
                
                For iCParejas = 0 To iParejas - 1
                    ' Solo las que tengan un determinado puesto
                    If iCPuesto = aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos Then
                        aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos = Redondea(iCPuesto + (iNumPuestos - 1) / 2, 1)
                    ElseIf iCPuesto + 1 = aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos Then
                        aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos = iCPuesto + iNumPuestos
                    End If
                Next iCParejas
                
                iCPuesto = iCPuesto + iNumPuestos
            Else
                Inc iCPuesto
            End If
        Loop
        'Despues de cada baile mostramos las posiciones
        addLog (mml_FRASE0494 & aBailes(iCBaile).DESCRIPCION)
        ' En el Puesto 0 en la base de datos en la tabla Cal_Baile está la posición final por baile
        For iCParejas = 0 To iParejas - 1
            addLog (mml_FRASE0300 & aCalBaile(iCParejas, 0, iCBaile).iNPos & mml_FRASE0495 & aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos)
            Debug.Print "INSERT INTO cal_baile VALUES(" & aCalBaile(iCParejas, 0, iCBaile).iNPos & "," & tbCodCat.Text & "," & aBailes(iCBaile).codigo & "," & 0 & ",'" & aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos & "'," & 0 & " ,1," & chkRep.Value & ")"
            db.Execute ("INSERT INTO cal_baile VALUES(" & aCalBaile(iCParejas, 0, iCBaile).iNPos & "," & tbCodCat.Text & "," & aBailes(iCBaile).codigo & "," & 0 & ",'" & aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos & "'," & 0 & " ,1," & chkRep.Value & ")")
        Next
    Next iCBaile
    
    'Ahora debemos pasar a la ordenación general **************************************
    'Primero calculamos la suma de los bailes por cada pareja
    addLog (mml_FRASE0496)
    For iCParejas = 0 To iParejas - 1
        For iCBaile = 0 To iBailes - 1
            aSumaTotalConjunto(iCParejas, POS_TOTAL_CONJUNTO) = aSumaTotalConjunto(iCParejas, POS_TOTAL_CONJUNTO) + _
                                aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBaile).iNPos
            aSumaTotalConjunto(iCParejas, POS_NUM_DORSAL) = aCalBaile(iCParejas, POS_NUM_DORSAL, iCBaile).iNPos
        Next iCBaile
        addLog (mml_FRASE0300 & aSumaTotalConjunto(iCParejas, POS_NUM_DORSAL) & mml_FRASE0497 & aSumaTotalConjunto(iCParejas, POS_TOTAL_CONJUNTO))
    Next iCParejas
    'Ahora calculamos los valores de desempate por la regla 11
Dim iMayAbsR11 As Integer
Dim iPos As Integer
Dim iPuestosAcum As Integer
Dim iSumaPuestosAcum As Integer
Dim iCJuez As Integer
Dim iPosAcumBaile As Integer

    iMayAbsR11 = Int((iJueces * iBailes) / 2)
    For iCParejas = 0 To iParejas - 1 ' Por cada pareja
        For iCPuesto = 1 To iParejas - 1 ' Por cada puesto
            iPuestosAcum = 0
            iSumaPuestosAcum = 0
            For iCBaile = 0 To iBailes - 1 ' Por cada baile
                For iCJuez = 1 To iJueces  ' Por cada resultadado de un juez ' 0= num_dorsal
                    iPosAcumBaile = aPosBaile(iCParejas, iCJuez, iCBaile)
                    If iPosAcumBaile <= iCPuesto Then
                        iPuestosAcum = iPuestosAcum + 1
                        iSumaPuestosAcum = iSumaPuestosAcum + iPosAcumBaile
                    End If
                Next iCJuez
            Next iCBaile
            ' Ahora comprobamos si supera la mayoría
            If iPuestosAcum > iMayAbsR11 Then
                aCalConjunto(iCParejas, iCPuesto).iNPos = iPuestosAcum
                aCalConjunto(iCParejas, iCPuesto).iSum = iSumaPuestosAcum
            Else
                aCalConjunto(iCParejas, iCPuesto).iNPos = 0
                aCalConjunto(iCParejas, iCPuesto).iSum = 0
            End If
            ' Grabamos el número de dorsal
            aCalConjunto(iCParejas, POS_NUM_DORSAL).iNPos = aPosBaile(iCParejas, POS_NUM_DORSAL, 0)
        Next iCPuesto
    Next iCParejas
    
    ' Inicializamos todos a la primera posición
    For iCParejas = 0 To iParejas - 1
        aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = 1
        aOrdenFinal(iCParejas, POS_NUM_DORSAL) = aCalConjunto(iCParejas, POS_NUM_DORSAL).iNPos
    Next iCParejas
    ' Realizamos el cálculo final de posiciones
Dim iEmpates As Integer
    iCPuesto = 1
        Do While iCPuesto <= iParejas
            iEmpates = R9(iCPuesto)
            GuardarPosiciones "R09", iCPuesto
            If iEmpates > 1 Then
                addLog (mml_FRASE0498 & iEmpates & mml_FRASE0499 & iCPuesto)
                iEmpates = R10(iCPuesto)
                GuardarPosiciones "R10", iCPuesto
                If iEmpates > 1 Then
                    addLog (mml_FRASE0498 & iEmpates & mml_FRASE0500 & iCPuesto)
                    'Hay Empates
                    iEmpates = R11(iCPuesto)
                    GuardarPosiciones "R11", iCPuesto
                    iCPuesto = iCPuesto + iEmpates
                    If iEmpates > 1 Then
                        addLog (mml_FRASE0498 & iEmpates & mml_FRASE0501 & iCPuesto)
                    End If
                Else
                    ' No hay empates
                    iCPuesto = iCPuesto + 1
                End If
            Else
                ' No hay empates
                iCPuesto = iCPuesto + 1
            End If
        Loop
    ' Presentamos el resultado de la ordenación
    For iCParejas = 0 To iParejas - 1
        addLog (mml_FRASE0300 & aOrdenFinal(iCParejas, POS_NUM_DORSAL) & mml_FRASE0495 & aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO))
        db.Execute "INSERT INTO cal_conjunto VALUES(" & aOrdenFinal(iCParejas, POS_NUM_DORSAL) & "," & tbCodCat.Text & "," & aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) & ",'FIN',0)", dbOpenSnapshot
    Next
End Sub
Private Sub GuardarPosiciones(sReglaAplicada As String, iPosicion As Integer)
Dim iCParejas As Integer
Dim sRegla As String
    For iCParejas = 0 To iParejas - 1
        
        If sDatosReglas(iCParejas) <> "" Then
            sRegla = sReglaAplicada & "-" & sDatosReglas(iCParejas)
        Else
            sRegla = sReglaAplicada
        End If
        db.Execute ("INSERT INTO cal_conjunto VALUES (" & aOrdenFinal(iCParejas, POS_NUM_DORSAL) & " , " & tbCodCat.Text & "," & aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) & ",'" & sRegla & "'," & iPosicion & ")")
    Next
    ReDim sDatosReglas(20)
End Sub

Private Function R10(iCPuesto As Integer) As Integer
Dim dValorMinimo As Integer
Dim iCValorMinimo As Integer
Dim iCParejas As Integer
Dim iBailesPorPuesto As Integer
    
    dValorMinimo = VALOR_MAXIMO
    For iCParejas = 0 To iParejas - 1
        ' Solo puesto actual a controlar
        If iCPuesto = aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) Then
            iBailesPorPuesto = iNumeroDePuestosPorSuma(iCPuesto, iCParejas)
            If dValorMinimo > iBailesPorPuesto Then
                dValorMinimo = iBailesPorPuesto
                iCValorMinimo = 1
            ElseIf dValorMinimo = iBailesPorPuesto Then
                iCValorMinimo = iCValorMinimo + 1
            End If
        End If
    Next iCParejas
    
    'Asignamos los puestos
    For iCParejas = 0 To iParejas - 1
        ' Solo puesto actual a controlar
        If iCPuesto = aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) Then
            iBailesPorPuesto = iNumeroDePuestosPorSuma(iCPuesto, iCParejas)
            If dValorMinimo = iBailesPorPuesto Then
                aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = iCPuesto
            Else
                aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = iCPuesto + iCValorMinimo
            End If
        End If
    Next iCParejas
    R10 = iCValorMinimo
End Function

Function iNumeroDePuestosPorSuma(iCPuesto As Integer, iCParejas As Integer) As Integer
Dim iSumBailes As Integer
Dim iNumBailes As Integer
Dim iCBailes As Integer
    For iCBailes = 0 To iBailes - 1
        If aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBailes).iNPos <= iCPuesto Then
            iNumBailes = iNumBailes + 1
            iSumBailes = iSumBailes + aCalBaile(iCParejas, POS_TOTAL_BAILE, iCBailes).iNPos
        End If
    Next iCBailes
    iNumeroDePuestosPorSuma = -(iNumBailes * OP_CRITERIOS - iSumBailes)
    
    sDatosReglas(iCParejas) = "@" & iCPuesto & "=" & iNumBailes & "(" & iSumBailes & ")"
End Function
Function iNumeroDePuestosAcumPorSuma(iCPuesto As Integer, iCParejas As Integer) As Long
    iNumeroDePuestosAcumPorSuma = -(aCalConjunto(iCParejas, iCPuesto).iNPos * OP_CRITERIOS - aCalConjunto(iCParejas, iCPuesto).iSum)
    
    sDatosReglas(iCParejas) = "@" & iCPuesto & "=" & aCalConjunto(iCParejas, iCPuesto).iNPos & "(" & aCalConjunto(iCParejas, iCPuesto).iSum & ")"
End Function
Private Function R11(iCPuesto As Integer) As Integer
Dim iValorMinimo As Long
Dim iCValorMinimo As Long
Dim iCPosicion As Integer
Dim iCParejas As Integer
Dim iPuestosAcum As Long
Dim bDatosAdicionales As Boolean
    
    
    iValorMinimo = 0
    iCValorMinimo = 0
    iValorMinimo = VALOR_MAXIMO
    ' Si hay empate seguimos con el acumulado del siguiente puesto
    For iCPosicion = iCPuesto To iParejas - 1
        bDatosAdicionales = False
        For iCParejas = 0 To iParejas - 1
            ' Solo puesto actual a controlar
            If iCPuesto = aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) Then
                iPuestosAcum = iNumeroDePuestosAcumPorSuma(iCPosicion, iCParejas)
                If sDatosReglas(iCParejas) <> "" And Mid$(sDatosReglas(iCParejas), 4, 1) <> "0" Then
                    bDatosAdicionales = True
                End If
                
                If iValorMinimo > iPuestosAcum Then
                    iValorMinimo = iPuestosAcum
                    iCValorMinimo = 1
                ElseIf iValorMinimo = iPuestosAcum Then
                    iCValorMinimo = iCValorMinimo + 1
                End If
            End If
        Next iCParejas
        ' Si no hay empate, localizamos la ganadora y salimos
        If iCValorMinimo = 1 Then
            Exit For
        ElseIf bDatosAdicionales Then
            GuardarPosiciones "R11", iCPuesto
            ReDim sDatosReglas(10)
        End If
    Next iCPosicion
    ' Si acabamos todos los puestos ordenamos por el último
    iCPosicion = IIf(iCPosicion > iParejas - 1, iParejas - 1, iCPosicion)
    'Asignamos los puestos
    For iCParejas = 0 To iParejas - 1
        ' Solo puesto actual a controlar
        If iCPuesto = aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) Then
            If iValorMinimo = iNumeroDePuestosAcumPorSuma(iCPosicion, iCParejas) Then
                aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = iCPuesto
            Else
                aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = iCPuesto + iCValorMinimo
            End If
        End If
    Next iCParejas
    R11 = iCValorMinimo
End Function


Private Function R9(iCPuesto As Integer) As Integer
Dim dValorMinimo As Double
Dim iCValorMinimo As Integer
Dim iCParejas As Integer
Dim iNumPuestos As Integer

    dValorMinimo = iParejas * iBailes + 10
    iNumPuestos = 0
    For iCParejas = 0 To iParejas - 1
        If iCPuesto = aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) Then
            If dValorMinimo > aSumaTotalConjunto(iCParejas, POS_TOTAL_CONJUNTO) Then
                dValorMinimo = aSumaTotalConjunto(iCParejas, POS_TOTAL_CONJUNTO)
                iCValorMinimo = 1
            ElseIf dValorMinimo = aSumaTotalConjunto(iCParejas, POS_TOTAL_CONJUNTO) Then
                iCValorMinimo = iCValorMinimo + 1
            End If
            iNumPuestos = iNumPuestos + 1
        End If
    Next iCParejas
    
    ' Si la posición solo la tiene un dorsal, no hay nada que ordenar
    If iNumPuestos > 1 Then
        'Asignamos los puestos
        For iCParejas = 0 To iParejas - 1
            ' Solo comprobamos el puesto actual
            If iCPuesto = aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) Then
                If dValorMinimo = aSumaTotalConjunto(iCParejas, POS_TOTAL_CONJUNTO) Then
                    aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = iCPuesto
                Else
                    aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = iCPuesto + iCValorMinimo
                End If
            End If
        Next iCParejas
    End If
    R9 = iCValorMinimo
End Function



Private Sub cmdGrupo_Click()
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM puntuaciones")

End Sub

Private Sub cmdImprimir_Click()
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    frmImprimirFinal.Imprimir tbCodComp.Text, tbCodCat.Text, tbCodFase.Text, chkRep.Value

End Sub

Private Sub cmdImprimirHojas_Click()
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If Val(tbCodFase.Text) = 1 Then
        MsgBox mml_FRASE0502, vbOKOnly Or vbInformation, mml_FRASE0086
        Exit Sub
    Else
        frmImpHojasPuntuaciones.ImprimirHojas tbCodComp.Text, tbCodCat.Text, Val(tbCodFase.Text) / 2, tbDescComp.Text, tbDescCat.Text
    End If
End Sub

Private Sub cmdPublicar_Click()
Dim iFaseSig As Integer
    If tbCodComp.Text = "" Or tbCodCat.Text = "" Or tbCodFase.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    If tbCodFase.Text = "1" And Not G_PUBLICAR_POSICION And Not MostrarPosicion(Val(tbCodCat.Text)) Then
        If MsgBox("La final no es eliminatoria y no debe publicarse. ¿Desea anular la publicación?", vbInformation Or vbYesNo, mml_FRASE0096) = vbYes Then
            Exit Sub
        End If
    End If
    If Val(tbCodFase.Text) > 1 Then
        iFaseSig = Val(tbCodFase.Text) / 2
    Else
        iFaseSig = 1
    End If
    frmPublicar.Publicar tbCodComp.Text, tbCodCat.Text, iFaseSig, tbDescCat.Text, chkRep.Value, tbDescCat.Text, tbDescComp.Text
    cmdImprimir.Default = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelComp_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    tbCodCat.Text = ""
    tbDescCat.Text = ""
    tbCodFase.Text = ""
    tbDescFase.Text = ""
End Sub

Private Sub CommandButton1_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCat.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text, "descripcion", " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCat.Text = sResultado(2)
    tbCodFase.Text = ""
    tbDescFase.Text = ""
End Sub

Private Sub CommandButton2_Click()
    If tbCodCat.Text = "" Then
        MsgBox mml_FRASE0504, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodFase.Text = sSeleccionar("SELECT DISTINCT fase FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " ORDER BY 1")
    DescFase
End Sub
Sub DescFase()
    Select Case tbCodFase.Text
        Case 1:
            tbDescFase.Text = mml_FRASE0329
        Case 2:
            tbDescFase.Text = "SEMI-FINAL"
        Case 4:
            tbDescFase.Text = "CUARTOS DE FINAL"
        Case 8:
            tbDescFase.Text = "OCTAVOS DE FINAL"
        Case "":
            tbDescFase.Text = ""
        Case Else
            tbDescFase.Text = tbCodFase.Text & "OS DE FINAL"
    End Select

End Sub

Private Function CalcularNoFinal() As Boolean
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim iMayAbs As Integer ' Valor que marca la mayoría
Dim iPos As Integer, iSum As Integer ' Variables donde se almacena los valores acumulados sin que  pasen de la mayoría
Dim iPos1 As Integer, iSum1 As Integer ' Variables donde se almacena los valores acumulados si pasan de la mayoría
Dim iCBaile As Integer, iCJuez As Integer, iCDorsal As Integer ' Contadores para crear los arrays
Dim iPuesto As Integer ' Posición final por vaile
Dim iFaseSig As Integer
Dim iRepesca As Integer
Dim iDorsalesHanPasado As Integer
Dim iJuecesCateg As Integer
    
    CalcularNoFinal = True
    
    iFaseSig = Val(tbCodFase.Text) / 2

    If tbCodCat.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Function
    End If
    'Borramos las puntuaciones de los no presentados
    db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal IN (SELECT DISTINCT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text & " AND no_presente>0) AND cod_categoria =" & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text)
    'Establecemos a 0 las puntuaciones de los no presentados
    'db.Execute ("UPDATE puntuaciones set puesto = 0 WHERE num_dorsal IN (SELECT DISTINCT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text & " AND no_presente>0) AND cod_categoria =" & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text)

    'Primero borramos cualquier calculo anterior
    tbLog.Text = ""
    addLog (mml_FRASE0468)
    db.Execute ("DELETE FROM cal_baile WHERE cod_categoria = " & tbCodCat.Text & " AND repesca = " & chkRep.Value & " AND fase=" & tbCodFase.Text)
    ' Si no calculamos una repesca borramos todos los dorsales
    If chkRep.Value = 0 Then
        db.Execute ("DELETE FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca = 0 AND fase=" & iFaseSig)
    End If
    ' Siempre borramos los dorsales de la repesca de esta fase no de la siguiente
    'db.Execute ("DELETE FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca = 1 AND fase=" & tbCodFase.Text)
    
    ' Comprobamos el número de jueces de la categoría
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCat.Text, dbOpenSnapshot)
    iJuecesCateg = rs.Fields(0)
    addLog (mml_FRASE0469 & iJuecesCateg)
    rs.Close
    ' Comprobamos el número de jueces en las puntuaciones
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iJueces = 0
    If Not rs.EOF Then
        rs.MoveLast
        iJueces = rs.RecordCount
    End If
    iNumJueces = iJueces
    addLog (mml_FRASE0470 & iJueces)
    rs.Close
    If iJueces <> iJuecesCateg Then
        If MsgBox(mml_FRASE0471 & iJueces & mml_FRASE0472 & iJuecesCateg & mml_FRASE0477, vbYesNo Or vbInformation, mml_FRASE0084) = vbNo Then
            CalcularNoFinal = False
            Exit Function
        End If
    End If
    ' Comprobamos el número de bailes
    If C_CALCULOS_PARCIALES Or BailesParciales(Val(tbCodCat.Text)) Then
        ' Comprobamos el número de bailes que hay en las puntuaciones
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ")", dbOpenSnapshot)
    Else
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCat.Text & " AND fase = " & BAILES_NO_FINAL, dbOpenSnapshot)
    End If
    iBailes = rs.Fields(0)
    iNumBailes = iBailes
    addLog (mml_FRASE0473 & iBailes)
    rs.Close
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    addLog (mml_FRASE0474 & iParejas)
    If rs.Fields(0) Mod iBailes * iJueces > 0 Then
        '"El número de puntuaciones no es correcta. ¿Continuar?"
        If MsgBox(mml_FRASE0977, vbYesNo Or vbCritical, mml_FRASE0096) = vbNo Then
            CalcularNoFinal = False
            Exit Function
        End If
    End If
    rs.Close
    ' Comprobamos si están todos los dorsales
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    If rs.Fields(0) <> iParejas Then
        If Not G_GEN_AUTO_RESULTADOS_PPC Then
            If MsgBox(mml_FRASE0476 & iParejas & mml_FRASE0472 & rs.Fields(0) & mml_FRASE0477, vbYesNo Or vbCritical, mml_FRASE0096) = vbNo Then
                CalcularNoFinal = False
                Exit Function
            End If
        End If
    End If
    rs.Close
    'Si es una repesca comprobamos los dorsales que ya han pasado
    If chkRep.Value = 1 Then
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=0 AND fase = " & iFaseSig, dbOpenSnapshot)
        iDorsalesHanPasado = rs.Fields(0)
        rs.Close
    Else
        iDorsalesHanPasado = 0
    End If
    
    
'Inicializamos el espacio para las tablas *********************************************
POS_MAY_ABS = iParejas
POS_TOTAL_BAILE = iParejas + 1
' Tablas para calcular posiciones en bailes individuales
' La primera columna se reserva para los dorsales
ReDim aPosBaile(iParejas, iJueces + 1, iBailes + 1) As Integer ' (Parejas, Jueces, Bailes)
ReDim aCalBaile(iParejas, iParejas + 2, iBailes + 1) As TCal ' (Parejas, Puestos, Bailes) ' Puestos + posición de mayoría abs de bailes + pos num_dorsales
' Public aPosTotalBaile(PAREJAS_FINAL, PAREJAS_FINAL, BAILES)   As Integer ' (Parejas, Puestos, Bailes)
' Tablas para calcular posiciones en el conjunto de bailes
ReDim aSumaTotalConjunto(iParejas, 1) As Double ' (Suma total , número de dorsal)
ReDim aCalConjunto(iParejas, iParejas) As TCal ' (Parejas, puestos) en la posición 0 el número de dorsal
ReDim aOrdenFinal(iParejas, 3)  ' (Parejas, Dorsales + Posición + cod_pareja)

ReDim aBailes(1 To iBailes) As TCodDesc
ReDim aDorsales(iParejas) As Integer
ReDim aJueces(iJueces) As String
'**************************************************************************************
Dim iNumDesc As Integer, iNumDescAnt As Integer
    
    If C_CALCULOS_PARCIALES Or BailesParciales(Val(tbCodCat.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCat.Text & " AND bc.fase =" & BAILES_NO_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & ") ORDER BY posicion"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        ' Primero comenzamos calculando las posiciones
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCat.Text & " AND fase = " & BAILES_NO_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBaile = 1
    While Not rsBailes.EOF
        addLog (mml_FRASE0478 & rsBailes.Fields(1) & "   ----------------------")
        aBailes(iCBaile).codigo = rsBailes.Fields(0)
        aBailes(iCBaile).DESCRIPCION = rsBailes.Fields(1)
        'Para cada dorsal
        iCDorsal = 0
        ' Recuperamos los dorsales a clasificar
        Debug.Print "SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1"
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            ' Por si hay descalificaciones cambiamos todas las puntuaciones de -1 a 1
            db.Execute ("UPDATE puntuaciones SET puesto = 1 WHERE puesto = -1 AND cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
            ' Comprobamos las descalificaciones de jueces normales------------------------------------
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones d, juez_categ j WHERE j.pasos = 0 AND d.id_juez = j.id_juez AND j.cod_categoria = d.cod_categoria AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & chkRep.Value & " AND d.fase = " & tbCodFase.Text & " AND d.cod_baile=" & rsBailes!cod_baile & " AND d.num_dorsal=" & rsDorsales!num_dorsal, dbOpenSnapshot)
            iNumDesc = rs.Fields(0)
            rs.Close
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones d, juez_categ j WHERE j.pasos = 0 AND d.id_juez = j.id_juez AND j.cod_categoria = d.cod_categoria AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & chkRep.Value & " AND d.fase < " & tbCodFase.Text & " AND d.cod_baile=" & rsBailes!cod_baile & " AND d.num_dorsal=" & rsDorsales!num_dorsal, dbOpenSnapshot)
            iNumDescAnt = rs.Fields(0)
            rs.Close
            ' Si ha sido descalificada antes
            If iNumDescAnt > 0 And iNumDesc > 1 Then
                If MsgBox(mml_FRASE0300 & rsDorsales!num_dorsal & mml_FRASE0480 & iNumDescAnt & mml_FRASE0506 & iNumDesc & mml_FRASE0507, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                    db.Execute ("UPDATE puntuaciones SET puesto = -1 WHERE puesto > 0 AND cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
                    If Not G_GEN_AUTO_RESULTADOS_PPC Then MsgBox mml_FRASE0508, vbOKOnly Or vbInformation, mml_FRASE0086
                End If
            End If
            ' Si estamos en la semifinal y es descalificado por una mayoría de jueces
            If tbCodFase.Text = 2 And iNumDesc > Int(iJueces / 2) Then
                If MsgBox(mml_FRASE0300 & rsDorsales!num_dorsal & mml_FRASE0480 & iNumDesc & mml_FRASE0509, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                    db.Execute ("UPDATE puntuaciones SET puesto = -1 WHERE  puesto > 0 AND cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
                    If Not G_GEN_AUTO_RESULTADOS_PPC Then MsgBox mml_FRASE0508, vbOKOnly Or vbInformation, mml_FRASE0086
                End If
            End If
            ' Comprobamos las descalificaciones del juez de pasos y figuras --------------------------
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones d, juez_categ j WHERE j.pasos = 1 AND d.id_juez = j.id_juez AND j.cod_categoria = d.cod_categoria AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & chkRep.Value & " AND d.fase = " & tbCodFase.Text & " AND d.cod_baile=" & rsBailes!cod_baile & " AND d.num_dorsal=" & rsDorsales!num_dorsal, dbOpenSnapshot)
            iNumDesc = rs.Fields(0)
            rs.Close
            ' Si ha sido descalificada por el juez de pasos y figuras
            If iNumDesc > 0 Then
                If MsgBox(mml_FRASE0300 & rsDorsales!num_dorsal & mml_FRASE0480 & iNumDesc & mml_FRASE0483, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                    db.Execute ("UPDATE puntuaciones SET puesto = -1 WHERE  puesto > 0 AND cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
                    If Not G_GEN_AUTO_RESULTADOS_PPC Then MsgBox mml_FRASE0508, vbOKOnly Or vbInformation, mml_FRASE0086
                End If
            End If
            '-----------------------------------------------------------------------------------------
            'Grabamos en la primera columna el número de dorsal
            aCalBaile(iCDorsal, POS_NUM_DORSAL, iCBaile).iNPos = rsDorsales!num_dorsal
            'Primero recuperamos todos los puestos (puntuaciones de cada juez) del dorsal en este baile
            Set rsPuestos = db.OpenRecordset( _
            "SELECT puesto,cod_juez FROM puntuaciones WHERE fase = " & tbCodFase.Text & _
            " AND repesca=" & chkRep.Value & _
            " AND cod_categoria = " & tbCodCat.Text & _
            " AND cod_baile = " & rsBailes!cod_baile & _
            " AND num_dorsal = " & rsDorsales!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            While Not rsPuestos.EOF
                If rsPuestos!Puesto > 0 Then
                    Inc aCalBaile(iCDorsal, POS_TOTAL_CONJUNTO, iCBaile).iNPos
                End If
                rsPuestos.MoveNext
            Wend
            addLog (mml_FRASE0225 & rsDorsales!num_dorsal & mml_FRASE0510 & aCalBaile(iCDorsal, POS_TOTAL_CONJUNTO, iCBaile).iNPos)
            db.Execute ("INSERT INTO cal_baile VALUES (" & rsDorsales!num_dorsal & "," & tbCodCat.Text & "," & rsBailes!cod_baile & ",0," & aCalBaile(iCDorsal, POS_TOTAL_CONJUNTO, iCBaile).iNPos & ",0," & tbCodFase.Text & "," & chkRep.Value & ")")
            iCDorsal = iCDorsal + 1
            rsDorsales.MoveNext
        Wend
        iCBaile = iCBaile + 1
        rsBailes.MoveNext
    Wend
    
    'Calculamos el conjunto de marcas en todos los bailes
Dim iCPuesto As Integer
Dim iCBailes As Integer
Dim iCParejas As Integer
Dim iCParejasSigFase As Integer
Dim iMaxParejas As Integer ' Número máx de parejas que pueden pasar
Dim iMinParejas As Integer
Dim iMaxParejasSemiyFinal As Integer
Dim iParejasEmpatadas As Integer
Dim bActivarRepesca As Boolean

    bActivarRepesca = False

    If iFaseSig = 1 Then
        iMinParejas = C_MIN_PAREJAS_SELEC_FINAL
    Else
        iMinParejas = Int(iParejas / 2 + 0.5)
    End If
    iMaxParejas = iFaseSig * 6
    iCParejasSigFase = iDorsalesHanPasado
    
    If iMaxParejas < iMinParejas Then
        iMinParejas = iMaxParejas
    End If
    iFaseSig = Val(tbCodFase.Text) / 2
    iParejasEmpatadas = VarCfg("parejas_por_empate")
    Select Case iFaseSig
        Case 1: ' Final
            iMaxParejasSemiyFinal = C_MAX_PAREJAS_FINAL
        Case 2: ' Semifinal
            iMaxParejasSemiyFinal = C_MAX_PAREJAS_SEMIFINAL
    End Select
    
    If Not (iMaxParejas = tbMaxParejas.Text And _
            iMinParejas = tbMinParejas.Text And _
            iParejasEmpatadas = tbEmpates.Text And _
            iMaxParejasSemiyFinal = tbMaxSemiOFinal.Text) Then
        'Los valores de parejas que pasan son distintos a los calculados
        If MsgBox(mml_FRASE0511 & iMaxParejas & mml_FRASE0512 & iMinParejas & mml_FRASE0513 & iParejasEmpatadas & mml_FRASE0514 & iMaxParejasSemiyFinal & mml_FRASE0515, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            'Dejamos valores menores que los normales para la fase
            bActivarRepesca = True
            iMaxParejas = tbMaxParejas.Text
            iMinParejas = tbMinParejas.Text
            iParejasEmpatadas = tbEmpates.Text
            iMaxParejasSemiyFinal = tbMaxSemiOFinal.Text
        Else
            tbMaxParejas.Text = iMaxParejas
            tbMinParejas.Text = iMinParejas
            tbEmpates.Text = iParejasEmpatadas
            tbMaxSemiOFinal.Text = iMaxParejasSemiyFinal
        End If
    End If
            
    
    'Se muestran las marcas totales de todos los dorsales
    If G_SELEC_DORSALES_SIG_FASE Then
        db.Execute "DELETE FROM SelecDorsales"
    End If
    addLog (mml_FRASE0516)
    For iCParejas = 0 To iParejas - 1
        For iCBailes = 1 To iBailes
            aCalBaile(iCParejas, POS_NUM_DORSAL, 0).iNPos = aCalBaile(iCParejas, POS_NUM_DORSAL, iCBailes).iNPos
            aCalBaile(iCParejas, POS_TOTAL_CONJUNTO, 0).iNPos = _
                   aCalBaile(iCParejas, POS_TOTAL_CONJUNTO, 0).iNPos + _
                   aCalBaile(iCParejas, POS_TOTAL_CONJUNTO, iCBailes).iNPos
            If G_SELEC_DORSALES_SIG_FASE Then
                db.Execute "INSERT INTO SelecDorsales VALUES (" & aCalBaile(iCParejas, POS_NUM_DORSAL, 0).iNPos & "," & aCalBaile(iCParejas, POS_TOTAL_CONJUNTO, 0).iNPos & ")"
            End If
        Next iCBailes
        addLog (mml_FRASE0225 & aCalBaile(iCParejas, POS_NUM_DORSAL, 0).iNPos & mml_FRASE0517 & aCalBaile(iCParejas, POS_TOTAL_CONJUNTO, 0).iNPos)
    Next iCParejas

    
Dim iValorMax As Integer
Dim iCValorMax As Integer
Dim iValor As Integer
Dim iCodPareja As Long

    For iCParejas = 0 To iParejas - 1
        aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = 1
        aOrdenFinal(iCParejas, POS_NUM_DORSAL) = aCalBaile(iCParejas, POS_NUM_DORSAL, 0).iNPos
    Next iCParejas
    ' Ahora ordenamos todos los puestos
    For iCPuesto = 1 To iParejas
        iValorMax = 0
        iCValorMax = 0
        For iCParejas = 0 To iParejas - 1
            ' Solo estudiamos el puesto actual
            If iCPuesto = aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) Then
                iValor = aCalBaile(iCParejas, POS_TOTAL_CONJUNTO, 0).iNPos
                If iValorMax < iValor Then
                    iValorMax = iValor
                    ' Las parejas no oficiales no se cuentan al limitar el número de las que pasan
                    If aOrdenFinal(iCParejas, POS_NUM_DORSAL) < iMinDorsalOficial(tbCodComp.Text) Then
                        iCValorMax = 0
                    Else
                        iCValorMax = 1
                    End If
                ElseIf iValorMax = iValor Then
                    If aOrdenFinal(iCParejas, POS_NUM_DORSAL) >= iMinDorsalOficial(tbCodComp.Text) Then
                        Inc iCValorMax
                    End If
                End If
            End If
        Next
        
        ' Comprobamos si ya tenemos parejas suficientes
        ' y con el puesto actual nos pasamos
        ' Grabamos la información de paso a la siguiente fase
        ' iCValorMax -> Parejas empatadas actuales
        If (iCParejasSigFase + iCValorMax > iMaxParejas And iRepesca = 0) Then
            If (iCParejasSigFase >= iMinParejas) Then
                addLog (mml_FRASE0518 & iFaseSig & " = " & iCParejasSigFase & mml_FRASE0519 & iMaxParejas & mml_FRASE0520 & iMinParejas)
                If bActivarRepesca Or C_PREGUNTAR_REPESCA_SIEMPRE Then
                    If MsgBox(mml_FRASE0521, vbDefaultButton2 Or vbYesNo Or vbQuestion, mml_FRASE0099) = vbNo Then
                        Exit Function
                    Else
                        iRepesca = 1
                    End If
                Else
                    Exit Function
                End If
            End If
            If (iFaseSig > 2 And iCParejasSigFase + iCValorMax > iMaxParejas + iParejasEmpatadas) Or _
               (iFaseSig <= 2 And iCParejasSigFase + iCValorMax > iMaxParejasSemiyFinal) Then
                addLog (mml_FRASE0518 & iFaseSig & " = " & iCParejasSigFase & mml_FRASE0519 & iMaxParejas & mml_FRASE0520 & iMinParejas)
                If (iFaseSig = 1 And iCParejasSigFase >= MIN_PAREJAS_FINAL) Then
                    addLog (mml_FRASE0522 & MIN_PAREJAS_FINAL)
                    If bActivarRepesca Then
                        If MsgBox(mml_FRASE0528, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                            Exit Function
                        Else
                            iRepesca = 1
                        End If
                    Else
                        Exit Function
                    End If
                End If
                addLog (mml_FRASE0523 & mml_FRASE0519 & iMaxParejas & mml_FRASE0520 & iMinParejas)
                'El numero de parejas no cumple los límites
                MsgBox mml_FRASE0524 & iCParejasSigFase + iCValorMax & " = " & iCParejasSigFase & " + " & iCValorMax & mml_FRASE0525 & mml_FRASE0519 & iMaxParejas & mml_FRASE0520 & iMinParejas, vbOKOnly Or vbCritical, mml_FRASE0084
                'Añadimos las parejas empatadas
                If MsgBox(mml_FRASE0526 & iCValorMax & mml_FRASE0527, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                    If C_PREGUNTAR_REPESCA_SIEMPRE Or C_PREGUNTAR_REPESCA Then
                        'Preguntar por la repesca
                        If MsgBox(mml_FRASE0528, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                            Exit Function
                        Else
                            iRepesca = 1
                        End If
                    Else
                        Exit Function
                    End If
                End If
            End If
        End If
        'Asignamos el puesto
        For iCParejas = 0 To iParejas - 1
            If iCPuesto = aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) Then
                iValor = aCalBaile(iCParejas, POS_TOTAL_CONJUNTO, 0).iNPos
                'aOrdenFinal(iCParejas, POS_NUM_DORSAL) = aCalBaile(iCParejas, POS_NUM_DORSAL, 0).iNPos
                If iValorMax = iValor Then
                    aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = iCPuesto
                    addLog (mml_FRASE0225 & aOrdenFinal(iCParejas, POS_NUM_DORSAL) & mml_FRASE0529 & iCPuesto)
                    ' Almacenamos la pareja en la siguiente fase
                    Set rs = db.OpenRecordset("SELECT DISTINCT cod_pareja FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND num_dorsal = " & aOrdenFinal(iCParejas, POS_NUM_DORSAL), dbOpenSnapshot)
                        If Not rs.EOF Then
                            iCodPareja = rs!cod_pareja
                        Else
                            iCodPareja = 1
                        End If
                    rs.Close
                    ' Siempre introducimos los dorsales en el grupo principal
                    If iRepesca = 1 Then
                        ' Los dorsales repescados no pasan de fase
                        Debug.Print "INSERT INTO dorsales VALUES (" & MaxCod("dorsales") & "," & aOrdenFinal(iCParejas, POS_NUM_DORSAL) & "," & tbCodCat.Text & "," & tbCodFase.Text & "," & iCodPareja & ",0,1)"
                        db.Execute ("INSERT INTO dorsales VALUES (" & MaxCod("dorsales") & "," & aOrdenFinal(iCParejas, POS_NUM_DORSAL) & "," & tbCodCat.Text & "," & tbCodFase.Text & "," & iCodPareja & ",0,1)")
                    Else
                        Debug.Print "INSERT INTO dorsales VALUES (" & MaxCod("dorsales") & "," & aOrdenFinal(iCParejas, POS_NUM_DORSAL) & "," & tbCodCat.Text & "," & iFaseSig & "," & iCodPareja & ",0,0)"
                        db.Execute ("INSERT INTO dorsales VALUES (" & MaxCod("dorsales") & "," & aOrdenFinal(iCParejas, POS_NUM_DORSAL) & "," & tbCodCat.Text & "," & iFaseSig & "," & iCodPareja & ",0,0)")
                    End If
                Else
                    aOrdenFinal(iCParejas, POS_TOTAL_CONJUNTO) = iCPuesto + iCValorMax
                End If
            End If
        Next
        iCParejasSigFase = iCParejasSigFase + iCValorMax
    Next
End Function


Private Sub Form_Activate()
    On Local Error Resume Next
    cmdCalcular.SetFocus

End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))
    
    If (VarCfg("publicacion") <> "publicacion_activa" And VarCfg("publicacion") <> "publicacion_auto") Then
        cmdPublicar.Visible = False
    End If
    cmdAutomatico.Enabled = False
End Sub


Sub CalcularNumeroParejas()
Dim iMaxParejas As Integer ' Número máx de parejas que pueden pasar
Dim iMinParejas As Integer
Dim iMaxParejasSemiyFinal As Integer
Dim iParejasEmpatadas As Integer
Dim iFaseSig As Integer
Dim iParejas As Integer
Dim iBailes As Integer
Dim iJueces As Integer
Dim rs As Recordset
Dim iFase As Integer
    
    If Val(tbCodCat.Text) = 0 Or Val(tbCodComp.Text) = 0 Or Val(tbCodFase.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
    iFaseSig = Val(tbCodFase.Text) / 2

    If iFaseSig = 0 Then
        frmSelecParejas.Visible = False
    Else
        frmSelecParejas.Visible = True
    End If
    
    iFase = IIf(Val(tbCodFase.Text) = 1, BAILES_FINAL, BAILES_NO_FINAL)
    ' Comprobamos el número de jueces
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCat.Text, dbOpenSnapshot)
    iJueces = rs.Fields(0)
    iNumJueces = iJueces
    addLog (mml_FRASE0470 & iJueces)
    rs.Close
    ' Comprobamos el número de bailes
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCat.Text & " AND fase = " & iFase, dbOpenSnapshot)
    iBailes = rs.Fields(0)
    iNumBailes = iBailes
    addLog (mml_FRASE0473 & iBailes)
    rs.Close
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    addLog (mml_FRASE0474 & iParejas)
    If BailesParciales(Val(tbCodCat.Text)) Then
        If rs.Fields(0) Mod iJueces > 0 Then
            MsgBox mml_FRASE0505, vbOKOnly Or vbCritical, mml_FRASE0096
        End If
    Else
        If rs.Fields(0) Mod iBailes * iJueces > 0 Then
            MsgBox mml_FRASE0505, vbOKOnly Or vbCritical, mml_FRASE0096
        End If
    End If
    rs.Close
    
    If iFaseSig = 1 Then
    iMinParejas = C_MIN_PAREJAS_SELEC_FINAL
    Else
        iMinParejas = Int(iParejas / 2 + 0.5)
    End If
    iMaxParejas = iFaseSig * 6
    
    If iMaxParejas < iMinParejas Then
        iMinParejas = iMaxParejas
    End If
    iFaseSig = Val(tbCodFase.Text) / 2
    iParejasEmpatadas = VarCfg("parejas_por_empate")
    Select Case iFaseSig
        Case 1: ' Final
            iMaxParejasSemiyFinal = C_MAX_PAREJAS_FINAL
        Case 2: ' Semifinal
            iMaxParejasSemiyFinal = C_MAX_PAREJAS_SEMIFINAL
    End Select
    tbMaxParejas.Text = iMaxParejas
    tbMinParejas.Text = iMinParejas
    tbEmpates.Text = iParejasEmpatadas
    tbMaxSemiOFinal.Text = iMaxParejasSemiyFinal
End Sub



Private Sub lblAutoPPC_dblClick()
    If lblAutoPPC.BackColor = vbRed Then
        lblAutoPPC.BackColor = vbGreen
    Else
        lblAutoPPC.BackColor = vbRed
    End If
End Sub

Private Sub tbCodCat_GotFocus()
    tbCodCat.SelStart = 0
    tbCodCat.SelLength = Len(tbCodCat.Text)
End Sub

Private Sub tbCodCat_KeyPress(KeyAscii As Integer)
    SoloNumero KeyAscii
End Sub

Private Sub tbCodCat_LostFocus()
    ComprobarCategyFase tbCodCat, tbDescCat, tbCodFase, tbDescFase
    
End Sub

Private Sub tbCodFase_Change()
    If IsNumeric(tbCodFase.Text) Then
        Call CalcularNumeroParejas
    End If
End Sub

Public Sub MostrarCalcular()
Dim rs As Recordset
On Local Error Resume Next
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    Set rs = db.OpenRecordset(" SELECT descripcion FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    tbDescComp.Text = rs!DESCRIPCION
    rs.Close
    Set rs = db.OpenRecordset(" SELECT descripcion FROM categorias WHERE codigo = " & tbCodCat.Text, dbOpenSnapshot)
    tbDescCat.Text = rs!DESCRIPCION
    rs.Close
    DescFase

    Me.Show vbModal
End Sub














































Private Sub CalcularFinal1()
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim iMayAbs As Integer ' Valor que marca la mayoría
Dim iPos As Integer, iSum As Integer ' Variables donde se almacena los valores acumulados sin que  pasen de la mayoría
Dim iPos1 As Integer, iSum1 As Integer ' Variables donde se almacena los valores acumulados si pasan de la mayoría
Dim iCBaile As Integer, iCJuez As Integer, iCDorsal As Integer ' Contadores para crear los arrays
Dim iPuesto As Integer ' Posición final por vaile

    If tbCodCat.Text = "" Or tbCodFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    'Borramos las puntuaciones de los no presentados
    db.Execute ("DELETE FROM puntuaciones WHERE num_dorsal IN (SELECT DISTINCT num_dorsal FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text & " AND no_presente>0) AND cod_categoria =" & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase=" & tbCodFase.Text)
    
    'Primero borramos cualquier calculo anterior
    tbLog.Text = ""
    addLog (mml_FRASE0468)
    db.Execute ("DELETE FROM cal_baile WHERE cod_categoria = " & tbCodCat.Text & " AND fase=" & tbCodFase.Text)
    db.Execute ("DELETE FROM cal_conjunto WHERE cod_categoria = " & tbCodCat.Text)
    
    ' Comprobamos el número de jueces
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCat.Text, dbOpenSnapshot)
    iJueces = rs.Fields(0)
    iNumJueces = iJueces
    addLog (mml_FRASE0470 & iJueces)
    rs.Close
    ' Comprobamos el número de bailes
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCat.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
    iBailes = rs.Fields(0)
    iNumBailes = iBailes
    addLog (mml_FRASE0473 & iBailes)
    rs.Close
    ' Comprobamos el número de parejas
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    iParejas = rs.Fields(0) / iBailes / iJueces
    addLog (mml_FRASE0474 & iParejas)
    If rs.Fields(0) Mod iBailes * iJueces > 0 Then
        If MsgBox(mml_FRASE0475, vbYesNo Or vbCritical, mml_FRASE0096) = vbNo Then
            Exit Sub
        End If
    End If
    rs.Close
    ' Comprobamos si están todos los dorsales
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & tbCodCat.Text & " AND repesca = " & chkRep.Value & " AND fase = " & tbCodFase.Text, dbOpenSnapshot)
    If rs.Fields(0) <> iParejas Then
        If MsgBox(mml_FRASE0476 & iParejas & mml_FRASE0472 & rs.Fields(0) & mml_FRASE0477, vbYesNo Or vbCritical, mml_FRASE0096) = vbNo Then
            Exit Sub
        End If
    End If
    rs.Close
    
    
    'Calculamos el valor de mayoría absoluta
    iMayAbs = Int(iJueces / 2)
    ' Primero comenzamos calculando las posiciones
    
'Inicializamos el espacio para las tablas *********************************************
POS_MAY_ABS = iParejas
POS_TOTAL_BAILE = iParejas + 1
' Tablas para calcular posiciones en bailes individuales
' La primera columna se reserva para los dorsales
ReDim aPosBaile(iParejas, iJueces + 1, iBailes) As Integer ' (Parejas, Jueces, Bailes)
ReDim aCalBaile(iParejas, iParejas + 2, iBailes) As TCal ' (Parejas, Puestos, Bailes) ' Puestos + posición de mayoría abs de bailes + pos num_dorsales
' Public aPosTotalBaile(PAREJAS_FINAL, PAREJAS_FINAL, BAILES)   As Integer ' (Parejas, Puestos, Bailes)
' Tablas para calcular posiciones en el conjunto de bailes
ReDim aSumaTotalConjunto(iParejas, 1) As Double ' (Suma total , número de dorsal)
ReDim aCalConjunto(iParejas, iParejas) As TCal ' (Parejas, puestos) en la posición 0 el número de dorsal
ReDim aOrdenFinal(iParejas, 2) ' (Parejas, Dorsales + Posición)

ReDim aBailes(iBailes) As TCodDesc
ReDim aDorsales(iParejas) As Integer
ReDim aJueces(iJueces) As String
'**************************************************************************************
Dim iNumDesc As Integer
    
    Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCat.Text & " AND fase = " & BAILES_FINAL & " ORDER BY posicion", dbOpenSnapshot)
    'Se calcula cada baile de modo idependiente
    iCBaile = 0
    While Not rsBailes.EOF
        addLog (mml_FRASE0478 & rsBailes.Fields(1) & "   ----------------------")
        aBailes(iCBaile).codigo = rsBailes.Fields(0)
        aBailes(iCBaile).DESCRIPCION = rsBailes.Fields(1)
        'Para cada dorsal
        iCDorsal = 0
        ' Recuperamos los dorsales a clasificar
        Set rsDorsales = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rsDorsales.EOF
            addLog (mml_FRASE0225 & rsDorsales!num_dorsal)
            
            addLog (mml_FRASE0479 & rsDorsales!num_dorsal)
            ' Por si hay descalificaciones cambiamos todas las puntuaciones > C_PUESTO_NEG
            db.Execute ("UPDATE puntuaciones SET puesto = puesto-" & C_PUESTO_NEG & " WHERE puesto >" & C_PUESTO_NEG & " AND cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
            ' Comprobamos las descalificaciones de jueces normales------------------------------------
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones d, juez_categ j WHERE j.pasos = 0 AND d.id_juez = j.id_juez AND j.cod_categoria = d.cod_categoria AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca = " & chkRep.Value & " AND d.fase = " & tbCodFase.Text & " AND d.cod_baile=" & rsBailes!cod_baile & " AND d.num_dorsal=" & rsDorsales!num_dorsal, dbOpenSnapshot)
            iNumDesc = rs.Fields(0)
            rs.Close
            ' Estamos en la final y es descalificado por más de un juez
            If iNumDesc >= G_MIN_JUECES_DESCALIFICACION Then
                If MsgBox(mml_FRASE0300 & rsDorsales!num_dorsal & mml_FRASE0480 & iNumDesc & mml_FRASE0481, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                    db.Execute ("UPDATE puntuaciones SET puesto = " & C_PUESTO_NEG & "+puesto WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
                    MsgBox mml_FRASE0482, vbOKOnly Or vbInformation, mml_FRASE0086
                End If
            End If
            ' Comprobamos las descalificaciones del juez de pasos y figuras --------------------------
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM descalificaciones d, juez_categ j WHERE j.pasos = 1 AND d.id_juez = j.id_juez AND j.cod_categoria = d.cod_categoria AND d.cod_categoria = " & tbCodCat.Text & " AND d.repesca=" & chkRep.Value & " AND d.fase = " & tbCodFase.Text & " AND d.cod_baile=" & rsBailes!cod_baile & " AND d.num_dorsal=" & rsDorsales!num_dorsal, dbOpenSnapshot)
            iNumDesc = rs.Fields(0)
            rs.Close
            ' Si ha sido descalificada por el juez de pasos y figuras
            If iNumDesc > 0 Then
                If MsgBox(mml_FRASE0300 & rsDorsales!num_dorsal & mml_FRASE0480 & iNumDesc & mml_FRASE0483, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
                    db.Execute ("UPDATE puntuaciones SET puesto = " & C_PUESTO_NEG & "+puesto WHERE cod_categoria = " & tbCodCat.Text & " AND repesca=" & chkRep.Value & " AND fase = " & tbCodFase.Text & " AND cod_baile=" & rsBailes!cod_baile & " AND num_dorsal=" & rsDorsales!num_dorsal)
                    MsgBox mml_FRASE0482, vbOKOnly Or vbInformation, mml_FRASE0086
                End If
            End If
            '-----------------------------------------------------------------------------------------
            
            'Grabamos en la primera columna el número de dorsal
            aCalBaile(iCDorsal, POS_NUM_DORSAL, iCBaile).iNPos = rsDorsales!num_dorsal
            'Primero recuperamos todos los puestos (puntuaciones de cada juez) del dorsal en este baile
            Set rsPuestos = db.OpenRecordset( _
            "SELECT puesto,cod_juez FROM puntuaciones " & _
            "WHERE fase = 1 " & _
            " AND repesca=" & chkRep.Value & _
            " AND cod_categoria = " & tbCodCat.Text & _
            " AND cod_baile = " & rsBailes!cod_baile & _
            " AND num_dorsal = " & rsDorsales!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            
            ' Almacenamos los puestos ordenados por la puntuación de cada juez
            Set rsPuestosJuez = db.OpenRecordset( _
            "SELECT puesto,cod_juez FROM puntuaciones " & _
            "WHERE fase = 1 " & _
            " AND repesca=" & chkRep.Value & _
            " AND cod_categoria = " & tbCodCat.Text & _
            " AND cod_baile = " & rsBailes!cod_baile & _
            " AND num_dorsal = " & rsDorsales!num_dorsal & " ORDER BY 2", dbOpenSnapshot)
            iCJuez = 1
            While Not rsPuestosJuez.EOF
                aPosBaile(iCDorsal, iCJuez, iCBaile) = rsPuestosJuez!Puesto
                'Control de puestos con valor 0
                If aPosBaile(iCDorsal, iCJuez, iCBaile) = 0 Then
                    addLog (mml_FRASE0484 & iCDorsal & mml_FRASE0485 & iCJuez & mml_FRASE0486 & iCBaile)
                End If
                iCJuez = iCJuez + 1
                rsPuestosJuez.MoveNext
            Wend
            'Grabamos en la primera columna el número de dorsal
            aPosBaile(iCDorsal, POS_NUM_DORSAL, iCBaile) = rsDorsales!num_dorsal
            rsPuestosJuez.Close
            
            
            'Calculamos las mejores posiciones y la suma de las mismas
            iPos = 0
            iSum = 0
            iPos1 = 0
            iSum1 = 0
            iCJuez = 1 ' La posición 0 queda reservada para el número de dorsal
            For j = 1 To iParejas - 1 ' Puestos
                Do While Not rsPuestos.EOF
                    If rsPuestos!Puesto <= j Then
                        iCJuez = iCJuez + 1
                        iPos = iPos + 1 ' Numero de puestos acumulados
                        iSum = iSum + rsPuestos!Puesto
                        rsPuestos.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                'Colocamos 0 a no ser que se supere la mayoría absoluta
                If iPos > iMayAbs Then 'iPos = Suma de puestos ant 0 = al estudiado
                    'Si en alguno pasamos de la mayoría podemos ordenarlo
                    iPos1 = iPos
                    iSum1 = iSum
                    'Comprobamos si es la primera vez que pasamos la mayoría abs
                    If aCalBaile(iCDorsal, POS_MAY_ABS, iCBaile).iNPos = 0 Then
                        aCalBaile(iCDorsal, POS_MAY_ABS, iCBaile).iNPos = j
                    End If
                End If
                aCalBaile(iCDorsal, j, iCBaile).iNPos = iPos1
                aCalBaile(iCDorsal, j, iCBaile).iSum = iSum1
                db.Execute ("INSERT INTO cal_baile VALUES(" & rsDorsales!num_dorsal & "," & tbCodCat.Text & "," & rsBailes!cod_baile & "," & j & "," & iPos1 & "," & iSum1 & " ,1," & chkRep.Value & ")")
            Next j
            While Not rsPuestos.EOF
                rsPuestos.MoveNext
            Wend
            If rsPuestos.RecordCount < iJueces Then
                MsgBox mml_FRASE0487 & rsDorsales!num_dorsal & mml_FRASE0488 & rsPuestos.RecordCount & mml_FRASE0489 & rsBailes!Nombre, vbOKOnly Or vbCritical, mml_FRASE0096
            End If
            
            'Si no hemos superado la mayoría en algún puesto, le corresponde el último
            If aCalBaile(iCDorsal, POS_MAY_ABS, iCBaile).iNPos = 0 Then
                aCalBaile(iCDorsal, POS_MAY_ABS, iCBaile).iNPos = iParejas
            End If
            rsPuestos.Close
            iCDorsal = iCDorsal + 1
            rsDorsales.MoveNext
        Wend ' Dorsales
        iCBaile = iCBaile + 1
        rsBailes.MoveNext
        rsDorsales.Close
    Wend 'Bailes
    rsBailes.Close
    
    ' Ahora tenemos todos los cuadros de todos los bailes y tenemos que calcular
    
    Call Ordenacion
End Sub


Function Calcular() As Boolean
    ImpresionHojas = False
    Me.Show vbModal
    Calcular = ImpresionHojas
End Function
