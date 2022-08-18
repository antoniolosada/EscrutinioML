VERSION 5.00
Begin VB.Form frmResultadosCombinados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1038"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmdCateg 
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
         Left            =   1920
         Picture         =   "frmResultadosCombinados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   870
         Width           =   450
      End
      Begin VB.TextBox tbDescCateg 
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
         Left            =   3210
         TabIndex        =   6
         Top             =   870
         Width           =   4935
      End
      Begin VB.TextBox tbCodCateg 
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
         Left            =   2370
         MaxLength       =   5
         TabIndex        =   0
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox tbCodCateg1 
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
         Left            =   1890
         TabIndex        =   5
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox tbDescCateg1 
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
         Left            =   2760
         TabIndex        =   4
         Top             =   270
         Width           =   5415
      End
      Begin VB.Label Label1 
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
         Left            =   270
         TabIndex        =   8
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label8 
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
         Left            =   270
         TabIndex        =   7
         Top             =   870
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCombinar 
      Caption         =   "mml_FRASE1038"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   1500
      Width           =   3285
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "mml_FRASE0029"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   1500
      Width           =   2265
   End
End
Attribute VB_Name = "frmResultadosCombinados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCodCategorias As String
Dim sCodCompeticion As String

Private Sub cmdCateg_Click()
    tbCodCateg.Text = sSeleccionar("SELECT * FROM categorias c WHERE cod_competicion =" & sCodCompeticion & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCateg.Text = sResultado(2)

End Sub

Private Sub cmdCombinar_Click()
Dim rs As Recordset
Dim sNuevaCat As String
Dim rsCat As Recordset
Dim rsDor As Recordset
Dim rsPunt As Recordset
Dim rsJueces As Recordset
Dim rsBailes As Recordset
Dim iMaxCat As Integer
Dim sSQLDorsales As String
Dim iNumJueces, iNumJueces1 As Integer
Dim iCodCateg As Integer
Dim iModalidad As Integer
Dim sCateg As String
Dim sSQL As String
Dim i As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    iCodCateg = tbCodCateg1.Text
    sSQL = ""
    
    'Recuperamos la información de los jueces
    'Comprobamos si los jueces no son los mismos
    Set rsJueces = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE cod_categoria = " & tbCodCateg1.Text & " AND NOT id_juez IN (SELECT id_juez FROM juez_categ jc2 WHERE jc2.cod_categoria = " & tbCodCateg.Text & ")", dbOpenSnapshot)
    iNumJueces = rsJueces.Fields(0)
    If iNumJueces > 0 Then
        MsgBox mml_FRASE1039, vbOKOnly Or vbCritical, G_MSG_ERROR
        Exit Sub
    End If
    rsJueces.Close
    
    'Si algún baile es común a las dos categorias no pueden combinarse
    'Recuperamos la información de todos los dorsales que bailan las dos categorias en la final
    sSQL = "SELECT COUNT(*) FROM bailes_categ bc WHERE bc.fase = 1 AND bc.cod_categoria = " & iCodCateg & " AND bc.cod_baile IN (SELECT cod_baile FROM bailes_categ bc2 WHERE bc2.fase = 1 AND bc2.cod_categoria = " & tbCodCateg.Text & ")"
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            MsgBox mml_FRASE1135, vbOKOnly Or vbCritical, G_MSG_ERROR
            Exit Sub
        End If
    rs.Close

    'Recuperamos la información de todos los dorsales que bailan las dos categorias en la final
    sSQLDorsales = "SELECT DISTINCT num_dorsal FROM dorsales d1 WHERE no_presente = 0 AND fase = 1 AND cod_categoria = " & iCodCateg & " AND num_dorsal = (SELECT num_dorsal FROM dorsales d2 WHERE d2.no_presente = 0 AND d2.cod_categoria = " & tbCodCateg.Text & " AND d2.num_dorsal = d1.num_dorsal AND fase = 1)"
    sSQL = sSQLDorsales
    Set rsDor = db.OpenRecordset(sSQLDorsales, dbOpenSnapshot)
    If Not rsDor.EOF Then
        'Insertamos la nueva categoria
        'sNuevaCat = mml_FRASE1120 + " " + Left$(tbDescCateg1.Text, 7) + "+" + Left$(tbDescCateg.Text, 7)
        sCateg = tbDescCateg.Text
        i = InStr(sCateg, " ")
        If i > 0 Then sCateg = Mid$(sCateg, i + 1)
        sNuevaCat = mml_FRASE1120 + " " + sCateg
        sSQL = "SELECT * FROM categorias WHERE codigo = " & tbCodCateg1.Text
        Set rsCat = db.OpenRecordset(sSQL, dbOpenSnapshot)
        iMaxCat = MaxCod("categorias")
        iModalidad = rsCat!cod_modalidad
        If G_CATEGORIA_COMBINACION > 0 Then
            If sDescModalidad(G_CATEGORIA_COMBINACION) <> "" Then
                iModalidad = G_CATEGORIA_COMBINACION
            End If
        End If
        
        sSQL = "INSERT INTO categorias VALUES(" & iMaxCat & ", '" & sNuevaCat & "','" & rsCat!id_categoria & "','" & rsCat!cod_grupoedad & "','" & rsCat!cod_competicion & "','" & iModalidad & "','" & rsCat!hora & "'," & rsCat!rec_parcial_bailes & "," & rsCat!mostrar_posicion & "," & rsCat!dorsales_tanda & "," & rsCat!combinar_dorsales & "," & ImpHojaUnica & " )"
        rsCat.Close
        db.Execute sSQL
        
        'Insertamos los dorsales
        rsDor.Close
        sSQL = "SELECT * FROM dorsales d1 WHERE fase = 1 AND cod_categoria = " & iCodCateg & " AND num_dorsal = (SELECT num_dorsal FROM dorsales d2 WHERE d2.cod_categoria = " & tbCodCateg.Text & " AND d2.num_dorsal = d1.num_dorsal AND fase = 1)"
        Set rsDor = db.OpenRecordset(sSQL, dbOpenSnapshot)
        While Not rsDor.EOF
            sSQL = "INSERT INTO dorsales VALUES (" & MaxCod("dorsales") & "," & rsDor!num_dorsal & "," & iMaxCat & "," & rsDor!fase & "," & rsDor!cod_pareja & "," & rsDor!no_presente & "," & rsDor!repesca & ")"
            db.Execute sSQL
            rsDor.MoveNext
        Wend
        rsDor.MoveFirst
        
        'Recuperamos la información de los jueces
        sSQL = "SELECT * FROM juez_categ WHERE cod_categoria = " & tbCodCateg1.Text
        Set rsJueces = db.OpenRecordset(sSQL, dbOpenSnapshot)
        While Not rsJueces.EOF
            sSQL = "INSERT INTO juez_categ VALUES(" & rsJueces!cod_juez & "," & iMaxCat & ",'" & rsJueces!id_juez & "'," & rsJueces!pasos & ")"
            db.Execute sSQL
            rsJueces.MoveNext
        Wend
        rsJueces.Close
        'Recuperamos la información de los bailes de la final
        Dim iPos As Integer
        iPos = 1
        sSQL = "SELECT * FROM bailes_categ WHERE (cod_categoria = " & tbCodCateg1.Text & " OR cod_categoria = " & tbCodCateg.Text & ") ORDER BY posicion"
        Set rsBailes = db.OpenRecordset(sSQL, dbOpenSnapshot)
        While Not rsBailes.EOF
            sSQL = "INSERT INTO bailes_categ VALUES(" & iMaxCat & "," & rsBailes!cod_baile & "," & rsBailes!fase & "," & iPos & ")"
            db.Execute sSQL
            rsBailes.MoveNext
            Inc iPos
        Wend
        rsBailes.Close
        'Recuperamos las puntuaciones de la final de todos los dorsales
        sSQL = "SELECT * FROM puntuaciones WHERE (cod_categoria = " & tbCodCateg1.Text & " OR cod_categoria = " & tbCodCateg.Text & ") AND fase = 1 AND num_dorsal IN (" & sSQLDorsales & ")"
        Set rsPunt = db.OpenRecordset(sSQL, dbOpenSnapshot)
        'Recuperamos solo las puntuaciones de los dorsales que bailan las dos categorias
        While Not rsPunt.EOF
            sSQL = "INSERT INTO puntuaciones VALUES (" & rsPunt!num_dorsal & "," & iMaxCat & "," & rsPunt!cod_baile & ",'" & rsPunt!cod_juez & "'," & rsPunt!puesto & "," & rsPunt!fase & "," & rsPunt!repesca & ")"
            db.Execute sSQL
            rsPunt.MoveNext
        Wend
        rsPunt.Close
    Else
        MsgBox mml_FRASE0589, vbOKOnly Or vbCritical, G_MSG_ERROR
    End If
    rsDor.Close
    
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE
    Exit Sub
error:
    ProcesarError "cmdCombinar_Click: " & sSQL
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub

Sub CombinarResultados(iCodCateg As Integer, lCodComp As Long)
    TraducirCadenas Me
    
    
    sCodCompeticion = Str$(lCodComp)
    tbCodCateg1.Text = iCodCateg
    tbDescCateg1.Text = sDescCategoria(iCodCateg)
    Me.Show vbModal
End Sub

Private Sub tbCodCateg_GotFocus()
    tbCodCateg.SelStart = 0
    tbCodCateg.SelLength = Len(tbCodCateg.Text)
End Sub

Private Sub tbCodCateg_KeyPress(Keyascii As Integer)
    SoloNumero Keyascii
End Sub

Private Sub tbCodCateg_LostFocus()
Dim sCateg As String
    
    If Val(tbCodCateg.Text) > 0 Then
        sCateg = sDescCategoria(tbCodCateg.Text)
        If sCateg = "" Then
            tbCodCateg.Text = ""
            tbDescCateg.Text = ""
        Else
            tbDescCateg.Text = sCateg
        End If
    Else
        tbCodCateg.Text = ""
        tbDescCateg.Text = ""
    End If
End Sub

