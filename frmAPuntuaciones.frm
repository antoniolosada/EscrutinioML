VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMAPuntuaciones 
   Caption         =   "mml_FRASE0057"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "mml_FRASE0296"
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
      Left            =   0
      TabIndex        =   23
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0422"
      Height          =   4095
      Left            =   60
      TabIndex        =   21
      Top             =   3660
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgPuntuaciones 
         Bindings        =   "frmAPuntuaciones.frx":0000
         Height          =   3735
         Left            =   180
         TabIndex        =   22
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
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
      Height          =   405
      Left            =   2025
      Picture         =   "frmAPuntuaciones.frx":001E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   585
      Width           =   495
   End
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
      Height          =   405
      Left            =   2025
      Picture         =   "frmAPuntuaciones.frx":0488
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   90
      Width           =   495
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
      Height          =   480
      Left            =   6795
      TabIndex        =   18
      Top             =   1080
      Width           =   1470
   End
   Begin VB.CheckBox chkRep 
      Caption         =   "mml_FRASE0418"
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
      Left            =   2880
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame frmPunt 
      Caption         =   "mml_FRASE0419"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   8295
      Begin MSComDlg.CommonDialog CD 
         Left            =   120
         Top             =   1500
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "mml_FRASE0065"
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
         Left            =   3240
         TabIndex        =   29
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox cbDorsal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   7335
      End
      Begin VB.ComboBox cbParametros 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3960
         TabIndex        =   26
         Top             =   600
         Width           =   1755
      End
      Begin VB.ComboBox cbBaile 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   24
         Top             =   600
         Width           =   2715
      End
      Begin VB.CommandButton cmdBorrarTodas 
         Caption         =   "mml_FRASE0383"
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
         Left            =   420
         TabIndex        =   17
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox tbValor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         MaxLength       =   3
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton mod 
         Caption         =   "mml_FRASE0420"
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
         Left            =   6900
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cbJuez 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton add 
         Caption         =   "mml_FRASE0297"
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
         Left            =   5520
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "mml_FRASE1265"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   27
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "mml_FRASE0185"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblBaile 
         Alignment       =   2  'Center
         Caption         =   "mml_FRASE1269"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   6060
         TabIndex        =   16
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0421"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.ComboBox cbFase 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmAPuntuaciones.frx":08F2
      Left            =   4440
      List            =   "frmAPuntuaciones.frx":090E
      TabIndex        =   8
      Text            =   "1, FINAL"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "mml_FRASE0295"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
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
      Left            =   3360
      TabIndex        =   4
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox tbCodCateg 
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
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   855
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
      Left            =   2520
      TabIndex        =   1
      Top             =   120
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
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin MSAdodcLib.Adodc adoPuntuaciones 
      Height          =   495
      Left            =   1080
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Escrutinio"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "mml_FRASE0033"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM MA_puntuaciones"
      Caption         =   "mml_FRASE0422"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
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
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   735
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
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label5 
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMAPuntuaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isPosY1 As Integer
Dim isPosY2 As Integer

Const C_ALTURA_FILA_TABLA = 400

Private Sub cmdBorrar_Click()
End Sub

Private Sub cbBaile_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub cbDorsal_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub cbFase_Click()
    Call ActualizarTodo

End Sub

Private Sub cbFase_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbJuez_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub ActualizarTodo()
Dim rs As Recordset
Dim sFase As String
Dim i As Integer
Dim sSQL As String

    frmPunt.Refresh
    If tbCodCateg.Text = "" Or tbCodComp.Text = "" Or Trim$(Mid$(cbFase.Text, 1, 2)) = "" Then
        Exit Sub
    End If
    ' Recuperamos los jueces y los bailes
    cbJuez.Clear
    Set rs = db.OpenRecordset("SELECT id_juez FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text & " ORDER BY 1", dbOpenSnapshot)
        While Not rs.EOF
            cbJuez.AddItem rs!id_juez
            rs.MoveNext
        Wend
    rs.Close
    sSQL = "SELECT num_dorsal, p.nombre_hombre, p.nombre_mujer FROM dorsales d, parejas p WHERE d.cod_pareja = p.codigo AND cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text) & " ORDER BY num_dorsal"
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
    cbDorsal.Clear
    While Not rs.EOF
        cbDorsal.AddItem rs.Fields("num_dorsal") & " - " & rs.Fields("nombre_hombre") & ", " & rs.Fields("nombre_mujer")
        rs.MoveNext
    Wend
    rs.Close
    ' Final = 1 o No final = 2
    sFase = IIf(Val(Mid$(cbFase.Text, 1, 2)) = 1, "1", "2")
    Set rs = db.OpenRecordset("SELECT cod_baile, nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE cod_categoria = " & tbCodCateg.Text & " AND fase =" & sFase & " AND bc.cod_baile = b.codigo ORDER BY posicion", dbOpenSnapshot)
        i = 0
        cbBaile.Clear
        While Not rs.EOF And i <= 9
            cbBaile.AddItem rs!cod_baile & " - " & rs!Nombre
            rs.MoveNext
        Wend
    rs.Close
    'Recuperamos la información de parámetros
    Set rs = db.OpenRecordset("SELECT codigo, descorta FROM MA_Parametros", dbOpenSnapshot)
        cbParametros.Clear
        While Not rs.EOF And i <= 9
            cbParametros.AddItem rs!codigo & " - " & rs!descorta
            rs.MoveNext
        Wend
    rs.Close
    Call cmdActualizar_Click
    cbJuez.Refresh
End Sub


Private Sub cbParametros_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

End Sub

Private Sub cmdActualizar_Click()
Dim sSQL As String
Dim rs As Recordset
    If tbCodCateg.Text <> "" Then
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        Sleep 100
        DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
        adoPuntuaciones.ConnectionString = "DSN=Escrutinio"
        
        sSQL = "SELECT * FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text) & " ORDER BY num_dorsal,cod_juez, cod_baile, parametro"
        Debug.Print sSQL
        adoPuntuaciones.RecordSource = sSQL
        adoPuntuaciones.Refresh
    End If

End Sub



Private Sub cmdBorrarTodas_Click()
    If tbCodCateg.Text = "" Or tbCodComp.Text = "" Or Trim$(Mid$(cbFase.Text, 1, 2)) = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If Val(cbDorsal.Text) > 0 And Val(cbBaile.Text) > 0 Then
        'Tiene seleccionado un dorsal, en este caso es obligatorio que seleccione un baile
        If MsgBox(mml_FRASE1274, vbYesNo Or vbQuestion, "") = vbYes Then
            Debug.Print "DELETE FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = '" & Trim$(Mid$(cbFase.Text, 1, 2)) & "'"
            db.Execute "DELETE FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text) & " AND num_dorsal = " & Val(cbDorsal.Text) & " AND cod_baile = " & Val(cbBaile.Text)
        End If
        
    ElseIf MsgBox(mml_FRASE0423, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
            If InputBox(mml_FRASE1273, "") = "ELIMINAR" Then
                Debug.Print "DELETE FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = '" & Trim$(Mid$(cbFase.Text, 1, 2)) & "'"
                db.Execute "DELETE FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text)
            End If
    End If
    Sleep 1000
    DoEvents
    adoPuntuaciones.Refresh
End Sub

Private Sub cmdCateg_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCateg.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCateg.Text = sResultado(2)
    Call ActualizarTodo

End Sub





Public Sub cmdImprimir_Click()
    If Val(cbDorsal.Text) = 0 Then
        MsgBox mml_FRASE1271, vbOKOnly Or vbCritical, "ERROR"
    Else
        CD.ShowPrinter
        ImprimirPuntuaciones Val(cbDorsal.Text)
    End If
End Sub

Private Sub cmdQuitar_Click()
Dim sNumDorsal As String
Dim sCodCategoria As String
Dim sCodCategorias As String
Dim sCodBaile As String
Dim sCodJuez As String
Dim sFase As String
Dim sRep As String
Dim sParametro As String

    If tbCodCateg.Text = "" Or cbFase.Text = "" Or Val(cbDorsal.Text) = 0 Then
        MsgBox mml_FRASE0424, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If cbJuez.Text = "" And cbBaile.Text = "" Then
        MsgBox mml_FRASE1291, vbOKOnly Or vbCritical, "AVISO"
        If InputBox(mml_FRASE1273, "") = "ELIMINAR" Then
            sNumDorsal = Val(cbDorsal.Text)
            db.Execute ("DELETE FROM MA_puntuaciones WHERE num_dorsal = " & sNumDorsal & " AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & Val(cbFase.Text) & " AND repesca=" & chkRep.Value)
        End If
    Else
        If MsgBox(mml_FRASE1268, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
            Exit Sub
        End If
        dgPuntuaciones.Col = 0
        sNumDorsal = dgPuntuaciones.Text
        dgPuntuaciones.Col = 1
        sCodCategoria = dgPuntuaciones.Text
        dgPuntuaciones.Col = 2
        sCodBaile = dgPuntuaciones.Text
        dgPuntuaciones.Col = 3
        sCodJuez = dgPuntuaciones.Text
        dgPuntuaciones.Col = 4
        sFase = dgPuntuaciones.Text
        dgPuntuaciones.Col = 5
        sRep = dgPuntuaciones.Text
        dgPuntuaciones.Col = 5
        sParametro = dgPuntuaciones.Text
        
        db.Execute ("DELETE FROM MA_puntuaciones WHERE num_dorsal = " & sNumDorsal & " AND cod_categoria = " & sCodCategoria & " AND cod_baile = " & sCodBaile & " AND cod_juez = '" & sCodJuez & "' AND fase = " & sFase & " AND repesca=" & sRep)
    End If
    
    DoEvents
    Call cmdActualizar_Click
End Sub

Private Sub CommandButton1_Click()

End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    Call cmdActualizar_Click
End Sub



Private Sub addPuntuaciones(bModificar As Boolean)
Dim i As Integer
Dim iValor As Integer
    iValor = Val(tbValor.Tag)
    If iValor > 20 Then
        MsgBox mml_FRASE1267, vbOKOnly Or vbInformation, ""
        Exit Sub
    End If
    
    If cbJuez.Text = "" Or cbDorsal.Text = "" Or tbCodCateg.Text = "" Or Trim$(Mid$(cbFase.Text, 1, 2)) = "" Or _
        cbBaile.Text = "" Or cbParametros.Text = "" Or Val(tbValor.Text) > 10 Or Not IsNumeric(tbValor.Text) Then
        CamposSinCubrir
        Exit Sub
    End If
    
    db.Execute "DELETE FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND num_dorsal=" & Val(cbDorsal.Text) & " AND cod_juez = '" & cbJuez.Text & "' AND repesca=" & chkRep.Value & " AND fase=" & Val(Trim$(Mid$(cbFase.Text, 1, 2))) & " AND cod_baile=" & Val(cbBaile.Text) & " AND parametro = " & Val(cbParametros.Text)
    db.Execute ("INSERT INTO MA_puntuaciones VALUES (" & Val(cbDorsal.Text) & "," & tbCodCateg.Text & "," & Val(cbBaile.Text) & ",'" & cbJuez.Text & "'," & Val(cbFase.Text) & "," & chkRep.Value & "," & Val(cbParametros.Text) & "," & iValor & ")")
    Sleep 500
    DoEvents
    
    Call cmdActualizar_Click
End Sub



Private Sub add_Click()
    If Not IsNumeric(tbValor.Text) Then
        MsgBox mml_FRASE0894, vbOKOnly Or vbCritical, ""
    Else
        addPuntuaciones False
    End If
End Sub

Private Sub dgPuntuaciones_Click()
    With dgPuntuaciones
        If .Row >= 0 Then
            .Col = 3
            cbJuez.Text = .Text
            .Col = 0
            cbDorsal.Text = .Text
            .Col = 2
            cbBaile.Text = .Text
            .Col = 6
            cbParametros.Text = .Text
            .Col = 7
            tbValor.Text = Format(.Text / 2, "#0.0")
        End If
    End With

    ActualizaCombos
End Sub

Private Sub mod_Click()
    If Not IsNumeric(tbValor.Text) Then
        MsgBox mml_FRASE0894, vbOKOnly Or vbCritical, ""
    Else
        addPuntuaciones True
    End If
End Sub

Private Sub dgPuntuaciones_AfterColUpdate(ByVal ColIndex As Integer)
Dim sNumDorsal As String
Dim sCodCategoria As String
Dim sCodCategorias As String
Dim sCodBaile As String
Dim sCodJuez As String
Dim sFase As String
Dim sRep As String
Dim sPuesto As String

    If tbCodCateg.Text = "" Or cbFase.Text = "" Then
        MsgBox mml_FRASE0424, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    dgPuntuaciones.Col = 0
    sNumDorsal = dgPuntuaciones.Text
    dgPuntuaciones.Col = 1
    sCodCategoria = dgPuntuaciones.Text
    dgPuntuaciones.Col = 2
    sCodBaile = dgPuntuaciones.Text
    dgPuntuaciones.Col = 3
    sCodJuez = dgPuntuaciones.Text
    dgPuntuaciones.Col = 4
    sPuesto = dgPuntuaciones.Text
    dgPuntuaciones.Col = 5
    sFase = dgPuntuaciones.Text
    dgPuntuaciones.Col = 6
    sRep = dgPuntuaciones.Text
    
    Debug.Print "DELETE FROM puntuaciones WHERE num_dorsal = " & sNumDorsal & " AND cod_categoria = " & sCodCategoria & " AND cod_baile = " & sCodBaile & " AND cod_juez = '" & sCodJuez & "' AND fase = " & sFase & " AND repesca=" & sRep
    db.Execute ("UPDATE puntuaciones SET puesto = " & sPuesto & " WHERE num_dorsal = " & sNumDorsal & " AND cod_categoria = " & sCodCategoria & " AND cod_baile = " & sCodBaile & " AND cod_juez = '" & sCodJuez & "' AND fase = " & sFase & " AND repesca=" & sRep)
    
    Call cmdActualizar_Click
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))

End Sub




Private Sub ImprimirPuntuaciones(iDorsal As Integer)
Dim iJueces As Integer
Dim iBailes As Integer
Dim iParejas As Integer
Dim i As Integer, j As Integer, k As Integer 'Contadores temporales
Dim rs As Recordset 'Recordset de uso general
Dim rsDorsales As Recordset 'Recordset de manejo de parejas/dorsales
Dim rsBailes As Recordset 'Recordset de manejo de bailes
Dim rsPuestos As Recordset 'Recordset de manejo de puestos
Dim rsPuestosJuez As Recordset 'Recordset de manejo de puestos ordenador por juez
Dim aBailes() As String
Dim aJueces() As String
Dim aDorsales() As String
Dim aParametros(MAX_PARAMETROS_MA) As Integer
Dim iCBailes As Integer
Dim iCJueces As Integer
Dim iCPuestos As Integer
Dim iCDorsales As Integer
Dim iEscala As Integer
Dim iTablas As Integer
Dim iPosY As Integer, iPosX As Integer, iPosFinBailesY As Integer
Dim dPuntosPuesto As Double
Dim iNumDorsales As Integer
Dim iNumRepPuesto As Integer
Dim X As Integer, Y As Integer
Dim iParametros As Integer
Dim rsParametros As Recordset
Dim iCParametros As Integer
Dim iBailesPag As Integer
Dim iHoja As Integer
Dim sPareja As String

    iHoja = 1
    
    If Not C_DEBUG Then On Local Error GoTo error

    If tbCodCateg.Text = "" Or cbFase.Text = "" Or tbCodComp.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbExclamation, mml_FRASE0084
        Exit Sub
    End If
    
    frmImprimirFinal.tbCodCateg.Text = tbCodCateg.Text
    frmImprimirFinal.ImprimirCabecera mml_FRASE0329
        
    Printer.CurrentX = iEscala * 8
    Printer.Print mml_FRASE0557 & iHoja
    
    'Recuperamos información de la pareja
    Set rs = db.OpenRecordset("SELECT nombre_hombre, nombre_mujer FROM parejas p, dorsales d WHERE d.num_dorsal = " & iDorsal & " AND d.cod_pareja = p.codigo AND d.cod_categoria = " & tbCodCateg.Text & " AND d.repesca=" & chkRep.Value & " AND d.fase = " & Val(cbFase.Text), dbOpenSnapshot)
    If Not rs.EOF Then
        Printer.Font = 16
        Printer.FontBold = True
        sPareja = rs.Fields("nombre_hombre") & "  &  " & rs.Fields("nombre_mujer")
        Centrado Printer, sPareja, Printer.Width
        Printer.FontBold = False
    End If
    rs.Close
    
    
    ' Comprobamos el número de jueces
    'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & tbCodCateg.Text, dbOpenSnapshot)
    'iJueces = rs.Fields(0)
    'rs.Close
    ' Comprobamos el número de jueces en las puntuaciones
    Set rs = db.OpenRecordset("SELECT DISTINCT cod_juez FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text), dbOpenSnapshot)
    iJueces = 0
    If Not rs.EOF Then
        rs.MoveLast
        iJueces = rs.RecordCount
    End If
    rs.Close
    
    ' Comprobamos el número de bailes
    If BailesParciales(Val(tbCodCateg.Text)) Then
        ' Comprobamos el número de bailes que hay en las puntuaciones
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes WHERE codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text) & ")", dbOpenSnapshot)
    Else
        ' Comprobamos el número de bailes
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL, dbOpenSnapshot)
    End If
    iBailes = rs.Fields(0)
    rs.Close
    
    'Comprobamos si tiene definidos bailes y jueces
    If iJueces = 0 Or iBailes = 0 Then
        MsgBox mml_FRASE0686, vbOKOnly Or vbCritical, mml_FRASE0096
        Exit Sub
    End If
    
    ' Comprobamos el número de parámetros totales
    iParametros = 0
    Set rs = db.OpenRecordset("SELECT * FROM MA_parametros WHERE activo = 'S' ORDER BY codigo", dbOpenSnapshot)
    While Not rs.EOF
        aParametros(iParametros) = rs.Fields("codigo")
        Inc iParametros
        rs.MoveNext
    Wend
    rs.Close
    iBailesPag = C_MAX_PARAMETROS_PAG_MA / iParametros
    
    ReDim aTabla(iParametros + 1, iJueces + 1)
    'Cargamos las etiquetas de los jueces
    Set rs = db.OpenRecordset("SELECT id_juez, nombre FROM juez_categ jc, jueces j WHERE pasos = 0 AND j.codigo = jc.cod_juez AND cod_categoria = " & tbCodCateg.Text & " ORDER BY id_juez", dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        ReDim aJueces(rs.RecordCount, 2)
        i = 0
        rs.MoveFirst
        While Not rs.EOF
            aJueces(i, 0) = rs!id_juez
            aJueces(i, 1) = rs!Nombre
            aTabla(0, i + 1) = rs!id_juez
            rs.MoveNext
            i = i + 1
        Wend
    End If
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text) & ") ORDER BY cod_baile"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY cod_baile", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    
    iTablas = 1
    aTabla(0, 0) = "Par"
    While Not rsBailes.EOF
        'Cargamos los parámetros
        iCParametros = 1
        Set rsParametros = db.OpenRecordset("SELECT codigo, descorta FROM ma_parametros WHERE activo = 'S' ORDER BY codigo", dbOpenSnapshot)
        While Not rsParametros.EOF
            aTabla(iCParametros, 0) = rsParametros!descorta
            iCPuestos = 1
            Dim iConPar As Integer
            iConPar = 0
            'Cargamos las puntuaciones
            Set rsPuestos = db.OpenRecordset("SELECT cod_baile, cod_juez, parametro, valor FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & _
                " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text) & " AND num_dorsal = " & iDorsal & _
                " AND parametro = " & rsParametros!codigo & " AND cod_baile = " & rsBailes.Fields(0) & _
                " ORDER BY cod_juez, cod_baile, parametro", dbOpenSnapshot)
            While Not rsPuestos.EOF
                ' Solo cargamos el valor, si el juez juzga este parametros
                While aTabla(0, iCPuestos) <> rsPuestos!cod_juez
                    iCPuestos = iCPuestos + 1
                Wend
                aTabla(iCParametros, iCPuestos) = rsPuestos!valor / 2
                rsPuestos.MoveNext
            Wend
            rsPuestos.Close
            rsParametros.MoveNext
            iCParametros = iCParametros + 1
            If Not rsParametros.EOF And iCParametros > iParametros Then
                MsgBox mml_FRASE1270, vbOKOnly Or vbCritical, mml_FRASE0096
                Exit Sub
            End If
        Wend
        rsParametros.Close
        
        ComprobarSaltoPagina iParametros, iEscala, iHoja, sPareja
        CambioLado iPosX, iPosY, iTablas, iJueces <= G_LIM_JUECES_PARA_TABLAS_FINAL_MA
        
        Printer.CurrentY = iPosY
        Printer.Print
        Printer.CurrentX = iPosX
        Printer.FontSize = 10
        Printer.Print rsBailes.Fields(1)
        Printer.Print
        Printer.FontSize = 7
        Printer.FontName = "Arial"
        'Imprimimos los bailes
        DibujarTablaExt Printer, iPosX, Printer.CurrentY, iParametros + 1, iJueces + 1, 350, 250, 350, 0, 0, 0, False
        rsBailes.MoveNext
        Inc iCBailes
        
        Inc iTablas
    
    Wend
    
    ComprobarSaltoPagina iJueces, iEscala, iHoja, sPareja
    CambioLado iPosX, iPosY, iTablas, iJueces <= G_LIM_JUECES_PARA_TABLAS_FINAL_MA
    Inc iTablas
    
    iPosFinBailesY = Printer.CurrentY
        
    'Imprimimos la tabla de los jueces
    ReDim aTabla(iJueces + 1, 2)
    
    
    Printer.CurrentY = iPosY
    Printer.Print
    Printer.Print
    Printer.CurrentX = iPosX
    Printer.FontSize = 10
    Printer.Print mml_FRASE0689
    Printer.Print
    Printer.FontSize = 7
    Printer.FontName = "Arial"
    For i = 0 To iJueces
        aTabla(i, 0) = aJueces(i, 0)
        aTabla(i, 1) = aJueces(i, 1)
    Next
    'DibujarTabla Printer, iPosX, Printer.CurrentY, iJueces, 2, 2500, 250
    DibujarTablaExt Printer, iPosX, Printer.CurrentY, iJueces, 2, 3500, 250, 300, 1, 0, 0
        
    ComprobarSaltoPagina iParametros, iEscala, iHoja, sPareja
    CambioLado iPosX, iPosY, iTablas, iJueces <= G_LIM_JUECES_PARA_TABLAS_FINAL_MA
    Inc iTablas
    
    'Imprimimos la tabla de los parámetros
    ReDim aTabla(20, 2)
    
    Printer.CurrentY = iPosY
    Printer.Print
    Printer.Print
    Printer.CurrentX = iPosX
    Printer.FontSize = 10
    Printer.Print mml_FRASE1287
    Printer.Print
    Printer.FontSize = 7
    Printer.FontName = "Arial"
    Set rs = db.OpenRecordset("SELECT * FROM MA_parametros ORDER BY 1", dbOpenSnapshot)
    i = 0
    While Not rs.EOF
        aTabla(i, 0) = rs.Fields("descorta")
        aTabla(i, 1) = rs.Fields("descripcion")
        Inc i
        rs.MoveNext
    Wend
    rs.Close
    'DibujarTabla Printer, iPosX, Printer.CurrentY, iJueces, 2, 2500, 250
    DibujarTablaExt Printer, iPosX, Printer.CurrentY, i, 2, 3500, 250, 300, 1, 0, 0
        
    ComprobarSaltoPagina iParametros, iEscala, iHoja, sPareja
    CambioLado iPosX, iPosY, iTablas, iJueces <= G_LIM_JUECES_PARA_TABLAS_FINAL_MA
    Inc iTablas
   
    'Tabla de totales ----------------------------------------------------------
    ReDim aTabla(iParametros + 1, 50)
    aTabla(0, 0) = "Par"
    'Ahora dibujamos las tablas de totales
    Set rs = db.OpenRecordset("SELECT descorta FROM MA_parametros WHERE activo = 'S' ORDER BY codigo", dbOpenSnapshot)
    rs.Close
    
    If BailesParciales(Val(tbCodCateg.Text)) Then
        sExecSQL = "SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & tbCodCateg.Text & " AND bc.fase =" & BAILES_FINAL & " AND  b.codigo IN (SELECT cod_baile FROM cal_baile WHERE cod_categoria = " & tbCodCateg.Text & " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text) & ") ORDER BY cod_baile"
        Debug.Print sExecSQL
        Set rsBailes = db.OpenRecordset(sExecSQL, dbOpenSnapshot)
    Else
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT bc.cod_baile, b.nombre, bc.posicion FROM bailes_categ bc, bailes b WHERE bc.cod_baile = b.codigo AND cod_categoria = " & tbCodCateg.Text & " AND fase = " & BAILES_FINAL & " ORDER BY cod_baile", dbOpenSnapshot)
    End If
    'Se calcula cada baile de modo independiente
    iCBailes = 0
    
    While Not rsBailes.EOF
        'Cargamos los parámetros
        iCParametros = 1
        Set rsParametros = db.OpenRecordset("SELECT codigo, descorta FROM MA_parametros WHERE activo = 'S' ORDER BY codigo", dbOpenSnapshot)
        While Not rsParametros.EOF
            aTabla(iCParametros, 0) = rsParametros!descorta
            'Cargamos posiciones por baile
            Set rsPuestos = db.OpenRecordset("SELECT cod_baile, SUM(valor)/2 as valorsuma, COUNT(*) as num_jueces FROM MA_puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & _
                " AND repesca=" & chkRep.Value & " AND fase = " & Val(cbFase.Text) & " AND num_dorsal = " & iDorsal & _
                " AND parametro = " & rsParametros!codigo & " AND cod_baile = " & rsBailes.Fields(0) & _
                " GROUP BY cod_baile ORDER BY cod_baile", dbOpenSnapshot)
            iCPuestos = 0
            aTabla(0, iCBailes + 1) = Mid$(rsBailes.Fields(1), 1, 1)
            aTabla(0, iBailes + 1) = mml_FRASE1288
            aTabla(iParametros + 1, 0) = mml_FRASE1288
                
            If Not rsPuestos.EOF Then
                If aTabla(iCParametros, iBailes + 1) = "" Then
                    aTabla(iCParametros, iBailes + 1) = "0"
                End If
                If NoPresente(iDorsal, tbCodCateg.Text, Val(cbFase.Text), chkRep.Value) Then
                    aTabla(iCParametros, iBailes + 1) = mml_FRASE0687
                    aTabla(iCParametros, iCBailes + 1) = mml_FRASE0687
                Else
                    aTabla(iCParametros, iCBailes + 1) = Trim$(Str$(Round(rsPuestos!valorsuma / rsPuestos!num_jueces, 2)))
                    ' Media de los parámetros
                    aTabla(iCParametros, iBailes + 1) = Trim(Str$(Val(aTabla(iCParametros, iBailes + 1)) + _
                                                    rsPuestos!valorsuma / rsPuestos!num_jueces))
                    'Media de los bailes
                    aTabla(iParametros + 1, iCBailes + 1) = Trim(Str$(Val(aTabla(iParametros + 1, iCBailes + 1)) + _
                                                             rsPuestos!valorsuma / rsPuestos!num_jueces))
                End If
            End If
            
            Dim sValor As String
            
            If iCParametros = iParametros Then
                sValor = aTabla(iParametros + 1, iCBailes + 1)
                If Val(sValor) > 0 Then
                    aTabla(iParametros + 1, iCBailes + 1) = Trim(Str$(Format(Val(sValor) / iParametros, "####.#")))
                Else
                    aTabla(iParametros + 1, iCBailes + 1) = ""
                End If
            End If
            
            If iCBailes = iBailes - 1 Then
                sValor = aTabla(iCParametros, iBailes + 1)
                If Val(sValor) > 0 Then
                    aTabla(iCParametros, iBailes + 1) = Trim(Str$(Format(Val(sValor) / iBailes, "####.#")))
                Else
                    aTabla(iCParametros, iBailes + 1) = ""
                End If
                'Total global
                sValor = aTabla(iCParametros, iBailes + 1)
                If Val(sValor) > 0 Then
                    aTabla(iParametros + 1, iBailes + 1) = Trim(Str$(Format(Val(sValor) + Val(sValor), "####.#")))
                Else
                    aTabla(iParametros + 1, iBailes + 1) = ""
                End If
            End If
            rsPuestos.Close
            
            rsParametros.MoveNext
            iCParametros = iCParametros + 1
            If Not rsParametros.EOF And iCParametros > iParametros Then
                MsgBox mml_FRASE0688, vbOKOnly Or vbCritical, mml_FRASE0096
            End If
        Wend
        
        rsParametros.Close
        rsBailes.MoveNext
        iCBailes = iCBailes + 1
    Wend
    sValor = aTabla(iParametros + 1, iBailes + 1)
    If Val(sValor) > 0 Then
        aTabla(iParametros + 1, iBailes + 1) = Trim(Str$(Format(Val(sValor) / iParametros, "####.#")))
    Else
        aTabla(iParametros + 1, iBailes + 1) = ""
    End If
    
    'Calculamos la media de los bailes
    'For i = 0 To iBailes - 1
    '    If aTabla(iParametros + 1, i + 1) = "" Then
    '        aTabla(iParametros + 1, i + 1) = ""
    '    Else
    '        aTabla(iParametros + 1, i + 1) = Trim(Str$(Format(Val(aTabla(i + 1, i + 1)) / iParametros, "####.#")))
    '    End If
    'Next
    
    iCPuestos = iCPuestos + 2
    Printer.Print
    Printer.FontSize = 14
    Printer.CurrentX = iPosX
    Printer.CurrentY = iPosY
    Printer.Print mml_FRASE0693
    Printer.FontSize = 8
    Printer.Print
    'DibujarTabla Printer, 100, Printer.CurrentY, iParejas + 1, iBailes + 2 + iCPuestos, 650, 350
    DibujarTablaExt Printer, iPosX, Printer.CurrentY, iParametros + 2, iBailes + 2, 650, 350, 500, iBailes + 1, 0, iBailes + 1
    Printer.Print
    On Local Error Resume Next
    If Mid$(C_LOGO_PATH, 1, 1) = "*" Then
        Printer.PaintPicture LoadPicture(Mid$(C_LOGO_PATH, 2)), (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    Else
        Printer.PaintPicture frmMenu.picEPADance.Picture, (Printer.Width - C_ANCHO_LOGO) / 2, Printer.CurrentY
    End If
    On Local Error GoTo error
    Printer.Print
    
    ' {·M} Imprimimos el hueco para la firma
    iPosY = Printer.CurrentY
    Printer.CurrentX = 0
    iPosY = Printer.CurrentY
    Printer.FontSize = 10
    Printer.Print mml_FRASE0694
    Printer.Line (0, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.CurrentX = 1800
    Printer.CurrentY = iPosY
    Printer.Print mml_FRASE0033
    Printer.Line (1800, Printer.CurrentY)-Step(1700, 1000), 0, B
    
    Printer.EndDoc
    
    Exit Sub

error:
    Dim Msj As String
    If Err.Number <> 0 Then
       Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
             & Err.Source & Chr(13) & Err.Description
       MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
    End If

End Sub


Private Sub tbValor_change()
    If Not IsNumeric(tbValor.Text) Then
        tbValor.Text = "0"
    End If
    tbValor.Tag = Val(tbValor.Text) * 2
End Sub

Sub ActualizaCombos()
    ActualizaCombo cbJuez
    ActualizaCombo cbBaile
    ActualizaCombo cbParametros
    ActualizaCombo cbDorsal
    
End Sub

Sub ActualizaCombo(cb As ComboBox)
    Dim sTmp As String
    Dim i As Integer
    
    sTmp = cb.Text
    
    For i = 0 To cb.ListCount - 1
        If Val(sTmp) = Val(cb.List(i)) Then
            cb.ListIndex = i
            Exit Sub
        End If
    Next
End Sub

Sub ComprobarSaltoPagina(iFilas As Integer, iEscala As Integer, iHoja As Integer, sPareja As String)
    If Printer.CurrentY + (iFilas + 1) * C_ALTURA_FILA_TABLA > Printer.Height Then
        Inc iHoja
        Printer.NewPage
        frmImprimirFinal.tbCodCateg.Text = tbCodCateg.Text
        frmImprimirFinal.ImprimirCabecera mml_FRASE0329
            
        Printer.CurrentX = iEscala * 8
        Printer.Print mml_FRASE0557 & iHoja
        Centrado Printer, sPareja, Printer.Width
        
        isPosY1 = Printer.CurrentY
        isPosY2 = Printer.CurrentY
        
    End If
End Sub

Sub CambioLado(iPosX As Integer, iPosY As Integer, iTablas As Integer, bCond As Boolean)

    Printer.Print
    If (iTablas Mod 2 = 0) And bCond Then
        iPosX = C_POS_COLUMNA_2
        isPosY1 = Printer.CurrentY
        Printer.CurrentX = iPosX
        Printer.CurrentY = isPosY2
    Else
        isPosY2 = Printer.CurrentY
        iPosY = Printer.CurrentY
        Printer.CurrentY = isPosY1
        iPosX = 100
    End If
    
End Sub

