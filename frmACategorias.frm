VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmACategorias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0248"
   ClientHeight    =   9000
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12090
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalirCateg 
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
      Left            =   10440
      TabIndex        =   37
      Top             =   4680
      Width           =   1500
   End
   Begin VB.CommandButton CommandButton3 
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
      Left            =   1920
      Picture         =   "frmACategorias.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4110
      Width           =   435
   End
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
      Height          =   405
      Left            =   1920
      Picture         =   "frmACategorias.frx":046A
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3630
      Width           =   435
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
      Left            =   2100
      Picture         =   "frmACategorias.frx":08D4
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   120
      Width           =   435
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
      Left            =   8640
      TabIndex        =   30
      Top             =   9120
      Width           =   1620
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   3720
      TabIndex        =   29
      Top             =   4695
      Width           =   1860
   End
   Begin VB.CommandButton cmdRecHoras 
      Caption         =   "mml_FRASE0249"
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
      Left            =   5670
      TabIndex        =   24
      Top             =   4680
      Width           =   3030
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton cmdAct 
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
      Left            =   8760
      TabIndex        =   12
      Top             =   4695
      Width           =   1620
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "mml_FRASE0251"
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
      Left            =   2445
      TabIndex        =   11
      Top             =   4710
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "mml_FRASE0252"
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
      Left            =   1155
      TabIndex        =   10
      Top             =   4695
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "mml_FRASE0253"
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
      Left            =   105
      TabIndex        =   9
      Top             =   4695
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   4035
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   12015
      Begin VB.CommandButton cmdCorregirNombres 
         Caption         =   "mml_FRASE1183"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10080
         TabIndex        =   36
         Top             =   780
         Width           =   1710
      End
      Begin VB.CheckBox chkImpUnaTanda 
         Caption         =   "mml_FRASE1108"
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
         Left            =   1920
         TabIndex        =   35
         Top             =   2610
         Width           =   6900
      End
      Begin VB.CheckBox chkCombinar 
         Caption         =   "mml_FRASE0972"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   9300
         TabIndex        =   34
         Top             =   1260
         Width           =   2445
      End
      Begin VB.TextBox tbDorsalesTanda 
         Alignment       =   2  'Center
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
         Left            =   8100
         TabIndex        =   27
         Top             =   1335
         Width           =   855
      End
      Begin VB.CheckBox chkMostrarPosicion 
         Caption         =   "mml_FRASE0254"
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
         Left            =   9315
         TabIndex        =   26
         Top             =   1830
         Width           =   2445
      End
      Begin VB.CheckBox chkRecParcialBailes 
         Caption         =   "mml_FRASE0255"
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
         Left            =   1920
         TabIndex        =   25
         Top             =   2250
         Width           =   4680
      End
      Begin VB.TextBox tbCodMod 
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
         Left            =   2370
         TabIndex        =   22
         Top             =   3510
         Width           =   855
      End
      Begin VB.TextBox tbDescMod 
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
         TabIndex        =   21
         Top             =   3510
         Width           =   8535
      End
      Begin VB.TextBox tbCodGrupoEdad 
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
         Left            =   2370
         TabIndex        =   19
         Top             =   3030
         Width           =   855
      End
      Begin VB.TextBox tbDescGrupoEdad 
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
         TabIndex        =   18
         Top             =   3030
         Width           =   8535
      End
      Begin VB.ComboBox cbCatMax 
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
         Height          =   405
         ItemData        =   "frmACategorias.frx":0D3E
         Left            =   1920
         List            =   "frmACategorias.frx":0D84
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1320
         Width           =   3975
      End
      Begin VB.TextBox tbHoraCat 
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
         Left            =   1920
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox tbDescCat 
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
         Left            =   1920
         MaxLength       =   35
         TabIndex        =   5
         Top             =   840
         Width           =   8055
      End
      Begin VB.TextBox tbCodCat 
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
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "mml_FRASE0256"
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
         Left            =   5985
         TabIndex        =   28
         Top             =   1335
         Width           =   2040
      End
      Begin VB.Label Label7 
         Caption         =   "mml_FRASE0187"
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
         Left            =   210
         TabIndex        =   23
         Top             =   3510
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "mml_FRASE0257"
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
         Left            =   210
         TabIndex        =   20
         Top             =   3030
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0258"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "mml_FRASE0259"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "mml_FRASE0260"
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
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "mml_FRASE0261"
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "mml_FRASE0032"
      Height          =   3660
      Left            =   0
      TabIndex        =   0
      Top             =   5280
      Width           =   12015
      Begin MSDataGridLib.DataGrid dgCat 
         Bindings        =   "frmACategorias.frx":0DF7
         Height          =   3315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   5847
         _Version        =   393216
         AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc adoCat 
      Height          =   495
      Left            =   540
      Top             =   9060
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
      RecordSource    =   "SELECT * FROM categorias ORDER BY descripcion"
      Caption         =   "mml_FRASE0032"
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
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmACategorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAct_Click()
    adoCat.ConnectionString = "DSN=Escrutinio"
    adoCat.RecordSource = "SELECT * FROM categorias WHERE cod_competicion LIKE '" & tbCodComp.Text & "%' ORDER BY descripcion"
    adoCat.Refresh
    
    If adoCat.Recordset.EOF Then
        dgCat.Enabled = False
    Else
        dgCat.Enabled = True
    End If
    
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    Sleep 500
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    adoCat.Refresh
    dgCat.Refresh
End Sub

Private Sub cmdBorrar_Click()
    If tbCodCat.Text = "" Then
        MsgBox mml_FRASE0262, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    db.Execute ("DELETE FROM cal_baile WHERE cod_categoria = " & tbCodCat.Text)
    db.Execute ("DELETE FROM cal_conjunto WHERE cod_categoria = " & tbCodCat.Text)
    db.Execute ("DELETE FROM puntuaciones WHERE cod_categoria = " & tbCodCat.Text)
    db.Execute ("DELETE FROM bailes_categ WHERE cod_categoria = " & tbCodCat.Text)
    db.Execute ("DELETE FROM juez_categ WHERE cod_categoria = " & tbCodCat.Text)
    db.Execute ("DELETE FROM dorsales WHERE cod_categoria = " & tbCodCat.Text)
    db.Execute ("DELETE FROM categorias WHERE codigo = " & tbCodCat.Text)
    Call cmdNuevo_Click
    Call cmdAct_Click
End Sub

Private Sub cmdCorregirNombres_Click()
Dim rs As Recordset
    
    If Val(tbCodComp.Text) = 0 Then
        CamposSinCubrir
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE1184, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbNo Then
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("SELECT codigo, descripcion, id_categoria FROM categorias WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    While Not rs.EOF
        db.Execute "UPDATE categorias SET descripcion = '" & Left$(CorregirNombre(SinNulos(rs.Fields("descripcion"))), C_MAX_LEN_DESC_CATEGORIA) & "' WHERE codigo = " & rs.Fields("codigo")
        db.Execute "UPDATE categorias SET id_categoria = '" & Trim$(SinNulos(rs.Fields("id_categoria"))) & "' WHERE codigo = " & rs.Fields("codigo")
        rs.MoveNext
    Wend
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT orden, grupo FROM horario WHERE cod_competicion = " & tbCodComp.Text, dbOpenSnapshot)
    While Not rs.EOF
        db.Execute "UPDATE horario SET grupo = '" & CorregirNombre(SinNulos(rs.Fields("grupo"))) & "' WHERE orden = " & rs.Fields("orden") & " AND cod_competicion = " & tbCodComp.Text
        rs.MoveNext
    Wend
    rs.Close
    
    MsgBox G_MSG_OPERACION_OK, vbOKOnly Or vbInformation, G_MSG_MENSAJE

End Sub

Private Sub cmdGrabar_Click()
Dim SelRow
    If Not IsDate(tbHoraCat.Text) Or tbDescCat.Text = "" Or cbCatMax.ListIndex = -1 Or tbCodGrupoEdad.Text = "" Or tbCodMod.Text = "" Then
        MsgBox mml_FRASE0264, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    
    On Local Error Resume Next
    SelRow = dgCat.Row + dgCat.FirstRow - 1
    On Local Error GoTo 0
    
    If tbCodCat.Text = "" Then
        sSQL = "INSERT INTO categorias VALUES(" & MaxCod("categorias") & ", '" & tbDescCat.Text & "','" & cbCatMax.List(cbCatMax.ListIndex) & "','" & tbCodGrupoEdad.Text & "','" & tbCodComp.Text & "','" & tbCodMod.Text & "','" & tbHoraCat.Text & "'," & chkRecParcialBailes.Value & "," & chkMostrarPosicion.Value & "," & tbDorsalesTanda.Text & "," & chkCombinar.Value & "," & chkImpUnaTanda.Value & " )"
        Debug.Print sSQL
        db.Execute (sSQL)
    Else
        db.Execute ("UPDATE categorias SET descripcion = '" & tbDescCat.Text & _
                    "', id_categoria='" & cbCatMax.List(cbCatMax.ListIndex) & _
                    "', cod_grupoedad=" & tbCodGrupoEdad.Text & _
                    ", cod_competicion=" & tbCodComp.Text & _
                    ", cod_modalidad=" & tbCodMod.Text & _
                    ", rec_parcial_bailes=" & chkRecParcialBailes.Value & _
                    ", mostrar_posicion=" & chkMostrarPosicion.Value & _
                    ", dorsales_tanda=" & Val(tbDorsalesTanda.Text) & _
                    ", combinar_dorsales=" & chkCombinar.Value & _
                    ", imprimir_una_hoja_puntuaciones = " & chkImpUnaTanda.Value & _
                    ", hora='" & tbHoraCat.Text & "' WHERE codigo = " & tbCodCat.Text)
    End If
    Call cmdAct_Click
    If SelRow > 0 Then dgCat.Row = SelRow - 1
End Sub

Private Sub cmdNuevo_Click()
    tbCodCat.Text = ""
    tbDescCat.Text = ""
    tbHoraCat.Text = ""
    cbCatMax.ListIndex = -1
    tbHoraCat.Text = ""
    tbCodGrupoEdad.Text = ""
    tbDescGrupoEdad.Text = ""
    tbCodMod.Text = ""
    tbDescMod.Text = ""
    chkRecParcialBailes.Value = 0
    chkMostrarPosicion.Value = 0
    chkCombinar.Value = 0
    tbDorsalesTanda.Text = iDorsalesPorTanda(Str$(Val(tbCodComp.Text)))
    chkImpUnaTanda.Value = ImpHojaUnica
End Sub

Private Sub cmdRecHoras_Click()
Dim rs As Recordset
    If tbCodComp.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If
    Set rs = db.OpenRecordset("SELECT h.hora,c.codigo FROM horario h, categorias c WHERE h.cod_competicion = " & tbCodComp.Text & " AND h.cod_categoria = c.codigo AND h.numfase = (SELECT MAX(numfase) FROM horario h1 WHERE h1.cod_categoria = h.cod_categoria)", dbOpenSnapshot)
    While Not rs.EOF
        db.Execute "UPDATE categorias SET hora = '" & Format$(rs!hora, "hh:mm") & "' WHERE codigo = " & rs!codigo
        rs.MoveNext
    Wend
    rs.Close
    
    MsgBox mml_FRASE0265, vbOKOnly Or vbInformation, mml_FRASE0086
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub cmdSelComp_Click()

End Sub

Private Sub cmdSalirCateg_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If tbCodComp.Text = "" Then
        MsgBox "Debe cubrir el código de la competición", vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If MsgBox("¿Seguro que desea borrar todas las categorías de la competición?", vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    BorrarDatosComp Val(tbCodComp.Text), False, False
    Call cmdNuevo_Click
    Call cmdAct_Click

End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
    Call cmdAct_Click
End Sub

Private Sub CommandButton2_Click()
    tbCodGrupoEdad.Text = sSeleccionar("SELECT * FROM gruposedad")
    tbDescGrupoEdad.Text = sResultado(2)

End Sub

Private Sub CommandButton3_Click()
    tbCodMod.Text = sSeleccionar("SELECT * FROM modalidad")
    tbDescMod.Text = sResultado(2)

End Sub

Private Sub dgCat_Click()
Dim i As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    dgCat.Col = 0
    tbCodCat.Text = dgCat.Text
    dgCat.Col = 1
    tbDescCat.Text = dgCat.Text
    dgCat.Col = 2
    If cbCatMax.ListCount > 0 Then
        cbCatMax.ListIndex = 0
    Else
        MsgBox mml_FRASE1216, vbOKOnly Or vbInformation, "ERROR"
    End If
    For i = 0 To cbCatMax.ListCount - 1
        If cbCatMax.List(i) = dgCat.Text Then
            cbCatMax.ListIndex = i
        End If
    Next
    dgCat.Col = 3
    tbCodGrupoEdad.Text = dgCat.Text
    tbDescGrupoEdad.Text = Buscar("GruposEdad", "nombre", dgCat.Text)
    dgCat.Col = 4
    tbCodComp.Text = dgCat.Text
    tbDescComp.Text = Buscar("competiciones", "descripcion", dgCat.Text)
    dgCat.Col = 5
    tbCodMod.Text = dgCat.Text
    tbDescMod.Text = Buscar("modalidad", "nombre", dgCat.Text)
    dgCat.Col = 6
    tbHoraCat.Text = dgCat.Text
    dgCat.Col = 7
    chkRecParcialBailes.Value = dgCat.Text
    dgCat.Col = 8
    chkMostrarPosicion.Value = dgCat.Text
    dgCat.Col = 9
    tbDorsalesTanda.Text = dgCat.Text
    dgCat.Col = 10
    chkCombinar.Value = Val(dgCat.Text)
    dgCat.Col = 11
    chkImpUnaTanda.Value = Val(dgCat.Text)
    Exit Sub
error:
    ProcesarError "dgCat_Click"
End Sub

Private Sub Form_Load()
Dim rs As Recordset
    
    TraducirCadenas Me
    On Local Error GoTo error
    
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))

    'Actualizamos las categorías
    Set rs = db.OpenRecordset("SELECT * FROM desccategoria ORDER BY 1", dbOpenSnapshot)
    cbCatMax.Clear
    While Not rs.EOF
        cbCatMax.AddItem SinNulos(rs!DESCRIPCION)
        rs.MoveNext
    Wend
    rs.Close
    
    ActualizarCategorias cbCatMax, Val(tbCodComp.Text)
    
    cmdAct_Click
    Exit Sub
error:
    ProcesarError "Form_load"
End Sub


Private Sub tbDescCat_KeyPress(KeyAscii As Integer)
    LimitarDescCateg KeyAscii
End Sub
