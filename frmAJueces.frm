VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAJueces 
   Caption         =   "mml_FRASE0364"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7980
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "mml_FRASE0029"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8550
      TabIndex        =   19
      Top             =   7470
      Width           =   1875
   End
   Begin VB.CommandButton cmdAct 
      Caption         =   "mml_FRASE0250"
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
      Left            =   6000
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
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
      Left            =   4680
      TabIndex        =   11
      Top             =   4320
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
      Left            =   3360
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "mml_FRASE0365"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   10455
      Begin VB.TextBox tbObs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2280
         Width           =   7695
      End
      Begin VB.TextBox tbCat 
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
         Left            =   6480
         TabIndex        =   16
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox tbTlf 
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
         TabIndex        =   13
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox tbDir 
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
         Top             =   1320
         Width           =   7695
      End
      Begin VB.TextBox tbNombre 
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
         TabIndex        =   5
         Top             =   840
         Width           =   7695
      End
      Begin VB.TextBox tbCodJuez 
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
      Begin VB.Label Label6 
         Caption         =   "mml_FRASE0366"
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
         TabIndex        =   18
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         Left            =   5160
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "mml_FRASE0367"
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
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "mml_FRASE0274"
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
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "mml_FRASE0266"
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
      Caption         =   "mml_FRASE0050"
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   10455
      Begin MSDataGridLib.DataGrid dgJuez 
         Bindings        =   "frmAJueces.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4048
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
   Begin MSAdodcLib.Adodc adoJuez 
      Height          =   495
      Left            =   240
      Top             =   7470
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
      RecordSource    =   "SELECT * FROM jueces ORDER BY nombre"
      Caption         =   "mml_FRASE0050"
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
End
Attribute VB_Name = "frmAJueces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAct_Click()
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    Sleep 500
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents:
    adoJuez.Refresh
    dgJuez.Refresh

    If adoJuez.Recordset.EOF Then
        dgJuez.Enabled = False
    Else
        dgJuez.Enabled = True
    End If
End Sub

Private Sub cmdBorrar_Click()
    If tbCodJuez.Text = "" Then
        MsgBox mml_FRASE0368, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If MsgBox(mml_FRASE0263, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    db.Execute ("DELETE FROM jueces WHERE codigo = " & tbCodJuez.Text)
    Call cmdNuevo_Click
    Call cmdAct_Click
End Sub

Private Sub cmdGrabar_Click()
Dim rs As Recordset
    If tbNombre.Text = "" Then
        MsgBox mml_FRASE0277, vbOKOnly Or vbInformation, mml_FRASE0084
        Exit Sub
    End If
    If tbCodJuez.Text = "" Then
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM jueces WHERE nombre Like '" & tbNombre.Text & "'", dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            If MsgBox("Juez duplicado, ¿Grabar de todos modos?", vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                rs.Close
                Exit Sub
            End If
        End If
        rs.Close
        db.Execute ("INSERT INTO jueces VALUES(" & MaxCod("jueces") & ", '" & tbNombre.Text & "','" & tbDir.Text & "','" & tbTlf.Text & "','" & tbCat.Text & "','" & tbObs.Text & "')")
    Else
        db.Execute ("UPDATE jueces SET nombre = '" & tbNombre.Text & _
                    "',direccion ='" & tbDir.Text & _
                    "',telefono='" & tbTlf.Text & _
                    "',categoria='" & tbCat.Text & _
                    "',observaciones='" & tbObs.Text & _
                    "' WHERE codigo = " & tbCodJuez.Text)
    End If
    Call cmdAct_Click
End Sub

Private Sub cmdNuevo_Click()
    tbCodJuez.Text = ""
    tbNombre.Text = ""
    tbDir.Text = ""
    tbTlf.Text = ""
    tbCat.Text = ""
    tbObs.Text = ""
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub



Private Sub dgJuez_Click()
    dgJuez.Col = 0
    tbCodJuez.Text = dgJuez.Text
    dgJuez.Col = 1
    tbNombre.Text = dgJuez.Text
    dgJuez.Col = 2
    tbDir.Text = dgJuez.Text
    dgJuez.Col = 3
    tbTlf.Text = dgJuez.Text
    dgJuez.Col = 4
    tbCat.Text = dgJuez.Text
    dgJuez.Col = 5
    tbObs.Text = dgJuez.Text
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub
