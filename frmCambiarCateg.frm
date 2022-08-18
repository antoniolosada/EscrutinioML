VERSION 5.00
Begin VB.Form frmCambiarCateg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0531"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopiarDorsales 
      Caption         =   "mml_FRASE1154"
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
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   2655
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
      Height          =   360
      Left            =   1980
      Picture         =   "frmCambiarCateg.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   225
      Width           =   450
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "mml_FRASE0532"
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
      Left            =   180
      TabIndex        =   5
      Top             =   960
      Width           =   2655
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
      Left            =   5820
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
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
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   855
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
         Left            =   3240
         TabIndex        =   1
         Top             =   240
         Width           =   4935
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
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCambiarCateg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCodCategorias As String
Dim sCodCompeticion As String

Sub Cambiar(sCodComp As String, sCodCateg As String)
    sCodCategorias = sCodCateg
    sCodCompeticion = sCodComp
    
    Me.Show 1
End Sub

Private Sub CambiarCateg(bCambiar As Boolean)
Dim iNuevaFase As Integer
Dim aDorsales(200) As Long
Dim iDorsal, i As Integer
Dim iCDorsal As Integer
Dim rs As Recordset
    
    If tbCodCateg.Text = "" Then
        CamposSinCubrir
        Exit Sub
    End If

    If Not C_DEBUG Then On Local Error GoTo error
    ' Recuperamos los dorsales
    frmADorsales.dgDorsales.Col = 0
    iCDorsal = 0
    For Each iDorsal In frmADorsales.dgDorsales.SelBookmarks
        frmADorsales.dgDorsales.Bookmark = iDorsal
        'frmADorsales.dgDorsales.Row = iDorsal - 1
        aDorsales(iCDorsal) = Val(frmADorsales.dgDorsales.Text)
        Inc iCDorsal
    Next
    
    'Si copiamos, solo son los seleccionados
    If Not bCambiar Then
        MsgBox mml_FRASE1155, vbOKOnly Or vbInformation, G_MSG_MENSAJE
        ' Copiamos los dorsales selecionados
        For i = 0 To iCDorsal - 1
            Set rs = db.OpenRecordset("SELECT * FROM dorsales WHERE codigo = " & aDorsales(i), dbOpenSnapshot)
            db.Execute ("INSERT INTO dorsales VALUES (" & MaxCod("dorsales") & "," & rs!num_dorsal & "," & tbCodCateg.Text & "," & rs!fase & "," & rs!cod_pareja & ",0," & rs!repesca & ")")
            rs.Close
        Next
    Else
        If MsgBox(mml_FRASE0533, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
            For i = 0 To iCDorsal - 1
                db.Execute ("UPDATE dorsales SET cod_categoria = " & tbCodCateg.Text & " WHERE cod_categoria = " & sCodCategorias & " And codigo = " & aDorsales(i))
            Next
        Else
            db.Execute ("UPDATE dorsales SET cod_categoria = " & tbCodCateg.Text & " WHERE cod_categoria = " & sCodCategorias)
        End If
    End If
    MsgBox mml_FRASE1055, vbOKOnly Or vbCritical, G_MSG_AVISO
    Unload Me
    Exit Sub
error:
    ProcesarError
End Sub

Private Sub cmdCambiar_Click()
    CambiarCateg True
End Sub

Private Sub cmdCateg_Click()
    tbCodCateg.Text = sSeleccionar("SELECT * FROM categorias c WHERE cod_competicion =" & sCodCompeticion & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCateg.Text = sResultado(2)

End Sub

Private Sub cmdCopiarDorsales_Click()
    CambiarCateg False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub

