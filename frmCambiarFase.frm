VERSION 5.00
Begin VB.Form frmCambiarFase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio/Copia de fase"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPublicar 
      Caption         =   "mml_FRASE0186"
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
      Left            =   5655
      TabIndex        =   10
      Top             =   1770
      Width           =   1305
   End
   Begin VB.CommandButton cmdCambiarCopiar 
      Caption         =   "mml_FRASE0001"
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
      Left            =   2910
      TabIndex        =   4
      Top             =   1770
      Width           =   2655
   End
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "mml_FRASE0000"
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
      Left            =   165
      TabIndex        =   3
      Top             =   1770
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
      Left            =   7080
      TabIndex        =   2
      Top             =   1770
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   8295
      Begin VB.ComboBox cbFase 
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
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   585
         Width           =   5700
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
         Height          =   330
         Left            =   2430
         TabIndex        =   8
         Top             =   1080
         Width           =   2265
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
         Left            =   3255
         TabIndex        =   6
         Top             =   180
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
         Left            =   2415
         TabIndex        =   5
         Top             =   180
         Width           =   855
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
         Left            =   255
         TabIndex        =   7
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "mml_FRASE0316"
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
         TabIndex        =   1
         Top             =   645
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCambiarFase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim g_iFase As Integer
Dim g_iRepesca As Integer
Dim g_CodComp As Long


Public Sub CambiarFase(sCodCateg As Long, iFase As Integer, iRepesca As Integer, iCodcomp As Integer)
Dim i As Integer, iIndex As Integer, j As Integer
    g_CodComp = iCodcomp
    tbCodCateg.Text = sCodCateg
    tbDescCat.Text = sDescCategoria(sCodCateg)
    chkRep.Value = iRepesca
    
    g_iFase = iFase
    g_iRepesca = iRepesca
    
    i = 1
    j = 0
    iIndex = 0
    cbFase.Clear
    While i <= 256
        cbFase.AddItem i & " - " & sDescFase(i)
        If i = iFase Then iIndex = j
        i = i * 2
        j = j + 1
    Wend
    cbFase.ListIndex = iIndex
    
    g_iFase = iFase
    
    Me.Show 1
End Sub

Private Sub CambiarCopiar(bCambiar As Boolean)
Dim iNuevaFase As Integer
Dim aDorsales(200) As Long
Dim iDorsal, i As Integer
Dim iCDorsal As Integer
Dim iRepesca As Integer
Dim rs As Recordset

    If Not C_DEBUG Then On Local Error GoTo error
    
    If chkRep.Value = vbChecked Then
        iRepesca = 1
    Else
        iRepesca = 0
    End If
    
    iNuevaFase = Val(cbFase.List(cbFase.ListIndex))
    
    If Val(cbFase.List(cbFase.ListIndex)) = g_iFase And chkRep.Value = g_iRepesca Then
        MsgBox mml_FRASE0970, vbInformation Or vbOKOnly, mml_FRASE0096
        Exit Sub
    End If
    
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & iNuevaFase & " AND repesca=" & chkRep.Value, dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            If MsgBox(mml_FRASE0317, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                Exit Sub
            End If
        End If
    rs.Close
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & iNuevaFase & " AND repesca=" & chkRep.Value, dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            If MsgBox(mml_FRASE0318, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
                Exit Sub
            End If
        End If
    rs.Close
    
    ' Recuperamos los dorsales
    frmADorsales.dgDorsales.Col = 0
    iCDorsal = 0
    For Each iDorsal In frmADorsales.dgDorsales.SelBookmarks
        frmADorsales.dgDorsales.Bookmark = iDorsal
        'frmADorsales.dgDorsales.Row = iDorsal - 1
        aDorsales(iCDorsal) = Val(frmADorsales.dgDorsales.Text)
        Inc iCDorsal
    Next
    
    If bCambiar Then
        If MsgBox(mml_FRASE0533, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
            ' Cambiamos los dorsales selecionados
            For i = 0 To iCDorsal - 1
                db.Execute ("UPDATE dorsales SET repesca = " & iRepesca & ", fase = " & iNuevaFase & " WHERE cod_categoria = " & tbCodCateg.Text & " And codigo = " & aDorsales(i))
            Next
        Else
            db.Execute ("UPDATE dorsales SET repesca = " & iRepesca & ", fase = " & iNuevaFase & " WHERE cod_categoria = " & tbCodCateg.Text & " AND fase = " & g_iFase & " AND repesca = " & g_iRepesca)
        End If
    Else
        If iCDorsal = 0 Then
            MsgBox "Debe seleccionar los dorsales que quiere copiar a la nueva fase. Para selecionar varios pulse la tecla Control mientrar pincha en el dorsal", vbCritical Or vbOKOnly, mml_FRASE0096
            Exit Sub
        End If
        
        ' Cambiamos los dorsales selecionados
        For i = 0 To iCDorsal - 1
            Set rs = db.OpenRecordset("SELECT * FROM dorsales WHERE codigo = " & aDorsales(i), dbOpenSnapshot)
            db.Execute ("INSERT INTO dorsales VALUES (" & MaxCod("dorsales") & "," & rs!num_dorsal & "," & rs!cod_categoria & "," & Val(cbFase.List(cbFase.ListIndex)) & "," & rs!cod_pareja & ",0," & chkRep.Value & ")")
            rs.Close
        Next
    End If
    
    'Preguntar si se borra la categoría del horario
    If MsgBox(mml_FRASE0971, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        db.Execute ("DELETE FROM horario WHERE cod_categoria = " & tbCodCateg.Text & " AND numfase = " & g_iFase & " AND repesca = " & g_iRepesca)
    End If
    
    Call frmADorsales.cmdActualizar_Click
    
    Exit Sub
error:
    ProcesarError
End Sub



Private Sub cmdCambiar_Click()
    CambiarCopiar True
End Sub

Private Sub cmdPublicar_Click()
    With frmPublicar
        .tbCodComp.Text = g_CodComp
        .tbCodCat.Text = tbCodCateg.Text
        .tbCodFase.Text = Val(cbFase.List(cbFase.ListIndex))
        .chkRep.Value = chkRep.Value
        
        .tbDescCat.Text = sDescCategoria(Val(.tbCodCat.Text))
        .tbDescFase.Text = sDescFase(Val(.tbCodFase.Text))
        .tbDescComp.Text = sDescCompeticion(g_CodComp)
        
        .AsignarComentario
        .Show vbModal
    End With
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdCambiarCopiar_Click()
    CambiarCopiar False
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub
