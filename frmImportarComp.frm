VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmImportarComp 
   Caption         =   "mml_FRASE0661"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6660
      TabIndex        =   4
      Top             =   4185
      Width           =   1890
   End
   Begin MSComDlg.CommonDialog DlgFicheros 
      Left            =   225
      Top             =   4455
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelecBD 
      Caption         =   "mml_FRASE0662"
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
      Left            =   225
      TabIndex        =   2
      Top             =   4185
      Width           =   3555
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "mml_FRASE0663"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   4185
      Width           =   2520
   End
   Begin VB.Frame Frame1 
      Height          =   4065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      Begin VB.ListBox lstComp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   45
         TabIndex        =   3
         Top             =   180
         Width           =   8700
      End
   End
End
Attribute VB_Name = "frmImportarComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database

Private Sub cmdImportar_Click()
Dim rs As Recordset, rs1 As Recordset
Dim iCodcomp As Integer
Dim iMinCod As Long, iMaxCod As Long
Dim aValores(5) As TValores

    iCodcomp = Val(lstComp.List(lstComp.ListIndex))
    If lstComp.ListIndex = -1 Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbExclamation, mml_FRASE0096
        Exit Sub
    End If
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    If MsgBox(mml_FRASE0664, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        If MsgBox(mml_FRASE0665, vbYesNo Or vbCritical, mml_FRASE0084) = vbYes Then
            'Importación de datos
            
                ' Localizamos el incremento del código de la competición
                iMinCod = iCodcomp
                iMaxCod = MaxCod("competiciones")
                
                aValores(3).Nombre = "cod_competicion"
                aValores(3).operacion = mml_FRASE0666
                aValores(3).valor = iMaxCod - iMinCod
                
                aValores(0).Nombre = "codigo"
                aValores(0).operacion = mml_FRASE0666
                aValores(0).valor = iMaxCod - iMinCod
            
            ImportarDatosConControl db1, "competiciones", "SELECT * FROM competiciones WHERE codigo = " & iCodcomp, aValores
            'Sumamos los mismo a cod_categoria y al codigo de la categoría
                ' Localizamos el incremento del código de las categoías
                Set rs = db1.OpenRecordset("SELECT MIN(codigo) FROM categorias WHERE cod_competicion = " & iCodcomp, dbOpenSnapshot)
                If Not IsNull(rs.Fields(0)) Then
                    iMinCod = rs.Fields(0)
                    rs.Close
                    iMaxCod = MaxCod("categorias")
                    
                    aValores(1).Nombre = "cod_categoria"
                    aValores(1).operacion = mml_FRASE0666
                    aValores(1).valor = iMaxCod - iMinCod
                    
                    aValores(0).Nombre = "codigo"
                    aValores(0).operacion = mml_FRASE0666
                    aValores(0).valor = iMaxCod - iMinCod
                ImportarDatosConControl db1, "categorias", "SELECT * FROM categorias WHERE cod_competicion = " & iCodcomp & " ORDER BY codigo", aValores
                End If
            'Sumamos los mismo a cod_pareja y al código de la pareja
                'Localizamos el incremento del código de parejas
                Set rs = db1.OpenRecordset("SELECT MIN(codigo) FROM parejas WHERE cod_competicion = " & iCodcomp, dbOpenSnapshot)
                iMinCod = rs.Fields(0)
                rs.Close
                iMaxCod = MaxCod("parejas")
                
                aValores(2).Nombre = "cod_pareja"
                aValores(2).operacion = mml_FRASE0666
                aValores(2).valor = iMaxCod - iMinCod
                
                aValores(0).Nombre = "codigo"
                aValores(0).operacion = mml_FRASE0666
                aValores(0).valor = iMaxCod - iMinCod
                
            ImportarDatosConControl db1, "parejas", "SELECT * FROM parejas WHERE cod_competicion = " & iCodcomp, aValores
             'Sumamos los mismo a cod_juez y al código del juez
                'Localizamos el incremento del código de juez
                Set rs = db1.OpenRecordset("SELECT MIN(codigo) FROM jueces WHERE codigo IN (SELECT DISTINCT cod_juez FROM Juez_Categ WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & "))", dbOpenSnapshot)
                If IsNull(rs.Fields(0)) Then
                    MsgBox "No hay jueces asignados a las categorías.", vbOKOnly Or vbCritical, "ERROR FATAL"
                Else
                    iMinCod = rs.Fields(0)
                    rs.Close
                    iMaxCod = MaxCod("jueces")
                    
                    aValores(4).Nombre = "cod_juez"
                    aValores(4).operacion = mml_FRASE0666
                    aValores(4).valor = iMaxCod - iMinCod
                    
                    aValores(0).Nombre = "codigo"
                    aValores(0).operacion = mml_FRASE0666
                    aValores(0).valor = iMaxCod - iMinCod
                    
                ImportarDatosConControl db1, "Jueces", "SELECT * FROM jueces WHERE codigo IN (SELECT DISTINCT cod_juez FROM Juez_Categ WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & "))", aValores
                End If
           ' A partir de ahora el código debe ser el máximo de la tabla
            aValores(0).operacion = "MaxCod"
            ImportarDatosConControl db1, "Juez_Categ", "SELECT * FROM Juez_Categ WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            aValores(4).Nombre = "cod_juez_anulado"
            ImportarDatosConControl db1, "resultadosfinales", "SELECT * FROM resultadosfinales WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "agrupaciones", "SELECT * FROM agrupaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "enlaceprobaile", "SELECT * FROM enlaceprobaile WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "resumenfinales", "SELECT * FROM resumenfinales WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "dorsalescombinados", "SELECT * FROM dorsalescombinados WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "bailes_Categ", "SELECT * FROM bailes_Categ WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "horario", "SELECT * FROM horario WHERE cod_competicion = " & iCodcomp, aValores
            ImportarDatosConControl db1, "cal_baile", "SELECT * FROM cal_baile WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "cal_conjunto", "SELECT * FROM cal_conjunto WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "dorsales", "SELECT * FROM dorsales dor WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "hojas_reconocidas", "SELECT * FROM hojas_reconocidas WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "descalificaciones", "SELECT * FROM descalificaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            ImportarDatosConControl db1, "puntuaciones", "SELECT * FROM puntuaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")", aValores
            
            MsgBox mml_FRASE0134, vbOKOnly Or vbInformation, mml_FRASE0086
        End If
    End If
    Exit Sub
error:
Dim Msj As String
   If Err.Number <> 0 Then
        Msj = "Error # " & Str(Err.Number) & sComentario & mml_FRASE0208 _
              & Err.Source & Chr(13) & Err.Description & mml_FRASE0668 & sTabla
        MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
   End If
End Sub



Private Sub cmdSalir_Click()
On Local Error Resume Next
    db1.Close
    Unload Me
End Sub

Private Sub cmdSelecBD_Click()
Dim rs As Recordset
    If Not C_DEBUG Then On Local Error GoTo error
    DlgFicheros.ShowOpen
    If DlgFicheros.FileName <> "" Then
        Set db1 = OpenDatabase(DlgFicheros.FileName)
        Set rs = db1.OpenRecordset("SELECT * FROM competiciones ORDER BY codigo")
        lstComp.Clear
        While Not rs.EOF
            lstComp.AddItem rs!codigo & vbTab & rs!DESCRIPCION
            rs.MoveNext
        Wend
    End If
    Exit Sub
error:
    ProcesarError
End Sub

Sub ImportarDatos(sTabla As String, sSQL As String)
Dim rs As Recordset, rs1 As Recordset
    Set rs1 = db1.OpenRecordset(sSQL, dbOpenSnapshot)
    Set rs = db.OpenRecordset(sTabla, dbOpenTable)
        While Not rs1.EOF
            rs.AddNew
                For i = 0 To rs.Fields.Count - 1
                    rs.Fields(i) = rs1.Fields(i)
                Next
            rs.Update
            rs1.MoveNext
        Wend
    rs.Close
    rs1.Close
End Sub


Private Sub Form_Load()
    TraducirCadenas Me

End Sub
