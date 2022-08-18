VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGenerarTablaAEBDC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1247"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5835
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   4920
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "mml_FRASE0288"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.ListBox lstMod 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmGenerarTablaAEBDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerar_Click()
    Dim rs As Recordset
    Dim rsCont As Recordset
    Dim i As Integer
    Dim sModalidades As String
    Dim sGrupoEdad As String
    Dim sSQL As String
    Dim sCateg(20) As String
    Dim iCateg As Integer
    
    If Not C_DEBUG Then On Local Error GoTo error
    
    i = lstMod.ListIndex
    If i >= 0 Then
        sModalidades = lstMod.ItemData(i)
    
        If sModalidades = "" Then
            MsgBox mml_FRASE1248, vbOKOnly Or vbCritical, "ERROR"
            Exit Sub
        Else
            Dim iFile As Integer
            
            CD.FileName = lstMod.List(i) & ".CSV"
            CD.ShowSave
            
            iFile = FreeFile
            
            'Recuperamos la cabecera de categorias
            sSQL = "SELECT * from desccategoria where not orden is null ORDER BY orden"
            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
            While Not rs.EOF
                sCateg(iCateg) = rs.Fields("descripcion")
                Inc iCateg
                rs.MoveNext
            Wend
            
            sSQL = "SELECT COUNT(*) as contador, grupoedad, ge.orden, p.categoria, d.orden FROM parejas p, desccategoria d, gruposedad ge WHERE ge.codigo = cod_grupoedad AND d.descripcion = p.categoria AND " & _
                       " cod_modalidad = " & sModalidades & " AND cod_competicion = " & CodCompActiva & " AND not ge.orden IS NULL AND NOT d.orden IS NULL GROUP BY grupoedad, ge.orden, p.categoria, d.orden ORDER BY ge.orden, d.orden"
            Debug.Print sSQL
            Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)
            If Not rs.EOF Then
                Open CD.FileName For Output As #iFile
                While Not rs.EOF
                    If sGrupoEdad <> rs.Fields("grupoedad") Then
                        'Si es el primer grupo imprimimos las cabeceras
                        If sGrupoEdad = "" Then
                            Print #iFile, ";";
                            For i = 0 To iCateg - 1
                                Print #iFile, sCateg(i) & ";";
                            Next
                        Else
                            While i < iCateg
                                Print #iFile, "0;";
                                Inc i
                            Wend
                        End If
                        i = 0
                        sGrupoEdad = rs.Fields("grupoedad")
                        Print #iFile, ""
                        Print #iFile, sGrupoEdad & ";";
                    End If
                    While sCateg(i) <> rs.Fields("categoria")
                        Print #iFile, "0;";
                        Inc i
                    Wend
                    Print #iFile, rs.Fields("contador") & ";";
                    Inc i
                    
                    rs.MoveNext
                Wend
                While i < iCateg
                    Print #iFile, "0;";
                    Inc i
                Wend
                Close #iFile
            End If
            rs.Close
            
            ShellExecute Me.hwnd, "Open", CD.FileName, "", "", 1
            
            MsgBox mml_FRASE0545, vbOKOnly Or vbInformation, ""
        End If
    End If
    Exit Sub
error:
    ProcesarError "frmGenerarTablaAEBDC"
    
End Sub

Private Sub Form_Load()
    Dim rs As Recordset
    Dim i As Integer
    
    TraducirCadenas Me
    
    Set rs = db.OpenRecordset("SELECT codigo, nombre FROM modalidad WHERE codigo IN (SELECT DISTINCT cod_modalidad FROM parejas WHERE cod_competicion = " & CodCompActiva & ") ORDER BY 2", dbOpenSnapshot)
    While Not rs.EOF
        lstMod.AddItem rs.Fields("nombre")
        lstMod.ItemData(i) = rs.Fields("codigo")
        rs.MoveNext
        Inc i
    Wend
    rs.Close
End Sub

