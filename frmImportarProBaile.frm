VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportarProBaile 
   Caption         =   "mml_FRASE0661"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportarPuestos 
      Caption         =   "mml_FRASE1080"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4455
      MaskColor       =   &H80000010&
      TabIndex        =   5
      Top             =   4725
      Width           =   4110
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
      Width           =   6345
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
      Height          =   420
      Left            =   225
      TabIndex        =   1
      Top             =   4725
      Width           =   4110
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
Attribute VB_Name = "frmImportarProBaile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim gl_sFichero As String

Function iCodGrupoEdadProbaile(iCodigo As Integer) As Integer
    iCodGrupoEdadProbaile = iCodigo + 1
End Function
Function sGrupoEdadProbaile(iCodigo As Integer) As String
    Select Case iCodigo
        Case 0
            sGrupoEdadProbaile = "Juvenil"
        Case 1
            sGrupoEdadProbaile = "Junior I"
        Case 2
            sGrupoEdadProbaile = "Junior II"
        Case 3
            sGrupoEdadProbaile = "Youth"
        Case 4
            sGrupoEdadProbaile = "Adulto I"
        Case 5
            sGrupoEdadProbaile = "Adulto II"
        Case 6
            sGrupoEdadProbaile = "Senior I"
        Case 7
            sGrupoEdadProbaile = "Senior II"
        Case 8
            sGrupoEdadProbaile = "Senior III"
    End Select
End Function
Function sCategoriaProbaile(iCodigo As Integer) As String
    sCategoriaProbaile = Chr$(Asc("A") + iCodigo)
End Function

Private Sub cmdExportarPuestos_Click()
Dim rs As Recordset, rs1 As Recordset
Dim db1 As Database
Dim iCodComp As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    
    iCodComp = Val(lstComp.List(lstComp.ListIndex))
    If lstComp.ListIndex = -1 Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbExclamation, mml_FRASE0096
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0664, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        Set rs = db.OpenRecordset("SELECT DISTINCT cod_competicion_probaile FROM enlaceprobaile WHERE cod_competicion = " & CodCompActiva, dbOpenSnapshot)
        If rs.EOF Then
            MsgBox mml_FRASE1083, vbOKOnly Or vbCritical, G_MSG_ERROR
            rs.Close
            Exit Sub
        Else
            rs.MoveLast
            If rs.RecordCount > 1 Or IsNull(rs.Fields("cod_competicion_probaile")) Then
                MsgBox mml_FRASE1083, vbOKOnly Or vbCritical, G_MSG_ERROR
                rs.Close
                Exit Sub
            End If
            If Val(rs.Fields("cod_competicion_probaile")) = 0 Or Val(rs.Fields("cod_competicion_probaile")) <> iCodComp Then
                rs.Close
                MsgBox mml_FRASE1083, vbOKOnly Or vbCritical, G_MSG_ERROR
                Exit Sub
            End If
        End If
        rs.Close
        
        'Tenemos que localizar todos los dorsales importados de probaile con sus parejas
        'Localizamos la categoria en la que bailan las parejas y su posicion y puntos
        'Actualizamos entradas de probaile según la categoría sea de estandar, latinos o open o open1
        
        Set db1 = OpenDatabase(gl_sFichero)
        'Tenemos que localizar todos los dorsales importados de probaile con sus parejas
        Set rs = db.OpenRecordset("SELECT * FROM enlaceprobaile WHERE cod_competicion = " & CodCompActiva, dbOpenSnapshot)
        While Not rs.EOF
            'Localizamos la categoria en la que bailan las parejas siempre que coincida su dorsal importado y su posicion y puntos
            sSQL = "SELECT DISTINCT posicion, puntos, c.descripcion, c.cod_modalidad FROM categorias c, dorsales d, resumenfinales rs WHERE d.cod_categoria = c.codigo AND d.cod_categoria = rs.cod_categoria AND d.num_dorsal = rs.dorsal AND d.cod_pareja = " & rs.Fields("cod_pareja") & " AND d.num_dorsal = " & rs.Fields("dorsal_probaile") & " AND d.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ")"
            Debug.Print sSQL
            Set rs1 = db.OpenRecordset(sSQL, dbOpenSnapshot)
            If rs1.EOF Then
                MsgBox mml_FRASE1090 & "Number/Dorsal: " & rs.Fields("dorsal_probaile") & " - Cod. Pareja/Couple: " & ", Cod.IRIS: " & rs.Fields("cod_pareja") & ", Cod.ProBaile: " & rs.Fields("cod_pareja_probaile"), vbOKOnly Or vbCritical, G_MSG_ERROR
            Else
                rs1.MoveLast
                If rs1.RecordCount > 1 Then
                    MsgBox mml_FRASE1085 & "Number/Dorsal: " & rs.Fields("dorsal_probaile") & " - Cod. Pareja/Couple: " & ", Cod.IRIS: " & rs.Fields("cod_pareja") & ", Cod.ProBaile: " & rs.Fields("cod_pareja_probaile"), vbOKOnly Or vbCritical, G_MSG_ERROR
                Else
                    'Actualizamos la información en probaile
                    If InStr(UCase(rs1.Fields("descripcion")), "OPEN") > 0 Then 'open
                    ElseIf rs1.Fields("cod_modalidad") = 1 Then ' estandar
                        db1.Execute "UPDATE entradas SET Std_Clasificacion = " & Val(rs1.Fields("posicion")) & ",Std_Puntuacion = " & Val(rs1.Fields("puntos")) & " WHERE competicion = " & iCodComp & _
                        " AND pareja = " & rs.Fields("cod_pareja_probaile") & " AND dorsal = " & rs.Fields("dorsal_probaile")
                    ElseIf rs1.Fields("cod_modalidad") = 2 Then ' latinos
                        db1.Execute "UPDATE entradas SET Lat_Clasificacion = " & Val(rs1.Fields("posicion")) & ",Lat_Puntuacion = " & Val(rs1.Fields("puntos")) & " WHERE competicion = " & iCodComp & _
                        " AND pareja = " & rs.Fields("cod_pareja_probaile") & " AND dorsal = " & rs.Fields("dorsal_probaile")
                    End If
                End If
            End If
            rs1.Close
            rs.MoveNext
        Wend
        db1.Close
        rs.Close
        
        'Por ultimo localizamos los dorsales que están en resumenfinales y no están en enlaceprobaile
        Dim sDorsales As String
        sDorsales = ""
        Set rs = db.OpenRecordset("SELECT DISTINCT dorsal FROM resumenfinales WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ") AND NOT dorsal IN (SELECT dorsal_probaile FROM enlaceprobaile WHERE cod_competicion = " & CodCompActiva & ")", dbOpenSnapshot)
        While Not rs.EOF
            sDorsales = sDorsales & " " & rs.Fields("dorsal")
            rs.MoveNext
        Wend
        If sDorsales <> "" Then
            MsgBox mml_FRASE1093 & ": " & sDorsales, vbOKOnly Or vbInformation, G_MSG_AVISO
        End If
        rs.Close
        MsgBox mml_FRASE1009, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
    Exit Sub
error:
    ProcesarError "cmdExportarPuestos_Click"
End Sub

Private Sub cmdImportar_Click()
Dim rs As Recordset, rs1 As Recordset
Dim db2 As Database
Dim iCodComp As Integer
Dim iGrupoProbaile  As Integer
Dim iCodCategoriaProbaile As Integer
Dim sSQL As String
Dim lCodPareja As Long
Dim sDorsal As String
Dim sUltimo As String

    If Not C_DEBUG Then On Local Error GoTo error
    
    iCodComp = CodCompActiva
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM parejas WHERE cod_competicion = " & CodCompActiva, dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        MsgBox mml_FRASE1051, vbOKOnly Or vbCritical, G_MSG_ERROR
        Exit Sub
    End If
    rs.Close
        
    iCodComp = Val(lstComp.List(lstComp.ListIndex))
    If lstComp.ListIndex = -1 Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbExclamation, mml_FRASE0096
        Exit Sub
    End If
    
    If MsgBox(mml_FRASE0664, vbYesNo Or vbQuestion, mml_FRASE0084) = vbYes Then
        db.Execute ("DELETE FROM enlaceprobaile WHERE cod_competicion = " & CodCompActiva)
        Set db1 = OpenDatabase(gl_sFichero)
        Set db2 = OpenDatabase(sExtraerPath(gl_sFichero) & "\" & "MdbParejas.mdb")
        Set rs = db1.OpenRecordset("SELECT * FROM entradas WHERE competicion = " & iCodComp, dbOpenSnapshot)
        While Not rs.EOF
            Set rs1 = db2.OpenRecordset("SELECT * FROM parejas WHERE pareja = " & rs!pareja)
            If Not rs1.EOF Then
                If rs!standard < 6 Then
                    iModalidad = 1
                    iGrupoProbaile = rs!grupo
                    iCodCategoriaProbaile = rs!standard
                    lCodPareja = MaxCod("parejas")
                    sSQL = "INSERT INTO parejas VALUES (" & lCodPareja & ",'','" & sQuitarCarProhibidosSQL(rs1!Nombre) & _
                    " " & sQuitarCarProhibidosSQL(rs1!apellidos) & "','" & SinNulos(rs1!fecha) & "','','','','" & sQuitarCarProhibidosSQL(rs1!nombre1) & " " & sQuitarCarProhibidosSQL(rs1!apellidos1) & _
                    "','" & SinNulos(rs1!fecha1) & "','','','" & sQuitarCarProhibidosSQL(rs1!direccion) & "-" & sQuitarCarProhibidosSQL(rs1!poblacion) & ", " & sQuitarCarProhibidosSQL(rs1!direccion1) & "-" & sQuitarCarProhibidosSQL(rs1!poblacion1) & _
                    "','" & SinNulos(rs1!telefono) & ", " & SinNulos(rs1!movil) & " - " & SinNulos(rs1!telefono1) & ", " & SinNulos(rs1!movil1) & "'," & _
                    CodCompActiva & ",'" & sGrupoEdadProbaile(iGrupoProbaile) & "',''," & _
                    iCodGrupoEdadProbaile(iGrupoProbaile) & "," & iModalidad & ",'" & _
                    sCategoriaProbaile(iCodCategoriaProbaile) & "','" & sQuitarCarProhibidosSQL(rs1!provincia) & "-" & _
                    sQuitarCarProhibidosSQL(rs1!provincia1) & "',0,'" & SinNulos(rs1!email) & "," & SinNulos(rs1!email1) & "','" & SinNulos(rs1!movil) & "',0,0,0,''," & IIf(rs!pago = "N", 0, 1) & ",0)"
                    Debug.Print sSQL
                    db.Execute (sSQL)
                
                    'Insertamos el dorsal de la pareja
                    sDorsal = IIf(SinNulos(rs!dorsal) = "", "0", rs!dorsal)
                    sSQL = "INSERT INTO enlaceprobaile VALUES (" & CodCompActiva & "," & rs!pareja & "," & sDorsal & "," & lCodPareja & "," & iCodComp & ")"
                    Debug.Print sSQL
                    db.Execute (sSQL)
                End If
                If rs!latino < 6 Then
                    iModalidad = 2
                    iGrupoProbaile = rs!grupo
                    iCodCategoriaProbaile = rs!latino
                    
                    lCodPareja = MaxCod("parejas")
                    sSQL = "INSERT INTO parejas VALUES (" & lCodPareja & ",'','" & sQuitarCarProhibidosSQL(rs1!Nombre) & _
                    " " & sQuitarCarProhibidosSQL(rs1!apellidos) & "','" & SinNulos(rs1!fecha) & "','','','','" & sQuitarCarProhibidosSQL(rs1!nombre1) & " " & sQuitarCarProhibidosSQL(rs1!apellidos1) & _
                    "','" & SinNulos(rs1!fecha1) & "','','','" & sQuitarCarProhibidosSQL(rs1!direccion) & "-" & sQuitarCarProhibidosSQL(rs1!poblacion) & ", " & sQuitarCarProhibidosSQL(rs1!direccion1) & "-" & sQuitarCarProhibidosSQL(rs1!poblacion1) & _
                    "','" & SinNulos(rs1!telefono) & ", " & SinNulos(rs1!movil) & " - " & SinNulos(rs1!telefono1) & ", " & SinNulos(rs1!movil1) & "'," & _
                    CodCompActiva & ",'" & sGrupoEdadProbaile(iGrupoProbaile) & "',''," & _
                    iCodGrupoEdadProbaile(iGrupoProbaile) & "," & iModalidad & ",'" & _
                    sCategoriaProbaile(iCodCategoriaProbaile) & "','" & sQuitarCarProhibidosSQL(rs1!provincia) & "-" & _
                    sQuitarCarProhibidosSQL(rs1!provincia1) & "',0,'" & SinNulos(rs1!email) & "," & SinNulos(rs1!email1) & "','" & SinNulos(rs1!movil) & "',0,0,0,''," & IIf(rs!pago = "N", 0, 1) & ",0)"
                    Debug.Print sSQL
                    db.Execute (sSQL)
                
                    'Insertamos el dorsal de la pareja
                    sDorsal = IIf(SinNulos(rs!dorsal) = "", "0", rs!dorsal)
                    sSQL = "INSERT INTO enlaceprobaile VALUES (" & CodCompActiva & "," & rs!pareja & "," & sDorsal & "," & lCodPareja & "," & iCodComp & ")"
                    Debug.Print sSQL
                    db.Execute (sSQL)
                End If
                If rs!Opens > 0 Then
                    iModalidad = 3
                    iGrupoProbaile = rs!grupo
                    
                    lCodPareja = MaxCod("parejas")
                    sSQL = "INSERT INTO parejas VALUES (" & lCodPareja & ",'','" & sQuitarCarProhibidosSQL(rs1!Nombre) & _
                    " " & sQuitarCarProhibidosSQL(rs1!apellidos) & "','" & SinNulos(rs1!fecha) & "','','','','" & sQuitarCarProhibidosSQL(rs1!nombre1) & " " & sQuitarCarProhibidosSQL(rs1!apellidos1) & _
                    "','" & SinNulos(rs1!fecha1) & "','','','" & sQuitarCarProhibidosSQL(rs1!direccion) & "-" & sQuitarCarProhibidosSQL(rs1!poblacion) & ", " & sQuitarCarProhibidosSQL(rs1!direccion1) & "-" & sQuitarCarProhibidosSQL(rs1!poblacion1) & _
                    "','" & SinNulos(rs1!telefono) & ", " & SinNulos(rs1!movil) & " - " & SinNulos(rs1!telefono1) & ", " & SinNulos(rs1!movil1) & "'," & _
                    CodCompActiva & ",'" & sGrupoEdadProbaile(iGrupoProbaile) & "',''," & _
                    iCodGrupoEdadProbaile(iGrupoProbaile) & "," & iModalidad & ",'Open','" & sQuitarCarProhibidosSQL(rs1!provincia) & "-" & _
                    sQuitarCarProhibidosSQL(rs1!provincia1) & "',0,'" & SinNulos(rs1!email) & "," & SinNulos(rs1!email1) & "','" & SinNulos(rs1!movil) & "',0,0,0,''," & IIf(rs!pago = "N", 0, 1) & ",0)"
                    Debug.Print sSQL
                    db.Execute (sSQL)
                
                    'Insertamos el dorsal de la pareja
                    sDorsal = IIf(SinNulos(rs!dorsal1) = "", "0", rs!dorsal1)
                    sSQL = "INSERT INTO enlaceprobaile VALUES (" & CodCompActiva & "," & rs!pareja & "," & sDorsal & "," & lCodPareja & "," & iCodComp & ")"
                    Debug.Print sSQL
                    db.Execute (sSQL)
                End If
                If rs!opens1 > 0 Then
                    iModalidad = 3
                    iGrupoProbaile = rs!grupo
                    iCodCategoriaProbaile = rs!latino
                    
                    lCodPareja = MaxCod("parejas")
                    sSQL = "INSERT INTO parejas VALUES (" & lCodPareja & ",'','" & sQuitarCarProhibidosSQL(rs1!Nombre) & _
                    " " & sQuitarCarProhibidosSQL(rs1!apellidos) & "','" & SinNulos(rs1!fecha) & "','','','','" & sQuitarCarProhibidosSQL(rs1!nombre1) & " " & sQuitarCarProhibidosSQL(rs1!apellidos1) & _
                    "','" & SinNulos(rs1!fecha1) & "','','','" & sQuitarCarProhibidosSQL(rs1!direccion) & "-" & sQuitarCarProhibidosSQL(rs1!poblacion) & ", " & sQuitarCarProhibidosSQL(rs1!direccion1) & "-" & sQuitarCarProhibidosSQL(rs1!poblacion1) & _
                    "','" & SinNulos(rs1!telefono) & ", " & SinNulos(rs1!movil) & " - " & SinNulos(rs1!telefono1) & ", " & SinNulos(rs1!movil1) & "'," & _
                    CodCompActiva & ",'" & sGrupoEdadProbaile(iGrupoProbaile) & "',''," & _
                    iCodGrupoEdadProbaile(iGrupoProbaile) & "," & iModalidad & ",'Open1','" & sQuitarCarProhibidosSQL(rs1!provincia) & "-" & _
                    sQuitarCarProhibidosSQL(rs1!provincia1) & "',0,'" & SinNulos(rs1!email) & "," & SinNulos(rs1!email1) & "','" & SinNulos(rs1!movil) & "',0,0,0,''," & IIf(rs!pago = "N", 0, 1) & ",0)"
                    Debug.Print sSQL
                    db.Execute (sSQL)
                
                    'Insertamos el dorsal de la pareja
                    sDorsal = IIf(SinNulos(rs!dorsal1) = "", "0", rs!dorsal1)
                    sSQL = "INSERT INTO enlaceprobaile VALUES (" & CodCompActiva & "," & rs!pareja & "," & sDorsal & "," & lCodPareja & "," & iCodComp & ")"
                    Debug.Print sSQL
                    db.Execute (sSQL)
                End If
            Else
                MsgBox mml_FRASE1052 & " " & rs!pareja, vbOKOnly Or vbCritical, G_MSG_ERROR
            End If
            
            sUltimo = lCodPareja & ",'','" & sQuitarCarProhibidosSQL(rs1!Nombre) & _
            " " & sQuitarCarProhibidosSQL(rs1!apellidos) & "','" & SinNulos(rs1!fecha) & "','','','','" & sQuitarCarProhibidosSQL(rs1!nombre1) & " " & sQuitarCarProhibidosSQL(rs1!apellidos1) & " - Dorsal: " & sDorsal
            
            rs1.Close
            rs.MoveNext
        Wend
        rs.Close
        db1.Close
        MsgBox mml_FRASE0134, vbOKOnly Or vbInformation, mml_FRASE0086
    End If
    Exit Sub
error:
Dim Msj As String
   If Err.Number <> 0 Then
        Msj = "Error # " & Str(Err.Number) & sComentario & mml_FRASE0208 _
              & Err.Source & Chr(13) & Err.Description & mml_FRASE0668 & sTabla & vbCrLf & _
              "Último Registro/Last Record: " & sUltimo
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

    On Local Error GoTo error
    If Not C_DEBUG Then On Local Error GoTo error
    DlgFicheros.ShowOpen
    If DlgFicheros.FileName <> "" Then
        gl_sFichero = DlgFicheros.FileName
        Set db1 = OpenDatabase(DlgFicheros.FileName)
        Set rs = db1.OpenRecordset("SELECT codigo,descripcion,fecha,cuota FROM competiciones ORDER BY codigo")
        lstComp.Clear
        While Not rs.EOF
            lstComp.AddItem rs!codigo & " - " & rs!DESCRIPCION
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
