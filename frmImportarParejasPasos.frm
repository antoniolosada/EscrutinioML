VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportarParejasPasos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE1278"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbLog 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4020
      Width           =   10275
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "Info. Formato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "mml_FRASE1278"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2580
      TabIndex        =   6
      Top             =   3180
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10275
      Begin VB.CheckBox chkOLa 
         Caption         =   "mml_FRASE1283"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   420
         TabIndex        =   5
         Top             =   2100
         Value           =   1  'Checked
         Width           =   9375
      End
      Begin VB.CheckBox chkOSt 
         Caption         =   "mml_FRASE1282"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   420
         TabIndex        =   4
         Top             =   1680
         Value           =   1  'Checked
         Width           =   9375
      End
      Begin VB.CheckBox chkGal 
         Caption         =   "mml_FRASE1281"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   420
         TabIndex        =   3
         Top             =   1260
         Value           =   1  'Checked
         Width           =   9375
      End
      Begin VB.CheckBox chkLat 
         Caption         =   "mml_FRASE1280"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   420
         TabIndex        =   2
         Top             =   840
         Value           =   1  'Checked
         Width           =   9375
      End
      Begin VB.CheckBox chkEst 
         Caption         =   "mml_FRASE1279"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   420
         TabIndex        =   1
         Top             =   420
         Value           =   1  'Checked
         Width           =   9375
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8940
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImportarParejasPasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAX_COLS = 28
Dim sCol(MAX_COLS) As String
Dim pCol(MAX_COLS) As Integer



Private Sub cmdImportar_Click()
Const NO_BAILO = "NO BAILO"
Dim categoriass As String
Dim categoriasl As String

Dim categoriasg As String

Dim categoriasostd As String
Dim categoriasolat As String

Dim nivels As String
Dim nivell As String

Dim bailesg As String

Dim iFile As Integer
Dim aCad() As String

Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim rs As Recordset

    
    'Inicializamos las columnas
    sCol(0) = ""
    sCol(1) = "categoriass"
    sCol(2) = "categoriasl"
    sCol(3) = "categoriasg"
    sCol(4) = "categoriasostd"
    sCol(5) = "categoriasolat"
    sCol(6) = "nivels"
    sCol(7) = "nivell"
    sCol(8) = "bailesg"
    sCol(9) = "bailesostd"
    sCol(10) = "bailesolat"
    sCol(11) = "nombreh"
    sCol(12) = "nombrem"
    sCol(13) = "apellidos1h"
    sCol(14) = "apellidos1m"
    sCol(15) = "apellidos2h"
    sCol(16) = "apellidos2m"
    sCol(17) = "poblacionh"
    sCol(18) = "poblacionm"
    sCol(19) = "telefonoh"
    sCol(20) = "telefonom"
    sCol(21) = "emailh"
    sCol(22) = "emailm"
    sCol(23) = "escuelah"
    sCol(24) = "escuelam"
    sCol(25) = "fechai"
    sCol(26) = "codpareja"
    sCol(27) = "codchico"
    sCol(28) = "codchica"
    
    If CodCompActiva = 0 Then
        Exit Sub
    End If
    If MsgBox("¿Quiere borrar todas las parejas de la competición?", vbYesNo Or vbQuestion, "AVISO") = vbYes Then
        'Comprobamos que no se hayan generado dorsales, ni categorias
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM categorias WHERE cod_competicion = " & CodCompActiva, dbOpenSnapshot)
        If rs.Fields(0) > 0 Then
            MsgBox "No puede eliminar las parejas, porque esta competición tiene grupos y dorsales asignados", vbOKOnly Or vbCritical, "ERROR"
            Exit Sub
        End If
        rs.Close
        db.Execute "DELETE FROM parejas WHERE cod_competicion = " & CodCompActiva
    End If
    
    CD.ShowOpen
    
    i = 1
    If CD.FileName <> "" Then
        iFile = FreeFile
        Open CD.FileName For Input As #iFile
        
        'Recuperamos la cabecera
        Line Input #iFile, sCad
        aCad = Split(sCad, ",")
        aCad(0) = ""
        'Recuperamos las posiciones de cada una de las columnas
        For j = 1 To UBound(aCad)
            For k = 1 To UBound(sCol)
                If SinComillas(aCad(j)) = sCol(k) Then
                    pCol(k) = j
                End If
            Next
        Next
        
        'Recuperamos las columnas no encontradas
        Dim sTmp As String
        sTmp = ""
        For j = 1 To UBound(pCol)
            If pCol(j) = 0 Then
                If sTmp <> "" Then sTmp = sTmp & ", "
                sTmp = sTmp & sCol(j)
            End If
        Next
        
        If sTmp <> "" Then
            MsgBox "Las siguientes columnas no se encontraron en el fichero: " & sTmp, vbOKOnly Or vbCritical, "ERROR"
        End If
        
        If EOF(iFile) Then
            MsgBox "Formato erróneo", vbOKOnly Or vbCritical, "ERROR"
            Exit Sub
            Close #iFile
        End If
        
        
        While Not EOF(iFile)
            Line Input #iFile, sCad
            aCad = Split(sCad, ",")
            aCad(0) = ""
            
            categoriass = SinComillas(aCad(pCol(1)))
            categoriasl = SinComillas(aCad(pCol(2)))
            categoriasg = SinComillas(aCad(pCol(3)))
            categoriasostd = SinComillas(aCad(pCol(4)))
            categoriasolat = SinComillas(aCad(pCol(5)))

            nivels = SinComillas(aCad(pCol(6)))
            nivell = SinComillas(aCad(pCol(7)))
            
            bailesg = SinComillas(aCad(pCol(8)))
            
            bailesostd = SinComillas(aCad(pCol(9)))
            bailesolat = SinComillas(aCad(pCol(10)))
            
            Dim res As Boolean
            res = True
            'Comprobamos si bailo estandar
            If categoriass <> NO_BAILO And Trim(categoriass) <> "" And chkEst.Value = 1 Then
                res = res And InsertarPareja(aCad, "ESTANDAR", categoriass, nivels)
            End If
            
            'Comprobamos latinos
            If categoriasl <> NO_BAILO And Trim(categoriasl) <> "" And chkLat.Value = 1 Then
                res = res And InsertarPareja(aCad, "LATINOS", categoriasl, nivell)
            End If
            
            'Comprobamos campeonato gallego
            If categoriasg <> NO_BAILO And Trim(categoriasg) <> "" And chkGal.Value = 1 Then
                res = res And InsertarPareja(aCad, "GALLEGO", categoriasg, bailesg)
            End If
            
            'Comprobamos open estandar
            If categoriasostd <> NO_BAILO And Trim(categoriasostd) <> "" And chkOSt.Value = 1 Then
                res = res And InsertarPareja(aCad, "OPEN_ESTANDAR", categoriasostd, "")
            End If
            
            'Comprobamos open latinos
            If categoriasolat <> NO_BAILO And Trim(categoriasolat) <> "" And chkOLa.Value = 1 Then
                res = res And InsertarPareja(aCad, "OPEN_LATINOS", categoriasolat, "")
            End If
            
            If Not res Then
                If MsgBox("¿CONTINUAR?" & vbCrLf & "Error cargando la línea: " & sCad, vbYesNo Or vbCritical, "ERROR") = vbNo Then
                    Exit Sub
                End If
            End If
            
        Wend
        Close #iFile
    End If
    
    MsgBox "Parejas importadas", vbOKOnly Or vbInformation, ""
    tbLog.Text = tbLog.Text & "Parejas importadas"

End Sub

Function InsertarPareja(aCad() As String, sMod As String, sGrupoEdad As String, sNivelp As String) As Boolean

Dim nivels As String
Dim nivell As String

Dim bailesg As String

Dim bailesostd As String
Dim bailesolat As String

Dim nombreh As String
Dim nombrem As String

Dim apellidos1h As String
Dim apellidos1m As String

Dim apellidos2h As String
Dim apellidos2m As String

Dim poblacionh As String
Dim poblacionm As String

Dim telefonoh As String
Dim telefonom As String

Dim emailh As String
Dim emailm As String

Dim escuelah As String
Dim escuelam As String

Dim fechai As String

Dim codpareja As String

Dim codchico As String
Dim codchica As String

Dim i As Integer

Dim iModalidad As Integer
Dim sSQL As String

Dim rs As Recordset
Dim lCodPareja As Long

Dim sNivel As String
    
    InsertarPareja = True
    If Not C_DEBUG Then On Local Error GoTo error
    
    Select Case sMod
        Case "ESTANDAR"
            iModalidad = COD_MODALIDAD_STD
            sNivel = DescCategoria(sNivelp)
        Case "LATINOS"
            iModalidad = COD_MODALIDAD_LAT
            sNivel = DescCategoria(sNivelp)
        Case "OPEN_ESTANDAR"
            iModalidad = COD_MODALIDAD_STD
            sNivel = "OPEN"
        Case "OPEN_LATINOS"
            iModalidad = COD_MODALIDAD_LAT
            sNivel = "OPEN"
        Case "GALLEGO"
            sNivel = sNivelp
            Select Case UCase(sNivel)
                Case "10 BAILES"
                    iModalidad = COD_MODALIDAD_COM
                Case "LATINO"
                    iModalidad = COD_MODALIDAD_LAT
                Case "STANDARD"
                    iModalidad = COD_MODALIDAD_STD
            End Select
            sNivel = "GALEGO"
    End Select
    
    If sNivel = "" Then
        tbLog.Text = tbLog.Text & "Falta el nivel: " & sNivelp & " en la base de datos" & vbCrLf & vbCrLf
        InsertarPareja = False
        Exit Function
    End If
    
    i = 1
    
    nombreh = SinComillas(aCad(pCol(11)))
    nombrem = SinComillas(aCad(pCol(12)))
    
    apellidos1h = SinComillas(aCad(pCol(13)))
    apellidos1m = SinComillas(aCad(pCol(14)))
    
    apellidos2h = SinComillas(aCad(pCol(15)))
    apellidos2m = SinComillas(aCad(pCol(16)))
    
    poblacionh = SinComillas(aCad(pCol(17)))
    poblacionm = SinComillas(aCad(pCol(18)))
    
    
    telefonoh = SinComillas(aCad(pCol(19)))
    telefonom = SinComillas(aCad(pCol(20)))
    
    emailh = SinComillas(aCad(pCol(21)))
    emailm = SinComillas(aCad(pCol(22)))
    
    escuelah = SinComillas(aCad(pCol(23)))
    escuelam = SinComillas(aCad(pCol(24)))
    
    fechai = SinComillas(aCad(pCol(25)))
    
    codpareja = SinComillas(aCad(pCol(26)))
    
    codchico = SinComillas(aCad(pCol(27)))
    codchica = SinComillas(aCad(pCol(28)))
    
    If CodigoGrupoEdad(sGrupoEdad) = 0 Then
        tbLog.Text = tbLog.Text & "Pareja: " & NombreParejasPasos(G_FORMATO_PAREJAS_PASOS, nombreh, apellidos1h, apellidos2h) & ", Grupo de edad no encontrado en la base de datos: " & sGrupoEdad & vbCrLf & vbCrLf
    End If
    
    sSQL = "'','" & NombreParejasPasos(G_FORMATO_PAREJAS_PASOS, nombreh, apellidos1h, apellidos2h) & "','','" & codchico & "','','','" & NombreParejasPasos(G_FORMATO_PAREJAS_PASOS, nombrem, apellidos1m, apellidos2m) & "','','" & codchica & "','','" & Combinar(poblacionh, poblacionm) & "','" & Combinar(telefonoh, telefonom) & "'," & CodCompActiva & ",'" & sGrupoEdad & "','" & Combinar(escuelah, escuelam) & "'," & CodigoGrupoEdad(sGrupoEdad) & "," & iModalidad & ",'" & sNivel & "','" & Combinar(poblacionh, poblacionm) & "',0,'" & Combinar(emailh, emailm) & "','',0,0,0,'" & Combinar(telefonoh, telefonom) & "',0,0," & Val(codpareja)
    
    Set rs = db.OpenRecordset("SELECT MAX(codigo) FROM parejas")
    lCodPareja = rs.Fields(0) + 1
    rs.Close
    sSQL = "INSERT INTO parejas VALUES (" & lCodPareja & "," & sSQL & ")"
    Debug.Print sSQL
    
    db.Execute sSQL

    Exit Function
error:
    ProcesarError "Insertando pareja:" & nombreh & ", " & nombrem
    InsertarPareja = False
End Function

Private Sub cmdInfo_Click()
    MsgBox "El fichero debe estar en formato ANSI, con retorno de carro y avance de línea como terminadores de cada línea." & vbCrLf & _
           "El primer campo será obligatoriamente Submitted. El resto de las columnas es opcional y puede encontrarse no ordenado" & vbCrLf & _
           "Las columnas deben denominarse:  categoriass, categoriasl, categoriasg, categoriasostd, categoriasolat, nivels, nivell, bailesg, bailesostd, bailesolat, nombreh, nombrem, apellidos1h, apellidos1m, apellidos2h, apellidos2m, poblacionh, poblacionm, telefonoh, telefonom, emailh, emailm, escuelah, escuelam, fechai, codpareja, codchico, codchica", _
           vbOKOnly Or vbInformation

End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub

Function Combinar(sCad1 As String, sCad2 As String) As String
    If Trim(sCad1) = "" Then
        sCad1 = Trim(sCad2)
    ElseIf Trim(sCad2) = "" Then
        sCad2 = Trim(sCad1)
    End If
    Combinar = sCad1 + " - " + sCad2
End Function


Function DescCategoria(sNivel As String) As String
    Dim rs As Recordset
    
    DescCategoria = ""
    Set rs = db.OpenRecordset("SELECT descripcion FROM desccategoria WHERE desc_importacion = '" + sNivel + "'", dbOpenSnapshot)
    If Not rs.EOF Then
        DescCategoria = rs.Fields("descripcion")
    End If
    rs.Close
End Function

Function SinComillas(sCad As String) As String
    SinComillas = SinCar(sCad, """")
End Function
Function SinComas(sCad As String) As String
    SinComas = SinCar(sCad, ",")
End Function
Function SinCar(sCad As String, sCar As String) As String
    Dim i As Integer
    Dim C As String
    
    SinCar = ""
    For i = 1 To Len(sCad)
        C = Mid(sCad, i, 1)
        If C <> sCar Then
            SinCar = SinCar + C
        End If
    Next
End Function

Function NombreParejasPasos(formato As String, Nombre As String, apellido1 As String, apellido2 As String) As String
    Dim i As Integer
    Dim C As String
    
    For i = 1 To Len(formato)
        C = Mid(formato, i, 1)
        Select Case C
            Case "1"
                NombreParejasPasos = NombreParejasPasos & PrimeraMayusculas(apellido1)
            Case "2"
                NombreParejasPasos = NombreParejasPasos & PrimeraMayusculas(apellido2)
            Case "N"
                NombreParejasPasos = NombreParejasPasos & PrimeraMayusculas(Nombre)
            Case Else
                NombreParejasPasos = NombreParejasPasos & C
        End Select
    Next
End Function

Function PrimeraMayusculas(sCad As String) As String
    PrimeraMayusculas = UCase(Left(sCad, 1)) & LCase(Mid(sCad, 2))
End Function
