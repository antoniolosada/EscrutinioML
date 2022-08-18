VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimirInternet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0716"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerarZIP 
      Caption         =   "mml_FRASE1295"
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
      Left            =   7320
      TabIndex        =   15
      Top             =   5160
      Width           =   3405
   End
   Begin VB.CommandButton cmdGenerarTablaAEBDC 
      Caption         =   "mml_FRASE1247"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   14
      Top             =   4200
      Width           =   3405
   End
   Begin VB.Frame Frame 
      Height          =   5295
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   6735
      Begin VB.CheckBox chkSelecTodas 
         Caption         =   "mml_FRASE1245"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Value           =   1  'Checked
         Width           =   5055
      End
      Begin VB.ListBox lstCat 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   12
         Top             =   840
         Width           =   6375
      End
   End
   Begin VB.CommandButton cmdGenXMLAEBDC 
      Caption         =   "mml_FRASE1229"
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
      Left            =   7320
      TabIndex        =   10
      Top             =   3480
      Width           =   3405
   End
   Begin VB.CheckBox chkSoloResumen 
      Caption         =   "mml_FRASE1211"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      TabIndex        =   9
      Top             =   3000
      Width           =   3165
   End
   Begin VB.CommandButton cmdImpLibro 
      Caption         =   "mml_FRASE1000"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   2280
      Width           =   3405
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   1320
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11415
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
         Height          =   360
         Left            =   1845
         Picture         =   "frmImprimitInternet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   450
      End
      Begin VB.TextBox tbInfo 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   11055
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
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   8055
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
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   855
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "mml_FRASE0029"
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
      Left            =   7320
      TabIndex        =   1
      Top             =   6000
      Width           =   3405
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "mml_FRASE0717"
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
      Left            =   7320
      TabIndex        =   0
      Top             =   1680
      Width           =   3405
   End
End
Attribute VB_Name = "frmImprimirInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gl_bGenerarInet As Boolean


Private Sub chkSelecTodas_Click()
    If chkSelecTodas.Value = 1 Then
        lstCat.Enabled = False
    Else
        lstCat.Enabled = True
    End If
    
End Sub

Private Sub cmdGenerarZIP_Click()
    'Recuperamos el directorio de ejecución de IRIS
    'Comprobamos si 7zip.exe y 7zip.dll se encuentran en el directorio
    'Si el fichero destino de empaquetado existe lo borramos
    'Comprobamos que ha sido borrado
    'Empaquetamos los ficheros
    'Comprobamos que se genera el fichero destino de empaquetado
    
    Dim sDir As String
    Dim sDirFichas As String
    Dim rs As Recordset
    
    sDirFichas = VarCfg("dir_fichas")
    If Dir(G_DIR_7ZIP & "\7zip.bat") And Dir(G_DIR_7ZIP & "\7z.exe") <> "" And Dir(G_DIR_7ZIP & "\7z.dll") <> "" Then
        Shell "cmd /c " & G_DIR_7ZIP & "\7zip.bat """ & tbDescComp.Text & """ Inet"
        
    Else
        MsgBox "Los archivos 7zip.exe y 7zip.dll deben estar en el directorio (sin espacios): " & G_DIR_7ZIP, vbOKOnly Or vbCritical, ""
    End If
    
    
End Sub

Private Sub cmdGenXMLAEBDC_Click()
Dim rs As Recordset
Dim rs1 As Recordset
Dim iFile As Integer
Dim iCont As Integer
Dim sPosicion As String
Dim sPosicion1 As String
Dim iCodCateg As Integer
Dim iTotalParejas As Integer
Dim iBajas As Integer
Dim iRondas As Integer
Dim sCodigoPareja As String

    iCont = 0
    iCodCateg = 0
    
    CDialog.CancelError = True
    On Local Error Resume Next
    CDialog.FileName = "AEBDC"
    CDialog.DefaultExt = "XML"
    CDialog.ShowSave
    If Err.Number > 0 Then Exit Sub
    
    If Not C_DEBUG Then
        On Local Error GoTo error
    Else
        On Local Error GoTo 0
    End If
    
    iFile = FreeFile
    
    Open CDialog.FileName For Output As iFile
    
    Print #iFile, "<?xml version=""1.0"" encoding=""ISO-8859-1"" standalone=""yes""?>"
    Print #iFile, "<Event>"
    Print #iFile, "<xs:schema id=""Event"" xmlns="""" xmlns:xs = ""http://www.w3.org/2001/XMLSchema"" xmlns:msdata=""urn:schemas-microsoft-com:xml-msdata"">"
    Print #iFile, "<xs:element name=""Event"" msdata:IsDataSet=""true"" msdata:UseCurrentLocale=""true"">"
    Print #iFile, "<xs:complexType>"
    Print #iFile, "<xs:choice minOccurs=""0"" maxOccurs=""unbounded"">"
    Print #iFile, "<xs:element name=""Results"">"
    Print #iFile, "<xs:complexType>"
    Print #iFile, "<xs:sequence>"
    Print #iFile, "<xs:element name=""EventCode"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""CoupleCode"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""CoupleNumber"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""Names"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""CompCode"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""CompLevel"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""CompType"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""Position1"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""Position2"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""Missing"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""TotalCouples"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "<xs:element name=""RoundsDanced"" type=""xs:string"" minOccurs=""0"" />"
    Print #iFile, "</xs:sequence>"
    Print #iFile, "</xs:complexType>"
    Print #iFile, "</xs:element>"
    Print #iFile, "</xs:choice>"
    Print #iFile, "</xs:complexType>"
    Print #iFile, "</xs:element>"
    Print #iFile, "</xs:schema>"
    
    
    sSQL = "SELECT DISTINCT m.aebdc_codigo as comptype, g.aebdc_codigo as compcode, c.id_categoria as complevel, " & _
            "e.aebdc_codigo as eventcode, p.aebdc_codigo as couplecode, p.nombre_hombre, p.nombre_mujer, d.num_dorsal,  " & _
            "f.posicion, c.codigo as cod_categ " & _
            "FROM resumenfinales f, competiciones e, parejas p, categorias c, dorsales d, gruposedad g,modalidad m " & _
            "WHERE c.id_categoria IN ( " & G_CATEG_XML_AEBDC & " ) and  e.codigo = " & tbCodComp.Text & " and c.cod_competicion = e.codigo AND d.cod_categoria = c.codigo " & _
            "AND p.codigo = d.cod_pareja AND f.dorsal = d.num_dorsal AND " & _
            "g.codigo = c.cod_grupoedad AND f.cod_categoria = c.codigo AND m.codigo = c.cod_modalidad " & _
            "ORDER BY c.codigo"
            
    If C_DEBUG Then Debug.Print sSQL
    
    Set rs = db.OpenRecordset(sSQL, dbOpenSnapshot)


    While Not rs.EOF
        If rs.Fields("eventcode") = 0 Then
            MsgBox "El código de competición está a 0.", vbOKOnly Or vbCritical, "ERRO"
            rs.Close
            Exit Sub
        End If
    
        If iCodCateg <> rs.Fields("cod_categ") Then
            iCodCateg = rs.Fields("cod_categ")
            'Set rs1 = db.OpenRecordset("SELECT COUNT(*) as total FROM dorsales WHERE cod_categoria = " & iCodCateg & " AND no_presente = 0", dbOpenSnapshot)
            Set rs1 = db.OpenRecordset("SELECT COUNT(*) as total FROM dorsales WHERE cod_categoria = " & iCodCateg, dbOpenSnapshot)
            If Not rs1.EOF Then
                iTotalParejas = rs1.Fields("total")
            End If
            rs1.Close
        End If
        Inc iCont
        Dim i As Integer
        sPosicion = rs.Fields("posicion")
        If sPosicion = "NO PRESENTADO" Or Val(sPosicion) = 0 Then
            iBajas = 1
            'Comprobamos si lo comunicó antes de la competición
            Set rs1 = db.OpenRecordset("SELECT no_presente FROM dorsales d WHERE d.cod_categoria = " & iCodCateg & " and num_dorsal = " & rs.Fields("num_dorsal"), dbOpenSnapshot)
            If Not rs1.EOF Then
                iBajas = rs1.Fields("no_presente")
            End If
            rs1.Close
            sPosicion = 0
            sPosicion1 = 0
            iRondas = 0
        Else
            iBajas = 0
            i = InStr(sPosicion, "-")
            If i > 0 Then
                sPosicion1 = Mid(sPosicion, i + 1)
                sPosicion = Mid(sPosicion, 1, i - 1)
            Else
                sPosicion1 = sPosicion
            End If
            
            'Comprobamos las rondas que ha bailado
            Set rs1 = db.OpenRecordset("SELECT COUNT(*) AS rondas FROM dorsales d WHERE d.cod_categoria = " & iCodCateg & " and num_dorsal = " & rs.Fields("num_dorsal") & " AND no_presente = 0", dbOpenSnapshot)
            If Not rs1.EOF Then
                iRondas = rs1.Fields("rondas")
            End If
            rs1.Close
        End If
        If Val(SinNulos(rs.Fields("couplecode"))) = 0 Then
            If MsgBox("La pareja formada por " & QuitarAcentos(rs.Fields("nombre_hombre")) & " - " & QuitarAcentos(rs.Fields("nombre_mujer")) & " con dorsal nº" & rs.Fields("num_dorsal") & " no tiene código de pareja. Figurará en el fichero con el código vacío. ¿Quiere continuar?", vbYesNo Or vbQuestion, "AVISO") = vbNo Then
                rs.Close
                Exit Sub
            End If
        End If
        
        sCodigoPareja = SinNulos(rs.Fields("couplecode"))
        If Val(sCodigoPareja) = 0 Then sCodigoPareja = ""
        
        Print #iFile, "<Results>"
        Print #iFile, "<EventCode>" & rs.Fields("eventcode") & "</EventCode>"
        Print #iFile, "<CoupleCode>" & sCodigoPareja & "</CoupleCode>"
        Print #iFile, "<CoupleNumber>" & rs.Fields("num_dorsal") & "</CoupleNumber>"
        Print #iFile, "<Names>" & QuitarAcentos(rs.Fields("nombre_hombre")) & " - " & QuitarAcentos(rs.Fields("nombre_mujer")) & "</Names>"
        Print #iFile, "<CompCode>" & rs.Fields("compcode") & "</CompCode>"
        Print #iFile, "<CompLevel>" & rs.Fields("complevel") & "</CompLevel>"
        Print #iFile, "<CompType>" & rs.Fields("comptype") & "</CompType>"
        Print #iFile, "<Position1>" & sPosicion & "</Position1>"
        Print #iFile, "<Position2>" & sPosicion1 & "</Position2>"
        Print #iFile, "<Missing>" & iBajas & "</Missing>"
        Print #iFile, "<TotalCouples>" & iTotalParejas & "</TotalCouples>"
        Print #iFile, "<RoundsDanced>" & iRondas & "</RoundsDanced>"
        Print #iFile, "</Results>"
        rs.MoveNext
    Wend
    Print #iFile, "</Event>"
    Close iFile
    
    MsgBox mml_FRASE1009
    
    Exit Sub
error:
    ProcesarError "cmdGenXMLAEBDC_Click"
End Sub

Private Sub cmdImpLibro_Click()
    gl_bGenerarInet = False
    ImprimirHojasDatos

End Sub

Private Sub cmdImprimir_Click()
    gl_bGenerarInet = True
    frmImprimirInternet.Enabled = False
    ImprimirHojasDatos
    cmdSalir.Enabled = True
    frmImprimirInternet.Enabled = True
End Sub

Private Sub ImprimirHojasDatos()
Dim rs As Recordset
Dim rsCont As Recordset
Dim rsBailes As Recordset
Dim rsFases As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iNumDorsales As Integer
Dim sDirFichas  As String
Dim iFichInet As Integer
Dim sCatAnt As String
Dim sModalidad As String
Dim scolor As String
Dim bComienzo As Boolean
Dim dateHora As Date

    bComienzo = True
    sDirFichas = VarCfg("dir_fichas")

    If MsgBox(mml_FRASE0718, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0719, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    
    MsgBox mml_FRASE0720 & sDirFichas & "\Inet\Comp-Resultados.htm" & Chr$(13) & Chr$(10) & _
           mml_FRASE0721 & Chr$(13) & Chr$(10) & _
           mml_FRASE0960 & C_FICHERO_INET, vbOKOnly Or vbInformation, mml_FRASE0147
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    On Local Error Resume Next
    Kill C_FICHERO_INET
    If Not C_DEBUG Then On Local Error GoTo error
    frmImprimirInternet.cmdImprimir.Enabled = False
    
    On Local Error Resume Next
    Kill sDirFichas & "\Inet\*.*"
    If Not C_DEBUG Then On Local Error GoTo error
    
    iFichInet = FreeFile
    Open sDirFichas & "\Inet\Comp-Resultados.htm" For Output As #iFichInet
    
    iEscala = Printer.Width / 10
    Set rs = db.OpenRecordset(" SELECT * FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    sTexto = rs!DESCRIPCION
    
    Print #iFichInet, "<HTML>" & G_CAB_INET
    Print #iFichInet, "<h1>" & Trim$(sTexto) & "</h1>"
    Print #iFichInet, "<h2>" & Trim$(sEscuela(tbCodComp.Text)) & "</h2>"
    Print #iFichInet, "<h2>" & Format$(rs!fecha, "dd/mm/yyyy") & "</h2><br><br>"
    
    rs.Close
        
    If chkSelecTodas.Value = 1 Then
        If C_COUNTRY Then
            'ordenamos según la configuración (solo para country)
            Set rs = db.OpenRecordset(" SELECT c.codigo, hora, id_categoria, ge.nombre, c.descripcion, m.nombre as mod FROM categorias c, gruposedad ge, modalidad m, desccategoria dc WHERE dc.descripcion = c.id_categoria AND c.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY m.orden, dc.orden, ge.orden, c.descripcion DESC", dbOpenSnapshot)
        Else
            'ordenamos por descripción
            Set rs = db.OpenRecordset(" SELECT c.codigo, hora, id_categoria, ge.nombre, descripcion, m.nombre as mod FROM categorias c, gruposedad ge, modalidad m WHERE c.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY descripcion DESC", dbOpenSnapshot)
        End If
    Else
        Dim sCateg As String
        Dim i As Integer
        For i = 0 To lstCat.ListCount - 1
            If lstCat.Selected(i) Then
                If sCateg = "" Then
                    sCateg = lstCat.ItemData(i)
                Else
                    sCateg = sCateg & "," & lstCat.ItemData(i)
                End If
            End If
        Next
        If sCateg <> "" Then
            ' Generamos las categorias seleccionadas
            If C_COUNTRY Then
                'ordenamos según la configuración (solo para country)
                Set rs = db.OpenRecordset(" SELECT c.codigo, hora, id_categoria, ge.nombre, c.descripcion, m.nombre as mod FROM categorias c, gruposedad ge, modalidad m, desccategoria dc " & _
                        "WHERE dc.descripcion = c.id_categoria AND c.codigo IN (" & sCateg & ") AND c.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY m.orden, dc.orden, ge.orden, c.descripcion DESC", dbOpenSnapshot)
            Else
                'Ordenamos según la descripción
                Set rs = db.OpenRecordset(" SELECT c.codigo, hora, id_categoria, ge.nombre, descripcion, m.nombre as mod FROM categorias c, gruposedad ge, modalidad m " & _
                        "WHERE c.codigo IN (" & sCateg & ") AND c.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY descripcion DESC", dbOpenSnapshot)
            End If
        Else
            MsgBox mml_FRASE1246, vbOKOnly Or vbCritical, "ERROR"
            Exit Sub
        End If
    End If
    
    Print #iFichInet, G_CAB_TABLA
    While Not rs.EOF
        If sModalidad <> Left$(rs!DESCRIPCION, 3) Then
            If sModalidad <> "" Then Print #iFichInet, "<tr border = 0><td>&nbsp;</td></tr>"
            If scolor = C_COLOR_INET2 Then
                scolor = C_COLOR_INET1
            Else
                scolor = C_COLOR_INET2
            End If
            sModalidad = Left$(rs!DESCRIPCION, 3)
        End If
        Print #iFichInet, "<TR><td nowrap bgcolor='" & scolor & "'><font size=4><b>" & rs!DESCRIPCION & "</b></font></TD>"
        sCatAnt = ""
        Set rsFases = db.OpenRecordset("SELECT DISTINCT fase, repesca FROM puntuaciones WHERE cod_categoria = " & rs!codigo & " ORDER BY 1 DESC")
        While Not rsFases.EOF
            tbInfo.Text = rs!DESCRIPCION & mml_FRASE0723 & rsFases!fase
            frmImprimirFinal.tbCodComp = tbCodComp.Text
            frmImprimirFinal.tbCodCateg = rs!codigo
            frmImprimirFinal.tbCodFase = rsFases!fase
            sCatAnt = rs!codigo
            If rsFases!repesca = 1 Then
                tbInfo.Text = tbInfo.Text & mml_FRASE0724
            End If
            frmImprimirFinal.chkRep.Value = rsFases!repesca
            DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
            'If Not bComienzo Then SendKeys "~"
            If chkSoloResumen.Value = 0 Then
                bComienzo = False
                frmImprimirFinal.ImprimirResultados
                
                Print #iFichInet, "<TD BGCOLOR='" & scolor & "'><A HREF = '" & rs!DESCRIPCION & mml_FRASE0723 & rsFases!fase & IIf(rsFases!repesca = 1, mml_FRASE0725, "") & ".PDF'>" & IIf(rsFases!fase = 1, mml_FRASE0726, IIf(rsFases!fase = 2, mml_FRASE0727, " 1/" & Trim$(Str$(rsFases!fase)) & " ")) & IIf(rsFases!repesca = 1, mml_FRASE0725, "") & "</A></TD>"
                'On Local Error GoTo fich_no_encontrado
                If gl_bGenerarInet Then
                    On Local Error Resume Next
                    Do
repetir:
                        Err.Clear
                        
                        dateHora = Time
                        While Dir$(C_FICHERO_INET) = "" And DateDiff("s", dateHora, Time) < G_ESPERA_FICH_INET
                            DoEvents
                            Sleep 500
                        Wend
                        FileCopy C_FICHERO_INET, sDirFichas & "\Inet\" & rs!DESCRIPCION & mml_FRASE0723 & rsFases!fase & IIf(rsFases!repesca = 1, mml_FRASE0725, "") & ".PDF"
                        If Err.Number > 0 Then
                            If MsgBox(mml_FRASE0340 & C_FICHERO_INET & mml_FRASE0999, vbYesNo Or vbCritical, mml_FRASE0096) = vbNo Then
                                Close #iFichInet
                                If MsgBox(mml_FRASE1136, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbNo Then
                                    Exit Sub
                                Else
                                    Err.Clear
                                End If
                            Else
                                GoTo repetir
                            End If
                        End If
                    Loop Until Err.Number = 0
                    On Local Error Resume Next
                    Kill C_FICHERO_INET
                    If Not C_DEBUG Then On Local Error GoTo error
                End If
            End If
            If Not C_DEBUG Then On Local Error GoTo error
            rsFases.MoveNext
        Wend
        rsFases.Close
        ' Al final de cada grupo imprimimos el resumen
        If sCatAnt <> "" Then
            tbInfo.Text = rs!DESCRIPCION & mml_FRASE0728
            frmImprimirFinal.tbCodComp = tbCodComp.Text
            frmImprimirFinal.tbCodCateg = sCatAnt
            DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
            'If Not bComienzo Then SendKeys "~"
            bComienzo = False
            frmImprimirFinal.ImprimirResumen
            Print #iFichInet, "<TD BGCOLOR='" & scolor & "'><A HREF = '" & rs!DESCRIPCION & mml_FRASE0729 & mml_FRASE0730
            'On Local Error GoTo fich_no_encontrado
            If gl_bGenerarInet Then
                On Local Error Resume Next
                Do
repetirr:
                    Err.Clear
                    dateHora = Time
                    While Dir$(C_FICHERO_INET) = "" And DateDiff("s", dateHora, Time) < G_ESPERA_FICH_INET
                        DoEvents
                        Sleep 500
                    Wend
                    FileCopy C_FICHERO_INET, sDirFichas & "\Inet\" & rs!DESCRIPCION & mml_FRASE0731
                    If Err.Number > 0 Then
                        If MsgBox(mml_FRASE0340 & C_FICHERO_INET & mml_FRASE0999, vbYesNo Or vbCritical, mml_FRASE0096) = vbNo Then
                            Close #iFichInet
                            If MsgBox(mml_FRASE1136, vbYesNo Or vbQuestion, G_MSG_AVISO) = vbNo Then
                                Exit Sub
                            Else
                                Err.Clear
                            End If
                        Else
                            GoTo repetirr
                        End If
                    End If
                Loop Until Err.Number = 0
                On Local Error Resume Next
                Kill C_FICHERO_INET
                If Not C_DEBUG Then On Local Error GoTo error
            End If
            If Not C_DEBUG Then On Local Error GoTo error
        End If
        Print #iFichInet, "</TR>" & Chr$(13) & Chr$(10)
        rs.MoveNext
    Wend
    Print #iFichInet, "</TABLE><br>"
    rs.Close
    Print #iFichInet, G_PIE_INET & "</HTML>"
    Close iFichInet
    If Not gl_bGenerarInet Then
        Printer.EndDoc
    End If
    MsgBox mml_FRASE0732, vbOKOnly Or vbInformation, mml_FRASE0086
    Exit Sub
fich_no_encontrado:
    MsgBox mml_FRASE0340 & C_FICHERO_INET & mml_FRASE0733, vbOKOnly Or vbCritical, mml_FRASE0096
    Exit Sub
error:
    Dim Msj As String
   If Err.Number <> 0 Then
        Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
              & Err.Source & Chr(13) & Err.Description
        MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
   End If

End Sub
Private Sub cmdImprimir1_Click()
Dim rs As Recordset
Dim rsCont As Recordset
Dim rsBailes As Recordset
Dim rsFases As Recordset
Dim sTexto As String
Dim iEscala As Integer
Dim iCCopias As Integer
Dim iNumDorsales As Integer
Dim sDirFichas  As String
Dim iFichInet As Integer
Dim sCatAnt As String

    If MsgBox(mml_FRASE0718, vbYesNo Or vbQuestion, mml_FRASE0084) = vbNo Then
        Exit Sub
    End If
    
    ComprobarImpresoraPorDefecto
    CDialog.Flags = cdlPDHidePrintToFile Or cdlPDNoSelection
    CDialog.CancelError = True
    On Local Error GoTo Pcancelar
    If G_IMPRESORA_DEFECTO = "NA" Then CDialog.ShowPrinter
    On Local Error GoTo 0
    CDialog.CancelError = False
    GoTo Pseguir
Pcancelar:
    Exit Sub
Pseguir:
    
    On Local Error GoTo error
    'frmImprimirInternet.cmdImprimir1.Enabled = False
    
    sDirFichas = VarCfg("dir_fichas")
    On Local Error Resume Next
    Kill sDirFichas & "\Inet\*.*"
    On Local Error GoTo 0
    
    iFichInet = FreeFile
    Open sDirFichas & "\Inet\Comp-Resultados.htm" For Output As #iFichInet
    
    iEscala = Printer.Width / 10
    Set rs = db.OpenRecordset(" SELECT * FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
    sTexto = rs!DESCRIPCION
    
    Print #iFichInet, "<HTML>" & G_CAB_INET
    Print #iFichInet, "<h1>" & Trim$(sTexto) & "</h1><br><br>"
    
    rs.Close
        
    Set rs = db.OpenRecordset(" SELECT c.codigo, hora, id_categoria, ge.nombre, descripcion, m.nombre as mod FROM categorias c, gruposedad ge, modalidad m WHERE c.cod_modalidad = m.codigo AND cod_competicion = " & tbCodComp.Text & " AND cod_grupoedad = ge.codigo ORDER BY descripcion", dbOpenSnapshot)
    While Not rs.EOF
        Print #iFichInet, G_CAB_TABLA
        Print #iFichInet, "<TR><TD><h3>" & rs!DESCRIPCION & "</h3></TD></TR>"
        
        sCatAnt = ""
        Set rsFases = db.OpenRecordset("SELECT DISTINCT fase, repesca FROM puntuaciones WHERE cod_categoria = " & rs!codigo & " ORDER BY 1 DESC")
        While Not rsFases.EOF
            tbInfo.Text = rs!DESCRIPCION & mml_FRASE0723 & rsFases!fase
            frmImprimirFinal.tbCodComp = tbCodComp.Text
            frmImprimirFinal.tbCodCateg = rs!codigo
            frmImprimirFinal.tbCodFase = rsFases!fase
            sCatAnt = rs!codigo
            If rsFases!repesca = 1 Then
                tbInfo.Text = tbInfo.Text & mml_FRASE0724
                frmImprimirFinal.chkRep.Value = rsFases!repesca
            End If
            DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
            'SendKeys "~"
            frmImprimirFinal.ImprimirResultados
            
            Print #iFichInet, "<TR><TD><A HREF = '" & rs!DESCRIPCION & mml_FRASE0723 & rsFases!fase & ".PDF'>" & rs!DESCRIPCION & mml_FRASE0723 & IIf(rsFases!fase = 1, mml_FRASE0329, IIf(rsFases!fase = 2, mml_FRASE0330, "1/" & Trim$(Str$(rsFases!fase)))) & "</A></TD></TR>"
            On Local Error GoTo fich_no_encontrado
            FileCopy C_FICHERO_INET, sDirFichas & "\Inet\" & rs!DESCRIPCION & mml_FRASE0723 & rsFases!fase & ".PDF"
            Kill C_FICHERO_INET
            On Local Error GoTo 0
            rsFases.MoveNext
        Wend
        rsFases.Close
        ' Al final de cada grupo imprimimos el resumen
        If sCatAnt <> "" Then
            tbInfo.Text = rs!DESCRIPCION & mml_FRASE0728
            frmImprimirFinal.tbCodComp = tbCodComp.Text
            frmImprimirFinal.tbCodCateg = sCatAnt
            DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
            'SendKeys "~"
            frmImprimirFinal.ImprimirResumen
            Print #iFichInet, "<TR><TD><A HREF = '" & rs!DESCRIPCION & mml_FRASE0729 & rs!DESCRIPCION & mml_FRASE0734
            On Local Error GoTo fich_no_encontrado
            FileCopy C_FICHERO_INET, sDirFichas & "\Inet\" & rs!DESCRIPCION & mml_FRASE0731
            Kill C_FICHERO_INET
            On Local Error GoTo 0
        End If
        Print #iFichInet, "</TABLE><br>"
        rs.MoveNext
    Wend
    rs.Close
    Print #iFichInet, G_PIE_INET & "</HTML>"
    Close iFichInet
    MsgBox mml_FRASE0732, vbOKOnly Or vbInformation, mml_FRASE0086
    Exit Sub
fich_no_encontrado:
    MsgBox mml_FRASE0340 & C_FICHERO_INET & mml_FRASE0733, vbOKOnly Or vbCritical, mml_FRASE0096
    Exit Sub
error:
Dim Msj As String
   If Err.Number <> 0 Then
        Msj = "Error # " & Str(Err.Number) & mml_FRASE0208 _
              & Err.Source & Chr(13) & Err.Description
        MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
   End If

End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
End Sub

Private Sub cmdGenerarTablaAEBDC_Click()
    frmGenerarTablaAEBDC.Show vbModal
End Sub

Private Sub Form_Load()
    TraducirCadenas Me
    tbCodComp.Text = VarCfg("horario_codcompeticion")
    tbDescComp.Text = sDescCompeticion(Val(tbCodComp.Text))

    Dim rs As Recordset
    Dim i As Integer
    
        i = 0
        'Localizamos las categorias que tienen puntuaciones
        Set rs = db.OpenRecordset("SELECT codigo,descripcion FROM categorias c WHERE cod_competicion = " & CodCompActiva & _
                  " AND (SELECT COUNT(*) FROM puntuaciones p WHERE p.cod_categoria = c.codigo)>0 ORDER BY descripcion DESC", dbOpenSnapshot)
        While Not rs.EOF
            lstCat.AddItem rs.Fields("descripcion")
            lstCat.ItemData(i) = rs.Fields("codigo")
            Inc i
            rs.MoveNext
        Wend
        rs.Close

End Sub
