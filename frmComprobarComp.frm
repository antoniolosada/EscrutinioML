VERSION 5.00
Begin VB.Form frmComprobarComp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "mml_FRASE1056"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
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
      Height          =   495
      Left            =   8190
      TabIndex        =   2
      Top             =   5490
      Width           =   1860
   End
   Begin VB.CommandButton cmdComprobarCompeticion 
      Caption         =   "mml_FRASE1056"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   5490
      Width           =   5355
   End
   Begin VB.TextBox tbComp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5385
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmComprobarComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdComprobarCompeticion_Click()
    Dim rs As Recordset
    Dim rs1 As Recordset
    Dim iTmp As Integer
    
    If Not C_DEBUG Then On Error GoTo error
    
    tbComp.Text = ""
    tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1060 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobamos si hay parejas Adicionales
    Set rs = db.OpenRecordset("SELECT min_dorsal_oficial FROM competiciones WHERE codigo = " & CodCompActiva, dbOpenSnapshot)
    If Not rs.EOF Then
        iTmp = rs.Fields("min_dorsal_oficial")
        rs.Close
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales d WHERE num_dorsal < " & iTmp & " AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ")", dbOpenSnapshot)
        If Not rs.EOF Then
            If rs.Fields(0) > 0 Then
                tbComp.Text = tbComp.Text & mml_FRASE1057 & iTmp & vbCrLf
            End If
        End If
    End If
    rs.Close
    
    Set rs = db.OpenRecordset("SELECT codigo, nombre_hombre, nombre_mujer FROM parejas WHERE pareja_adicional = 1 AND cod_competicion = " & CodCompActiva & " ORDER BY 1", dbOpenSnapshot)
    If Not rs.EOF Then
        tbComp.Text = tbComp.Text & mml_FRASE1134 & vbCrLf
        While Not rs.EOF
            tbComp.Text = tbComp.Text & "(" & rs.Fields("codigo") & ") " & rs.Fields("nombre_hombre") & ", " & rs.Fields("nombre_mujer") & vbCrLf
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1061 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobamos que no hay bailes repetidos
    Set rs = db.OpenRecordset("SELECT bc.cod_baile, cod_categoria, fase, c.descripcion, b.nombre, COUNT(*) as numero FROM bailes_categ bc, categorias c, bailes b WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = c.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ") GROUP BY bc.cod_baile, cod_categoria, fase, c.descripcion, b.nombre HAVING COUNT(*) > 1", dbOpenSnapshot)
    While Not rs.EOF
        tbComp.Text = tbComp.Text & mml_FRASE1058 & " (" & rs.Fields("cod_categoria") & ") " & rs.Fields("descripcion") & ", " & sDescFase(rs.Fields("fase")) & " - " & rs.Fields("nombre") & vbCrLf
        rs.MoveNext
    Wend
    rs.Close
    
    ' LA comprobación de los números de bailes solo es para Salon
    If Not G_COUNTRY Then
        tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1062 & vbCrLf
        tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
        'Comprobamos y alguna categoria tiene un número distinto de bailes y no es un Open de 10 bailes
        Set rs = db.OpenRecordset("SELECT id_categoria, cod_categoria, fase, c.descripcion, COUNT(*) as numero FROM bailes_categ bc, categorias c WHERE bc.cod_categoria = c.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ") GROUP BY id_categoria, cod_categoria, fase, c.descripcion HAVING COUNT(*) <> 5", dbOpenSnapshot)
        If rs.EOF Then
            tbComp.Text = tbComp.Text & mml_FRASE1081 & vbCrLf
        Else
            While Not rs.EOF
                'Solo comprobamos las categorias <> 10 bailes que no tengan un 10 en la descripción
                'Comprobamos las categorias que tienen más de 1 caracter TMAtch, OpenX, IDSF, Prof <> 5 bailes
                'Comprobamos categorias inferiores o iguales a la D con más de 5 bailes
                'Comprobamos categorias superiores a la D con un número distinto a 5 bailes
                If Not (rs.Fields("numero") = 10 And InStr(UCase(rs.Fields("descripcion")), "10") > 0) Then
                    If rs.Fields("numero") < 2 Or _
                       (Len(rs.Fields("id_categoria")) > 1 And rs.Fields("numero") <> 5) Or _
                       (rs.Fields("id_categoria") >= "F" And (rs.Fields("numero") > 5 Or rs.Fields("numero") < 2)) Or _
                       ((rs.Fields("id_categoria") = "D" Or rs.Fields("id_categoria") = "E") And (rs.Fields("numero") > 5 Or rs.Fields("numero") < 3)) Or _
                       ((rs.Fields("id_categoria") < "D" And rs.Fields("numero") <> 5) And _
                        (rs.Fields("id_categoria") <> "C" And rs.Fields("numero") <> 4)) Then
                        tbComp.Text = tbComp.Text & mml_FRASE1059 & rs.Fields("numero") & " -> " & rs.Fields("descripcion") & ", " & sDescFase(rs.Fields("fase")) & vbCrLf
                    End If
                End If
                rs.MoveNext
            Wend
        End If
        rs.Close
    End If
    
    'Comprobamos que si hay fases eliminatorias debe haber bailes definidos para final y eliminatorias
    tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1131 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobamos que todas las categorias tengan bailes definidos para la final
    Set rs = db.OpenRecordset("SELECT c.codigo, c.descripcion FROM categorias c WHERE 0=(SELECT COUNT(*) FROM bailes_categ bc WHERE bc.cod_categoria = c.codigo AND bc.fase = 1) AND c.cod_competicion =" & CodCompActiva, dbOpenSnapshot)
    If Not rs.EOF Then
        tbComp.Text = tbComp.Text & mml_FRASE1132 & vbCrLf
        While Not rs.EOF
            tbComp.Text = tbComp.Text & rs.Fields("codigo") & " - " & rs.Fields("descripcion") & vbCrLf
            rs.MoveNext
        Wend
    End If
    rs.Close
    'Comprobamos que todas las categorias con más de 7 dorsales tengan bailes definidos para las eliminatorias
    Set rs = db.OpenRecordset("SELECT c.codigo, c.descripcion FROM categorias c WHERE 1<(SELECT MAX(FASE) FROM dorsales d WHERE c.codigo = d.cod_categoria ) AND 0=(SELECT COUNT(*) FROM bailes_categ bc WHERE bc.cod_categoria = c.codigo AND bc.fase = 2) AND c.cod_competicion =" & CodCompActiva, dbOpenSnapshot)
    If Not rs.EOF Then
        tbComp.Text = tbComp.Text & mml_FRASE1133 & vbCrLf
        While Not rs.EOF
            tbComp.Text = tbComp.Text & rs.Fields("codigo") & " - " & rs.Fields("descripcion") & vbCrLf
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1063 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobamos que todas las categorias tengan jueces distintos
    Set rs = db.OpenRecordset("SELECT id_juez, cod_categoria, c.descripcion, COUNT(*) as numero FROM juez_categ jc, categorias c WHERE jc.cod_categoria = c.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ") GROUP BY id_juez, cod_categoria, c.descripcion HAVING COUNT(*) > 1", dbOpenSnapshot)
    While Not rs.EOF
        tbComp.Text = tbComp.Text & mml_FRASE1066 & " (" & rs.Fields("cod_categoria") & ") " & rs.Fields("descripcion") & vbCrLf
        rs.MoveNext
    Wend
    rs.Close
    
    tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1064 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobar que solo hay un juez de pasos en cada categoría
    Set rs = db.OpenRecordset("SELECT pasos, cod_categoria, c.descripcion, COUNT(*) as numero FROM juez_categ jc, categorias c WHERE jc.cod_categoria = c.codigo AND pasos = 1 AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ") GROUP BY pasos, cod_categoria, c.descripcion HAVING COUNT(*) > 1", dbOpenSnapshot)
    While Not rs.EOF
        tbComp.Text = tbComp.Text & mml_FRASE1067 & " (" & rs.Fields("cod_categoria") & ") " & rs.Fields("descripcion") & vbCrLf
        rs.MoveNext
    Wend
    rs.Close

    tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1065 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobar los ditintos números de jueces
    Dim iJueces As Integer
    iJueces = 0
    Set rs = db.OpenRecordset("SELECT c.descripcion, cod_categoria, COUNT(*) as numero FROM juez_categ jc, categorias c WHERE c.codigo = jc.cod_categoria AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ") GROUP BY c.descripcion, cod_categoria ORDER BY 3,2", dbOpenSnapshot)
    If rs.EOF Then
        tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1082 & vbCrLf
    Else
        While Not rs.EOF
            If iJueces <> rs.Fields("numero") Then
                iJueces = rs.Fields("numero")
                tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1069 & iJueces & vbCrLf
            Else
                tbComp.Text = tbComp.Text & ", "
            End If
            tbComp.Text = tbComp.Text & rs.Fields("descripcion")
            rs.MoveNext
        Wend
    End If
    rs.Close
    
    'Comprobación de dorsales repetidos en categorias Est o Lat
    tbComp.Text = tbComp.Text & vbCrLf & vbCrLf & mml_FRASE1074 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobando el numero de dorsales que aparecen mas de una vez y que corresponde alguna de sus apariciones a parejas distintas siendo categorias que comienzan por Lat o Est
    'Por cada dorsal se busca si hay otro igual en la competición con nombre de pareja distinto
    Set rs = db.OpenRecordset( _
        "SELECT DISTINCT num_dorsal, fase, c.descripcion FROM dorsales d, categorias c, parejas p WHERE " & _
          " (Mid(c.descripcion,1,3) = 'Lat' OR Mid(c.descripcion,1,3) = 'Est' OR Mid(c.descripcion,1,3) = 'Com') AND p.codigo = d.cod_pareja AND d.cod_categoria = c.codigo AND " & _
          "0<(SELECT COUNT(*) FROM dorsales d1, parejas p1 WHERE d.num_dorsal = d1.num_dorsal AND " & _
          "d1.cod_pareja = p1.codigo AND (p1.nombre_hombre <> p.nombre_hombre OR p1.nombre_mujer <> p.nombre_mujer) " & _
          "AND d1.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ")) " & _
          "AND d.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ") ORDER BY 1" _
          , dbOpenSnapshot)
    While Not rs.EOF
        tbComp.Text = tbComp.Text & mml_FRASE1075 & ": " & rs.Fields("num_dorsal") & " -> " & rs.Fields("descripcion") & " , " & sDescFase(rs.Fields("fase")) & vbCrLf
        rs.MoveNext
    Wend
    
    'Comprobar horario: Todas las categorias y fases generadas no vacias deben estar en el horario
    'Todas las categorias y fases mayores del horario deben existir
    tbComp.Text = vbCrLf & tbComp.Text & vbCrLf & mml_FRASE1071 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Buscamos las categorias y fases generadas de esta competición que no están en el horario de la competición y que tienen dorsales
    Set rs = db.OpenRecordset("SELECT DISTINCT d.cod_categoria, d.fase, c.descripcion FROM dorsales d, categorias c WHERE d.cod_categoria = c.codigo AND NOT (cod_categoria & fase) IN (SELECT (cod_categoria & numfase) FROM horario WHERE cod_competicion = " & CodCompActiva & ") AND d.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ") AND (SELECT COUNT(*) FROM dorsales d1 WHERE d1.cod_categoria = d.cod_categoria AND d1.fase = d.fase)>0", dbOpenSnapshot)
    While Not rs.EOF
        tbComp.Text = tbComp.Text & mml_FRASE1072 & ": " & rs.Fields("descripcion") & " , " & sDescFase(rs.Fields("fase")) & vbCrLf
        rs.MoveNext
    Wend
    rs.Close
    'Buscamos las categorias y fases generadas que estan en el horario y no están en la competición
    Set rs = db.OpenRecordset("SELECT grupo, fase FROM horario h WHERE h.numfase <> " & C_FASE_GENERAL_LOOK & " AND h.cod_competicion = " & CodCompActiva & " AND 0=(SELECT COUNT(*) FROM dorsales d WHERE d.cod_categoria = h.cod_categoria AND d.fase >= h.numfase AND d.cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & "))", dbOpenSnapshot)
    While Not rs.EOF
        tbComp.Text = tbComp.Text & mml_FRASE1073 & ": " & rs.Fields("grupo") & " , " & rs.Fields("fase") & vbCrLf
        rs.MoveNext
    Wend
    rs.Close
    
    tbComp.Text = tbComp.Text & vbCrLf & vbCrLf & mml_FRASE1091 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobamos si hay alguna inscripción en una categoria  con modalidad equivocada
    Set rs = db.OpenRecordset("SELECT d.num_dorsal, c.descripcion, p.codigo, p.nombre_hombre, p.nombre_mujer, m.nombre FROM parejas p, dorsales d, categorias c, modalidad m WHERE p.cod_modalidad = m.codigo AND c.codigo = d.cod_categoria AND p.cod_competicion = " & CodCompActiva & " AND p.codigo = d.cod_pareja AND p.cod_modalidad <> c.cod_modalidad", dbOpenSnapshot)
    While Not rs.EOF
        tbComp.Text = tbComp.Text & mml_FRASE1092 & ": Number/Dorsal " & rs.Fields("num_dorsal") & " , " & rs.Fields("descripcion") & ", Cod.Couple/Pareja " & rs.Fields("codigo") & " , " & rs.Fields("nombre_hombre") & " , " & rs.Fields("nombre_mujer") & "(" & rs.Fields("nombre") & ")" & vbCrLf
        rs.MoveNext
    Wend
    rs.Close
    
    
    tbComp.Text = tbComp.Text & vbCrLf & vbCrLf & mml_FRASE1086 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobamos que si hay pistas, todas las categorias tienen asociada su pista
    Dim i As Integer
    Dim iCont As Integer
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM categorias WHERE descripcion LIKE '*(P?)*' AND cod_competicion = " & CodCompActiva, dbOpenSnapshot)
    iCont = rs.Fields(0)
    rs.Close
    'Hay varias pistas
    If iCont > 0 Then
        'Hay categoria con indicación de pista, comprobamos si alguna no la tiene
        iCont = 0
        Set rs = db.OpenRecordset("SELECT COUNT(*) FROM categorias WHERE NOT descripcion LIKE '*(P?)*' AND cod_competicion = " & CodCompActiva, dbOpenSnapshot)
        iCont = rs.Fields(0)
        rs.Close
        If iCont > 0 Then
            tbComp.Text = tbComp.Text & mml_FRASE1087 & vbCrLf
        End If
        
        'Comprobamos si cada pista tiene un panel de jueces totalmente diferente
        'For i = 1 To 9
        '    Set rs = db.OpenRecordset("SELECT DISTINCT id_juez FROM juez_categ jc, categorias c WHERE c.cod_competicion = " & CodCompActiva & " AND c.codigo = jc.cod_categoria AND c.descripcion LIKE '*(P" & Trim$(Str$(i)) & ")*' AND id_juez IN (SELECT DISTINCT id_juez FROM juez_categ jc1, categorias c1 WHERE c1.cod_competicion = c.cod_competicion AND c1.codigo = jc1.cod_categoria AND NOT c1.descripcion LIKE '*(P" & Trim$(Str$(i)) & ")*')", dbOpenSnapshot)
        '    If Not rs.EOF Then
        '        tbComp.Text = tbComp.Text & mml_FRASE1089 & " " & i & vbCrLf
         
        '        While Not rs.EOF
        '            tbComp.Text = tbComp.Text & mml_FRASE1088 & ": " & rs.Fields("id_juez") & vbCrLf
        '            rs.MoveNext
        '        Wend
        '        rs.Close
        '    End If
        'Next
    End If
    
    tbComp.Text = tbComp.Text & vbCrLf & vbCrLf & mml_FRASE1223 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobando si hay dorsales calificados como no oficiales
    Set rs = db.OpenRecordset("SELECT COUNT(*) as numero FROM dorsales WHERE num_dorsal < " & iMinDorsalOficial(CodCompActiva) & " AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ")", dbOpenSnapshot)
    If rs.Fields("numero") > 0 Then
        tbComp.Text = tbComp.Text & mml_FRASE1224 & rs.Fields("numero") & vbCrLf
    Else
        tbComp.Text = tbComp.Text & mml_FRASE1225 & vbCrLf
    End If
    rs.Close
    
    
    tbComp.Text = tbComp.Text & vbCrLf & vbCrLf & mml_FRASE1070 & vbCrLf
    tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
    'Comprobando el número de parejas por categoria y fase y número de fases si no ha comenzado la competición
    'Comprobamos si ha comenzado la competición
    Dim bCompIniciada As Boolean
    Set rs = db.OpenRecordset("SELECT COUNT(*) as numero FROM puntuaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ")", dbOpenSnapshot)
    If rs.Fields("numero") > 0 Then
        bCompIniciada = True
    Else
        bCompIniciada = False
    End If
    rs.Close
    ComprobarFases bCompIniciada
    ComprobarDobleFase bCompIniciada

        
    If C_COUNTRY Then
        'Comprobamos si todos los grupos que bailan juntos en country tienen el mismo número de bailes y de jueces
        tbComp.Text = tbComp.Text & vbCrLf & vbCrLf & mml_FRASE1255 & vbCrLf
        tbComp.Text = tbComp.Text & "----------------------------------------------------------------------" & vbCrLf
        Dim sJueces As String
        Dim sJuecesPasos As String
        Dim sBailes As String
        Dim sJuecesCab As String
        Dim sJuecesPasosCab As String
        Dim sBailesCab As String
        
        Dim sGrupoOrden As String
        Dim sGrupoCodCateg As String
        Dim sGrupoNombre As String
        Dim sGrupoFase As String
        Dim sGrupoRepesca As String
        Set rs = db.OpenRecordset("SELECT h.cod_categoria, h.numfase, h.grupo, h.fase, h.repesca, h.inicio_grupo, h.orden FROM horario h WHERE h.cod_competicion = " & CodCompActiva & " ORDER BY orden", dbOpenSnapshot)
        While Not rs.EOF
            'Recuperamos los jueces
            Set rs1 = db.OpenRecordset("SELECT id_juez, pasos FROM juez_categ WHERE cod_categoria = " & rs.Fields("cod_categoria"), dbOpenSnapshot)
            sJueces = ""
            sJuecesPasos = ""
            While Not rs1.EOF
                sJueces = sJueces & Trim(rs1.Fields("id_juez"))
                sJuecesPasos = sJuecesPasos & Trim(rs1.Fields("pasos"))
                rs1.MoveNext
            Wend
            'Recuperamos los bailes
            Set rs1 = db.OpenRecordset("SELECT cod_baile, fase, posicion FROM bailes_categ WHERE cod_categoria = " & rs.Fields("cod_categoria") & " AND fase = " & IIf(rs.Fields("numfase") = 1, "1", "2"), dbOpenSnapshot)
            sBailes = ""
            While Not rs1.EOF
                sBailes = Trim(rs1.Fields("cod_baile")) & Trim(rs1.Fields("fase")) & Trim(rs1.Fields("posicion"))
                rs1.MoveNext
            Wend
            
            If rs.Fields("inicio_grupo") = 1 Then
                'Almacenamos la información para compararla con el resto de los grupos que salen juntos
                sJuecesCab = sJueces
                sJuecesPasosCab = sJuecesPasos
                sBailesCab = sBailes
            
                'Recuperamos la información
                sGrupoOrden = rs.Fields("orden")
                sGrupoCodCateg = rs.Fields("cod_categoria")
                sGrupoNombre = rs.Fields("grupo")
                sGrupoFase = rs.Fields("fase")
                sGrupoRepesca = IIf(rs.Fields("repesca") = 1, "Repesca", "")
            Else
                Dim sCad As String
                sCad = ""
                If sJuecesCab <> sJueces Then
                    'Los jueces son diferentes
                    sCad = mml_FRASE1256
                End If
                If sJuecesPasosCab <> sJuecesPasos Then
                    'Los jueces de pasos son diferentes
                    sCad = mml_FRASE1257
                End If
                If sBailesCab <> sBailes Then
                    'Los bailes o el orden son diferentes
                    sCad = mml_FRASE1258
                End If
                If sCad <> "" Then
                    tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE1259 & ": (" & sGrupoOrden & ") " & sGrupoNombre & ", " & sGrupoFase & " " & sGrupoRepesca & vbCrLf
                    tbComp.Text = tbComp.Text & sCad & vbCrLf
                    tbComp.Text = tbComp.Text & mml_FRASE1260 & ": (" & rs.Fields("orden") & ") " & rs.Fields("grupo") & ", " & rs.Fields("fase") & IIf(rs.Fields("repesca") = 1, " Repesca", "") & vbCrLf
                End If
            End If
            rs1.Close
            rs.MoveNext
        Wend
        rs.Close
    End If
    
    tbComp.Text = tbComp.Text & vbCrLf & mml_FRASE0545
    Exit Sub
error:
    ProcesarError "cmdComprobarCompeticion"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub

Private Sub ComprobarFases(bCompIniciada As Boolean)
Dim rs As Recordset, rs1 As Recordset, iNumFases As Integer

    tbComp.Text = tbComp.Text & mml_FRASE0539 & Chr$(13) & Chr$(10)
    
    Set rs = db.OpenRecordset("SELECT COUNT(*) as numdorsales,cod_categoria,descripcion, fase, repesca FROM dorsales d, categorias c WHERE d.cod_categoria = c.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & frmADorsales.tbCodComp.Text & ") GROUP BY cod_categoria, descripcion, fase, repesca", dbOpenSnapshot)
    While Not rs.EOF
        ' comprobamos si es la correcta para el número de dorsales
        If (rs!fase = 1 And rs!numdorsales <= 7) Or (rs!fase = 2 And rs!numdorsales > 7 And rs!numdorsales <= 13) Or (rs!fase > 2 And CalcularFase(rs!numdorsales) = rs!fase) Then
        'Fase correcta
        Else
            tbComp.Text = tbComp.Text & rs!DESCRIPCION & mml_FRASE0541 & rs!fase & "(r" & rs!repesca & ") " & mml_FRASE0542 & CalcularFase(rs!numdorsales) & mml_FRASE0543 & rs!numdorsales & mml_FRASE0544 & Chr$(13) & Chr$(10)
        End If
        
        rs.MoveNext
    Wend
    rs.Close
End Sub

Private Sub ComprobarDobleFase(bCompIniciada As Boolean)
Dim rs As Recordset, rs1 As Recordset, iNumFases As Integer

    If Not bCompIniciada Then
        tbComp.Text = tbComp.Text & mml_FRASE0539 & Chr$(13) & Chr$(10)
        
        Set rs = db.OpenRecordset("SELECT cod_categoria,descripcion FROM dorsales d, categorias c WHERE d.cod_categoria = c.codigo AND cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & frmADorsales.tbCodComp.Text & ") GROUP BY cod_categoria, descripcion, fase", dbOpenSnapshot)
        While Not rs.EOF
            'Contamos el número de fases
            Set rs1 = db.OpenRecordset("SELECT DISTINCT fase FROM dorsales WHERE cod_categoria = " & rs!cod_categoria, dbOpenSnapshot)
            iNumFases = 0
            If Not rs1.EOF Then
                rs1.MoveLast
                iNumFases = rs1.RecordCount
            End If
            rs1.Close
            
            If iNumFases > 1 Then
                tbComp.Text = tbComp.Text & rs!DESCRIPCION & " (" & rs!cod_categoria & mml_FRASE0540 & Chr$(13) & Chr$(10)
            End If
            
            rs.MoveNext
        Wend
        rs.Close
    
    End If
End Sub

