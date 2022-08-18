VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimirOrdenCombinado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mml_FRASE0735"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImpOrdenCombinado 
      Caption         =   "mml_FRASE0736"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton cmdCateg1 
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
         Picture         =   "frmImprimitOrdenCombinado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1170
         Width           =   450
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
         Left            =   1845
         Picture         =   "frmImprimitOrdenCombinado.frx":046A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   450
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
         Height          =   360
         Left            =   1845
         Picture         =   "frmImprimitOrdenCombinado.frx":08D4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   450
      End
      Begin VB.TextBox tbDescCategEst 
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
         TabIndex        =   10
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox tbCodCategEst 
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
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   840
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox tbCodCategLat 
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
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox tbDescCategLat 
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
         TabIndex        =   5
         Top             =   720
         Width           =   4935
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
         TabIndex        =   3
         Top             =   240
         Width           =   4935
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
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "mml_FRASE0737"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "mml_FRASE0738"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1575
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
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
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
      Left            =   6960
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmImprimirOrdenCombinado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCateg_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCategLat.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCategLat.Text = sResultado(2)

End Sub

Private Sub cmdCateg1_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCategEst.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCategEst.Text = sResultado(2)
End Sub

Private Sub cmdImpOrdenCombinado_Click()
Dim iEscala As Integer
Dim rs As Recordset
Dim rsDorsal As Recordset
Dim rsElim As Recordset
Dim rsMarcas As Recordset
Dim aPosicion() As Integer
Dim iParejas As Integer
Dim iParejasPart As Integer
Dim iCPareja As Integer
Dim iUltimaFaseLat As Integer
Dim iUltimaFaseEst As Integer
Dim iUltimaFase As Integer
Dim iPosValorMax As Integer
Dim iCValorMax As Integer
Dim lValorMax As Long
Dim lValor As Long
Dim iCParejas As Integer
Dim iCPuestos As Integer
Dim aTabla() As Integer
Dim iPuestoIni As Integer
Dim iRepesca As Integer
Dim dPuestoLat As Double
Dim dPuestoEst As Double
Dim iCCopias As Integer

Const C_NUM_DORSAL = 0
Const C_ELIM_SUP_LAT = 1
Const C_ELIM_SUP_EST = 2
Const C_CONT_MARCAS_LAT = 3
Const C_CONT_MARCAS_EST = 4
Const C_PUESTO_LAT = 5
Const C_PUESTO_EST = 6
Const C_PUESTO_FIN_LAT = 7
Const C_PUESTO_FIN_EST = 8
Const C_PUNTOS_LAT = 9
Const C_PUNTOS_EST = 10

Const C_COD_PAREJA = 11
Const C_MAX_DIM = 12
Const C_NO_PRESENTADO = -1

    
    If tbCodComp.Text = "" Or tbCodCategLat.Text = "" Or tbCodCategEst.Text = "" Then
        CamposSinCubrir
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
    
    For iCCopias = 1 To CDialog.Copies
        'Imprimimos la cabecera
        Printer.PaintPicture frmMenu.picEPA.Picture, Printer.Width - frmMenu.picEPA.Width - G_MARGEN_EPA, G_MARGEN_EPA_Y
        iEscala = Printer.Width / 10
        Printer.FontSize = 8
        If G_COUNTRY Then
            Printer.Print mml_FRASE1034;
        Else
            Printer.Print mml_FRASE0684;
        End If
        Printer.FontSize = 13
        Printer.Print
        Printer.CurrentX = 0
        Set rs = db.OpenRecordset(" SELECT descripcion, fecha FROM competiciones WHERE codigo = " & tbCodComp.Text, dbOpenSnapshot)
        Printer.Print rs!DESCRIPCION & "  (" & rs!fecha & ")"
        rs.Close
        Printer.Line Step(0, 0)-Step(Printer.Width, 0)
        Set rs = db.OpenRecordset(" SELECT hora, id_categoria, c.codigo, ge.nombre, descripcion FROM categorias c, gruposedad ge WHERE c.codigo = " & tbCodCategLat.Text & " AND cod_grupoedad = ge.codigo ORDER BY hora", dbOpenSnapshot)
        Printer.FontSize = 12
        Printer.CurrentX = 0
        Printer.Print mml_FRASE0739;
        Printer.CurrentX = iEscala
        Printer.Print rs!id_categoria;
        Printer.CurrentX = iEscala * 2
        Printer.Print rs!Nombre;
        Printer.CurrentX = iEscala * 4
        Printer.Print rs!DESCRIPCION & mml_FRASE0740
        Printer.Line -Step(Printer.Width, 0)
        Printer.FontSize = 2
        Printer.Print
        Printer.FontSize = 10
        rs.Close
        
        ' Seleccionamos las parejas que tienen puntuación por ser oficiales
        ' Comprobamos el número de parejas anotadas
        Debug.Print "SELECT DISTINCT num_dorsal FROM dorsales WHERE num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND (cod_categoria = " & tbCodCategLat.Text & " OR cod_categoria = " & tbCodCategLat.Text & ") ORDER BY 1"
        Set rsDorsal = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM dorsales WHERE num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND (cod_categoria = " & tbCodCategLat.Text & " OR cod_categoria = " & tbCodCategLat.Text & ") ORDER BY 1", dbOpenSnapshot)
        If Not rsDorsal.EOF Then
            rsDorsal.MoveLast
            iParejas = rsDorsal.RecordCount
            rsDorsal.MoveFirst
        Else
            MsgBox mml_FRASE0696, vbOKOnly Or vbCritical, mml_FRASE0096
            rsDorsal.Close
            Exit Sub
        End If
        ' Comprobamos el número de parejas oficiales que participaron
        Debug.Print "SELECT DISTINCT num_dorsal FROM puntuaciones WHERE  num_dorsal >= " & iMinDorsalOficial(tbCodComp.Text) & " AND cod_categoria=" & tbCodCategLat.Text
        Set rs = db.OpenRecordset("SELECT DISTINCT num_dorsal FROM puntuaciones WHERE cod_categoria=" & tbCodCategLat.Text, dbOpenSnapshot)
        If Not rs.EOF Then
            rs.MoveLast
            iParejasPart = rs.RecordCount
            rs.Close
        Else
            MsgBox mml_FRASE0697, vbOKOnly Or vbCritical, mml_FRASE0096
            rs.Close
            Exit Sub
        End If
        ReDim aTabla(iParejas, C_MAX_DIM)
        iCPareja = 0
        iCPuestos = 1
        'Recorremos todas las parejas anotadas
        While Not rsDorsal.EOF
            aTabla(iCPareja, C_NUM_DORSAL) = rsDorsal!num_dorsal
            Set rs = db.OpenRecordset("SELECT cod_pareja FROM dorsales WHERE cod_categoria = " & tbCodCategLat.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If Not rs.EOF Then
                aTabla(iCPareja, C_COD_PAREJA) = rs!cod_pareja
            Else
                aTabla(iCPareja, C_COD_PAREJA) = 1
            End If
            rs.Close
            ' Primero comprobamos si tenemos un puesto en la final
            'Latinos
            Set rsElim = db.OpenRecordset("SELECT puesto FROM cal_conjunto WHERE regla = 'FIN' AND cod_categoria = " & tbCodCategLat.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            If Not rsElim.EOF Then
                aTabla(iCPareja, C_CONT_MARCAS_LAT) = 10 - rsElim!puesto
            Else
                aTabla(iCPareja, C_CONT_MARCAS_LAT) = 0
            End If
            rsElim.Close
            'Estandar
            Set rsElim = db.OpenRecordset("SELECT puesto FROM cal_conjunto WHERE regla = 'FIN' AND cod_categoria = " & tbCodCategEst.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            If Not rsElim.EOF Then
                aTabla(iCPareja, C_CONT_MARCAS_EST) = 10 - rsElim!puesto
            Else
                aTabla(iCPareja, C_CONT_MARCAS_EST) = 0
            End If
            rsElim.Close
            ' Por cada uno de los dorsales comprobamos el número de eliminatorias superadas
            'Latinos
            Set rsElim = db.OpenRecordset("SELECT DISTINCT fase FROM puntuaciones WHERE cod_categoria = " & tbCodCategLat.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            If Not rsElim.EOF Then
                iUltimaFaseLat = rsElim.Fields(0)
                rsElim.MoveLast
                aTabla(iCPareja, C_ELIM_SUP_LAT) = rsElim.RecordCount - 1
            Else
                aTabla(iCPareja, C_ELIM_SUP_LAT) = C_NO_PRESENTADO
                aTabla(iCPareja, C_PUESTO_LAT) = C_NO_PRESENTADO
                iUltimaFase = -1
            End If
            rsElim.Close
            'Estandar
            Set rsElim = db.OpenRecordset("SELECT DISTINCT fase FROM puntuaciones WHERE cod_categoria = " & tbCodCategLat.Text & " AND num_dorsal = " & rsDorsal!num_dorsal & " ORDER BY 1", dbOpenSnapshot)
            If Not rsElim.EOF Then
                iUltimaFaseEst = rsElim.Fields(0)
                rsElim.MoveLast
                aTabla(iCPareja, C_ELIM_SUP_EST) = rsElim.RecordCount - 1
            Else
                aTabla(iCPareja, C_ELIM_SUP_EST) = C_NO_PRESENTADO
                aTabla(iCPareja, C_PUESTO_EST) = C_NO_PRESENTADO
                iUltimaFase = -1
            End If
            rsElim.Close
            ' y el número de marcas de la última eliminatoria
            ' Primero comprobamos si en la última eliminatoria contamos con una repesca, ya que las marcas serán las de la repesca
            'Latinos
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE repesca = 1 AND fase = " & iUltimaFaseLat & " AND cod_categoria=" & tbCodCategLat.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If rs.Fields(0) > 0 Then
                iRepesca = 1
            Else
                iRepesca = 0
            End If
            rs.Close
            Set rsMarcas = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE puesto > 0 AND repesca=" & iRepesca & "AND fase = " & iUltimaFaseLat & " AND cod_categoria=" & tbCodCategLat.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If aTabla(iCPareja, C_CONT_MARCAS_LAT) = 0 Then
                If Not rsMarcas.EOF Then
                    aTabla(iCPareja, C_CONT_MARCAS_LAT) = rsMarcas.Fields(0)
                Else
                    aTabla(iCPareja, C_CONT_MARCAS_LAT) = 0
                End If
            End If
            rsMarcas.Close
            'Estandar
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE repesca = 1 AND fase = " & iUltimaFaseEst & " AND cod_categoria=" & tbCodCategEst.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If rs.Fields(0) > 0 Then
                iRepesca = 1
            Else
                iRepesca = 0
            End If
            rs.Close
            Set rsMarcas = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE puesto > 0 AND repesca=" & iRepesca & "AND fase = " & iUltimaFaseLat & " AND cod_categoria=" & tbCodCategLat.Text & " AND num_dorsal=" & rsDorsal!num_dorsal, dbOpenSnapshot)
            If aTabla(iCPareja, C_CONT_MARCAS_EST) = 0 Then
                If Not rsMarcas.EOF Then
                    aTabla(iCPareja, C_CONT_MARCAS_EST) = rsMarcas.Fields(0)
                Else
                    aTabla(iCPareja, C_CONT_MARCAS_EST) = 0
                End If
            End If
            rsMarcas.Close
            
            rsDorsal.MoveNext
            Inc iCPareja
        Wend
        rsDorsal.Close
        
        Printer.CurrentX = 100
        Printer.Print mml_FRASE0300;
        Printer.CurrentX = iEscala * 1
        Printer.Print mml_FRASE0741;
        Printer.CurrentX = iEscala * 2
        Printer.Print mml_FRASE0329;
        Printer.CurrentX = iEscala * 3
        Printer.Print mml_FRASE0691;
        Printer.CurrentX = iEscala * 4
        Printer.Print mml_FRASE0742;
        Printer.CurrentX = iEscala * 6
        Printer.Print mml_FRASE0743 & iParejasPart & mml_FRASE0472 & iParejas & mml_FRASE0744
        
        ' Ahora asignamos los puestos Latinos
        iCPuestos = 1
        Do While iCPuestos <= iParejasPart
            lValorMax = 0
            iCValorMax = 0
            'Recorremos todas las parejas
            For iCParejas = 0 To iParejas - 1
                ' Solo con los puestos no asignados
                If aTabla(iCParejas, C_PUESTO_LAT) = 0 Then
                    lValor = aTabla(iCParejas, C_ELIM_SUP_LAT) * 10 + aTabla(iCParejas, C_CONT_MARCAS_LAT)
                    If lValor > lValorMax Then
                        lValorMax = lValor
                        iPosValorMax = iCParejas
                        iCValorMax = 1
                    ElseIf lValor = lValorMax Then
                        Inc iCValorMax
                    End If
                End If
            Next iCParejas
            
            ' Asignamos el puesto y los puntos
            'Recorremos todas las parejas
            For iCParejas = 0 To iParejas - 1
                ' Solo con los puestos no asignados
                If aTabla(iCParejas, C_PUESTO_LAT) = 0 Then
                    lValor = aTabla(iCParejas, C_ELIM_SUP_LAT) * 10 + aTabla(iCParejas, C_CONT_MARCAS_LAT)
                    If lValor = lValorMax Then
                        aTabla(iCParejas, C_PUESTO_LAT) = iCPuestos
                        aTabla(iCParejas, C_PUESTO_FIN_LAT) = iCPuestos + iCValorMax - 1
                        'Puntuación entre los inscritos al cierre
                        aTabla(iCParejas, C_PUNTOS_LAT) = Redondea((iParejas - iCPuestos) / (iParejas - 1) / iCValorMax * 1000, 0)
                        
                    End If
                End If
            Next iCParejas
            
            If iCValorMax = 0 Then
                Exit Do
            End If
            iCPuestos = iCPuestos + iCValorMax
        Loop
        ' Ahora asignamos los puestos Estandar
        iCPuestos = 1
        Do While iCPuestos <= iParejasPart
            lValorMax = 0
            iCValorMax = 0
            'Recorremos todas las parejas
            For iCParejas = 0 To iParejas - 1
                ' Solo con los puestos no asignados
                If aTabla(iCParejas, C_PUESTO_EST) = 0 Then
                    lValor = aTabla(iCParejas, C_ELIM_SUP_EST) * 10 + aTabla(iCParejas, C_CONT_MARCAS_EST)
                    If lValor > lValorMax Then
                        lValorMax = lValor
                        iPosValorMax = iCParejas
                        iCValorMax = 1
                    ElseIf lValor = lValorMax Then
                        Inc iCValorMax
                    End If
                End If
            Next iCParejas
            
            ' Asignamos el puesto y los puntos
            'Recorremos todas las parejas
            For iCParejas = 0 To iParejas - 1
                ' Solo con los puestos no asignados
                If aTabla(iCParejas, C_PUESTO_EST) = 0 Then
                    lValor = aTabla(iCParejas, C_ELIM_SUP_EST) * 10 + aTabla(iCParejas, C_CONT_MARCAS_EST)
                    If lValor = lValorMax Then
                        aTabla(iCParejas, C_PUESTO_EST) = iCPuestos
                        aTabla(iCParejas, C_PUESTO_FIN_EST) = iCPuestos + iCValorMax - 1
                        'Puntuación entre los inscritos al cierre
                        aTabla(iCParejas, C_PUNTOS_EST) = Redondea((iParejas - iCPuestos) / (iParejas - 1) / iCValorMax * 1000, 0)
                        
                    End If
                End If
            Next iCParejas
            
            If iCValorMax = 0 Then
                Exit Do
            End If
            iCPuestos = iCPuestos + iCValorMax
        Loop
        
        'Imprimilos los dorsales
        For iCParejas = 0 To iParejas - 1
            ' Solo con los N.P.
            
            Printer.CurrentX = 100
            Printer.Print aTabla(iCParejas, C_NUM_DORSAL);
            Printer.CurrentX = iEscala * 1
            Printer.Print Trim$(aTabla(iCParejas, C_PUESTO_LAT));
            If aTabla(iCParejas, C_PUESTO_LAT) <> aTabla(iCParejas, C_PUESTO_FIN_LAT) Then
                Printer.Print "-" & aTabla(iCParejas, C_PUESTO_FIN_LAT);
            End If
            Printer.Print " , ";
            Printer.Print Trim$(aTabla(iCParejas, C_PUESTO_EST));
            If aTabla(iCParejas, C_PUESTO_EST) <> aTabla(iCParejas, C_PUESTO_FIN_EST) Then
                Printer.Print "-" & aTabla(iCParejas, C_PUESTO_FIN_EST);
            End If
            Printer.CurrentX = iEscala * 2
            dPuestoLat = Val(aTabla(iCParejas, C_PUESTO_LAT)) + (Val(aTabla(iCParejas, C_PUESTO_FIN_LAT)) - Val(aTabla(iCParejas, C_PUESTO_LAT))) / 2
            dPuestoEst = Val(aTabla(iCParejas, C_PUESTO_EST)) + (Val(aTabla(iCParejas, C_PUESTO_FIN_EST)) - Val(aTabla(iCParejas, C_PUESTO_EST))) / 2
            Printer.Print Redondea((dPuestoLat + dPuestoEst) / 2, 1);
            Printer.CurrentX = iEscala * 3
            Printer.Print Format((Val(aTabla(iCParejas, C_PUNTOS_EST)) + Val(aTabla(iCParejas, C_PUNTOS_LAT))) / 2, "####");
            
            Set rs = db.OpenRecordset("SELECT nombre_hombre, nombre_mujer FROM parejas WHERE codigo =" & aTabla(iCParejas, C_COD_PAREJA), dbOpenSnapshot)
            If Not rs.EOF Then
                Printer.CurrentX = iEscala * 4
                Printer.Print rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer;
            End If
            Printer.Print
            rs.Close
            If aTabla(iCParejas, C_PUESTO_LAT) = C_NO_PRESENTADO Then
                Printer.CurrentX = 100
                Printer.Print aTabla(iCParejas, C_NUM_DORSAL);
                Printer.CurrentX = iEscala * 1
                Printer.Print mml_FRASE0704;
                Set rs = db.OpenRecordset("SELECT nombre_hombre, nombre_mujer FROM parejas WHERE codigo =" & aTabla(iCParejas, C_COD_PAREJA), dbOpenSnapshot)
                If Not rs.EOF Then
                    Printer.CurrentX = iEscala * 3
                    Printer.Print rs!nombre_hombre & mml_FRASE0035 & rs!nombre_mujer;
                End If
                Printer.Print
                rs.Close
            End If
        Next iCParejas
        
        Printer.EndDoc
    Next iCCopias
End Sub




Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCateg_Click()
    tbCodComp.Text = sSeleccionar("SELECT * FROM competiciones ORDER BY 1")
    tbDescComp.Text = sResultado(2)
End Sub


Function CadPuesto(iPuesto As Integer) As String
    If iPuesto > C_PUESTO_NEG Then
        CadPuesto = Trim$(Str$(iPuesto - C_PUESTO_NEG)) & "d"
    Else
        CadPuesto = iPuesto
    End If
End Function

Private Sub CommandButton1_Click()
    If tbCodComp.Text = "" Then
        MsgBox mml_FRASE0320, vbOKOnly Or vbInformation, mml_FRASE0096
        Exit Sub
    End If
    tbCodCategEst.Text = sSeleccionar("SELECT * FROM Categorias c WHERE cod_modalidad = 2 AND cod_competicion = " & tbCodComp.Text & " ORDER BY " & G_ORDEN_CATEGORIAS)
    tbDescCategEst.Text = sResultado(2)
End Sub

Private Sub Form_Load()
    TraducirCadenas Me

End Sub
