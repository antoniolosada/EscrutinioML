Attribute VB_Name = "modFrmGeneral"
Global m_sSelec As String
Global m_sGroupBy As String

Public Function sSeleccionar(sSQL As String, Optional sSelec As String = "", Optional sGroupBy As String = "") As String

    If Not C_DEBUG Then On Local Error GoTo error
    m_sSelec = sSelec
    m_sGroupBy = sGroupBy
    
    frmSelec.adoSelec.ConnectionString = "DSN=Escrutinio"
#If PROTECCION Then
    If InStr(sSQL, mml_FRASE0049) > 0 Then
        sSQL = "SELECT * FROM competiciones WHERE (codigo = " & G_PCOD_COMP & " AND descripcion = '" & CalcularResultado(G_PDESCRIPCION) & "' AND escuela = '" & CalcularResultado(G_PESCUELA) & "' AND fecha = #" & Format$(CDate(CalcularResultado(G_PFECHA)), "mm/dd/yyyy") & "#) or" & _
                            " (codigo = " & G_5PCOD_COMP & " AND descripcion = '" & CalcularResultado(G_5PDESCRIPCION) & "' AND escuela = '" & CalcularResultado(G_5PESCUELA) & "' AND fecha = #" & Format$(CDate(CalcularResultado(G_5PFECHA)), "mm/dd/yyyy") & "#) or" & _
                            " (codigo = " & G_4PCOD_COMP & " AND descripcion = '" & CalcularResultado(G_4PDESCRIPCION) & "' AND escuela = '" & CalcularResultado(G_4PESCUELA) & "' AND fecha = #" & Format$(CDate(CalcularResultado(G_4PFECHA)), "mm/dd/yyyy") & "#) or" & _
                            " (codigo = " & G_3PCOD_COMP & " AND descripcion = '" & CalcularResultado(G_3PDESCRIPCION) & "' AND escuela = '" & CalcularResultado(G_3PESCUELA) & "' AND fecha = #" & Format$(CDate(CalcularResultado(G_3PFECHA)), "mm/dd/yyyy") & "#) or" & _
                            " (codigo = " & G_2PCOD_COMP & " AND descripcion = '" & CalcularResultado(G_2PDESCRIPCION) & "' AND escuela = '" & CalcularResultado(G_2PESCUELA) & "' AND fecha = #" & Format$(CDate(CalcularResultado(G_2PFECHA)), "mm/dd/yyyy") & "#) or" & _
                            " (codigo = " & G_1PCOD_COMP & " AND descripcion = '" & CalcularResultado(G_1PDESCRIPCION) & "' AND escuela = '" & CalcularResultado(G_1PESCUELA) & "' AND fecha = #" & Format$(CDate(CalcularResultado(G_1PFECHA)), "mm/dd/yyyy") & "#) ORDER BY 1"
        Debug.Print sSQL
    End If
#End If

    sSelecSQL = UCase(sSQL)
    frmSelec.adoSelec.RecordSource = sSQL & sGroupBy
    frmSelec.adoSelec.Refresh
    frmSelec.cbOrden.ListIndex = Val(VarCfg("orden_sel")) - 1

    frmSelec.Show 1
    sSeleccionar = sResultado(1)
    
    Exit Function
error:
    ProcesarError "sSeleccionar"

End Function

Sub DesplazarVisualizacionHoja(Y As Integer)
Static iPos As Integer
    If Y = 0 Then
        If iPos > 0 Then frmRecOptico.pbFicha.Top = iPos
    Else
        iPos = frmRecOptico.pbFicha.Top
        If Y > G_MAX_FILA_VIS_HOJA Then
            frmRecOptico.pbFicha.Top = frmRecOptico.pbFicha.Top - G_DESPLAZ_VIS_HOJA
        Else
            frmRecOptico.pbFicha.Top = G_VIS_HOJA_POS_INIC
        End If
    End If
    DoEvents: DoEvents: DoEvents: DoEvents:
End Sub

Sub Timers(bOpcion As Boolean)
    
    If bOpcion Then
        'Activar Timers - Solo se activan los que estaban activos
        If frmEnlacePPC.lblActTimer.Tag = "1" Then
            frmEnlacePPC.Timer1.Enabled = True
            frmEnlacePPC.lblActTimer.BackColor = vbGreen
        Else
            frmEnlacePPC.lblActTimer.BackColor = vbRed
        End If
    
        If frmEnlacePPC1.lblActTimer.Tag = "1" Then
            frmEnlacePPC1.Timer1.Enabled = True
            frmEnlacePPC1.lblActTimer.BackColor = vbGreen
        Else
            frmEnlacePPC1.lblActTimer.BackColor = vbRed
        End If
    
        If frmEnlacePPC2.lblActTimer.Tag = "1" Then
            frmEnlacePPC2.Timer1.Enabled = True
            frmEnlacePPC2.lblActTimer.BackColor = vbGreen
        Else
            frmEnlacePPC2.lblActTimer.BackColor = vbRed
        End If
    
        If frmEnlacePPC3.lblActTimer.Tag = "1" Then
            frmEnlacePPC3.Timer1.Enabled = True
            frmEnlacePPC3.lblActTimer.BackColor = vbGreen
        Else
            frmEnlacePPC3.lblActTimer.BackColor = vbRed
        End If
        
    Else 'Cuando los desactiva esta función pasan a color amarillo
        'Desactivar Timers
        frmEnlacePPC.Timer1.Enabled = False
        frmEnlacePPC.lblActTimer.BackColor = vbYellow
    
        frmEnlacePPC1.Timer1.Enabled = False
        frmEnlacePPC1.lblActTimer.BackColor = vbYellow
    
        frmEnlacePPC2.Timer1.Enabled = False
        frmEnlacePPC2.lblActTimer.BackColor = vbYellow
    
        frmEnlacePPC3.Timer1.Enabled = False
        frmEnlacePPC3.lblActTimer.BackColor = vbYellow

    End If

    frmEnlacePPC.lblActTimer.Refresh
    frmEnlacePPC1.lblActTimer.Refresh
    frmEnlacePPC2.lblActTimer.Refresh
    frmEnlacePPC3.lblActTimer.Refresh
End Sub

Function DistintaPista(sPista As String, Optional sForm As String = "frmEnlacePPC_HTML") As Boolean
Dim iCont As Integer
Dim i As Integer
Dim frm As Form


    iCont = 0
    For i = 0 To Forms.Count - 1
    
        If Forms(i).Name = sForm Then
            If Forms(i).Controls("cbPista").Text = sPista Then
                Inc iCont
            End If
        End If
    Next

    If iCont = 1 Then
        DistintaPista = True
    Else
        DistintaPista = False
    End If
End Function

