'----------------------------------------------------------------------------------------------
'Nombre de función:funAbrirWord
'   Descripción: Abrir un documento con word
'   Parámetros: NombreArchivo -> Nombre del documento ha abrir.
'   Devuelve: Boolean -> True si consigue abrir el documeto y false en cosa contrario
'   Dependencias: FileSysteObject -> para comprobar que existe el fichero y operaciones con archivos
'		  NombreFecha()-> Devuelve un string aleatorio para usar como nombre del documento.
'		  cRutaTemp -> Constante con Path donde se ubicara una copia del documento.	
'		  cRutaAcceso-> Constante con Path donde se ubican las plantillas.		  	
'		  gWord -> Variable global de tipo Word.Application
'----------------------------------------------------------------------------------------------

Function funAbrirWord(NombreArchivo As String) As Boolean
funAbrirWord = True 
On Error GoTo ControlErrores

Dim sNombreArchivoTemp As String
'// Crear una nombre de documento con ruta de acceso incluida
sNombreArchivoTemp = cRutaTemp & NombreFecha() & ".doc"

Dim fsofArchivo As File
Set fsoArchivo = New FileSystemObject

'// Crear una instancia del word
Set gWord = CreateObject("Word.Application")

'// Comprobar si existe la plantalli 
If Not fsoArchivo.FileExists(cRutaAcceso & NombreArchivo) Then
    MsgBox mml_FRASE0952, vbCritical, mml_FRASE0084
    funAbrirWord = False
    Exit Function
End If



'// Setear para crear una copia
Set fsofArchivo = fsoArchivo.GetFile(cRutaAcceso & NombreArchivo)

'// Comprobar si existe un archivo con el mismo nombre que se le dara a la copia del documento
If fsoArchivo.FileExists(sNombreArchivoTemp) Then fsoArchivo.DeleteFile sNombreArchivoTemp, True

'// Setear para crear una copia (crea el documento)
fsofArchivo.Copy sNombreArchivoTemp
    

gWord.Application.WindowState = wdWindowStateMaximize '// Minimizar el word
gWord.Application.Visible = True
gWord.Application.Documents.Open (sNombreArchivoTemp)

ControlErrores:
Select Case Err.Number
    Case 0
        '// no se ha generado ningún error.
    Case 70
        MsgBox mml_FRASE0953 & Chr(13) & _
               mml_FRASE0954, vbCritical
        funAbrirWord = False
    Case Else
        MsgBox mml_FRASE0955 & Chr(13) & _
                mml_FRASE0867, vbCritical
        End
End Select
On Error GoTo 0

End Function

//////////////////////////////////////////////////////////////////////////////////////////////////////

'----------------------------------------------------------------------------------------------
'Nombre de función:funBuscarRemplazarTextWordMarcos
'   Descripción: Busca y remplaza el text de un campo, incluso dentro de marcos 
'   Parámetros: CampoBusqueda ->Text a buscar 
'   	 	TextoInsertar -> Text a remplazar 	
'		MarcoSN ->Indica si la busqueda es dentro de un marco, para forzar la salida
'   Dependencias: gWord -> Variable global de tipo Word.Application
'----------------------------------------------------------------------------------------------
Function funBuscarRemplazarTextWordMarcos(CampoBusqueda As String, TextoInsertar As String, Optional MarcoSN As Boolean)
gWord.Selection.Find.ClearFormatting
gWord.Selection.Find.Replacement.ClearFormatting
gWord.Selection.Find.Text = CampoBusqueda

If TextoInsertar = Space(0) Then
        gWord.Selection.Find.Replacement.Text = ""
    Else
        gWord.Selection.Find.Replacement.Text = TextoInsertar
End If

gWord.Selection.Find.Execute Replace:=wdReplaceAll

gWord.Selection.MoveRight Unit:=wdCharacter, Count:=1 '//Deseleccionar el campo, para q no realize una busqueda en una selección

If MarcoSN Then
    '// Hay que salir del marco para que pueda seguir buscando en el documento
    gWord.Selection.Find.Forward = True
    gWord.Selection.Find.Wrap = wdFindContinue
End If
End Function

//////////////////////////////////////////////////////////////////////////////////////////////////////

'----------------------------------------------------------------------------------------------
'Nombre de función:funBuscarRemplazarTextWord
'   Descripción: Busca y remplaza texto
'   Parámetros: CampoBusqueda ->Text a buscar 
'   	 	TextoInsertar -> Text a remplazar 	
'   Dependencias: gWord -> Variable global de tipo Word.Application
'----------------------------------------------------------------------------------------------
Function funBuscarRemplazarTextWord(CampoBusqueda As String, TextoInsertar As String)
gWord.Selection.Find.Text = CampoBusqueda
gWord.Selection.Find.Execute '// Ejecutar el commando de busqueda
If TextoInsertar = Space(0) Then
        gWord.Selection.Text = " "
    Else
        
        gWord.Selection.Text = TextoInsertar
End If
gWord.Selection.MoveRight Unit:=wdCharacter, Count:=1 '//Deseleccionar el texto, para q no realize una busqueda en una selección
End Function

//////////////////////////////////////////////////////////////////////////////////////////////////////

'----------------------------------------------------------------------------------------------
'Nombre de función:wordInsLineaCampo
'   Descripción: Inserta un nuevo campo y puede crear una nueva fila de una tabla
'   Parámetros: NewLine -> Boolean donde se indica si se creara una nueva fíla
'   	 	indice -> Indice de la fila de la tabla pra añadir al nombre de campo a insertar.
		NameText -> Nombre del campo a insertar
'   Dependencias: gWord -> Variable global de tipo Word.Application
'----------------------------------------------------------------------------------------------
Function wordInsLineaCampo(NewLine As Boolean, indice As Integer, NameText As String)

If NewLine Then
      gWord.Selection.InsertRowsBelow 1 '// Insertar una nueva linea
      If gWord.Selection.Cells.Count > 1 Then gWord.Selection.MoveRight Unit:=wdCell
   Else
      gWord.Selection.MoveRight Unit:=wdCell '// Deselección
End If

'// Insertar Campo Formulario
gWord.Selection.FormFields.Add Range:=gWord.Selection.Range, Type:=wdFieldFormTextInput
gWord.Selection.PreviousField.Select
'//Poner nombre al campo y valor por defecto
With gWord.Selection.FormFields(1)
      .Name = NameText & indice
      .EntryMacro = ""
      .ExitMacro = ""
      .Enabled = True
      .OwnHelp = False
      .HelpText = ""
      .OwnStatus = False
      .StatusText = ""
      With .TextInput
            .EditType wdRegularText, NameText, "", Enabled
            .Width = 0
      End With
End With
'// Seleccionar el campo de la celda
gWord.Selection.Cells(1).Select
end Function

//////////////////////////////////////////////////////////////////////////////////////////////////////

'----------------------------------------------------------------------------------------------
'Nombre de función: EncabezadoPie 
'   Descripción: Activa o desactiva el encabezado o el píe de página
'   Parámetros: PieSN -> Boolean donde se indica si es el píe o el encabezado
'   	 	ActivarSN -> Boolean donde se indica si si hay que activar o desactivar
'   Dependencias: gWord -> Variable global de tipo Word.Application
'----------------------------------------------------------------------------------------------
Function EncabezadoPie( PieSN As Boolean, ActivarSN As Boolean)
If PieSN then
	If ActivarSN then
		gWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter  '//Activar el pie
	   Else
		gWord.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument 'Desactivar la cabecera y el pie
	Endif
    Else
	If ActivarSN then
		gWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '//Activar la cabecera
	   Else
   	 	gWord.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument 'Desactivar la cabecera 
   	 Endif
Endif
End Function


Sub CargarDatosListFact()
'// Posicionarse en la celda(1,1)
For i = 0 To 6
    gWord.Selection.Cells(1).Next.Select
Next
sqlFacturas.MoveFirst '// Posicionarse al principio del recordset 
'//Iniciar Proceso de pegado en word
For i = 1 To frmGestionFra.lvPresupuestos.ListItems.Count '// Recorrer toda el listview
    gWord.Selection.TypeText (sqlFacturas!CodFactura) '//Código Fáctura
    gWord.Selection.Cells(1).Next.Select '// Siguiente Campo
    gWord.Selection.Cells(1).Next.Select '// Siguiente Campo este siempre es blanco
    gWord.Selection.TypeText (sqlFacturas!NomCliente)
    gWord.Selection.Cells(1).Next.Select '// Siguiente Campo
    gWord.Selection.TypeText (sqlFacturas!direccion)
    gWord.Selection.Cells(1).Next.Select '// Siguiente Campo
    gWord.Selection.TypeText (sqlFacturas!Poblacion)
    gWord.Selection.Cells(1).Next.Select '// Siguiente Campo
    gWord.Selection.TypeText (sqlFacturas!FechaFactura)
    gWord.Selection.Cells(1).Next.Select '// Siguiente Campo
    gWord.Selection.TypeText (frmGestionFra.lvPresupuestos.ListItems(i).ListSubItems(10)) 'Campo que indica si está pagado
    gWord.Selection.Cells(1).Next.Select '// Siguiente Campo
    gWord.Selection.TypeText (frmGestionFra.lvPresupuestos.ListItems(i).ListSubItems(9) & " €")
    '// enviar información a la línea de status
    frmGestionFra.Informacion.Caption = mml_FRASE0868 & sqlFacturas!NomCliente
    DoEvents
    sqlFacturas.MoveNext
    '// Si no es fin del recordset insertar una nueva fila
    If Not sqlFacturas.EOF Then gWord.Selection.MoveRight Unit:=wdCell
        
Next
End Sub

//////////////////////////////////////////////////////////////////////////////////////////////////////

Sub EstablecerParametrosDocumento()
gWord.Application.Selection.Font.Size = 18
gWord.Application.ActiveDocument.PageSetup.Orientation = wdOrientLandscape
End Sub


