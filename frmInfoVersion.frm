VERSION 5.00
Begin VB.Form frm1InfoVersion 
   Caption         =   "I.R.I.S."
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13860
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   13860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "mml_FRASE0886"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5280
      TabIndex        =   1
      Top             =   9120
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   8925
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   13785
      Begin VB.TextBox tbMejoras 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8715
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   150
         Width           =   13635
      End
   End
End
Attribute VB_Name = "frm1InfoVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    TraducirCadenas Me
    
    Me.Caption = " I.R.I.S. v" & App.Major & "." & App.Minor & " b" & App.Revision
    tbMejoras.Text = ""
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.33 (17/04/18) {�233}      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Eliminaci�n de errores con menos de tres participantes en modo medallas." & vbCrLf & vbCrLf & _
    "2. Cambiar la marcha C.Honkey a modalidad." & vbCrLf & vbCrLf & _
    "3. Cambio de los rangos de categor�a de edad." & vbCrLf & vbCrLf
        
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.31 (10/10/17) {�220}      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Se a�ade la exportaci�n de puntuaciones. la variable de configuraci�n 'categorias_exportar_puntuaciones' define las categor�as a exportar" & vbCrLf & vbCrLf

    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.30 (10/10/17) {�220}      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Permitir seleccionar la pista en el horario del coordinador de pista" & vbCrLf & vbCrLf & _
    "2. El horario del coordinador de pista se imprime en horizontal" & vbCrLf & vbCrLf & _
    "3. En el listado de participantes por nombre se ha ampliado el campo categor�a para que no solape con los nombres y se imprime en horizontal" & vbCrLf & vbCrLf & _
    "4. Se cambia el listado de resumen de clasificaci�n para que muestre el club" & vbCrLf & vbCrLf & _
    "5. Se corrige error en agrupaciones por la que desaparec�an grupos debido a coincidencias parciales de la categor�a" & vbCrLf & vbCrLf & _
    "6. Se mejora la funcionalidad de movimiento de grupos en el horario. Si no hay espacio en el destino avisa y cancela la operaci�n" & vbCrLf & vbCrLf & _
    "7. Corregido el error del fallo de b�squeda en la pantalla de competiciones (PTE)" & vbCrLf & vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.20 (04/10/17) {�220}      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Incorporaci�n de nuevo sistema de c�lculo con medias de los jueces que genera �nicamente 3 posiciones en las finales para country" & vbCrLf & vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.17 (03/11/16)      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Modificaciones internas que permiten ampliar el c�digo de las parejas por encima de 32768 parejas introducidas" & vbCrLf & vbCrLf

    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.16 (03/11/16)      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Se incorpora el par�metro 'impresora_por_defecto' para fijar la impresora por defecto. Colocar NA para que abra el cuadro de selecci�n de impresoras" & vbCrLf & _
    "2. Permitir cambiar la fase de un grupo en el horario a general look" & vbCrLf & _
    "3. Permitir desplazar correctamente grupos en el horario cuando este ocupe varias p�ginas" & vbCrLf & _
    "4. Corregir el error de impresi�n de horario de m�ltiples pistas" & vbCrLf & _
    "5. Corregir el error de asignaci�n de pista a fases ficticias de separaci�n. Actualmente cambiaba de pistas todas las fases ficticias" & vbCrLf & _
    "6. Corregir el error de fecha en la incorporaci�n de fases de separaci�n en el horario"

    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.15      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Cambiar tama�os de columnas de cuatro informes (OK)" & vbCrLf & _
    "2. Generar resumen con las parejas que no pasan de ronda (OK)" & vbCrLf & _
    "3. En las agrupaciones donde repite el grupo de edad para distintas categor�as, se unifica (OK)" & vbCrLf & _
    "4. Cuando se realizan agrupaciones de categorias debe permitir agrupar las que comienzan por la misma letra (OK)" & vbCrLf & _
    "5. Generar fichero ZIP con el nombre de la competici�n y dentro una carpeta Inet y dentro los PDF de resultados y el HTML (OK)" & vbCrLf & _
    "6. Permitir configurar el n�mero m�nimo de jueces para descalificar una pareja (OK)" & vbCrLf & _
    "7. Listado por rondas con informaci�n de dorsales y rondas en las que baila cada dorsal (PTE)" & vbCrLf & _
    vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.14      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Nueva funcionalidad que informa de la categor�a, fase, repesca, juez y baile que tiene cargado cada PDA (PTE)" & vbCrLf & _
    "2. Nueva funcionalidad que almacena en cookies todas las puntuaciones de las �ltimas 300 categor�as (Este valor dependde del navegador utilizado) (PTE)" & vbCrLf & _
    vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.13      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Corregido el error por el que calcula mal las medias de par�metros y bailes en el MatchAnalysis" & vbCrLf & _
    "2. Cambiado el literal Media de las hojas de puntuciones de bailes por total" & vbCrLf & _
    vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.12      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Modificada la importaci�n de parejas de pasos para que localice la posici�n de cada una de las columnas y permita columnas que no existan columnas" & vbCrLf & _
    "2. Corregido el error por el que introduc�a parejas en importaci�n pasos para modalidades para las que no exist�an las columnas" & vbCrLf & _
    "3. En importaci�n de parejas pasos se introduce una nueva pantalla de visualizaci�n de errores de importaci�n para grupos de edad y niveles" & vbCrLf & _
    "4. El sistema recupera la informaci�n de n�mero de segundos m�nimo entre c�lculos en la pantalla de enlace con PDA para que sea necesaria salir de IRIS al realizar cambios" & vbCrLf & _
    "5. MatchAnalysis. Mejora de la pantalla de puntuaciones. Se visualiza el nombre del dorsal y se recargan con todos los valores los combos para modificar puntuaciones" & vbCrLf & _
    "6. MatchAnalysis. Se modifica el m�dulo web para que permita elegir el alto de las filas de las barras de puntuaciones hasta ajustarlas con el alto de la barra" & vbCrLf & _
    "7. MatchAnalysis. Se cambia la cabecera Suma por Media en la tabla de puntuaciones totales" & vbCrLf & _
    "8. MatchAnalysis. Se incorpora la tabla de par�metros y descripciones en la hoja de puntuaciones" & vbCrLf & _
    vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.10.7      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Importaciones Pasos campeonato gallego, selecciona la modalidad concreta dependiento de la participaci�n de la pareja" & vbCrLf & _
    "2. Correcci�n de importaci�n de direcciones en importaci�n de pasos" & vbCrLf & _
    "3. Contabilizaci�n de las puntuaciones de dorsales no presentes para c�lculo autom�tico. Esta nueva funci�n permite realizar un c�lculo autom�tico, aunque no se hayan sincronizado los no presentes, se activa con la variable de configuraci�n 'control_puntuaciones_no_presentes'" & _
    "4. Corregido el efecto colateral por el que fallaba en control de asignaci�n de todas las puntuaciones antes del env�o en finales" & _
    "5. Se elimina la funcionalidad de no presente en el PDA. En caso de no sincornizar los no presentes en una final debe indicarse a los jueces que se marque como �ltimo, IRIS elimina sus puntuaciones al recuperarlas" & _
    vbCrLf
    
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.10      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Permitir varias competiciones con varias pistas de modo simult�neo con PDAs. Las pistas de todas las competiciones deben ser distintas" & vbCrLf & _
    "2. La suma de puntuaciones no puede contabilizar el juez de pasos" & vbCrLf & _
    "3. La PDA no puede escoger el juez de pasos" & vbCrLf & _
    "4. Corregir el desajuste est�tico por el que cuando la barra de botones de los PDAs es mayor que la zona de puestos se desajusta la pantalla" & vbCrLf & _
    "5. Si se produce un error en el env�o y se vuelve a la pantalla de puntuaciones, el sistema puede recuperar la informaci�n de pantalla en el momento del error. Se introducen nuevas opciones de recarga (ca: cookie actual, se graba siempre que cambian puntuaciones, ce: cookie enviada, se guarda siempre que se envian puntuaciones, re, ra, recargan la pantalla con la informaci�n de las cookies" & vbCrLf & _
    "6. Nueva funcionalidad de importaci�n de parejas de pasos. El fichero debe estar en formato ANSI. Se ha a�adido una variable de cfg formato_impt_pasos en el que se introduce el formato del nombre de los participantes" & vbCrLf & _
    "7. Filtro de variables de configuraci�n por texto" & vbCrLf & _
    "8. Calcular autom�ticamente el �ltimo grupo " & vbCrLf & _
    "9. Se prepara IRISMobile para 10 bailes que se deben bailar dos grupos de 5 bailes" & vbCrLf & _
    "10. Nuevo par�metros para dispositivos m�viles que permite establecer el ancho m�nimo de la tabla de puntuaciones para evitar desajustes con el �ltimo dorsal por tama�o excesivo del nombre del juez o categor�a" & vbCrLf & _
    vbCrLf

    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Mejoras versi�n 2.7      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf & _
    "1. Se a�ade la nueva funcionalidad de impresi�n de diplomas adaptada para bailes de sal�n. Se a�aden variables de cfg para imprimir o no, la posici�n de los bailes 'imprimir_posicion_bailes_diplomas' y definir el n�mero del primer puesto 'diplomas_primer_dorsal_salon'" & vbCrLf & _
    "2. Se a�ade una nueva funci�n a la pantalla de diplomas para imprimir s�lamente las categor�as agrupadas en un grupo, empleable para las m�ltiples entregas de dorsales de sal�n." & vbCrLf & _
    "3. Se a�ade la funcionalidad de soporte de MAthAnalisys con conexi�n tablets y m�dulo de puntuaciones directo e impresi�n" & vbCrLf & _
    "4. Se a�ade nueva variable de cfg que permite definir la altura de los gr�ficos en la cabecera 'cabecera_marge_y_imagenes'" & vbCrLf & _
    "5. Se modifica el m�dulo de puntuaciones de las tablets para que permita la recepci�n �nicamente de los dorsales con puntuaci�n. Se crea la variable de configuraci�n que activa esta opci�n 'html_solo_dorsales_presentes'" & vbCrLf & _
    "6. Se A�ade la posibilidad visualizaci�n de resultados en alto contraste (blanco o amarillo sobre negro), configurable con la variable 'publicacion_alto_contraste'" & vbCrLf & _
    vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Pruebas realizadas y mejoras en la Versi�n 2.5      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    " 1. (PDA) Nueva aplicaci�n de gesti�n de puntuaciones para PDA basada en Web y adaptada a Pocket 2003, Mobile 5 y 6 y iPod Touch (OK)" & vbCrLf & _
    " 2. (PC) Nuevas pantallas de control de puntuaciones de PDAs con el nuevo sistema Web (OK)" & vbCrLf & _
    " 2. (PDA) Nuevo sistema de gesti�n de informaci�n de conexi�n y bateria (OK)"

    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Pruebas realizadas y mejoras en la Versi�n 2.4      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    " 1. (PC) Bot�n para colocar el combo de jueces al principio en la pantalla de Bailes y Jueces (OK)" & vbCrLf & _
    " 2. (PC) Correcci�n del error por el que fallaba la aplicaci�n si intentabas quitar un juez y no hab�a ninguno (OK)" & vbCrLf & _
    " 3. (PC) Se ha a�adido un contador de jueces en la pantalla de Bailes y Jueces (OK)" & vbCrLf & _
    " 4. (PC) Se ha modificado el campo que almacena el �rden actual del horario para que soporte m�s de tres d�gitos (OK)" & vbCrLf & _
    " 5. (PC) Corregido el error por el que no se pod�a mover un grupo en el horario despu�s de la primera posici�n (OK)" & vbCrLf & _
    " 6. (PC) Corregido el error por el que ped�a confirmaci�n al imprimir las hojas de jueces ordenadas seg�n el horario (OK)" & vbCrLf & _
    " 7. (PC) Opci�n que permite introducir directamente en dorsales la ultima pareja introducida en la competici�n (OK)" & vbCrLf & _
    " 8. (PC) Opci�n que permite cambiar un juez incluso si ya hay puntuaciones introducidas (OK)" & vbCrLf & _
    " 9. (PC) Corregido el error en la introducci�n de los nombres de las parejas (OK)" & vbCrLf & _
    " 10.(PC) Cada vez que se elimina un dorsal de Dorsales se genera una l�nea de anotaci�n en el fichero indicado por la variable de configuraci�n 'fichero_eliminados' (OK)" & vbCrLf & _
    " 11.(PC) Se a�ade una opci�n al men� Archivo para ver el archivo de eliminaciones de dorsales (OK)" & vbCrLf & _
    " 12.(PC) Opci�n que permite a�adir una leyenda a los dorsales para diferencias los dorsales de distintas competiciones definidas bajo un mismo nombre (OK)" & vbCrLf & _
    " 13.(PC) Se han inclu�do nuevas opciones en los listados para poder discernir por categoria y pode obtener paneles de jueces por categoria o listados de participantes de una solo categor�a (OK)" & vbCrLf & _
    " 14.(PC) Se ha a�adido una pantalla de b�squeda de participantes por nombre o dorsal enlazada en la pantalla de dorsales. Al hacer doble clic enlaza con dorsales (OK)" & vbCrLf & _
    " 15.(PC) Se ha inclu�do una pantalla de gesti�n de paneles de jueces (OK)" & vbCrLf & _
    " 16.(PC) Se ha introducido el control de los datos transmitidos entre PC y PDAs en las pantallas de enlace con los PDAs (OK)" & vbCrLf & _
    " 17.(PC) Se ha introducido un nuevo bot�n que permite cerrar todos los ficheros abiertos por IRIS en caso de fallo de funcionamiento y de quedar alg�n fichero bloqueado (OK)" & vbCrLf & _
    " 18.(PC) En las pantalla de enlace con los PDAs se ha a�adido un bot�n en la esquina superior derecha que permite bloquear la pantalla para que no se puede tocar accidentalmente (OK)" & vbCrLf & _
    " 19.(PC) En la pantalla inicial hay un nuevo bot�n en la esquina superior derecha que permite pasar a modo PDA de forma que se diferencia perfectamente la aplicaci�n que est� en modo PDA y la que no (OK)" & vbCrLf & _
    " 20.(PDA) Se ha verificado la modificaci�n en los PDAs por la que no es posible enviar dos veces el mismo baile. Antes, al transmitir el baile se mostraba un cuadro que indicaba que los datos estaban transmitidos, si no pulsabamos Ok y pulsabamos fuera del cuadro, �ste desaparec�a y permit�a volver a transmitir el mismo baile. Ahora se bloquea la aplicaci�n hasta pulsar el Ok (OK)" & vbCrLf & _
    " 21.(PC) Nuevo bot�n en dorsales que permite imprimir un solo dorsal en la impresora por defecto solo preguntando por el n�mero de dorsal y una leyenda(OK)" & vbCrLf & _
    " 22.(PC) En la pantalla de introducci�n de datos manual y de PDA se ha inclu�do en el bot�n que muestra las puntuaciones introducidas, el n�mero de puntuaciones introducidas por cada juez y baile para comprobar desajustes de no presentes o eliminados (OK)" & vbCrLf & _
    " 23.(PC) Nueva funcionalidad de recarga autom�tica seleccionable en la configuraci�n inicial del PDA (OK)" & vbCrLf & _
    vbCrLf

    tbMejoras.Text = tbMejoras.Text & _
    " 23.(PC) Se ha incrementado el refresco de los controles de bater�a para que en caso de coincidir con un refresco del PDA y desaparecer la informaci�n de bater�a, que solo desaparezca por 4s (OK)" & vbCrLf & _
    " 24.(PC) Se ha solucionado el problema de solapamiento de las tablas de c�lculos en las impresiones de finales con 9 jueces y 6 puestos o m�s (OK)" & vbCrLf & _
    " 25.(PC) Se ha solucionado el problema de tama�o de las tablas de c�lculos en las impresiones de las semifinales (OK)" & vbCrLf & _
    vbCrLf


    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Pruebas realizadas y mejoras en la Versi�n 2.3 (10/12/2007)     " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    " 1.  (PC) Errores en cambio de Tama�o y recolocaci�n de la pantalla principal corregido. (OK)" & vbCrLf & _
    " 2.  (PC) Se incorpora la fecha y hora de generaci�n en listados de puntuaciones y resumen. (OK)" & vbCrLf & _
    " 3.  (PC) Opci�n 'Copiar Datos' para realizar copia de seguridad de la base de datos y ficheros de puntuaciones preguntando el nombre del directorio de copia. (OK)" & vbCrLf & _
    " 4.  (PDA)  Crear un prog en pda para comprobar la conexion (CheckPPC), q diga si esta activo y contra q pc (Opci�n 'i' del men� de salida de IRIS_PPC)(OK)" & vbCrLf & _
    " 5.  (PDA)  Crear un prog en pda q sea capaz de indicar si hay algun fichero pendiente de envio al pc, q pueda reenviarlo o visualizar su contenido" & vbCrLf & _
    " 6.  (PC) Nueva funcion en pc q permita chequear la conexi�n con todos los pdas (OK, control bateria)" & vbCrLf & _
    " 7.  (PDA) Al pasar automatico al mosaico vuelve a vertical en el sig baile sin comprobar 8 dorsales" & vbCrLf & _
    " 8.  (PC) Si imprimes las que pasaron sin calcular da una div por cero. Corregido (OK)" & vbCrLf & _
    " 9.  (PC) A�adir la asignacion de fase en el horario (OK)" & vbCrLf & _
    " 10. (PC) En el textbox de control de jueces sustituir los espacios por guiones bajos (OK)" & vbCrLf & _
    " 11. (PC) Colocar una pantalla de log en enlaceppc para errores controlados (OK)" & vbCrLf & _
    " 12. (PDA) Evitar q un juez puedan transmitir el mismo baile varias veces.  (OK)" & vbCrLf & _
    " 13. (PC) Solo se comprueban los ficheros de bateria del panel activo y los q se actualizan bien (OK)" & vbCrLf & _
    " 14. (PC) Copiar dorsales a otra categoria (OK)" & vbCrLf & _
    " 15. (PC) Bot�n que permita que se inserte directamente al final del horario una categoria y fase (OK)" & vbCrLf & _
    " 16. (PC) Opci�n de borrado de datos de prueba de una competici�n" & vbCrLf & _
    " 17. (PC) Borrado de multiples categorias en el horario (OK)" & vbCrLf & _
    " 18. (PC) Opci�n 'Copia Completa' Boton para generar una copia de la base de datos, preguntar si se quiere hacer copia de ficheros de pda. Esto debe ser un script que no interfiera con el programa. En cada copia se genera un nuevo directorio con el codigo y descripcion de la competicion sustituyendo lOs espacios por guiones y la fecha y la hora (OK)" & vbCrLf & _
    " 19. (PC) Opci�n 'Borrar Todo' Cuando eliminas todas las puntuaciones, eliminar publicaciones ,ficheros inet y ficheros pda y pedir si se quiere copia. Tambien eliminar publicaciones" & vbCrLf & _
    " 20. (PC) Imprimir la fecha y la hora en las cabeceras de las hojas de puntuacion y de resumen (OK)" & vbCrLf & _
    " 21. (PC) Cuando una pantalla de Enlace con Pda's comprueba ficheros desactiva todas las pantallas de pda (OK)" & vbCrLf & _
    " 22. (PDA) La pantalla de selecci�n de jueces debe generar el fichero de bateria, pero sin identificar el juez pero identificando (OK)" & vbCrLf & _
    " 23. (PDA) Crear un fichero de configuracion con los modos de visualizacion, quitar dorsal,repetir dorsal y el identificador del pda (OK)" & vbCrLf & _
    " 24. (PENDIENTE) Comprobar todos los modos para una final con m�s de dos dorsales no presentes"
    tbMejoras.Text = tbMejoras.Text & vbCrLf & _
    " 25. (PDA) Cambiar la 'r' de rondas por 'h' de hits" & vbCrLf & _
    " 26. (PDA) En la pantalla de pda de selecci�n de tandas debe aparecer se�aladas las que tienen punteos amarillos en amarillo, los azules en azul y las demas q tienen verdes en verde, las q no tienen nada en rojo claro. Debe aparecer informacion de marcas por tanda, totales y restantes" & vbCrLf & _
    " 27. (PC) Comprobar la introduccion de puntuaciones manual para q no cambie juez y categoria cuando se usa el grid. Comprobar lo q hay que hacer para introducir todos los jueces de un solo baile, utilizando el check del menu por juez (OK)" & vbCrLf & _
    " 28. (PENDIENTE) Comprobar el enlace con prodance" & vbCrLf & _
    " 29. (PC) Boton de asignar puesto en el grid de puntuacion de la introduccion de puntuacion manual.pc" & vbCrLf & _
    " 32. (PC) Cerrar los archivos abierto en caso de error de lectura para evitar posibles bloqueos (OK)" & vbCrLf & _
    " 33. (PDA) Al entrar en el pda comprobat si hay algun fichero temporal pendiente de tx, si lo hay solicita la transmision, y si no se quiere transmitir lo borra (OK)" & vbCrLf & _
    " 34. (PDA) Al tx, ahora comprueba q no haya nada pendiente, pero obliga a transmitirlo" & vbCrLf & _
    " 35. (PDA) Checkppc debe comprobar si hay ficheros c_tmp pendientes de envio (OK)" & vbCrLf & _
    " 36. (PDA) Boton de sincronizacion de hora con los pda, al pulsarlo en los pda Leen el fichero de sincronizaci�n de hora y salen de la aplicaci�n (OK)" & vbCrLf & _
    " 37. (PC) Opcion para no imprimir los puntos (OK)" & vbCrLf & _
    " 38. (PC) A�adir la categoria S a la base de datos (OK)" & vbCrLf & _
    " 39. (PC) Generar en \TMP\PuntuacionesManuales ficheros con todas las puntuaciones manuales introducidas (OK)" & vbCrLf & _
    " 40. (PC) Bot�n en Dorsales que permita a�adir una categor�a y una fase al horario (OK)" & vbCrLf & _
    " 41. (PC) Boton de 'Mover despues de' que permite mover en el horario un grupo de categorias detr�s de otra (OK)" & vbCrLf & _
    " 42. (PC) Boton que permita modificar el sistema de introducci�n de puntuaciones de por baile a por juez en la pantalla de puntuaciones (OK)" & vbCrLf & _
    " 43. (PC) El bot�n Quitar de la pantalla de introducci�n de puntuaciones se cambia de funci�n y ahora quita todas las puntuaciones un juez y un baile (OK)" & vbCrLf & _
    " 44. (PC) Se a�ade el control en la pantalla de Definici�n de jueces/bailes para que el bot�n de quitar no quite un baile de una categor�a si ya tiene puntuaciones (OK)" & vbCrLf & _
    " 45. (PC) Se a�ade un bot�n en la pantalla de Definici�n de jueces/bailes para poder cambiar las puntuaciones de un baile introducido por error en una categor�a a otro (OK)" & vbCrLf & _
    " 46. (PDA) Opci�n de los PDA que obligue a introducir el n�mero exacto de puntuaciones solicitadas (OK)" & vbCrLf & _
    " 47. (PDA) Generaci�n del fichero de configuraci�n de los PDA. La primera l�nea contiene el c�digo de identificaci�n de dos letras y la segunda el nombre largo del PDA. Cuando el PDA se encuentre en la pantalla de selecci�n de jueces, su fichero de control de bater�a contendr� la identificaci�n del PDA en vez del nombre del juez (OK)" & vbCrLf & _
    " 48. (PDA) Corregido el error por el que dejaba enviar puntuaciones con muy pocas marcas en eliminatorias. Ahora en caso de no obligar a marcar todas, solo permite un env�o de un 25% menos de marcas en eliminatorias (OK)" & vbCrLf & _
    " 49. (PDA) Corregido el error que se produce en el contador de marcas de eliminatorias cuando se seleccionan dorsales por n�mero"
    tbMejoras.Text = tbMejoras.Text & vbCrLf & _
    " 50. (PDA) Corregido el error que se producia cuando se seleccionaba un dorsal con la pantalla de selecci�n por n�mero de dorsal y no se reflejaba en la de selecci�n conjunta (OK)" & vbCrLf & _
    " 51. (PDA) Opci�n de configuraci�n que permita que se deseleccione el punteo autom�ticamente al marcar un dorsal (OK)" & vbCrLf & _
    " 52. (PDA) Se coloca un mensaje de 'Enviando datos' mientras se transmiten los datos, para informarles del proceso y de que no deben realizar ninguna acci�n (OK)" & vbCrLf & _
    " 53. (PDA) Eliminar autom�ticamente el estado de anulado o no presente cuando se selecciona el dorsal (OK)" & vbCrLf & _
    " 54. (PDA) Sustituir el bot�n de baile anterior por el de doble marcado (OK)" & vbCrLf & _
    " 55. (PDA) Generar un fichero temporal de selecci�n en el raiz del directorio del PDA con las marcas actuales cada vez que se avanza de tanda (OK)" & vbCrLf & _
    " 56. (PDA) Recalcular din�micamente el n�mero de tandas que quedan con m�s marcas que la media para visualizar en la pantalla de tandas (OK)" & vbCrLf & _
    " 57. (PC) Modificamos el tama�o de la columna de categor�a en Agrupaci�nManual para que se visualicen enteros los nombres de los Open (OK)" & vbCrLf & _
    " 58. (PC) A�adir la posibilidad de seleccionar fase 1/512 en introducci�n de puntuaciones (OK)" & vbCrLf & _
    " 59. (PC) Corregido el error por el que no se pod�an seleccionar en introducci�n de puntuaciones manuales por dorsal, dorsales no visibles en pantalla (OK)" & vbCrLf & _
    " 60. (PC) Nueva funcionalidad en introducci�n parcial de puntuaciones de un juez, de forma que el sistema admita introducir parte de las puntuaciones de un baile por PDA y parte por PC con solo realizar dobles clics (OK)" & vbCrLf & _
    " 61. (PDA) Avisar si en una tanda se selecciona un n�mero de marcas distinto al solicitado (OK)" & vbCrLf & _
    " 62. (PDA) Colocar en un color distinto el bot�n de avance con salvado de datos temporal (OK)" & vbCrLf & _
    " 63. (PC) Se ha modificado la pantalla de resultados para que posibilite la visualizaci�n de hasta 1000 parejas que pasen una fase, antes solo soportaba 60 (OK)" & vbCrLf & _
    " 64. (PC) La variable G_LIM_SEMI_UNA_TANDA modifica din�micamente el n�mero de dorsales por tanda para una semifinal para convertir p.e. dos tandas con 13 dorsales porque pasa de 12 por tanda a una tanda de 13 (OK)" & vbCrLf & _
    vbCrLf

    tbMejoras.Text = tbMejoras.Text & _
    "******************************************************************************************************" & vbCrLf & _
    "Pruebas realizadas y mejoras en la Versi�n 2.2      " & vbCrLf & _
    "******************************************************************************************************" & vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    " 1. (A) Control de Consumo de bateria de los PDA" & vbCrLf & _
    " 2. (A) Al recargar una categor�a permite elegir juez en el PDA" & vbCrLf & _
    " 3. (B) Permitir cambiar los dorsales de cada grupo en reparto de dorsales por grupo" & vbCrLf & _
    " 4. Cuando se utilizan pistas el PDA y el PC deben poner y recoger los ficheros en el directorio de la pista adecuada" & vbCrLf & _
    " 5. Cada pista utiliza su propio directorio en PDA y PC" & vbCrLf & _
    " 6. Pantalla de recogida de dorsales y boton que controla si hay alg�n cambio de fase en el horario" & vbCrLf & _
    " 7. Se realizan varias pruebas con paneles con m�s de 13 jueces" & vbCrLf & _
    "    (Se adaptan las hojas de c�lculo de final para 25 jueces) (OK)" & vbCrLf & _
    "    (Se prueban fases eliminatorias con 17 jueces)            (OK)" & vbCrLf & _
    " 8. Se realizan pruebas con categorias con m�s de 400 dorsales" & vbCrLf & _
    " 9. Se realizan las modificaciones para soportar nombres de juez de dos caracteres, con lo que ahora se soportan" & vbCrLf & _
    "    paneles de hasta 125 jueces distintos" & vbCrLf & _
    "10. Se cambia la generaci�n de PDF's de internet para adaptarlo al driver CutePDF. Se a�ade espera configurable por la generaci�n de los PDF." & vbCrLf & _
    "11. Modificamos la visualizaci�n de la publicidad para que no genere errores aunque se le quiten ficheros de publicidad en funcionamiento" & vbCrLf & _
    "12. A�adimos la funcionalidad de exportaci�n de posiciones a ProBaile" & vbCrLf & _
    "13. Se modifica la importaci�n de dorsales de ProBaile (OK)" & vbCrLf & _
    "14. A�adimos el agente de supervisi�n de competiciones que realiza todo tipo de comprobaciones sobre una competici�n para comprobar si se ha definido correctamente (OK)" & vbCrLf
    
    tbMejoras.Text = tbMejoras.Text & _
    "15. A�adimos la opci�n de impresi�n de diplomas por modalidad (Se utiliza solo en Country) (OK)" & vbCrLf & _
    "16. Se permite seleccionar la impresi�n de una �nica hoja de puntuaciones con una sola tanda por categoria (OK)" & vbCrLf & _
    "17. Comprobamos que en dorsales avise si se introduce una incripci�n con modalidad distinta a la de la categor�a (OK)" & vbCrLf

    
End Sub

