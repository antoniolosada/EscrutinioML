Attribute VB_Name = "modGeneral"
Option Explicit
'*********************************
'*********************************
Public Const C_DEBUG = True
'*********************************
Global Const C_ERROR_FECHA = True
'*********************************



Public Const PROTECCION = False
#Const PROTECCION = False
Global Const RETRASO_LICENCIA = 1200

#Const CONTROL_LICENCIA = False

' En caso de utilizar conexión ODBC el caracter del LIKE es %, sino *
Global Const scarLike = "*"

Global Const MAX_NUMERO = 2000000000

Global Const MAX_PDAS_BATERIAS = 19

Global Const C_MAX_SEG_CON = 40
Global Const C_BATERIA_MEDIA = 30
Global Const C_BATERIA_BAJA = 10

Global Const C_CONTROL_GRABACION_ACTIVO = False

Public Const COD_MODALIDAD_STD = 1
Public Const COD_MODALIDAD_LAT = 2
Public Const COD_MODALIDAD_COM = 3

Public Const MAX_BAILES = 10

Public Const C_DORSAL_INI_RENUMERACION = 10000
Public Const C_CONT_RENUM_INI = 200000
Public Const MAX_DORSALES = 702
Public Const COD_MAX_BAILES = 999

Public Const FORMATO_PROG_GRANDE = "FORMATO_PROG_GRANDE"
Public Const FORMATO_PROG_PEQUE = "FORMATO_PROG_PEQUE"

Global Const C_LIM_MSGBOX = 700

Global Const TIPO_LIC_TOTAL = "69"
Global Const TIPO_LIC_REC_OPTICO_Y_RESULTADOS = "70"
Global Const TIPO_LIC_RESULTADOS = "71"
Global Const TIPO_LIC_REC_OPT_SIN_RESULTADOS = "72"
Global Const TIPO_LIC_ESTANDAR = "73"

Global Const C_INI_CALCULO = -1
Global Const C_CARGANDO = 1

Global Const C_BASE_ORDEN_COUNTRY = "orden_country_"

Const G_PCOD_COMP = 10
Const G_PESCUELA = "Fu`yjfa!bh$=gikg#Lfyot"
Const G_PDESCRIPCION = "J""Qvt`epFsiieknrik Btqinro"
Const G_PFECHA = "14/02/07"

Const G_1PCOD_COMP = 11
Const G_1PESCUELA = "Fu`yjfa!bh$=gikg#Lfyot"
Const G_1PDESCRIPCION = "JKF$Ylogcr$_k Mcuïs"
Const G_1PFECHA = "27/02/07"

Const G_2PCOD_COMP = 12
Const G_2PESCUELA = "Fu`yjfa!bh$=gikg#Lfyot"
Const G_2PDESCRIPCION = "JKXwiffm#h`&Snwwkrgipt"
Const G_2PFECHA = "13/03/07"

Const G_3PCOD_COMP = 9
Const G_3PESCUELA = "Fu`yjfa!bh$=gikg#Lfyot"
Const G_3PDESCRIPCION = "JKF$Ylogcr$>unbgoht&df""Mswliòm"
Const G_3PFECHA = "17/01/07"

Const G_4PCOD_COMP = 13
Const G_4PESCUELA = "Fu`yjfa!bh$=gikg#Lfyot"
Const G_4PDESCRIPCION = "J""Qvt`epgi\ikcqk{g"
Const G_4PFECHA = "04/04/07"


Const G_5PCOD_COMP = 15
Const G_5PESCUELA = "Fu`yjfa!bh$=gikg#Lfyot"
Const G_5PDESCRIPCION = "JXXwiffm#h`&Pntueöu"
Const G_5PFECHA = "15/05/07"

Global g_lCodUltimaPareja As Long

Global CR As String
Global LF As String
Global dBase As Date
Global CreandoHorario As Integer

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long


Public Const vbAmarilloClaro = &HC0FFFF
Public Const vbRojoClaro = &HC0C0FF
Public Const vbNomodal = 0
Public Const C_TITULO_VENTANA_PRINCIPAL = "I.R.I.S."
Public Const C_TEAM_MATCH = "TMatch"

Public db As Database
Public ws As Workspace
Public sResultado(50) As String
Public sFecha As String

Public Const C_MAX_INTEGER = 32000

Public Const C_MIN_POR_BAILE_POR_DEFECTO = 7

Public Const C_TEAMMATCH_POR_CATEGORIA = 0
Public Const C_TEAMMATCH_TOTAL = 1
Public Const C_MAX_POS_TEAMMATCH = 5
'Constantes para el horario -> inicio_Sesion
Public Const C_INICIO_SESION = 1
Public Const C_NO_ACT_HORA = 2

Public Const C_MARGEN_RESULTADOS_Y = 70

Public Const C_TAM_FUENTE_TITULO_LIBRETA = 10

'Márgenes mínimos para las tablas de puntuaciones de las semifinales
Public Const G_MARGEN_ANCHO_TABLAS = 1500

Public Const C_SALIDA_PROXIMA = mml_FRASE0191
Public Const C_RESULTADOS_ELIMINATORIA = mml_FRASE0928
Public Const C_REC_OPTICO = mml_FRASE0929

Public Const C_DORSALES_TANDA_DEFECTO = 12

'Constantes de posición de la hoja extendida
Public Const C_LIM_HOJA_EXT = 740
Public Const C_FASE_GENERAL_LOOK = 99

Public Const INIC_JUEZ_EXT = 14
Public Const C_MAX_JUECES_EXT = 28
Public Const C_POS_CONTROL_CALIDAD_EXT = 25
Public Const C_MAX_CATEG_POR_LINEA_EXT = 24
Public Const C_ANCHO_ESPACIO = 150
Public Const C_MAX_MARCAS_X_EXT = 27
Public Const C_POS_FALLO_EXT = 17
Public Const C_POS_CUADRO_FIRMA_EXT = 20
Public Const C_POS_FIRMA_EXT = 20
Public Const C_POS_CONTROL_EXT = 34
Public Const C_POS_REPESCA_EXT = 36
Public Const C_POS_HOJA2_EXT = 38
Public Const C_MAX_TANDAS_EXT = 24
Public Const C_POS_BORRADOR_EXT = 22
Public Const C_POS_HOJA_EXT_EXT = 18
Public Const C_REC_POS_X_MARCA_BAILE_EXT = 25
Public Const C_REC_POS_X_MARCA_BAILE = 19
Public Const C_POS_HOJA_EXT = 26
Public Const C_MAX_DORSALES_ESPACIADOS = 36
Public Const C_MAX_LEN_DESC_CATEGORIA = 31 ' Tamaño máximo de la descripción de un grupo
Public Const C_LIM_TABLA_CON_PUESTOS = 14

Public Const noModal = 0
Global Const C_MAX_DORSALES_HOJA_OPTICA_FINAL = 7
Global Const C_MAX_MSGLOG = 800
Global Const C_MAX_PART_CATEG = 1000
Global Const C_ALTO_CELDA_TABLAA = 280
Global Const C_MARGEN_IZQ_TABLAA = 200
Global Const C_PUNTOS_NO_PRESENTADO = "---"
Global Const C_FUENTE_DEFECTO_TABLAA = 10

Global Const C_CODIGO_GENERAL_LOOK = 99

Global Const C_POS_MAY_ABS_NO_PRESENTADOS = 10

Global Const C_MIN_PUNTOS_MARCA_REFERENCIA = 10

Global Const C_MAX_DORSALES_LIST_PART = 12

Public Const C_ALTURA_LECT_OPTICA = 3015
Public Const C_ANCHURA_LECT_OPTICA = 10950

Public Const C_POS_BORRADOR = 16.5

Public Const C_HOJAS_PROC = mml_FRASE0930

Public Const C_COMBINAR_EDAD = 0
Public Const C_NO_COMBINAR_EDAD = 1

Public Const C_MARGEN_CONTROL_ERROR_MARCA = 40

Public Const G_MARGEN_SUPER = 20

Public Const C_POS_COLUMNA_2 = 5900

Public Const C_ULTIMO_PUESTO = 9

Public Const C_COLOR_INET1 = "Silver"
Public Const C_COLOR_INET2 = "Silver"

Public Const G_INFANTIL = 10
Public Const G_JUVENIL = 1
Public Const G_JUNIOR1 = 2
Public Const G_JUNIOR2 = 3
Public Const G_YOUTH = 4
Public Const G_ADULTO1 = 5
Public Const G_ADULTO2 = 6
Public Const G_SENIOR1 = 7
Public Const G_SENIOR2 = 8
Public Const G_SENIOR3 = 9

Public Const C_MAX_PUESTO_FINAL_NO_PRESENTADO = 10

Public Const C_MIN_PAREJAS_GRUPO = 5
Public Const C_NUM_GRUPOS_AGRUPADOS = 3
Public Const C_NO_HAY_GRUPO_ANTERIOR = -1

Public C_MARGEN_BUSQUEDA As Integer
Public Const C_MARGEN_BUSQUEDA_INIC = 6 ' Marcas de ancho al buscar la primera marca
Public Const C_MARGEN_BUSQUEDA_NORMAL = 3 ' Marcas de ancho

Public Const C_TAM_MARCA = 20
Public Const c_MARGEN_MARCA = 2
Public Const G_MARGEN_EPA = 760
Public Const G_MARGEN_EPA_Y = 240

Public Const C_MAX_MARCAS_X_NORMAL = 21
Public Const C_UMBRAL = 400
Public Const C_NUM_PUNTOS_BLANCOS = 3
Public Const C_TAM_MAX_MARCA_X = 25
Public Const C_TAM_MIN_MARCA_X = 15
Public Const C_TAM_MIN_MARCA_Y = 8
Public Const C_TAM_MARCA_X = 18
Public Const C_TAM_MARCA_Y = 10
Public Const C_MAX_MARCAS_Y = 52
Public Const C_MARGEN = 2
Public Const C_ANCHO_MARCAS_BAILE = 5
Public Const C_ANCHO_MARCAS_BAILE_JUEZ_PASOS = 5
Public Const C_ANCHO_MARCAS_BAILE_FINAL = 8

Global C_MAX_MARCAS_X As Integer

Public Const C_MAX_DORSAL_FINAL = 7
Public Const C_MAX_DORSAL_FINAL_EXTRA = 8
Public Const C_MIN_PAREJAS_SELEC_FINAL = 6

Public Const C_FUENTE_GRANDE_BOTON = 18
Public Const C_FUENTE_PEQUE_BOTON = 9

Public Const C_MARGEN_DESC = 50
Public Const POS_CONTROL = 24
Public Const POS_REPESCA = 26
Public Const POS_HOJA2 = 28
Public Const C_LIM_MIN_PTOS_NEGROS = 7

Public Const C_MARGEN_CELDA = 80

'Public Const C_COLOR_MARCA = &HC0FFC0 'Verde
Public Const C_COLOR_MARCA = &HC00000    ' Azul oscuro
'Public Const C_COLOR_X_MARCA = &HFFFFFF  ' Color de la X blanco
Public Const C_COLOR_X_MARCA = &H80FF80 ' Color de la X verde
Public Const C_COLOR_ROJO = &HFF&
Public Const C_COLOR_VERDE = &HFF00&
Public Const C_COLOR_VERDE_MEDIO = &H80FF80
Public Const C_COLOR_AMARILLO_MEDIO = &H80FFFF
Public Const C_COLOR_NEGRO = 0
Public Const C_COLOR_AZUL_OSCURO = &HC00000
Public Const C_COLOR_ROJO_OSCURO = &H80&


Global sSelecSQL As String

'Posición del formulario de lectura optica
Global iFormularioTop
Global iFormularioLeft
Global iFormularioHeight
Global iFormularioWidth

Global sExecSQL As String
Global sSQL As String
Global C_MARCA As Integer
Global C_BLANCO As Integer
Global C_MEDIO As Integer
Global ScaleFactor As Single
Global MAX_LIN_PAG As Integer
Global HAY_JUEZ_PASOS As Boolean
Global HAY_CONTROL As Boolean
Global C_CAR_DESC_MOD As Integer
Global C_ORDEN_PAREJAS As String
Global C_MIN_DORSAL_OFICIAL As Integer
Global C_SALTO_PUNTOS_MARCA As Integer
Global C_EXTENSION_FICHEROS As String
Global C_FICHERO_INET As String
Global C_LOGO_PATH As String
Global SALTAR_PUESTO_SIG As String
Global CONTROL_HORA As String
Global PANEL_RESULTADOS As String
Global C_PREGUNTAR_REPESCA As Boolean
Global C_PREGUNTAR_REPESCA_SIEMPRE As Boolean
Global C_REFRESCO_PUBLICIDAD As Boolean
Global G_MARCA_CONTROL As Boolean
Global G_MARCAR_BUS_MARCA As Boolean
Global G_MARCA_MAYOR As Boolean
Global G_ARCH_SOCIOS_ANULADOS As String
Global G_DORSALES_COMBINADOS As Boolean
Global G_PUBLICAR_POSICION As Boolean
Global G_PUB_RES_DEC_LETRAS As Integer
Global G_REC_OPTICO_PARCIAL As Boolean
Global G_LOGO_ESCUELA As String
Global G_ORDEN_CATEGORIAS As String
Global G_HOJA_EXTENDIDA As Boolean
Global G_DEC_POSICIONES_POR_DESCALIFICACION As Boolean
Global G_LINEAS_DIVISION_FINAL As Integer
Global G_SALTO_ORDEN As Integer
Global G_PUBLICAR_HORA_ESTIMADA As Boolean
Global G_RESULTADOS_UNO_A_UNO As Boolean
Global G_ESPERA_ENTRE_PART As Integer
Global C_MINUTOS_MARGEN_CALCULAR As Integer
Global G_DESPLAZAR_SI_CONTROL_NO_LOCALIZADO As Boolean
Global G_NO_CONTAR_HOJAS As Boolean
Global G_TAM_FUENTE_TABLA_SEMI As Integer
Global G_ANCHO_COL_JUEZ As Integer
Global G_REC_HOJA_EXT As Boolean
Global C_TOLERANCIA_X As Integer
Global C_TOLERANCIA_Y As Integer
Global C_REC_OPTICO_RAPIDO As Boolean
Global C_CAT_UNICA_HOJA_POR_BAILE As Boolean
Global C_MAX_DORSALES_BAILE_POR_COL As Integer
Global C_MAX_COLS_HOJA_POR_BAILE As Integer
Global C_MAX_DORSALES_BAILE_POR_COL_HOJA_UNICA As Integer
Global C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA As Integer
Global G_FUENTE_GRANDE_DORSAL As Integer
Global G_FUENTE_PEQUE_DORSAL As Integer
Global C_NUM_GRUPOS As Integer
Global G_ORDENAR_DESCALIFICADOS_FINAL As String
Global G_NOMBRE_FUENTE_DORSAL As String
Global G_DORSAL_POR_Y As Integer
Global G_MARGEN_DORSAL_X As Integer
Global G_LOGO_DORSAL_DER As String
Global G_LOGO_DORSAL_IZQ As String
Global G_PUBLICAR_NUM_GRUPOS_RESULTADOS As Integer
Global G_PUNTEO_ANULACION As Boolean
Global G_AUTO_IMP_HOJAS_PUNTUACION As Boolean
Global G_IMP_HOJAS_BAILE_EN_FINALES As Boolean
Global G_PISTAS_HOJAS_OPTICAS As String
Global C_DESC_SIN_PUESTO As Boolean
Global C_ORDEN_HORARIO_MODALIDAD As String
Global C_MAX_PAREJAS_FINAL As Integer
Global C_MAX_PAREJAS_SEMIFINAL As Integer
Global G_ORDEN_CATEG_COM As String
Global G_ORDEN_CATEG_EST As String
Global G_ORDEN_CATEG_LAT As String
Global G_AUTO_POR_DORSAL As Boolean
Global C_CALCULOS_PARCIALES As Boolean
Global G_MAX_DORSALES_HOJA_SEMI As Integer
Global G_ELIMINATORIAS_PAGINADAS As Boolean
Global G_IMPRIMIR_AVISO_GENERAL_LOOK As Boolean
Global G_MARGEN_CONTROL_PTE As Integer
Global G_SALTO_CATEG As Integer
Global G_PREGUNTAR_IMPRESION_AUTO As Boolean
Global G_MINUTOS_POR_CATEG As Integer
Global G_FICHERO_SALIDA As String
Global G_FICHERO_ENTRADA As String
Global G_FICHERO_ENTRADA_P2 As String
Global G_CAMBIO_AUTO As Boolean
Global G_CALCULO_AUTO_PPC As Boolean
Global G_INTERVALO_TIMER_PPC As Integer
Global G_APP_GRAFICA As String
Global C_PREGUNTAR_EDIC_HOJA As Boolean
Global G_IMAGENES_EPA_PEQUE As String
Global G_DIR_PRODANCE As String
Global G_DATOS_ORG_PRODANCE As String
Global G_SELEC_HOJA_EXT_AUTO As Boolean
Global G_MARGEN_RETRASO_INICIAL As Integer
Global C_SISTEMA_TANDAS_VIEJO As Boolean
Global C_NUM_JUECES_ACEPTAR_NO_PRESENTES As Integer
Global C_PREGUNTA_ACEPTAR_NO_PRESENTES As Boolean
Global G_PISTAS_PPC As String
Global G_GEN_AUTO_RESULTADOS_PPC As Boolean
Global G_TIEMPO_ESPERA_JPASOS_PPC As Integer
Global G_NO_MARCAR_BAILES As Boolean
Global C_RESET_ULTIMOS_5_BAILES As Boolean
Global G_ORDEN_10B_LAT_EST As String
Global G_VALORES_MULTIPLES_DBLCLIK As Boolean
Global G_IMPRIMIR_TODOS_LOS_CUADROS As Boolean
Global G_PATH_ESCRUTINIO As String
Global G_NO_PROC_HOJAS_ERROR As Boolean
Global G_PARAR_REC_SI_FALLO As Boolean
Global G_ORDEN_SEL As Integer
Global G_MARGEN_X_MARCA_CONTROL As Integer
Global G_PATH_GRAFICO_HOJAS As String
Global G_SELEC_DORSALES_SIG_FASE As Boolean
Global G_PPC_GEN_DORSALES_COMBINADOS As Boolean
Global G_COUNTRY As Boolean
Global C_BAILES_POR_HOJA As Integer
Global C_BAILES_POR_HOJA_UNICA As Integer
Global MAX_DORSALES_HOJA_UNICA  As Integer
Global MAX_DORSLES_HOJA_PUNT_COL_DOBLE As Integer
Global C_NO_COPIAS_COMBINACION As Integer
Global G_IMAGEN_FONDO_DIPLOMA As String
Global G_MARGEN_IZQ_TABLA_DIPLOMAS As Integer
Global G_MARGEN_SUP_TABLA_DIPLOMAS As Integer
Global G_MARGEN_IZQ_IMAGEN_DIPLOMAS As Integer
Global G_MARGEN_SUP_IMAGEN_DIPLOMAS As Integer
Global G_MARGEN_SUP_NOMBRE_COMP As Integer
Global G_MARGEN_IZQ_NOMBRE_COMP As Integer
Global G_ANCHO_IMAGEN_DIPLOMAS As Integer
Global G_ALTO_IMAGEN_DIPLOMAS As Integer
Global G_DIPLOMA_TITULO_FUENTE As String
Global G_ESPERA_FICH_INET As Integer
Global G_LIM_JUECES_PARA_TABLAS_FINAL_DOBLES As Integer
Global G_FICHERO_BATERIA As String
Global G_FICHERO_CONTROL_JUECES As String
Global G_FICHERO_HORA As String
Global G_TIEMPO_PARA_PERDIDA_DE_CONEXION As Long
Global G_CATEGORIA_COMBINACION As Integer
Global G_PUESTOS_CON_DIPLOMA As Integer
Global G_CADENA_SEPARADOR_CAMPOS As String
Global G_CADENA_DELIMITADOR_VALORES As String
Global G_TIMER_CONTROL_BATERIA As Integer
Global G_DIR_COPIA_BD As String
Global G_MOSTRAR_PUNTOS As Boolean
Global G_NO_REPETIR_PRIMERA_TANDA As Boolean
Global G_NO_PROCESAR_FICH_DE_OTRA_CATEG As Boolean
Global G_NO_TANDAS_NO_REPETIR_PRIMERA_TANDA As Integer
Global G_NO_PUBLICAR_COMO_ANT_PANEL_DERECHO As Boolean
Global G_RETARDO_RESULTADOS_MULTIPLES_PANTALLAS As Integer
Global G_LIM_SEMI_UNA_TANDA As Integer
Global G_ACTUALIZAR_ULTIMA_PUBLICACION As Integer
Global G_VELOCIDAD_LIBRETA As Integer
Global G_RETRASO_RESULTADOS As Integer
Global G_SOLO_UN_PC As Boolean
Global G_ASIGNAR_AUTOMATICAMENTE_LETRA_A_JUEZ As Boolean
Global G_FICHERO_ELIMINADOS As String
Global G_PREGUNTA_OPERACION As String
Global G_ADM_EQUIPOS As String
Global G_MOVER_FICHEROS_PDA As Boolean
Global G_RUTA_COPIA_FICH_PDA As String
Global G_SEG_MAX_CONTROL_BATERIA As Integer
Global G_UNIDADES_NIVEL_MIN_CONTROL As Integer
Global G_NO_PRESENTES_AUTO As Boolean
Global G_PAREJAS_AEBDC_UN_APELLIDO As String

Global G_ESPERA_NO_PUBLIC As Integer
Global G_SALTO_PUBLIC As Integer

Global G_MARGEN_MARCA_Y As Integer

Global G_PANEL_LAPIZ_OPTICO As Boolean
Global G_ULTIMO_PUESTO_AUTOMATICO As Boolean
Global G_POS_INIC_CATEG As Integer
Global G_ARCH_HORARIO As String
Global G_CAB_RESULTADOS As String
Global G_CAB1_RESULTADOS As String
Global G_DIR_PUBLICIDAD As String

Global G_AVISO_NUM_MARCAS As Boolean

Global G_CAB_INET As String
Global G_PIE_INET As String
Global G_CAB_TABLA As String
Global G_COLOR_MARCAS As String

Global G_ESCUELA As String
Global G_BUSCAR_NOMBRE As String

Global G_MARCAR_PUNTOS As Boolean
Global G_MOSTRAR_NUM_PUNTOS As Boolean
Global G_MAX_FILAS_POR_PAG As Integer

Global Const G_VIS_HOJA_POS_INIC = 16
Global G_DESPLAZ_VIS_HOJA As Integer
Global G_MAX_FILA_VIS_HOJA As Integer

Public Const C_ANCHO_LOGO = 3000
Public Const C_ANCHO_TABLA_NO_FINAL = 2100
Public Const C_POS_TABLA_JUECES_NO_FINAL = 6200

Public Const C_BAILES_POR_PAG_NO_FINAL = 5
Public Const C_BAILES_POR_PAG_FINAL = 5

Public Const C_PUESTO_NEG = 10

'Presentación de resultados

Public Const C_FUENTE_GRANDE_LIBRETA = 40
Public Const C_FUENTE_MEDIANA_LIBRETA = 30
Public Const C_FUENTE_PEQUE_LIBRETA = 22
Public Const C_FUENTE_MUY_PEQUE_LIBRETA = 10
Public Const C_MAX_INFO_RESULTADOS_LINEA_LIBRETA = 23
Public Const C_MAX_PAREJAS_PARA_FUENTE_GRANDE_LIBRETA = 6
Public Const C_MAX_PAREJAS_POR_COLUMNA_LIBRETA = 13
Public Const C_MAX_PAREJAS_POR_COLUMNA_LIBRETA_PEQUE = 27
Public Const C_MAX_DORSALES_PANT_COMPLETA_LIBRETA = 8

Public Const C_FUENTE_COMENTARIO = 14
Public Const C_FUENTE_TITULO_RESULTADOS = 17

Public Const C_FUENTE_GRANDE = 22
Public Const C_FUENTE_MEDIANA = 18
Public Const C_FUENTE_PEQUE = 10
Public Const C_MAX_INFO_RESULTADOS_LINEA = 23
Public Const C_MAX_PAREJAS_PARA_FUENTE_GRANDE = 10
Public Const C_MAX_PAREJAS_POR_COLUMNA = 25
Public Const C_MAX_DORSALES_PANT_COMPLETA = 14
Public Const C_MAX_CAR_LINEA_LIBRETA = 26
Public Const C_MAX_PAREJAS_PANTALLA = 52

Public Const C_LON_MAX_NOMBRE = 9
Public Const C_LONG_NOMBRE_LIBRETA = 9
Public Const C_MAX_CAR_POR_FILA = 22

Public Const C_MAX_DORSALES_HOJA_OPTICA = 18
Public Const C_MAX_DORSALES_HOJA_OPTICA_EXT = 24

Public Const C_MAX_JUECES_HOJA_OPTICA = 10
Public Const C_MAX_JUECES_HOJA_OPTICA_EXT = 24


Public Const POS_NUM_DORSAL = 0
Public Const OP_CRITERIOS = 1000
Public Const POS_TOTAL_CONJUNTO = 1
Public Const POS_COD_PAREJA = 2
Public Const VALOR_MAXIMO = 32000
Public Const BAILES_FINAL = 1
Public Const BAILES_NO_FINAL = 2
Public Const BD_POS_POSICION = 0
Public Const MARGEN_PAGINA = 2000
Public Const MARGEN_PAGINA_INF = 600
Public Const MIN_PAREJAS_FINAL = 5
Global MARGEN_SUPERIOR

Dim iCFichas As Integer

Enum EJustificado
    ecizquierda = 0
    eccentro = 1
End Enum

Type TValores
    Nombre As String
    valor As String
    operacion As String
End Type

Type TCelda
    Ancho As Integer
    Justificado As EJustificado
End Type

Type TDirectorio
    sDirectorio As String
    bNuevo As Boolean
End Type

Type TGrupo
    iGrupoEdad As Integer
    sCategoria As String
    iModalidad As Integer
    iTodasParejas As Integer
End Type

Type TMarca
    iXi As Integer
    iXf As Integer
    iYi As Integer
    iYf As Integer
End Type

Type TCal
    iNPos As Double
    iSum As Double
End Type

Type TCodDesc
    codigo As Integer
    DESCRIPCION As String
End Type

Public aTabla() As String
Public aMarcas(C_MAX_MARCAS_X_EXT + 1, C_MAX_MARCAS_Y + 1) As TMarca


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Function Inc(vVar As Variant)
    vVar = vVar + 1
End Function

Public Function Dec(vVar As Variant)
    vVar = vVar - 1
End Function

Public Function MaxCod(sTabla As String) As Long
Dim rs As Recordset

    Set rs = db.OpenRecordset("SELECT MAX(codigo) FROM " & sTabla, dbOpenSnapshot)
    If IsNull(rs.Fields(0)) Then
        MaxCod = 1
    Else
        MaxCod = rs.Fields(0) + 1
    End If
    rs.Close
End Function

Sub CamposSinCubrir()
    MsgBox mml_FRASE0277, vbOKOnly Or vbInformation, mml_FRASE0084
End Sub


Function VarCfg(sVar As String, Optional sValor As String, Optional sDesc As String = "", Optional sDescIngles As String = "") As String
Dim rs As Recordset
Dim sSQL As String

    
    Set rs = db.OpenRecordset("SELECT valor FROM cfg WHERE variable='" & sVar & "'")
        If Not rs.EOF Then
            VarCfg = IIf(IsNull(rs.Fields(0)), "", rs.Fields(0))
            rs.Close
        Else
            rs.Close
            If IsMissing(sValor) Then
                VarCfg = ""
                MsgBox mml_FRASE0931 & sVar & mml_FRASE0932, vbOKOnly Or vbCritical, mml_FRASE0096
            Else
                VarCfg = sValor
                If sDescIngles = "" Then sDescIngles = sDesc
                sSQL = "INSERT INTO cfg VALUES ('" & sVar & "','" & sValor & "','" & sDesc & "','" & sDescIngles & "','" & sValor & "')"
                Debug.Print sSQL
                db.Execute sSQL
            End If
        End If
End Function
Public Sub DibujarTablaA(f As Printer, iPosX As Integer, iPosY As Integer, iFilas As Integer, iCols As Integer, aTabla() As String, aDefCelda() As TCelda, iAltoCelda As Integer)
Dim iFila As Integer
Dim iCol As Integer
Dim sValor As String
Dim iSuper As Integer
Dim iSuper2 As Integer
Dim iTamFuente As Integer
Dim iArroba As Integer
Dim sTipoLetra As String

    
    For iFila = 0 To iFilas
        f.CurrentY = iPosY + (iFila * iAltoCelda)
        f.CurrentX = iPosX
        f.Line -Step(PosicionColX(iCols - 1, aDefCelda()) + aDefCelda(iCols - 1).Ancho, 0)
    Next iFila
    For iCol = 0 To iCols
        If iCol = 0 Then
            f.CurrentX = iPosX
        Else
            f.CurrentX = iPosX + (PosicionColX(iCol, aDefCelda()))
        End If
        f.CurrentY = iPosY
        f.Line -Step(0, iFilas * iAltoCelda)
    Next iCol
    For iFila = 0 To iFilas - 1
        For iCol = 0 To iCols - 1
            sValor = aTabla(iFila, iCol)
            If Left$(sValor, 1) = "." And Len(sValor) >= 4 Then
                iTamFuente = Val(Mid$(sValor, 2, 2))
                sTipoLetra = Mid$(sValor, 4, 1)
                sValor = Mid$(aTabla(iFila, iCol), 5)
            Else
                sTipoLetra = ""
                iTamFuente = C_FUENTE_DEFECTO_TABLAA
            End If
            
            f.FontSize = iTamFuente
            Select Case sTipoLetra
                Case "b"
                    f.FontBold = True
                Case "i"
                    f.FontItalic = True
                Case Else
                    f.FontBold = False
                    f.FontItalic = False
            End Select
            
            If aDefCelda(iCol).Justificado = eccentro Then
                    f.CurrentX = iPosX + PosicionColX(iCol, aDefCelda()) + (aDefCelda(iCol).Ancho - f.TextWidth(sValor)) / 2
            Else
                    f.CurrentX = iPosX + PosicionColX(iCol, aDefCelda()) + C_MARGEN_CELDA
            End If
            f.CurrentY = iPosY + (iFila * iAltoCelda) + (iAltoCelda - f.TextHeight(sValor)) / 2
            
            If aDefCelda(iCol).Ancho > 0 Then
                f.Print sValor
            End If
        Next iCol
    Next iFila
End Sub
Function PosicionColX(iCol As Integer, aDefCelda() As TCelda) As Integer
Dim i As Integer
    PosicionColX = 0
    For i = 0 To iCol - 1
        PosicionColX = PosicionColX + aDefCelda(i).Ancho
    Next
End Function

Public Sub DibujarTabla(f As Printer, iPosX As Integer, iPosY As Integer, iFilas As Integer, iCols As Integer, iAnchoCelda As Integer, iAltoCelda As Integer, Optional bColor As Boolean = True, Optional iFilaIni As Integer = 0)
Dim iFila As Integer
Dim iCol As Integer
Dim sValor As String
Dim iSuper As Integer
Dim iSuper2 As Integer
Dim iTamFuente As Integer
Dim iArroba As Integer
Dim iFilaTmp As Integer
    
    If bColor Then
        f.ForeColor = Val(VarCfg("color_columnas_tabla"))
        f.Line (iPosX, iPosY)-(iPosX + iAnchoCelda, iPosY + iAltoCelda * iFilas), , BF
        f.Line (iPosX, iPosY)-(iPosX + iAnchoCelda * iCols, iPosY + iAltoCelda), , BF
    End If
    f.ForeColor = 0
    For iFila = 0 To iFilas
        f.CurrentY = iPosY + (iFila * iAltoCelda)
        f.CurrentX = iPosX
        
        f.Line -Step(iCols * iAnchoCelda, 0)
    Next iFila
    For iCol = 0 To iCols
        f.CurrentX = iPosX + (iCol * iAnchoCelda)
        f.CurrentY = iPosY
        f.Line -Step(0, iFilas * iAltoCelda)
    Next iCol
    For iFilaTmp = iFilaIni To iFilaIni + iFilas - 1
        iFila = iFilaTmp - iFilaIni
        For iCol = 0 To iCols - 1
            sValor = aTabla(iFilaTmp, iCol)
            iSuper = InStr(sValor, "(")
            iArroba = InStr(sValor, "@")
            If iSuper > 0 Then
                sValor = Mid$(sValor, 1, Len(sValor) - 1)
            End If
            
            iTamFuente = f.FontSize
            If iArroba > 0 Then
                f.FontSize = 7
            End If
            f.CurrentX = iPosX + (iCol * iAnchoCelda) + (iAnchoCelda - f.TextWidth(sValor)) / 2
            f.CurrentY = iPosY + (iFila * iAltoCelda) + (iAltoCelda - f.TextHeight(sValor)) / 2
            
            If iSuper > 0 Then
                f.CurrentX = f.CurrentX + G_MARGEN_SUPER
                iSuper2 = InStr(Mid$(sValor, iSuper + 1), ")")
                f.Print Mid$(sValor, 1, iSuper - 1);
                f.FontSize = 5
                f.Print Mid$(sValor, iSuper + 1)
                f.FontSize = iTamFuente
            ElseIf Val(sValor) - Int(Val(sValor)) = 0.5 Then
                f.Print Trim$(Str$(Int(Val(sValor))));
                f.Print "½"
            Else
                f.Print sValor
            End If
        Next iCol
    Next iFilaTmp
End Sub
Public Sub DibujarTablaExt(f As Printer, iPosX As Integer, iPosY As Integer, iFilas As Integer, iCols As Integer, iAnchoCelda As Integer, iAltoCelda As Integer, iAnchoCelda2 As Integer, iColsAnchoCelda2 As Integer, iColNegrita1 As Integer, iColNegrita2 As Integer, Optional bResumen As Boolean = False)
Dim iFila As Integer
Dim iCol As Integer
Dim sValor As String
Dim iSuper As Integer
Dim iSuper2 As Integer
Dim iTamFuente As Integer
Dim iArroba As Integer
Dim iColsPuestos As Integer
Dim iColActual As Integer
    
    iColsPuestos = 0
    If bResumen Then
        'Imprimimos tabla de finales resumida
        'Contamos las columnas de los puestos
        For iCol = 0 To iCols
            If InStr(aTabla(0, iCol), "º") > 0 Then
                Inc iColsPuestos
            End If
        Next
    End If
    
    f.ForeColor = Val(VarCfg("color_columnas_tabla"))
    f.Line (iPosX + iAnchoCelda2 * (iColNegrita1), iPosY)-Step(iAnchoCelda2, iAltoCelda * iFilas), , BF
    If Not bResumen And iAnchoCelda = iAnchoCelda2 Then
        f.Line (iPosX + iAnchoCelda * (iColNegrita2), iPosY)-Step(iAnchoCelda, iAltoCelda * iFilas), , BF
        f.Line (iPosX, iPosY)-Step(iAnchoCelda * iCols, iAltoCelda), , BF
    End If
    f.ForeColor = 0
    For iFila = 0 To iFilas
        f.CurrentY = iPosY + (iFila * iAltoCelda)
        f.CurrentX = iPosX
        f.Line -Step((iColsAnchoCelda2 * iAnchoCelda2) + (iCols - iColsPuestos - iColsAnchoCelda2) * iAnchoCelda, 0)
    Next iFila
    For iCol = 0 To iCols - iColsPuestos
        If iCol <= iColsAnchoCelda2 Then
            f.CurrentX = iPosX + (iCol * iAnchoCelda2)
        Else
            f.CurrentX = iPosX + (iColsAnchoCelda2 * iAnchoCelda2) + ((iCol - iColsAnchoCelda2) * iAnchoCelda)
        End If
        f.CurrentY = iPosY
        f.Line -Step(0, iFilas * iAltoCelda)
    Next iCol
    For iFila = 0 To iFilas - 1
        iCol = 0
        For iColActual = 0 To iCols - 1
            If Not bResumen Or InStr(aTabla(0, iColActual), "º") = 0 Then
                If iColNegrita1 = iCol Or iColNegrita2 = iCol Then
                    f.FontBold = True
                Else
                    f.FontBold = False
                End If
                sValor = aTabla(iFila, iColActual)
                iSuper = InStr(sValor, "(")
                iArroba = InStr(sValor, "@")
                If iSuper > 0 Then
                    sValor = Mid$(sValor, 1, Len(sValor) - 1)
                End If
                
                iTamFuente = f.FontSize
                If iArroba > 0 Then
                    f.FontSize = 7
                End If
                If iCol < iColsAnchoCelda2 Then
                    f.CurrentX = iPosX + (iCol * iAnchoCelda2) + (iAnchoCelda2 - f.TextWidth(sValor)) / 2
                Else
                    f.CurrentX = iPosX + ((iColsAnchoCelda2 * iAnchoCelda2) + (iCol - iColsAnchoCelda2) * iAnchoCelda) + (iAnchoCelda - f.TextWidth(sValor)) / 2
                End If
                f.CurrentY = iPosY + (iFila * iAltoCelda) + (iAltoCelda - f.TextHeight(sValor)) / 2
                
                If iSuper > 0 Then
                    f.CurrentX = f.CurrentX + G_MARGEN_SUPER
                    iSuper2 = InStr(Mid$(sValor, iSuper + 1), ")")
                    f.Print Mid$(sValor, 1, iSuper - 1);
                    f.FontSize = 5
                    f.Print Mid$(sValor, iSuper + 1)
                    f.FontSize = iTamFuente
                ElseIf Val(sValor) - Int(Val(sValor)) = 0.5 Then
                    f.Print Trim$(Str$(Int(Val(sValor))));
                    f.Print "½";
                    If InStr(sValor, "d") > 0 Then
                        f.Print "d"
                    Else
                        f.Print
                    End If
                Else
                    f.Print sValor
                End If
                
                Inc iCol
            End If
        Next iColActual
    Next iFila
    f.FontBold = False
End Sub
Public Sub RetrasaSeg(iSeg As Integer)
Dim i As Integer
Dim j As Integer
    i = 0
    Do While i < iSeg
        For j = 1 To 9
            Sleep 100
            DoEvents
            Inc i
        Next
    Loop

End Sub
Public Sub PararS(iSeg As Integer)
Dim i As Integer
Dim j As Integer
    i = 0
    Do While i < iSeg
        For j = 1 To 9
            Sleep 100
            DoEvents
            Inc i
        Next
    Loop

End Sub
Public Sub MensajeSeg(iSeg As Integer)
Dim i As Integer
Dim j As Integer
    i = 0
    Do While i < iSeg
        For j = 1 To 9
            Sleep 100
            DoEvents
            Inc i
        Next
    Loop

End Sub

Public Sub Espera(iMs As Integer)
Dim i As Integer, iFactorDivisor As Integer
    iFactorDivisor = Val(VarCfg("factor_divisor_tiempo"))
    If iFactorDivisor = 0 Then iFactorDivisor = 1
    
    Do While i < iMs / iFactorDivisor
        Sleep 1
        DoEvents
        Inc i
    Loop
End Sub

Public Function ValorColor(lColor As Long) As Long
Dim iRojo As Integer, iVerde As Integer, iAzul As Integer, iTono As Integer

        
    iTono = Int(lColor / 16777216)
    lColor = lColor Mod 16777216
    iRojo = Int(lColor / 65536)
    lColor = lColor Mod 65536
    iVerde = Int(lColor / 256)
    iAzul = lColor Mod 256
    
    ValorColor = iRojo + iVerde + iAzul + iTono
End Function

Public Function Redondea(dNum As Double, iDec As Integer) As Double
    Redondea = Round(dNum + IIf(Mid$(Trim$(Str$(dNum - Int(dNum))), iDec + 2, 1) = "5", (1 / 10 ^ (iDec + 1)), 0), iDec)
End Function

Public Function Buscar(sTabla As String, sCampo As String, sCodigo As String) As String
Dim rs As Recordset

    Set rs = db.OpenRecordset("SELECT " & sCampo & " FROM " & sTabla & " WHERE codigo = " & sCodigo, dbOpenSnapshot)
        If Not rs.EOF Then
            Buscar = rs.Fields(0)
        Else
            Buscar = ""
        End If
    rs.Close
End Function

Public Function LiteralFase(sCodFase As String) As String
    Select Case Val(sCodFase)
        Case 1
            LiteralFase = mml_FRASE0329
        Case 2
            LiteralFase = "Semi-Final"
        Case 4
            LiteralFase = "Cuartos de Final"
        Case 8
            LiteralFase = "Octavos de Final"
        Case 16
            LiteralFase = mml_FRASE0933
        Case 32
            LiteralFase = mml_FRASE0934
        Case 64
            LiteralFase = mml_FRASE0935
        Case Else
            LiteralFase = "1/" & sCodFase & mml_FRASE0708
    End Select
End Function

Public Function SinNulos(sCad As Variant) As String
    If IsNull(sCad) Then
        SinNulos = ""
    Else
        SinNulos = sCad
    End If
End Function

Public Function IdentificarGrupoEdad(sCat As String) As Integer
Dim sPrimerCateg As String
Dim iCateg As Integer
    sPrimerCateg = Mid$(sCat, 7, 5)
    sPrimerCateg = UCase(sPrimerCateg)
    If InStr(sPrimerCateg, mml_FRASE0936) > 0 Then
        IdentificarGrupoEdad = G_INFANTIL
    ElseIf InStr(sPrimerCateg, mml_FRASE0937) > 0 Then
        IdentificarGrupoEdad = G_JUVENIL
    ElseIf InStr(sPrimerCateg, mml_FRASE0938) > 0 Then
        IdentificarGrupoEdad = G_JUNIOR1
    ElseIf InStr(sPrimerCateg, mml_FRASE0939) > 0 Then
        IdentificarGrupoEdad = G_JUNIOR1
    ElseIf InStr(sPrimerCateg, mml_FRASE0940) Then
        IdentificarGrupoEdad = G_JUNIOR2
    ElseIf InStr(sPrimerCateg, mml_FRASE0941) Then
        IdentificarGrupoEdad = G_YOUTH
    ElseIf InStr(sPrimerCateg, mml_FRASE0942) Then
        IdentificarGrupoEdad = G_ADULTO1
    ElseIf InStr(sPrimerCateg, mml_FRASE0943) Then
        IdentificarGrupoEdad = G_ADULTO1
    ElseIf InStr(sPrimerCateg, mml_FRASE0944) Then
        IdentificarGrupoEdad = G_ADULTO2
    ElseIf InStr(sPrimerCateg, mml_FRASE0945) Then
        IdentificarGrupoEdad = G_SENIOR1
    ElseIf InStr(sPrimerCateg, mml_FRASE0946) Then
        IdentificarGrupoEdad = G_SENIOR1
    ElseIf InStr(sPrimerCateg, mml_FRASE0947) Then
        IdentificarGrupoEdad = G_SENIOR2
    ElseIf InStr(sPrimerCateg, mml_FRASE0948) Then
        IdentificarGrupoEdad = G_SENIOR3
    Else
        IdentificarGrupoEdad = 0
    End If

End Function

Public Sub GrabarPosicion(frm As Form)
    iFormularioTop = frm.Top
    iFormularioLeft = frm.Left
    iFormularioHeight = frm.Height
    iFormularioWidth = frm.Width
End Sub

Public Sub RestaurarPosicion(frm As Form)
    frm.Top = iFormularioTop
    frm.Left = iFormularioLeft
    frm.Height = iFormularioHeight
    frm.Width = iFormularioWidth
End Sub

Public Sub SaltoLinea(Disp As Printer, iTamFuente As Integer)
Dim iTamFuenteAnt As Integer
    iTamFuenteAnt = Disp.FontSize
    Disp.FontSize = iTamFuente
    Disp.Print
    Disp.FontSize = iTamFuenteAnt
End Sub

Sub Centrado(Disp As Printer, sCad As String, iAncho As Integer)
    On Local Error Resume Next
    Disp.CurrentX = iAncho / 2 - Disp.TextWidth(sCad) / 2
    Disp.Print sCad
End Sub
Public Function sEscuela(iCodcomp As Integer)
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT escuela FROM competiciones WHERE codigo = " & iCodcomp, dbOpenSnapshot)
    If Not rs.EOF Then
        sEscuela = rs!escuela
    Else
        sEscuela = ""
    End If
End Function

Public Function CalcularCodGrupoEdad(sGrupoEdad As Variant) As Integer
    If IsNull(sGrupoEdad) Then
        CalcularCodGrupoEdad = 0
        Exit Function
    End If
    Select Case UCase(sGrupoEdad)
        Case UCase(mml_FRASE0949)
            CalcularCodGrupoEdad = G_INFANTIL
        Case UCase(mml_FRASE0135)
            CalcularCodGrupoEdad = G_JUVENIL
        Case UCase(mml_FRASE0136)
            CalcularCodGrupoEdad = G_JUNIOR1
        Case UCase(mml_FRASE0137)
            CalcularCodGrupoEdad = G_JUNIOR2
        Case UCase(mml_FRASE0138)
            CalcularCodGrupoEdad = G_YOUTH
        Case UCase(mml_FRASE0125)
            CalcularCodGrupoEdad = G_ADULTO1
        Case UCase(mml_FRASE0126)
            CalcularCodGrupoEdad = G_ADULTO2
        Case UCase(mml_FRASE0141)
            CalcularCodGrupoEdad = G_SENIOR1
        Case UCase(mml_FRASE0140)
            CalcularCodGrupoEdad = G_SENIOR2
        Case UCase(mml_FRASE0139)
            CalcularCodGrupoEdad = G_SENIOR3
        Case Else
            CalcularCodGrupoEdad = 0
    End Select
End Function


Public Function iMinDorsalOficial(sCodComp As String) As Integer
Dim rs As Recordset
    
    sCodComp = Val(sCodComp)
    
    Set rs = db.OpenRecordset("SELECT min_dorsal_oficial FROM competiciones WHERE codigo = " & sCodComp, dbOpenSnapshot)
    If rs.EOF Then
        iMinDorsalOficial = C_MIN_DORSAL_OFICIAL
    Else
        iMinDorsalOficial = rs!min_dorsal_oficial
    End If
    rs.Close
End Function

Public Function iDorsalesPorTanda(sCodComp As String) As Integer
Dim rs As Recordset
    
    sCodComp = Val(sCodComp)
    
    Set rs = db.OpenRecordset("SELECT dorsales_tanda FROM competiciones WHERE codigo = " & sCodComp, dbOpenSnapshot)
    If rs.EOF Then
        iDorsalesPorTanda = C_DORSALES_TANDA_DEFECTO
    Else
        iDorsalesPorTanda = rs!dorsales_tanda
    End If
    rs.Close
End Function
Public Function iDorsalesPorTandaCateg(iCateg As Integer) As Integer
Dim rs As Recordset
    
    Set rs = db.OpenRecordset("SELECT dorsales_tanda FROM categorias WHERE codigo = " & iCateg, dbOpenSnapshot)
    If rs.EOF Then
        iDorsalesPorTandaCateg = C_DORSALES_TANDA_DEFECTO
    Else
        iDorsalesPorTandaCateg = rs!dorsales_tanda
    End If
    rs.Close
End Function


Public Function NoPresente(iDorsal As Integer, iCodCateg As Integer, iFase As Integer, iRep As Integer) As Boolean
Dim rs As Recordset
    
    Set rs = db.OpenRecordset("SELECT no_presente FROM dorsales WHERE num_dorsal = " & iDorsal & " AND cod_categoria = " & iCodCateg & " AND fase = " & iFase & " AND repesca = " & iRep, dbOpenSnapshot)
    If Not rs.EOF Then
        NoPresente = IIf(rs!no_presente > 0, True, False)
    Else
        NoPresente = True
    End If
    rs.Close
End Function


Public Function ProcesarError(Optional sComentario As String = "", Optional bMensaje As Boolean = True) As String
Dim Msj As String
Dim sSQL As String
   ProcesarError = ""
   If Err.Number <> 0 Then
        Msj = "Error # " & Str(Err.Number) & "-" & sComentario & mml_FRASE0208 _
              & Err.Source & Chr(13) & Err.Description
        ProcesarError = Msj
        If Not C_DEBUG Then On Local Error Resume Next
        sSQL = "INSERT INTO errores VALUES ('" & App.Major & "." & App.Minor & "." & App.Revision & "','" & sQuitarCarProhibidosSQL(sComentario) & "','" & sQuitarCarProhibidosSQL(Msj) & "','" & Format$(Now, "dd/mm/yyyy") & " " & Format$(Time, "hh:mm:ss") & "')"
        Debug.Print sSQL
        db.Execute sSQL
        If bMensaje Then
            MsgBox Msj, , mml_FRASE0096, Err.HelpFile, Err.HelpContext
        End If
   End If
End Function

Public Function MinPorBaile(iFase As Integer, sIdCat As String, iCodcomp As Integer, iRepesca As Integer, iCodCateg As Integer) As Integer
Dim rs As Recordset, iMaxTandas As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    If iFase = C_FASE_GENERAL_LOOK Then
        MinPorBaile = ConsultaMinPorBaile(mml_FRASE0645 & sEquivCateg(UCase(sIdCat)), iCodcomp)
    ElseIf iFase = 1 Then
        MinPorBaile = ConsultaMinPorBaile(mml_FRASE0329 & sEquivCateg(UCase(sIdCat)), iCodcomp)
    Else
        MinPorBaile = ConsultaMinPorBaile(mml_FRASE0950 & sEquivCateg(UCase(sIdCat)), iCodcomp)
        If iFase > 1 Then
            'Multiplicamos la duración por el número de tandas
            Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCateg & " AND fase = " & iFase & " AND repesca = " & iRepesca, dbOpenSnapshot)
            'comprobar si es fase inicial y tiene dorssales
            If rs.Fields(0) > 0 Then
                iMaxTandas = 0
                CalcularDorsalesPorTandaCat iCodCateg, iFase, iRepesca, 1, iMaxTandas
                MinPorBaile = (iMaxTandas) * MinPorBaile
            ElseIf iFase > 2 Then
                MinPorBaile = Val(VarCfg("factor_mult_tiempo_fase" & Trim$(Str$(iFase)))) * MinPorBaile
            End If
            rs.Close
        End If
    End If
    Exit Function
error:
    MinPorBaile = C_MIN_POR_BAILE_POR_DEFECTO
End Function

Public Function ConsultaMinPorBaile(sCatFase As String, iCodcomp As Integer) As Integer
Dim rs As Recordset
    ConsultaMinPorBaile = 0
    Set rs = db.OpenRecordset("SELECT " & sCatFase & " FROM competiciones WHERE codigo = " & iCodcomp, dbOpenSnapshot)
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then
            ConsultaMinPorBaile = rs.Fields(0)
        End If
    End If
    rs.Close
End Function

Public Function sEstimacion(sHora As String) As String
    If IsDate(sHora) Then
        sEstimacion = Format$(DateAdd("n", Val(VarCfg("dif_hora")), CDate(sHora)), "hh:mm")
    Else
        sEstimacion = sHora
    End If
End Function

Public Function sEstimacionInversa(sHora As String) As String
    If IsDate(sHora) Then
        sEstimacionInversa = Format$(DateAdd("n", -Val(VarCfg("dif_hora")), CDate(sHora)), "hh:mm")
    Else
        sEstimacionInversa = sHora
    End If
End Function

Public Function sDescFase(iFase As Integer) As String
    Select Case iFase
        Case 1:
            sDescFase = mml_FRASE0329
        Case 2:
            sDescFase = mml_FRASE0872
        Case 4:
            sDescFase = "1/" & Trim$(Str$(iFase))
        Case 8:
            sDescFase = "1/" & Trim$(Str$(iFase))
        Case 16:
            sDescFase = "1/" & Trim$(Str$(iFase))
        Case 32:
            sDescFase = "1/" & Trim$(Str$(iFase))
        Case 64:
            sDescFase = "1/" & Trim$(Str$(iFase))
        Case 128:
            sDescFase = "1/" & Trim$(Str$(iFase))
        Case 256:
            sDescFase = "1/" & Trim$(Str$(iFase))
        Case 512:
            sDescFase = "1/" & Trim$(Str$(iFase))
        Case C_FASE_GENERAL_LOOK:
            sDescFase = mml_FRASE0645
        Case Else
            sDescFase = ""
    End Select
End Function

Public Function sDescModalidad(iMod As Integer)
Dim rs As Recordset

    Set rs = db.OpenRecordset("SELECT nombre FROM modalidad WHERE codigo = " & iMod, dbOpenSnapshot)
    If Not rs.EOF Then
        sDescModalidad = rs!Nombre
    Else
        sDescModalidad = ""
    End If
    rs.Close
End Function

Public Sub ProcesarEventos()
    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
End Sub
Public Function GenerarControl(sCodComp As Integer) As Boolean
Dim rs As Recordset
    GenerarControl = False
    Set rs = db.OpenRecordset("SELECT control FROM competiciones WHERE codigo = " & sCodComp, dbOpenSnapshot)
    If Not rs.EOF Then
        GenerarControl = IIf(rs.Fields(0) = 1, True, False)
    End If
    rs.Close
End Function

Public Function BailesParciales(sCodCat As Integer) As Boolean
Dim rs As Recordset
    BailesParciales = False
    Set rs = db.OpenRecordset("SELECT rec_parcial_bailes FROM categorias WHERE codigo = " & sCodCat, dbOpenSnapshot)
    If Not rs.EOF Then
        BailesParciales = IIf(rs.Fields(0) = 1, True, False)
    End If
    rs.Close
End Function
Public Function MostrarPosicion(sCodCat As Integer) As Boolean
Dim rs As Recordset
    MostrarPosicion = False
    Set rs = db.OpenRecordset("SELECT mostrar_posicion FROM categorias WHERE codigo = " & sCodCat, dbOpenSnapshot)
    If Not rs.EOF Then
        MostrarPosicion = IIf(rs.Fields(0) = 1, True, False)
    End If
    rs.Close
End Function

Public Sub AbrirBaseDeDatos(Optional Exclusivo As Boolean = False)
Dim sBD As String * 255, iBDLon As Integer
    If Not db Is Nothing Then Exit Sub
    iBDLon = GetPrivateProfileString(mml_FRASE0033, "DataBase", "", sBD, 255, "Escrutinio.ini")
    If iBDLon = 0 Then
        MsgBox mml_FRASE0951, vbOKOnly Or vbCritical, mml_FRASE0096
        End
    End If
    Set db = OpenDatabase(sBD, Exclusivo)
    'Acceso a la base de datos mediante origen de datos ODBC
    'Set ws = CreateWorkspace("IRIS", "admin", "", dbUseODBC)
    'Set db = ws.OpenDatabase("Escrutinio", dbDriverNoPrompt, , "ODBC;DSN=Escrutinio")
End Sub

Public Function CalcularRetrasoAlCalcular(iCateg As Integer, iFase As Integer, iRepesca As Integer) As Integer
Dim rs As Recordset, dHora As Date
    ' La hora actual es la hora del cálculo, y sabemos la hora aproxima a la que han
    ' Salido a bailar
    'Localizamos la hora a la que debian salir a bailar
    Set rs = db.OpenRecordset("SELECT hora FROM horario WHERE cod_categoria = " & iCateg & " AND numfase = " & iFase & " AND repesca = " & iRepesca, dbOpenSnapshot)
    If rs.EOF Then
        dHora = Now
    Else
        dHora = rs!hora
    End If
    rs.Close
    Set rs = db.OpenRecordset("SELECT id_categoria FROM categorias WHERE codigo = " & iCateg, dbOpenSnapshot)
    If Not rs.EOF Then
        'Direferencia entre la hora de entrada oficial y la hora calculada de entrada= hora actual-tiempo de presencia en pista
        CalcularRetrasoAlCalcular = DateDiff("n", dHora, DateAdd("n", -(MinPorBaile(iFase, rs!id_categoria, VarCfg("horario_codcompeticion"), iRepesca, iCateg) + C_MINUTOS_MARGEN_CALCULAR), CDate(Format$(Now, "hh:nn"))))
    Else
        CalcularRetrasoAlCalcular = 0
    End If
    rs.Close
    db.Execute ("UPDATE cfg SET valor = " & CalcularRetrasoAlCalcular & " WHERE variable = 'dif_hora'")
End Function


Public Function CalcularResultado(sCadCod As String)
Dim sCad As String
Dim i As Integer
    If IsDate(sCadCod) Then
        sCad = Format$(DateAdd("d", -1060, CDate(sCadCod)), "dd/mm/yy")
    Else
        For i = 1 To Len(sCadCod)
            sCad = sCad & Chr$(Asc(Mid$(sCadCod, i, 1)) - (i Mod 7) * IIf(i Mod 3 = 0, -1, 1))
        Next
    End If
    CalcularResultado = sCad

End Function

Public Function sDescCortaFase(iFase As Integer) As String
    If Not C_DEBUG Then On Local Error GoTo error
    Select Case iFase
        Case 1:
            sDescCortaFase = "F"
        Case 2:
            sDescCortaFase = "SF"
        Case C_FASE_GENERAL_LOOK:
            sDescCortaFase = "GL"
        Case Else
            sDescCortaFase = Trim$(Str$(iFase)) & "F"
    End Select
    Exit Function
error:
    ProcesarError
End Function

Public Function bAddLog(sLog As String, sFichero As String) As Boolean
Dim iFichero As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    bAddLog = True
    iFichero = FreeFile
    Open sFichero For Append As #iFichero
    Print #iFichero, sLog & Chr$(13) & Chr$(10)
    Close #iFichero
    Exit Function
error:
    bAddLog = False
    Err.Clear
End Function

Public Function DividirCampo(sCampo As String, aeMail() As String, sSeparador As String) As Integer
Dim j As Integer

    DividirCampo = 1
    j = InStr(sCampo, sSeparador)
    While j > 0
        aeMail(DividirCampo - 1) = LTrim$(RTrim$(Mid$(sCampo, 1, j - 1)))
        sCampo = Mid$(sCampo, j + Len(sSeparador))
        j = InStr(sCampo, sSeparador)
        Inc DividirCampo
    Wend
    aeMail(DividirCampo - 1) = sCampo

End Function

Public Function CalcularFase(iNumDorsales) As Integer
    If iNumDorsales <= 7 Then
        CalcularFase = 1
    ElseIf iNumDorsales <= 13 Then
        CalcularFase = 2
    Else
        CalcularFase = 2 ^ (Int(Log((iNumDorsales - 1) / 6) / Log(2)) + 1)
    End If
End Function

Public Function QuitarCadena(sQuitarCad As String, sCad As String) As String
Dim i As Integer
    i = InStr(sCad, sQuitarCad)
    While i > 0
        sCad = Mid$(sCad, 1, i - 1) & Mid$(sCad, i + 1)
        i = InStr(sCad, sQuitarCad)
    Wend
    QuitarCadena = sCad
End Function
Public Function CambiarCadena(sCad1 As String, sCad2 As String, sCad As String) As String
Dim i As Integer
    i = InStr(sCad, sCad1)
    While i > 0
        sCad = Mid$(sCad, 1, i - 1) & sCad2 & Mid$(sCad, i + 1)
        i = InStr(sCad, sCad1)
    Wend
    CambiarCadena = sCad
End Function
Public Function sEquivCateg(sCat As String) As String
    Select Case UCase(sCat)
        Case "A"
            sEquivCateg = "A"
        Case "B"
            sEquivCateg = "B"
        Case "C"
            sEquivCateg = "C"
        Case "D"
            sEquivCateg = "D"
        Case "E"
            sEquivCateg = "E"
        Case "F"
            sEquivCateg = "F"
        Case "G"
            sEquivCateg = "G"
        Case "H"
            sEquivCateg = "H"
        Case "I"
            sEquivCateg = "I"
        Case "J"
            sEquivCateg = "I"
        Case "K"
            sEquivCateg = "I"
        Case Else
            sEquivCateg = "A"
    End Select
End Function

Public Function sDescCortaGrupoEdad(iCodGrupo As Integer)
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT abreviatura FROM gruposedad WHERE codigo = " & iCodGrupo, dbOpenSnapshot)
    If Not rs.EOF Then
        sDescCortaGrupoEdad = rs!abreviatura
    Else
        sDescCortaGrupoEdad = ""
    End If
    rs.Close
End Function


Public Sub ImportarDatosConNuevoCodigo(sTabla As String, sSQL As String, aValores() As TValores)
Dim rs As Recordset, rs1 As Recordset
Dim i As Integer, j As Integer
    Set rs1 = db.OpenRecordset(sSQL, dbOpenSnapshot)
    Set rs = db.OpenRecordset(sTabla, dbOpenTable)
        While Not rs1.EOF
            rs.AddNew
                For i = 0 To rs.Fields.Count - 1
                    rs.Fields(i) = rs1.Fields(i)
                    For j = 0 To UBound(aValores) - 1
                        If rs.Fields(i).Name = aValores(j).Nombre Then
                            If aValores(j).valor = "MaxCod" Then
                                rs.Fields(i) = MaxCod(sTabla)
                            Else
                                rs.Fields(i) = aValores(j).valor
                            End If
                        End If
                    Next
                Next
            rs.Update
            rs1.MoveNext
        Wend
    rs.Close
    rs1.Close
End Sub

Public Sub ImportarDatosConControl(db1 As Database, sTabla As String, sSQL As String, aValores() As TValores)
Dim rs As Recordset, rs1 As Recordset
Dim i As Integer, j As Integer

    If Not C_DEBUG Then On Error GoTo error

    Set rs1 = db1.OpenRecordset(sSQL, dbOpenSnapshot)
    Set rs = db.OpenRecordset(sTabla, dbOpenTable)
        While Not rs1.EOF
            rs.AddNew
                For i = 0 To rs.Fields.Count - 1
                    rs.Fields(i) = rs1.Fields(i)
                    For j = 0 To UBound(aValores) - 1
                        If UCase(rs.Fields(i).Name) = UCase(aValores(j).Nombre) Then
                            Select Case aValores(j).operacion
                                Case "MaxCod"
                                    rs.Fields(i) = MaxCod(sTabla)
                                Case mml_FRASE0666
                                    rs.Fields(i) = rs1.Fields(i) + aValores(j).valor
                                Case Else
                                    rs.Fields(i) = aValores(j).valor
                            End Select
                        End If
                    Next
                Next
            rs.Update
            rs1.MoveNext
        Wend
    rs.Close
    rs1.Close
error:
    ProcesarError
End Sub

Public Sub AsignarParametro(sVar As String, sValor As String)
    db.Execute "UPDATE cfg SET valor ='" & sValor & "' WHERE variable = '" & sVar & "'"
End Sub


Public Function sDescCategoria(ByVal lCodCat As Long, Optional ByVal lCodComp As Long = 0)
Dim rs As Recordset
    If lCodComp = 0 Then
        Set rs = db.OpenRecordset("SELECT descripcion FROM categorias WHERE codigo = " & lCodCat, dbOpenSnapshot)
    Else
        Set rs = db.OpenRecordset("SELECT descripcion FROM categorias WHERE cod_competicion = " & lCodComp & " AND codigo = " & lCodCat, dbOpenSnapshot)
    End If
    If rs.EOF Then
        sDescCategoria = ""
    Else
        sDescCategoria = rs.Fields(0)
    End If
    rs.Close
End Function

Public Function sDescCompeticion(iCodcomp As Long)
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT descripcion FROM competiciones WHERE codigo = " & iCodcomp, dbOpenSnapshot)
    If rs.EOF Then
        sDescCompeticion = ""
    Else
        sDescCompeticion = rs.Fields(0)
    End If
    rs.Close
End Function

Function ComprobarSiEstanTodasPuntuaciones(iCodCat As Long, iHojaRepesca As Integer, iFase As Integer) As Boolean
Dim iPuestos As Integer
Dim iJueces As Integer
Dim iBailes As Integer
Dim iDorsales As Integer
Dim rs As Recordset
        
    ' Comprobamos si hemos recopilado la información de todos los dorsales
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase = " & iFase, dbOpenSnapshot)
    iPuestos = rs.Fields(0)
    rs.Close
    ' Comprobamos el número de jueces
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE pasos = 0 AND cod_categoria = " & iCodCat, dbOpenSnapshot)
    iJueces = rs.Fields(0)
    rs.Close
    ' Comprobamos el número de bailes
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & iCodCat & " AND fase = " & IIf(iFase > 1, 2, 1), dbOpenSnapshot)
    iBailes = rs.Fields(0)
    rs.Close
    ' Comprobamos el número de dorsales
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND repesca=" & iHojaRepesca & " AND fase = " & iFase, dbOpenSnapshot)
    iDorsales = rs.Fields(0)
    rs.Close

    If iPuestos = iDorsales * iJueces * iBailes And iPuestos > 0 Then
        ComprobarSiEstanTodasPuntuaciones = True
    Else
        ComprobarSiEstanTodasPuntuaciones = False
    End If
    
End Function


Public Function MostrarCompeticionActiva() As String
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT * FROM competiciones WHERE codigo = " & Val(VarCfg("horario_codcompeticion")), dbOpenSnapshot)
    If rs.EOF Then
        MostrarCompeticionActiva = ""
    Else
        MostrarCompeticionActiva = rs!DESCRIPCION
    End If
    rs.Close
End Function


Function sExtraerPath(sFichero As String) As String
Dim i As Integer, sCad As String
    sExtraerPath = sCad
    i = InStr(sFichero, "\")
    While i > 0
        sExtraerPath = Mid$(sFichero, 1, i - 1)
        i = InStr(i + 1, sFichero, "\")
    Wend
End Function
Function sExtraerFichero(sFichero As String) As String
Dim i As Integer, sCad As String
    sExtraerFichero = sCad
    i = InStr(sFichero, "\")
    While i > 0
        sExtraerFichero = Mid$(sFichero, i + 1)
        i = InStr(i + 1, sFichero, "\")
    Wend
End Function

Sub CargarPistas(cbPista As ComboBox)
    cbPista.Clear
    cbPista.AddItem ""
    cbPista.AddItem "(P1)"
    cbPista.AddItem "(P2)"
    cbPista.AddItem "(P3)"
    cbPista.AddItem "(P4)"
    cbPista.AddItem "(P5)"
    cbPista.AddItem "(P6)"
    cbPista.AddItem "(P7)"
    cbPista.AddItem "(P8)"
    cbPista.AddItem "(P9)"
End Sub

Private Sub CalcularDorsalesPorTanda(iTotalDorsales As Integer, iMaxDorsalesTandaPedidos As Integer, iMinDorsalesTandaCalculados As Integer, iMaxTandas As Integer, iTandaSolicitada As Integer, iDorsalesTandaSolicitada As Integer, iTandasConMasDorsales As Integer)
    If Not C_DEBUG Then On Error GoTo error
    If iMaxTandas = 0 Then
        iMaxTandas = iTotalDorsales \ iMaxDorsalesTandaPedidos
        If iTotalDorsales Mod iMaxDorsalesTandaPedidos > 0 Then
            iMaxTandas = iMaxTandas + 1
        End If
    End If
    If iMaxTandas = 0 Then iMaxTandas = 1
    iMinDorsalesTandaCalculados = iTotalDorsales \ iMaxTandas
    iTandasConMasDorsales = iTotalDorsales Mod iMaxTandas
    If iTandaSolicitada <= iTandasConMasDorsales Then
        iDorsalesTandaSolicitada = iMinDorsalesTandaCalculados + 1
    Else
        iDorsalesTandaSolicitada = iMinDorsalesTandaCalculados
    End If
    Exit Sub
error:
    ProcesarError
End Sub

'El dosal primero es el 1
Private Function CalcularDorsalInicialTanda(iTotalDorsales As Integer, iMaxDorsalesTandaPedidos As Integer, iMaxTandas As Integer, iTandaSolicitada As Integer, iDorsalesTandaSolicitada As Integer, iTandasConMasDorsales As Integer) As Integer
Dim iMinDorsalesTandaCal As Integer
Dim iTandasMasDorsales As Integer
Dim iTandasMenosDorsales As Integer
    
    CalcularDorsalesPorTanda iTotalDorsales, iMaxDorsalesTandaPedidos, iMinDorsalesTandaCal, iMaxTandas, iTandaSolicitada, iDorsalesTandaSolicitada, iTandasConMasDorsales
    
    If iTandaSolicitada - 1 > iTandasConMasDorsales Then
        iTandasMasDorsales = iTandasConMasDorsales
        iTandasMenosDorsales = (iTandaSolicitada - 1) - iTandasConMasDorsales
    Else
        ' Si TandasConMasDorsales = 0 iMinDorsales contiene los dorsales de todas las tandas
        If iTandasConMasDorsales = 0 Then
            iTandasMasDorsales = 0
            iTandasMenosDorsales = iTandaSolicitada - 1
        Else
            iTandasMasDorsales = iTandaSolicitada - 1
            iTandasMenosDorsales = 0
        End If
    End If
    
    CalcularDorsalInicialTanda = iTandasMasDorsales * (iMinDorsalesTandaCal + 1) + iTandasMenosDorsales * iMinDorsalesTandaCal + 1
End Function

Private Function CalcularDorsalInicialTandaCatExt(iCodCat As Integer, iFase As Integer, iRepesca As Integer, iTanda As Integer, iMaxTandas As Integer, iDorsalesTandaSolicitada As Integer, iTandasMasDorsales As Integer, iTotalDorsales As Integer)
Dim rs As Recordset, iMaxDorsalesPorTanda As Integer

    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND fase = " & iFase & " AND repesca = " & iRepesca, dbOpenSnapshot)
    iTotalDorsales = rs.Fields(0)
    rs.Close

    iMaxDorsalesPorTanda = iDorsalesPorTandaCateg(iCodCat)

    'Controla solo en la semifinal si queremos única tanda de un número determinado de dorsales
    'Siempre que MaxTandas no venga prefijado como en RecOptico
    If iMaxTandas = 0 And InStr(VarCfg("max_dorsales_tanda_semi_excepcion"), Format$(iCodCat, "0####")) > 0 And iFase = 2 And iTotalDorsales <= Val(VarCfg("max_dorsales_tanda_semi")) Then
        iTandasMasDorsales = 0
        iDorsalesTandaSolicitada = iTotalDorsales
        CalcularDorsalInicialTandaCatExt = 1
        iMaxTandas = 1
    Else
        If C_SISTEMA_TANDAS_VIEJO Then
            iTandasMasDorsales = 0
            If iTanda = iMaxTandas Then
                iDorsalesTandaSolicitada = iTotalDorsales \ iMaxTandas + iTotalDorsales Mod iMaxTandas
            Else
                iDorsalesTandaSolicitada = iTotalDorsales \ iMaxTandas
            End If
            CalcularDorsalInicialTandaCatExt = (iTanda - 1) * (iTotalDorsales \ iMaxTandas) + 1
        Else
            CalcularDorsalInicialTandaCatExt = CalcularDorsalInicialTanda(iTotalDorsales, iMaxDorsalesPorTanda, iMaxTandas, iTanda, iDorsalesTandaSolicitada, iTandasMasDorsales)
        End If
    End If
End Function
Function CalcularDorsalInicialTandaCat(iCodCat As Integer, iFase As Integer, iRepesca As Integer, iTanda As Integer, iMaxTandas As Integer, iDorsalesTandaSolicitada As Integer) As Integer
Dim iTandasConMasDorsales As Integer, iTotalDorsales As Integer
    CalcularDorsalInicialTandaCat = CalcularDorsalInicialTandaCatExt(iCodCat, iFase, iRepesca, iTanda, iMaxTandas, iDorsalesTandaSolicitada, iTandasConMasDorsales, iTotalDorsales)
End Function

'Si iMaxTandas entra a 0 calcula también el número de tandas
Function CalcularDorsalesPorTandaCat(iCodCat As Integer, iFase As Integer, iRepesca As Integer, iTanda As Integer, iMaxTandas As Integer) As Integer
    CalcularDorsalInicialTandaCat iCodCat, iFase, iRepesca, iTanda, iMaxTandas, CalcularDorsalesPorTandaCat
End Function

Function CalcularDorsalesPorTandaCatExt(iCodCat As Integer, iFase As Integer, iRepesca As Integer, iTanda As Integer, iMaxTandas As Integer, iTandasConMasDorsales As Integer, iTotalDorsales As Integer) As Integer
    CalcularDorsalInicialTandaCatExt iCodCat, iFase, iRepesca, iTanda, iMaxTandas, CalcularDorsalesPorTandaCatExt, iTandasConMasDorsales, iTotalDorsales
End Function


Sub MostrarDatosIntroducidos(iCodCat As Integer, iFase As Integer, iRep As Integer)
Dim rs As Recordset, sMsj As String, sBaile As String, sJueces As String

    If Not C_DEBUG Then On Local Error GoTo error
    If iCodCat > 0 And iFase > 0 Then
        Set rs = db.OpenRecordset("SELECT COUNT(*) as punt, b.nombre,p.cod_juez FROM puntuaciones p, bailes b WHERE p.cod_baile = b.codigo AND cod_categoria = " & iCodCat & " AND fase = " & iFase & " AND repesca = " & iRep & " GROUP BY b.nombre,p.cod_juez ORDER BY 2,3", dbOpenSnapshot)
            While Not rs.EOF
                If sBaile <> rs!Nombre Then
                    If sBaile <> "" Then
                        sMsj = sMsj & mml_FRASE0436 & sBaile & "  " & vbTab & mml_FRASE0440 & sJueces & Chr$(13) & Chr$(10)
                    End If
                    sBaile = rs!Nombre
                    sJueces = ""
                End If
                sJueces = sJueces & rs!cod_juez & " (" & rs.Fields("punt") & ") "
                rs.MoveNext
            Wend
            If sBaile <> "" Then
                sMsj = sMsj & mml_FRASE0436 & sBaile & vbTab & mml_FRASE0440 & sJueces & Chr$(13) & Chr$(10)
            End If
        rs.Close
        If sMsj <> "" Then
            sMsj = mml_FRASE0439 & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & sMsj
            MsgBox sMsj, vbOKOnly Or vbInformation, mml_FRASE0147
        Else
            MsgBox mml_FRASE0441, vbOKOnly Or vbInformation, mml_FRASE0147
        End If
    Else
        CamposSinCubrir
    End If
    Exit Sub
error:
    ProcesarError "MostrarDatosIntroducidos"
End Sub

Function ComprobarSiTeamMatch(iCodCat As Integer) As Boolean
Dim rs As Recordset
    
    'Primero comprobamos si la hoja corresponde a un TeamMatch
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM categorias WHERE codigo = " & iCodCat & " AND id_categoria = '" & C_TEAM_MATCH & "'", dbOpenSnapshot)
    ComprobarSiTeamMatch = IIf(rs.Fields(0) > 0, True, False)
    rs.Close
End Function

Function Ptos2CadTeamMatch(fNum As Double) As String
    Ptos2CadTeamMatch = Str$(Int(fNum))
    If fNum - Int(fNum) = 0.5 Then
        Ptos2CadTeamMatch = Ptos2CadTeamMatch & "½"
    Else
        Ptos2CadTeamMatch = " " & Ptos2CadTeamMatch
    End If
End Function

'Nombre Abreviado
Public Function sNombreBaileAbreviado(sBaile As String) As String
    If InStr(sBaile, " ") Then
        sNombreBaileAbreviado = Left$(sBaile, 1) & Mid$(sBaile, InStr(sBaile, " ") + 1, 1)
    Else
        sNombreBaileAbreviado = Left$(sBaile, 2)
    End If
End Function

Public Sub BorrarCompeticion(iCodcomp As Integer, Optional bBorrarParticipantes As Boolean = True, Optional bBorrarRegComp As Boolean = True, Optional bBorrarAgrupaciones As Boolean = True)
        'Borrar agrupaciones
        If bBorrarAgrupaciones Then
            db.Execute "DELETE FROM agrupaciones WHERE cod_competicion = " & iCodcomp
        End If
        'Borrar los resultados finales
        db.Execute "DELETE FROM resultadosfinales WHERE cod_competicion = " & iCodcomp
        'Borrar todas las puntuaciones, incluída la fase seleccionada
        db.Execute "DELETE FROM puntuaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar todas las descalificaciones, incluída la fase seleccionada
        db.Execute "DELETE FROM descalificaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar todas las hojas reconocidas, incluída la fase seleccionada
        db.Execute "DELETE FROM hojas_reconocidas WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar todas los dorsales de fases anteriores a la seleccionada
        db.Execute "DELETE FROM dorsales dor WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar todas los datos de cal_conjunto (solo FINAL)
        db.Execute "DELETE FROM cal_conjunto WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar todas los datos de cal_baile
        db.Execute "DELETE FROM cal_baile WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar el horario
        db.Execute "DELETE FROM horario WHERE cod_competicion = " & iCodcomp
        'Borrar los jueces y bailes
        db.Execute "DELETE FROM Juez_Categ WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar Jueces de paneles de la competición
        db.Execute "DELETE FROM Juez_panel WHERE cod_competicion = " & iCodcomp
        'Borrar paneles
        db.Execute "DELETE FROM paneles WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar bailes
        db.Execute "DELETE FROM bailes_Categ WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar los dorsales combinados
        db.Execute "DELETE FROM dorsalescombinados WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar PArejas
        If bBorrarParticipantes Then
            db.Execute "DELETE FROM parejas WHERE cod_competicion = " & iCodcomp
            'Borrar enlaceprobaile
            db.Execute "DELETE FROM enlaceprobaile WHERE cod_competicion = " & iCodcomp
        End If
        'Borrar los resumenes finales
        db.Execute "DELETE FROM resumenfinales WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar Categorias
        db.Execute "DELETE FROM categorias WHERE cod_competicion = " & iCodcomp
        'Borrar los resumenes finales
        db.Execute "DELETE FROM publicar WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & iCodcomp & ")"
        'Borrar la competición
        If bBorrarRegComp Then db.Execute ("DELETE FROM competiciones WHERE codigo = " & iCodcomp)
End Sub

Sub BorrarDatosComp(iCodcomp As Integer, bBorrarRegComp As Boolean, Optional bBorrarParejas As Boolean = True)
    BorrarCompeticion iCodcomp, bBorrarParejas, bBorrarRegComp
End Sub

Public Function sNombreBaile(iCodBaile) As String
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT nombre FROM bailes WHERE codigo = " & iCodBaile, dbOpenSnapshot)
    If Not rs.EOF Then
        sNombreBaile = rs!Nombre
    Else
        sNombreBaile = "First"
    End If
    rs.Close
End Function

Public Function iBuscarOrden(ByVal iCodCat As Long, ByVal iCodFase As Integer, iRep As Integer, Optional lOrden As Long = -1, Optional iGrupo As Integer = -1) As Long
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT orden, num_grupo FROM horario WHERE cod_categoria = " & iCodCat & " AND numfase = " & iCodFase & " AND repesca = " & iRep, dbOpenSnapshot)
    If Not rs.EOF Then
        iBuscarOrden = rs!orden
        lOrden = rs!orden
        iGrupo = Val(SinNulos(rs!num_grupo))
    Else
        iBuscarOrden = -1
        lOrden = -1
        iGrupo = -1
    End If
    rs.Close

End Function

Sub LimitarDescCateg(iKey As Integer)
Dim sKey As String
    sKey = LCase(Chr$(iKey))
    If Not ((sKey >= "0" And sKey <= "9") Or (sKey >= "a" And sKey <= "z") Or sKey = "ñ" Or sKey = " " Or _
        sKey = "á" Or sKey = "é" Or sKey = "í" Or sKey = "ó" Or sKey = "ú" Or _
        sKey = "+" Or sKey = "-" Or sKey = "_" Or sKey = "," Or sKey = Chr$(8)) Then
        iKey = 0
    End If
    
End Sub


Public Function TipoClave() As String
Dim lValor As Long, lLong As Long, sValor As String * 100
Dim sNumero As String, sNum1 As String, sNum2 As String
Dim ComputerInfo As cComputerInfo
Dim sCodigo As String
  
  Set ComputerInfo = New cComputerInfo

    TipoClave = ""
    RegOpenKey HKEY_CURRENT_USER, "Applications\GenNumber" & Chr$(0), lValor
    If lValor = 0 Then Exit Function
    lLong = 80
    RegQueryValue lValor, "" & Chr$(0), sValor, lLong
    RegCloseKey lValor
    
    sNumero = Mid$(sValor, 1, InStr(sValor, Chr$(0)) - 1)
    
    sNum1 = Mid$(sNumero, 1, InStr(sNumero, "-") - 1)
    sNum2 = Mid$(sNumero, InStr(sNumero, "-") + 1)
    
    sNum1 = Val("&H" & sNum1) Xor 69
    sNum2 = Val("&H" & sNum2) Xor 69
    
    sCodigo = Trim$(Str$(Val(sNum1) Xor ComputerInfo.ProcessorRevision))

    If (Left$(sCodigo, 2) = "69" Or Left$(sCodigo, 2) = "70" Or Left$(sCodigo, 2) = "71" Or Left$(sCodigo, 2) = "72" Or Left$(sCodigo, 2) = "73") And _
        sCodigo = Trim$(Str$(Val(sNum2) Xor GetHDDSerialNumber("C:\"))) And _
        CDate("01/" & Mid$(sCodigo, 3, 2) & "/" & "20" & Right(sCodigo, 2)) > Now Then
        TipoClave = Left$(sCodigo, 2)
    End If

End Function

Function GetHDDSerialNumber(strDrive As String) As Long
Dim SerialNum As Long
Dim Res As Long
Dim Temp1 As String
Dim Temp2 As String

Temp1 = String$(255, Chr$(0))
Temp2 = String$(255, Chr$(0))
Res = GetVolumeInformation(strDrive, Temp1, _
Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))

GetHDDSerialNumber = SerialNum
End Function

Function CombinarDorsalesCateg(ByVal iCodCat As Long) As Boolean
Dim rs As Recordset

    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM categorias WHERE codigo = " & iCodCat & " AND combinar_dorsales = 1", dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        CombinarDorsalesCateg = True
    Else
        CombinarDorsalesCateg = False
    End If
    rs.Close
End Function

Function SoloNumero(KeyAscii As Integer) As Integer
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then
        SoloNumero = 0
    Else
        SoloNumero = KeyAscii
    End If
    KeyAscii = SoloNumero
End Function

Sub MensajeError(sCad As String)
    MsgBox sCad, vbOKOnly Or vbCritical, mml_FRASE0096
End Sub

Public Sub CombinarDorsales(iCodCat As Integer, iCodFase As Integer, iRepesca As Integer, iMaxTandas As Integer, Optional iRecombinar As Integer = 0, Optional bAgruparTandas As Boolean = True)
Dim rs As Recordset, rsBailes As Recordset
Dim i As Integer, iValor As Integer, iNumDorsales As Integer, iCDorsales As Integer
Dim iCTandas As Integer

    If Not C_DEBUG Then On Local Error GoTo error

    'Si hay una combinación generada y no marcamos recombinar,mantenemos la recombinación anterior
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsalescombinados WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iRepesca, dbOpenSnapshot)
    If iRecombinar = 0 And rs.Fields(0) > 0 Then
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    'Primero borramos la posible ordenación de dorsales
    db.Execute ("DELETE FROM dorsalescombinados WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iRepesca)
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iRepesca, dbOpenSnapshot)
        iNumDorsales = rs.Fields(0)
    rs.Close
    Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & iCodCat & " AND bc.fase = 2 ORDER BY posicion", dbOpenSnapshot)
    Randomize
    While Not rsBailes.EOF
        'Primero desordenamos los dorsales
        Set rs = db.OpenRecordset("SELECT num_dorsal FROM dorsales WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iRepesca & " ORDER BY 1", dbOpenSnapshot)
        While Not rs.EOF
            ' Primero los desordenamos
            If G_DORSALES_COMBINADOS And CombinarDorsalesCateg(iCodCat) Then
                iValor = Int(Rnd() * 1000)
            Else
                iValor = rs!num_dorsal
            End If
            db.Execute ("INSERT INTO dorsalescombinados VALUES (" & MaxCod("dorsalescombinados") & "," & rs!num_dorsal & "," & iCodCat & "," & iCodFase & "," & iRepesca & "," & rsBailes!codigo & "," & iValor & ")")
            rs.MoveNext
        Wend
        
        If bAgruparTandas Then
            ' Ahora los agrupamos en tandas
            iCDorsales = 0
            iCTandas = 1
            Set rs = db.OpenRecordset("SELECT d.num_dorsal, dc.orden, dc.codigo FROM dorsales d,dorsalescombinados dc WHERE d.num_dorsal = dc.num_dorsal AND dc.cod_categoria = d.cod_categoria AND dc.fase = d.fase AND dc.repesca = d.repesca AND d.cod_categoria = " & iCodCat & " AND d.repesca=" & iRepesca & " AND d.fase =" & iCodFase & " AND dc.cod_baile = " & rsBailes!codigo & " ORDER BY 2,1", dbOpenSnapshot)
            
        Dim iDorsalesTanda As Integer, iDorsalInicial As Integer
            iDorsalInicial = CalcularDorsalInicialTandaCat(iCodCat, iCodFase, iRepesca, iCTandas, iMaxTandas, iDorsalesTanda)
            While Not rs.EOF
                ' Despues los separamos en tandas
                If iCDorsales = iDorsalInicial - 1 + iDorsalesTanda Then
                    Inc iCTandas
                    iDorsalInicial = CalcularDorsalInicialTandaCat(iCodCat, iCodFase, iRepesca, iCTandas, iMaxTandas, iDorsalesTanda)
                End If
                
                db.Execute ("UPDATE dorsalescombinados SET orden =" & iCTandas & " WHERE codigo = " & rs!codigo)
                rs.MoveNext
                Inc iCDorsales
            Wend
            rs.Close
        End If
        rsBailes.MoveNext
    Wend
    
    'si hay que generar dorsales recombinados y agrupados por tandas, recuperamos los dorsales de la primera tanda
    If G_NO_REPETIR_PRIMERA_TANDA Then
        NoRepetirUltimaYPrimeraTanda iCodCat, iCodFase, iRepesca, iMaxTandas, iRecombinar, bAgruparTandas
    End If
    
    Exit Sub
error:
    ProcesarError "RecombinarDorsales"
End Sub

Public Sub NoRepetirUltimaYPrimeraTanda(iCodCat As Integer, iCodFase As Integer, iRepesca As Integer, iMaxTandas As Integer, Optional iRecombinar As Integer = 0, Optional bAgruparTandas As Boolean = True)
Dim rsRepetidosUltima As Recordset
Dim rsNoRepetidosUltima As Recordset
Dim rsBailes As Recordset
Dim iCodBaileAnt As Integer
Dim iTandaDestino As Integer

    If Not C_DEBUG Then On Local Error GoTo error
    If iMaxTandas >= G_NO_TANDAS_NO_REPETIR_PRIMERA_TANDA Then
        iCodBaileAnt = 0
        Set rsBailes = db.OpenRecordset("SELECT DISTINCT b.codigo, b.nombre, bc.posicion FROM bailes b, bailes_categ bc WHERE bc.cod_baile = b.codigo AND bc.cod_categoria = " & iCodCat & " AND bc.fase = 2 ORDER BY posicion", dbOpenSnapshot)
        While Not rsBailes.EOF
            If iCodBaileAnt <> 0 Then
                'Comprobamos los dorsales del baile actual que se encuentran en la última tanda del baile anterior y en la primera del actual
                Set rsRepetidosUltima = db.OpenRecordset("SELECT * FROM dorsalescombinados WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iRepesca & " AND cod_baile = " & rsBailes.Fields("codigo") & " AND orden = 1 AND num_dorsal IN (SELECT num_dorsal FROM dorsalescombinados WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iRepesca & " AND cod_baile = " & iCodBaileAnt & " AND orden = " & iMaxTandas & ")", dbOpenSnapshot)
                'Buscamos dorsales del baile actual que no se encuentren en la última tanda del baile anterior y no se encuentran en la primera del actual
                Set rsNoRepetidosUltima = db.OpenRecordset("SELECT * FROM dorsalescombinados WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iRepesca & " AND cod_baile = " & rsBailes.Fields("codigo") & " AND orden <> 1 AND NOT num_dorsal IN (SELECT num_dorsal FROM dorsalescombinados WHERE cod_categoria = " & iCodCat & " AND fase = " & iCodFase & " AND repesca = " & iRepesca & " AND cod_baile =" & iCodBaileAnt & " AND orden = " & iMaxTandas & ") ORDER BY orden", dbOpenSnapshot)
                'Intercambiamos los dorsales uno a uno
                While Not rsRepetidosUltima.EOF
                    db.Execute "UPDATE dorsalescombinados SET orden = " & rsNoRepetidosUltima.Fields("orden") & " WHERE codigo = " & rsRepetidosUltima.Fields("codigo")
                    db.Execute "UPDATE dorsalescombinados SET orden = 1 WHERE codigo = " & rsNoRepetidosUltima.Fields("codigo")
                    rsRepetidosUltima.MoveNext
                    rsNoRepetidosUltima.MoveNext
                Wend
                rsRepetidosUltima.Close
                rsNoRepetidosUltima.Close
            End If
            iCodBaileAnt = rsBailes.Fields("codigo")
            rsBailes.MoveNext
        Wend
        rsBailes.Close
    End If
    Exit Sub
error:
    ProcesarError "NoRepetirUltimaYPrimeraTanda"
End Sub

Sub SeleccionarCampo(tb As TextBox)
    tb.SelStart = 0
    tb.SelLength = Len(tb.Text)
    
End Sub


Function sQuitarCarProhibidosSQL(sCad As Variant) As String
Dim sCad1 As String
    If IsNull(sCad) Then
        sCad1 = ""
    Else
        sCad1 = sCad
    End If
    CambiarCadena "'", "´", sCad1
    CambiarCadena Chr$(10), "", sCad1
    CambiarCadena Chr$(13), ".", sCad1
    sQuitarCarProhibidosSQL = sCad1
End Function

Function CompeticionIniciada() As Boolean
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & CodCompActiva & ")", dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        CompeticionIniciada = True
    Else
        CompeticionIniciada = False
    End If
    rs.Close
End Function

Function CodCompActiva() As Long
    CodCompActiva = Val(VarCfg("horario_codcompeticion"))
End Function

Function ImpHojaUnica() As Integer
    If C_CAT_UNICA_HOJA_POR_BAILE Then
        ImpHojaUnica = 1
    Else
        ImpHojaUnica = 0
    End If
End Function

Function ImpTandaUnicaCateg(lCodCat As Long) As Boolean
Dim rs As Recordset
    
    If Not C_DEBUG Then On Local Error GoTo error
    ImpTandaUnicaCateg = False
    ' Si se utilizan hojas ópticas no es posible la tanda única
    Set rs = db.OpenRecordset("SELECT imprimir_una_hoja_puntuaciones FROM categorias WHERE codigo = " & lCodCat, dbOpenSnapshot)
    If IsNull(rs.Fields("imprimir_una_hoja_puntuaciones")) Then
        ImpTandaUnicaCateg = C_CAT_UNICA_HOJA_POR_BAILE
    ElseIf rs.Fields("imprimir_una_hoja_puntuaciones") = 1 Then
        ImpTandaUnicaCateg = True
    Else
        ImpTandaUnicaCateg = False
    End If
    rs.Close
    Exit Function
error:
    ProcesarError "ImpTandaUnicaCateg"
End Function


Function sMes(iMes As Integer) As String
    Select Case iMes
        Case 1
            sMes = "January"
        Case 2
            sMes = "February"
        Case 3
            sMes = "March"
        Case 4
            sMes = "April"
        Case 5
            sMes = "May"
        Case 6
            sMes = "June"
        Case 7
            sMes = "July"
        Case 8
            sMes = "August"
        Case 9
            sMes = "September"
        Case 10
            sMes = "October"
        Case 11
            sMes = "November"
        Case 12
            sMes = "December"
        Case Else
            sMes = ""
    End Select
End Function

Function EsFaseDeLaCategoria(lCodCateg As Long, iFase As Integer) As Boolean
Dim rs As Recordset

    If Not C_DEBUG Then On Local Error GoTo error

    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & lCodCateg & " AND fase = " & iFase, dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        EsFaseDeLaCategoria = True
    Else
        EsFaseDeLaCategoria = False
    End If
    rs.Close
    Exit Function
error:
    ProcesarError "EsFaseDeLaCategoria"
    
End Function

Sub ComprobarCategyFase(tbCodCat As TextBox, tbDescCat As TextBox, tbCodFase As TextBox, tbDescFase As TextBox)
Dim sCateg As String
    
    If Not C_DEBUG Then On Local Error GoTo error
    sCateg = sDescCategoria(Val(tbCodCat.Text))
    If Val(tbCodCat.Text) > 0 And sCateg <> "" Then
        tbDescCat.Text = sCateg
    
        If EsFaseDeLaCategoria(Val(tbCodCat.Text), Val(tbCodFase.Text)) Then
            tbDescFase.Text = sDescFase(tbCodFase.Text)
            Exit Sub
        End If
    Else
        tbCodCat.Text = ""
        tbDescCat.Text = ""
    End If

    tbCodFase.Text = ""
    tbDescFase.Text = ""
error:
    ProcesarError "ComprobarCategyFase"
    

End Sub

Sub SelecCBFase(cbFase As ComboBox, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("1")
            cbFase.ListIndex = 1
        Case Asc("2")
            cbFase.ListIndex = 2
        Case Asc("4")
            cbFase.ListIndex = 3
        Case Asc("8")
            cbFase.ListIndex = 4
        Case Asc("6")
            cbFase.ListIndex = 5
        Case Asc("3")
            cbFase.ListIndex = 6
    End Select
    KeyAscii = 0
End Sub

Sub EsperaGrabacionDeFichero(sFichero As String, Optional iEspera As Integer = 200)
Dim lTam As Long

    If Not C_CONTROL_GRABACION_ACTIVO Then
        Exit Sub
    End If
    
    If Not C_DEBUG Then On Local Error GoTo error
    'Comprobamos si todavía se está grabando el fichero
    Sleep iEspera
    lTam = 0
    While lTam <> FileLen(sFichero)
        lTam = FileLen(sFichero)
        DoEvents
        Sleep iEspera
        DoEvents
    Wend
    Exit Sub
error:
    ProcesarError "EsperaGrabadoDeFichero"
End Sub

Sub BorrarPuntuaciones(lCodComp As Long)
    'Borrar todas las puntuaciones si es TeamMatch
    db.Execute "DELETE FROM ResultadosTeamMatch WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & lCodComp & ")"
    'Borrar todas las puntuaciones, incluída la fase seleccionada
    db.Execute "DELETE FROM puntuaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & lCodComp & ")"
    'Borrar todas las descalificaciones, incluída la fase seleccionada
    db.Execute "DELETE FROM descalificaciones WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & lCodComp & ")"
    'Borrar todas las hojas reconocidas, incluída la fase seleccionada
    db.Execute "DELETE FROM hojas_reconocidas WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & lCodComp & ")"
    'Borrar todas los dorsales de fases anteriores a la seleccionada
    db.Execute "DELETE FROM dorsales dor WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & lCodComp & ") AND ((fase < (SELECT MAX(fase) FROM dorsales d WHERE d.cod_categoria = dor.cod_categoria)) OR (repesca = 1 AND fase = (SELECT MAX(fase) FROM dorsales d WHERE d.cod_categoria = dor.cod_categoria)) )"
    'Borrar las combinaciones de dorsales
    db.Execute "DELETE FROM dorsalescombinados dor WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & lCodComp & ") AND ((fase < (SELECT MAX(fase) FROM dorsalescombinados d WHERE d.cod_categoria = dor.cod_categoria)) OR (repesca = 1 AND fase = (SELECT MAX(fase) FROM dorsalescombinados d WHERE d.cod_categoria = dor.cod_categoria)) )"
    'Borrar todas los datos de cal_conjunto (solo FINAL)
    db.Execute "DELETE FROM cal_conjunto WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & lCodComp & ")"
    'Borrar todas los datos de cal_baile
    db.Execute "DELETE FROM cal_baile WHERE cod_categoria IN (SELECT codigo FROM categorias WHERE cod_competicion = " & lCodComp & ")"
    'Borrar todas los datos de publicaciones
    db.Execute "DELETE FROM publicar"

End Sub


Function ExisteCompeticion(lCodComp As Long) As Boolean
Dim rs As Recordset

    If Not C_DEBUG Then On Local Error GoTo error
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM competiciones WHERE codigo = " & lCodComp, dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        ExisteCompeticion = True
    Else
        ExisteCompeticion = False
    End If
    rs.Close
    Exit Function
    
error:
    ProcesarError "ExisteCompeticion"
End Function

Function EsJuezDelPanel(ByVal lCodCateg As Long, sIdJuez As String) As Boolean
    Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM juez_categ WHERE id_juez = '" & sIdJuez & "' AND cod_categoria = " & lCodCateg, dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        EsJuezDelPanel = True
    Else
        EsJuezDelPanel = False
    End If
    rs.Close
End Function


Function NumeroBaileCateg(lCodCat As Long, ByVal iCodFase As Integer, iCodBaile As Integer) As Integer
Dim rs As Recordset

    If C_DEBUG Then On Local Error GoTo error
    iCodFase = IIf(iCodFase = 1, 1, 2)

    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM bailes_categ WHERE cod_categoria = " & lCodCat & " AND fase = " & iCodFase & " AND posicion <= (SELECT posicion FROM bailes_categ WHERE cod_categoria = " & lCodCat & " AND fase = " & iCodFase & " AND cod_baile = " & iCodBaile & ")", dbOpenSnapshot)
        If rs.EOF Then
            MsgBox mml_FRASE1175, vbOKOnly Or vbCritical, G_MSG_ERROR
            NumeroBaileCateg = 6
            Exit Function
        End If
        If rs.Fields(0) > MAX_BAILES Then
            MsgBox mml_FRASE1175, vbOKOnly Or vbCritical, G_MSG_ERROR
            NumeroBaileCateg = 6
            Exit Function
        End If
            
        NumeroBaileCateg = rs.Fields(0)
    rs.Close
    Exit Function
error:
    ProcesarError "NumeroBaileCateg, Cat " & lCodCat & " Phase " & iCodFase & " Dance " & iCodBaile
End Function

Sub RecuperarCatActualHorario(rs As Recordset, sPista As String, iCodcomp As Integer)
    Set rs = db.OpenRecordset("SELECT TOP 1 * FROM horario h WHERE grupo LIKE '" & scarLike & "" & sPista & "" & scarLike & "' AND numfase <> 99 AND cod_competicion = " & iCodcomp & " AND (SELECT COUNT(*) FROM puntuaciones WHERE ((h.cod_baile < 0 AND cod_baile " & G_ORDEN_10B_LAT_EST & " -h.cod_baile) OR h.cod_baile = 0 OR cod_baile = h.cod_baile) AND cod_categoria = h.cod_categoria AND fase = h.numfase AND repesca = h.repesca) = 0 ORDER BY orden", dbOpenSnapshot)
End Sub

Function CorregirNombre(ByVal sNombre As String) As String
Dim i As Integer
Dim bEspacio As Boolean
Dim sCar As String

    bEspacio = True
    CorregirNombre = ""
    For i = 1 To Len(sNombre)
        sCar = Mid$(sNombre, i, 1)
        If sCar = " " Or sCar = "," Or sCar = "+" Or sCar = "(" Then
            bEspacio = True
        Else
            If bEspacio Then
                sCar = UCase(sCar)
            Else
                sCar = LCase(sCar)
            End If
            bEspacio = False
        End If
        CorregirNombre = CorregirNombre & sCar
    Next
    sNombre = CorregirNombre
    CorregirNombre = ""
    'Eliminamos los espacios múltiples
    i = InStr(sNombre, " ")
    While i > 0
        CorregirNombre = CorregirNombre & Mid$(sNombre, 1, i)
        sNombre = Trim$(Mid$(sNombre, i + 1))
        i = InStr(sNombre, " ")
    Wend
    CorregirNombre = CorregirNombre & Trim$(sNombre)
End Function

Function NumeroParejas(ByVal lCategoria As Long, ByVal iFase As Integer, ByVal iRep As Integer) As Integer
Dim rs As Recordset

    NumeroParejas = 0
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM dorsales WHERE cod_categoria = " & lCategoria & " AND fase = " & iFase & " AND repesca = " & iRep, dbOpenSnapshot)
    NumeroParejas = rs.Fields(0)
    rs.Close
    Exit Function
End Function

Function AsignarDorsalesTanda(ByVal lCategoria, ByVal iLimUnaTanda As Integer, bGrabarSoloSiMenor As Boolean) As Boolean
Dim rs As Recordset

    AsignarDorsalesTanda = True
        
    If bGrabarSoloSiMenor Then
        Set rs = db.OpenRecordset("SELECT dorsales_tanda FROM categorias WHERE codigo = " & lCategoria, dbOpenSnapshot)
        If Not rs.EOF Then
            If Val(SinNulos(rs.Fields("dorsales_tanda"))) < iLimUnaTanda Then
                db.Execute "UPDATE categorias SET dorsales_tanda = " & iLimUnaTanda & " WHERE codigo= " & lCategoria
            End If
        Else
            AsignarDorsalesTanda = False
            Exit Function
        End If
        rs.Close
    Else
        db.Execute "UPDATE categorias SET dorsales_tanda = " & iLimUnaTanda & " WHERE codigo= " & lCategoria
    End If
End Function


Sub ActualizarCategorias(tbCat As ComboBox, ByVal lCodComp As Long)
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT * FROM desccategoria ORDER BY 1", dbOpenSnapshot)
    tbCat.Clear
    While Not rs.EOF
        tbCat.AddItem rs!DESCRIPCION
        rs.MoveNext
    Wend
    rs.Close

    If lCodComp > 0 Then
        'Añadimos las distintas categorias que tienen las parejas de la competición no presentes en DescCategoria
        Set rs = db.OpenRecordset("SELECT DISTINCT categoria FROM parejas WHERE cod_competicion = " & lCodComp & " AND NOT categoria IN (SELECT descripcion FROM desccategoria) ORDER BY 1", dbOpenSnapshot)
        While Not rs.EOF
            tbCat.AddItem rs.Fields("categoria")
            rs.MoveNext
        Wend
        rs.Close
    End If
End Sub

Function LimSemiUnaTanda()
    LimSemiUnaTanda = Val(VarCfg("lim_semi_una_tanda"))
End Function

Sub CargarVariablesConfiguracion()
    C_MARCA = Val(VarCfg("puntos_marca"))
    C_BLANCO = Val(VarCfg("puntos_blanco"))
    MAX_LIN_PAG = Val(VarCfg("max_lin_pag_part"))
    C_MEDIO = Int((C_MARCA + C_BLANCO) / 2)
    HAY_JUEZ_PASOS = IIf(VarCfg("juez_pasos") = "S", True, False)
    HAY_CONTROL = IIf(VarCfg("control") = "S", True, False)
    ScaleFactor = Val(VarCfg("escala"))
    C_CAR_DESC_MOD = Val(VarCfg("tam_desc_modalidad"))
    MARGEN_SUPERIOR = Val(VarCfg("margen_superior"))
    C_ORDEN_PAREJAS = VarCfg("orden_parejas")
    C_MIN_DORSAL_OFICIAL = Val(VarCfg("min_dorsal_oficial"))
    C_SALTO_PUNTOS_MARCA = Val(VarCfg("salto_puntos_marca"))
    C_EXTENSION_FICHEROS = VarCfg("extension_ficheros")
    C_FICHERO_INET = VarCfg("fichero_inet")
    C_LOGO_PATH = VarCfg("logo")
    SALTAR_PUESTO_SIG = VarCfg("saltar_puesto_sig")
    C_PREGUNTAR_REPESCA = IIf(VarCfg("preguntar_repesca") = "S", True, False)
    C_PREGUNTAR_REPESCA_SIEMPRE = IIf(VarCfg("preguntar_repesca_siempre") = "S", True, False)
    G_PANEL_LAPIZ_OPTICO = IIf(VarCfg("panel_lapiz_optico") = "0", True, False)
    G_ULTIMO_PUESTO_AUTOMATICO = IIf(VarCfg("ult_puesto_automatico") = "S", True, False)
    G_POS_INIC_CATEG = VarCfg("pos_inic_categ")
    G_ARCH_HORARIO = VarCfg("archivo_horario")
    G_ARCH_SOCIOS_ANULADOS = VarCfg("archivo_socios_anulados")
    G_DIR_PUBLICIDAD = VarCfg("dir_publicidad")
    C_REFRESCO_PUBLICIDAD = VarCfg("refresco_publicidad")
    G_ESPERA_NO_PUBLIC = VarCfg("espera_no_publicidad")
    G_SALTO_PUBLIC = VarCfg("velocidad_publicidad")
    G_COLOR_MARCAS = VarCfg("color_marcas")
    'G_ESCUELA = VarCfg("escuela")
    G_MAX_FILAS_POR_PAG = Val(VarCfg("max_filas_por_pag_tablaa"))
    G_BUSCAR_NOMBRE = VarCfg("buscar_nombre")
    G_AVISO_NUM_MARCAS = IIf(VarCfg("aviso_num_marcas") = "S", True, False)
    G_MARGEN_MARCA_Y = Val(VarCfg("margen_marca_y"))
    G_MARGEN_X_MARCA_CONTROL = Val(VarCfg("margen_x_marca_control"))
    G_MARGEN_CONTROL_PTE = Val(VarCfg("margen_control_pendiente"))
    G_MARCA_CONTROL = IIf(VarCfg("marca_control") = "S", True, False)
    G_MARCAR_BUS_MARCA = IIf(VarCfg("marcar_bus_marca") = "S", True, False)
    G_DORSALES_COMBINADOS = IIf(VarCfg("dorsales_combinados") = "S", True, False)
    G_REC_OPTICO_PARCIAL = IIf(VarCfg("rec_optico_parcial") = "S", True, False)
    G_LOGO_ESCUELA = VarCfg("logo_escuela")
    G_ORDEN_CATEGORIAS = VarCfg("orden_categorias")
    G_HOJA_EXTENDIDA = IIf(VarCfg("hoja_optica_extendida") = "S", True, False)
    G_DEC_POSICIONES_POR_DESCALIFICACION = IIf(VarCfg("dec_posiciones_por_descalificacion") = "S", True, False)
    G_SALTO_ORDEN = Val(VarCfg("salto_orden_horario"))
    G_PUBLICAR_HORA_ESTIMADA = IIf(VarCfg("publicar_hora_estimada") = "S", True, False)
    G_PUBLICAR_POSICION = IIf(VarCfg("publicar_posicion") = "S", True, False)
    G_RESULTADOS_UNO_A_UNO = IIf(VarCfg("resultados_uno_a_uno") = "S", True, False)
    G_ESPERA_ENTRE_PART = Val(VarCfg("espera_entre_part"))
    C_MINUTOS_MARGEN_CALCULAR = Val(VarCfg("min_margen_calcular"))
    G_NO_CONTAR_HOJAS = IIf(VarCfg("no_contar_hojas") = "S", True, False)
    G_TAM_FUENTE_TABLA_SEMI = Val(VarCfg("tam_fuente_tabla_semi"))
    G_ANCHO_COL_JUEZ = Val(VarCfg("ancho_col_juez"))
    G_REC_HOJA_EXT = IIf(VarCfg("rec_hojas_ext") = "S", True, False)
    C_REC_OPTICO_RAPIDO = IIf(VarCfg("modo_rapido_rec_optico") = "S", True, False)
    C_TOLERANCIA_X = Val(VarCfg("tolerancia_x"))
    C_TOLERANCIA_Y = Val(VarCfg("tolerancia_y"))
    C_MAX_PAREJAS_FINAL = Val(VarCfg("max_parejas_final"))
    C_MAX_PAREJAS_SEMIFINAL = Val(VarCfg("max_parejas_semifinal"))
    G_ORDENAR_DESCALIFICADOS_FINAL = VarCfg("ordenar_descalificados_final")
    C_CAT_UNICA_HOJA_POR_BAILE = IIf(VarCfg("cat_unica_hoja_por_baile", "S", mml_FRASE0158) = "S", True, False)
    C_MAX_DORSALES_BAILE_POR_COL = Val(VarCfg("max_dorsales_baile_por_col", "18", mml_FRASE0159))
    C_MAX_COLS_HOJA_POR_BAILE = Val(VarCfg("max_cols_hoja_por_baile", "4", mml_FRASE0160))
    C_MAX_DORSALES_BAILE_POR_COL_HOJA_UNICA = Val(VarCfg("max_dorsales_baile_por_col_hoja_unica", "6", mml_FRASE0159))
    C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA = Val(VarCfg("max_cols_hoja_por_baile_hoja_unica", "13", mml_FRASE0160))
    G_FUENTE_GRANDE_DORSAL = Val(VarCfg("fuente_grande_dorsal", "30", mml_FRASE0161))
    G_FUENTE_PEQUE_DORSAL = Val(VarCfg("fuente_peque_dorsal", "14", mml_FRASE0162))
    G_NOMBRE_FUENTE_DORSAL = VarCfg("nombre_fuente_dorsal", "Arial", mml_FRASE0163)
    G_DORSAL_POR_Y = Val(VarCfg("dorsal_por_y", 1000, mml_FRASE0164))
    G_MARGEN_DORSAL_X = Val(VarCfg("dorsal_margen_x", 200, mml_FRASE0165))
    G_LOGO_DORSAL_IZQ = VarCfg("logo_dorsal_izq", "logo_dorsal.bmp", mml_FRASE0166)
    G_LOGO_DORSAL_DER = VarCfg("logo_dorsal_der", "logo_dorsal_der.bmp", mml_FRASE0167)
    G_PUBLICAR_NUM_GRUPOS_RESULTADOS = Val(VarCfg("publicar_num_grupos_resultados", 5, mml_FRASE0168))
    G_PUNTEO_ANULACION = IIf(VarCfg("circulo_punteo_anulacion", "N", mml_FRASE0169) = "S", True, False)
    G_AUTO_IMP_HOJAS_PUNTUACION = IIf(VarCfg("auto_imp_hojas_puntuaciones", "S", mml_FRASE0170) = "S", True, False)
    G_IMP_HOJAS_BAILE_EN_FINALES = IIf(VarCfg(mml_FRASE0171, "N", mml_FRASE0172) = "S", True, False)
    G_PISTAS_HOJAS_OPTICAS = VarCfg("pistas_con_lectura_optica", "", mml_FRASE0173)
    C_DESC_SIN_PUESTO = IIf(VarCfg("descalificar_sin_puesto", "S", mml_FRASE0174) = "S", True, False)
    C_ORDEN_HORARIO_MODALIDAD = VarCfg("orden_modalidad_horario", "312", mml_FRASE0175)
    G_ORDEN_CATEG_COM = VarCfg("orden_categorias_com", "(""G"",""H"",""I"",""J"")", mml_FRASE0176)
    G_ORDEN_CATEG_EST = VarCfg("orden_categorias_est", "(""D"",""E"",""F"",""G"",""H"",""I"",""J"",""O"",""Open"",""Open I"",""OpInt""):(""A"",""B"",""C"")", mml_FRASE0176)
    G_ORDEN_CATEG_LAT = VarCfg("orden_categorias_lat", "(""D"",""E"",""F"",""G"",""H"",""I"",""J"",""O"",""Open"",""Open I"",""OpInt""):(""A"",""B"",""C"")", mml_FRASE0177)
    G_AUTO_POR_DORSAL = IIf(VarCfg("auto_por_dorsal", "S", mml_FRASE0178) = "S", True, False)
    C_CALCULOS_PARCIALES = IIf(VarCfg("calculos_parciales", "S", mml_FRASE0179) = "S", True, False)
    G_MAX_DORSALES_HOJA_SEMI = Val(VarCfg("max_dorsales_hojas_semi", "40", mml_FRASE0180))
    G_ELIMINATORIAS_PAGINADAS = IIf(VarCfg("eliminatorias_paginadas", "S", mml_FRASE0181) = "S", True, False)
    G_IMPRIMIR_AVISO_GENERAL_LOOK = IIf(VarCfg("imprimir_aviso_general_look", "S", mml_FRASE0193) = "S", True, False)
    G_SALTO_CATEG = Val(VarCfg("salto_fases_categ_horario", "20", mml_FRASE0194))
    G_PREGUNTAR_IMPRESION_AUTO = IIf(VarCfg("preguntar_impresion_auto", "S", mml_FRASE0195) = "S", True, False)
    G_FICHERO_SALIDA = VarCfg("fichero_salida_ppc", "C:\Escrutinio\PocketIRIS", mml_FRASE0196)
    G_FICHERO_ENTRADA = VarCfg("fichero_entrada_ppc", "C:\Escrutinio\fichas\ResultadosIRIS", mml_FRASE0196)
    G_FICHERO_BATERIA = VarCfg("fichero_bateria_ppc", "C:\Escrutinio\fichas\BateriaPPC", "Fichero de control de carga de bateria de los PPC", "File of control of load of battery of the PPC")
    G_FICHERO_CONTROL_JUECES = VarCfg("fichero_control_presencia_jueces", "C:\Escrutinio\fichas\ControlJueces", "Fichero de control de presencia de los jueces", "File of control of the judges´s presence")
    G_FICHERO_HORA = VarCfg("fichero_sincronizacion_hora", "C:\Escrutinio\fichas\SyncTime", "Fichero de sincronización de hora con los PDAs", "File of hour synchronization with the PDAs")
    G_FICHERO_ENTRADA_P2 = VarCfg("fichero_entrada_ppc_p2", "C:\Escrutinio\fichas\P2\ResultadosIRIS", mml_FRASE0196)
    G_CAMBIO_AUTO = IIf(VarCfg("cambio_auto_categ_ppc", "S", mml_FRASE0197) = "S", True, False)
    G_CALCULO_AUTO_PPC = IIf(VarCfg("calculo_auto_ppc", "N", mml_FRASE0198) = "S", True, False)
    G_INTERVALO_TIMER_PPC = Val(VarCfg("intervalo_timer_PPC", "4000", mml_FRASE0199))
    G_APP_GRAFICA = VarCfg("app_grafica", "C:\Archivos de programa\Jasc Software Inc\Paint Shop Pro 7\psp.exe", mml_FRASE0200)
    C_PREGUNTAR_EDIC_HOJA = IIf(VarCfg("preguntar_edicion_hoja", "S", mml_FRASE0201) = "S", True, False)
    G_IMAGENES_EPA_PEQUE = VarCfg("path_epa_peque", "C:\Escrutinio\", mml_FRASE0202)
    G_DIR_PRODANCE = VarCfg("dir_prodance", "C:\PD23", mml_FRASE0203)
    G_DATOS_ORG_PRODANCE = VarCfg("datos_org_prodance", "Miguel Abreu" & Chr$(13) & Chr$(10) & "Mónica" & Chr$(13) & Chr$(10) & "Antonio Losada", mml_FRASE0204)
    G_SELEC_HOJA_EXT_AUTO = IIf(VarCfg("selec_hoja_ext_auto", "S", mml_FRASE0205) = "S", True, False)
    C_SISTEMA_TANDAS_VIEJO = IIf(VarCfg("sistema_viejo_tandas", "N", "Calcula las tandas según el algoritmo antiguo. Todas iguales menos la última que tiene más") = "S", True, False)
    C_NUM_JUECES_ACEPTAR_NO_PRESENTES = Val(VarCfg("num_jueces_aceptar_no_presentes", "2", "Nº de jueces iguales para aceptar los dorsales no presentes"))
    C_PREGUNTA_ACEPTAR_NO_PRESENTES = IIf(VarCfg("pregunta_aceptar_no_presentes", "N", "Pregunta al usuario si acepta los no presentes") = "S", True, False)
    G_PISTAS_PPC = VarCfg("pistas_ppc", "", "Pistas gestionadas integramente por PocketPCs")
    G_GEN_AUTO_RESULTADOS_PPC = IIf(VarCfg("ppc_gen_auto_resultados", "N", "Si es afirmatico genera resultados y hojas de puntuaciones automáticamente desde la pantalla de enlacePPC") = "S", True, False)
    G_TIEMPO_ESPERA_JPASOS_PPC = Val(VarCfg("ppc_espera_descalificaciones_jpasos", "10", "Tiempo de margen para que el juez de pasos mande sus datos"))
    G_NO_MARCAR_BAILES = IIf(VarCfg("no_marcar_bailes", "N", "Si es afirmatico deja vacias todas las marcas de baile presente") = "S", True, False)
    C_RESET_ULTIMOS_5_BAILES = IIf(VarCfg("reset_marca_ultimos_5_bailes", "S", "Indica si la marca de los últimos 5 bailes debe controlarse automáticamente - Enlace PPC") = "S", True, False)
    G_ORDEN_10B_LAT_EST = VarCfg("orden_lat_est_10bailes", "<", "Indica el orden de las modalidades en los Open para localizar en el horario los sig. 10 bailes (por defecto Est-Lat)")
    G_VALORES_MULTIPLES_DBLCLIK = IIf(VarCfg("valores_multiples_dbl_clic", "N", "Activa el modo de selección múltiple al pulsar doble clic") = "S", True, False)
    G_IMPRIMIR_TODOS_LOS_CUADROS = IIf(VarCfg("imprimir_todos_los_cuadros", "N", "Imprime todos los cuadros aunque no haya dorsales") = "S", True, False)
    G_PATH_ESCRUTINIO = VarCfg("path_escrutinio", "C:\Escrutinio\", "Directorio de la base de datos y el fichero de licencia")
    G_SELEC_DORSALES_SIG_FASE = IIf(VarCfg("selec_dorsales_sig_fase", "N", "Permite seleccionar los dorsales para la siguiente fase", "Allows to select the numbers for the following phase") = "S", True, False)
    G_NO_PROC_HOJAS_ERROR = IIf(VarCfg("no_proc_hojas_error", "S", "No inserta como procesadas las hojas en las que hay errores") = "S", True, False)
    G_PARAR_REC_SI_FALLO = IIf(VarCfg("parar_rec_si_fallo", "S", "Para el reconocimiento de hojas si se detecta un fallo en una") = "S", True, False)
    G_ORDEN_SEL = Val(VarCfg("orden_sel", "2", "Ordenación de la selección de parejas", "Order for the couple selection"))
    G_PATH_GRAFICO_HOJAS = VarCfg("path_graficos_hojas", "C:\Escrutinio\grafico_hojas.bmp", "Gráfico que aparece al final de las hojas de puntuaciones", "Graph that appears at the end of the leaves of punctuations")
    G_PPC_GEN_DORSALES_COMBINADOS = IIf(VarCfg("ppc_gen_dorsales_combinados", "S", "Generacion de dorsales desordenados para los PPC", "Generate the numbers disordered for the PPC") = "S", True, False)
    G_COUNTRY = IIf(VarCfg("country", "N", "Baile Country Activo", "Country Active") = "S", True, False)
    C_BAILES_POR_HOJA = Val(VarCfg("bailes_por_hoja", 5, "Bailes por hoja", "Dances for sheet"))
    C_BAILES_POR_HOJA_UNICA = IIf(VarCfg("bailes_por_hoja_unica", "N", "Imprime todos los bailes en una hoja de puntuaciones", "Print All dances in one sheet") = "S", True, False)
    MAX_DORSALES_HOJA_UNICA = C_MAX_DORSALES_BAILE_POR_COL_HOJA_UNICA * C_MAX_COLS_HOJA_POR_BAILE_HOJA_UNICA
    MAX_DORSLES_HOJA_PUNT_COL_DOBLE = Val(VarCfg("max_dorsles_hoja_punt_col_doble", 78, "Nº máximo de dorsales para mostrar columnas dobles en hojas de puntuaciones", "Maximum numbers to show double columns in leaves of punctuations"))
    C_NO_COPIAS_COMBINACION = Val(VarCfg("no_copias_combinacion", "3", "Nº de copias de la hoja de combinación de tandas", "Number of copies of the sheet of hits combination"))
    G_IMAGEN_FONDO_DIPLOMA = VarCfg("imagen_fondo_diploma", "C:\Escrutinio\diploma.bmp", "Imagen de fondo de los diplomas", "Background image of the doploma")
    G_MARGEN_IZQ_TABLA_DIPLOMAS = Val(VarCfg("margen_izq_datos_diploma", "3000", "Márgen izquierdo de los datos del diploma", "Left margin of the data of the diploma"))
    G_MARGEN_SUP_TABLA_DIPLOMAS = Val(VarCfg("margen_sup_datos_diploma", "4000", "Márgen superior de los datos del diploma", "Upper margin of the data of the diploma"))
    G_MARGEN_IZQ_IMAGEN_DIPLOMAS = Val(VarCfg("margen_izq_imagen_diploma", "400", "Márgen izquierdo de la imagen del diploma", "Left margin of the image of the diploma"))
    G_MARGEN_SUP_IMAGEN_DIPLOMAS = Val(VarCfg("margen_sup_imagen_diploma", "400", "Márgen superior de la imagen del diploma", "Upper margin of the image of the diploma"))
    G_MARGEN_SUP_NOMBRE_COMP = Val(VarCfg("margen_sup_nombre_competicion", "10000", "Márgen superior del nombre de la competición en el diploma (0 si no debe aparecer)", "Upper margin of the name of the competition in the diploma (0 if should not appear)"))
    G_MARGEN_IZQ_NOMBRE_COMP = Val(VarCfg("margen_izq_nombre_competicion", "4000", "Márgen izquierdo del nombre de la competición en el diploma (0 si no debe aparecer)", "Left margin of the name of the competition in the diploma (0 if should not appear)"))
    G_ANCHO_IMAGEN_DIPLOMAS = Val(VarCfg("ancho_imagen_diploma", "16000", "Ancho de la imagen del diploma", "Witdh of the image of the diploma"))
    G_ALTO_IMAGEN_DIPLOMAS = Val(VarCfg("alto_imagen_diploma", "11000", "Alto de la imagen del diploma", "Height of the image of the diploma"))
    G_DIPLOMA_TITULO_FUENTE = VarCfg("diploma_titulo_fuente", "Algerian", "Fuente del título del diploma", "Font of the diploma title")
    G_ESPERA_FICH_INET = Val(VarCfg("espera_fichero_inet", "20", "Espera en segundos por la generación del fichero de internet por el driver de impresora", "Time of Wait for the generation of the internet file"))
    G_LIM_JUECES_PARA_TABLAS_FINAL_DOBLES = Val(VarCfg("max_jueces_tablas_final_dobles", "13", "Máximo número de jueces con los que se imprimirán tablas de baile en paralelo en la hoja de resultados de la final", "Maximum number of judges with those that dance charts will be printed in parallel in the leaf of results of the final one"))
    G_TIEMPO_PARA_PERDIDA_DE_CONEXION = Val(VarCfg("tiempo_perdida_conexion", "50", "Segundos transcurridos sin comunicación del PDA para indicar fallo de red", "Seconds lapsed without communication of the PDA to indicate net failure"))
    G_CATEGORIA_COMBINACION = Val(VarCfg("categoria_combinacion", "38", "Código de categoria del resultado de la combinación de bailes", "Code of category of the result of the combination of dances"))
    G_PUESTOS_CON_DIPLOMA = Val(VarCfg("puestos_con_diploma", "5", "Número de puestos con diploma", "Number of positions with diploma"))
    G_CADENA_SEPARADOR_CAMPOS = VarCfg("cadena_separador_campos", ",", "Cadena de separación de campos en la importación de dorsales desde fichero", "String of separation of fields in the import of numbers from file")
    G_CADENA_DELIMITADOR_VALORES = VarCfg("cadena_delimitador_valores", """", "Cadena entre la que tienen que encerrarse los campos en la importación de dorsales desde fichero", "String among which you/they have to lock the fields in the import of numbers from file")
    G_TIMER_CONTROL_BATERIA = Val(VarCfg("timer_control_bateria", "2", "Unidades de 4 segundos para recarga del estado de bateria", "Units of 4 seconds for recharge of the battery state"))
    G_DIR_COPIA_BD = VarCfg("dir_copia_bd", "C:\Escrutinio\COPIA_BD", "Directorio donde se realizan las copias de seguridad de la base de datos y ficheros de puntuaciones", "Directory where they are carried out the copies of security of the database and files of punctuations")
    G_MOSTRAR_PUNTOS = IIf(VarCfg("mostrar_puntos_en_hojas_finales", "N", "Muestra los puntos en las hojas de puntuaciones", "Show the points in the leaf of punctuations") = "S", True, False)
    G_NO_REPETIR_PRIMERA_TANDA = IIf(VarCfg("no_repetir_primera_tanda", "S", "El sistema Genera combinaciones siempre distintas para la última tanda y la primera tanda de cada baile", "The system always Generates combinations different for the last shift and the first shift of each dance") = "S", True, False)
    G_NO_TANDAS_NO_REPETIR_PRIMERA_TANDA = Val(VarCfg("num_tandas_no_repetir_primera_tanda", "3", "El sistema Genera combinaciones siempre distintas para la última tanda y la primera tanda de cada baile a partir de este número de tandas", "The system always Generates combinations different for the last shift and the first shift of each dance starting from this number of hits"))
    G_NO_PROCESAR_FICH_DE_OTRA_CATEG = IIf(VarCfg("pda_no_procesar_fich_cat_no_actual", "S", "Si se detecta un fichero de una categoría que no es la activa, no se procesa", "If a file of a category is detected that is not the active one, it is not processed") = "S", True, False)
    G_NO_PUBLICAR_COMO_ANT_PANEL_DERECHO = IIf(VarCfg("res_no_publicar_como_anterior_panel_derecho", "S", "En el panel de resultados anteriores no se publica el resultado actual", "In the panel of previous results the current result is not published") = "S", True, False)
    G_RETARDO_RESULTADOS_MULTIPLES_PANTALLAS = Val(VarCfg("res_retardo_resultados_multiples_pantalla", "20", "Décimas de segundo de retardo en una publicación de resultados de multiples pantallas", "Tenth of second of retard in a publication of results of multiple screens"))
    G_LIM_SEMI_UNA_TANDA = Val(VarCfg("lim_semi_una_tanda", "13", "Nº máximo de dorsales que puede tener una semifinal de múltiples tandas para convertirla en una tanda (0 deshabilita la conversión automática en EnlacePPC)", "Maximum Nº of numbers that can have a semifinal of multiple shifts to transform it into a semifinal of one shift (0 disable the automatic conversion in EnlacePPC)"))
    G_ACTUALIZAR_ULTIMA_PUBLICACION = Val(VarCfg("res_periodo_actualizar_ultima_publicacion", "4", "Actualiza periódicamente el panel de derecha de la última publicación (0 no actualiza)", "The system Upgrades the panel of right of the last publication periodically (0 not upgrade)"))
    G_SOLO_UN_PC = IIf(VarCfg("pda_solo_un_pc", "N", "Indica que pistas múltiples se gestionan desde el mismo ordenador", "This parameter indicates that multiple hints are managed from the same computer") = "S", True, False)
    G_ASIGNAR_AUTOMATICAMENTE_LETRA_A_JUEZ = IIf(VarCfg("asignar_juez_automatico", "S", "El sistema asigna automáticament el juez a la letra si ya ha sido asignado en la competición", "The system assigns the judge automatically to the letter if it has already been assigned in the competition") = "S", True, False)
    G_FICHERO_ELIMINADOS = VarCfg("fichero_eliminados", "C:\Escrutinio\ELIMINADOS.TXT", "Fichero que contiene la información de los dorsales eliminados", "File that contains the information of the eliminated numbers")
    G_ADM_EQUIPOS = VarCfg("administrador_equipos", "C:\WINDOWS\SYSTEM32\MMC.EXE c:\windows\system32\compmgmt.msc", "Administrador de equipos", "Computer Manager")
    G_MOVER_FICHEROS_PDA = IIf(VarCfg("pda_mover_ficheros", "S", "Indica si los ficheros de los pdas se copian a otro directorio antes de procesarlos", "This Indian parameters if the files of the pdas are copied to another directory before processing them") = "S", True, False)
    G_RUTA_COPIA_FICH_PDA = VarCfg("pda_copia_fich", "C:\Escrutinio\fichas_copia", "Directorio de copia de los ficheros de los pdas", "Copy Path of the PDAs Files")
    G_SEG_MAX_CONTROL_BATERIA = Val(VarCfg("bateria_seg_max_valided_control", "900", "Nº máximo de segundos que se considera válido una medición de carga de bateria", "Maximum Nº of seconds that is considered valid a mensuration of battery load"))
    G_UNIDADES_NIVEL_MIN_CONTROL = Val(VarCfg("bateria_unidades_minimas_control", "2", "Nº de unidades de porcentaje de carga de bateria para realizar la actualización de tiempo de descarga", "Nº of units of percentage of battery load to carry out the upgrade of time of discharge"))
    G_NO_PRESENTES_AUTO = IIf(VarCfg("no_presentes_auto", "S", "Indica si se detectan automáticamente los dorsales no presentes por las indicaciones de los jueces", "The numbers don`t present they are detected automatically") = "S", True, False)
    
    G_MARCA_MAYOR = IIf(VarCfg("marca_mayor") = "S", True, False)
    
    G_DESPLAZ_VIS_HOJA = Val(VarCfg("desplaz_vis_hoja"))
    G_MAX_FILA_VIS_HOJA = Val(VarCfg("max_fila_vis_hoja"))
    G_LINEAS_DIVISION_FINAL = IIf(VarCfg("lineas_division_final") = "S", True, False)
    
    G_DESPLAZAR_SI_CONTROL_NO_LOCALIZADO = IIf(VarCfg("desplazar_si_control_no_localizado") = "S", True, False)
    
    G_CAB_INET = VarCfg("cab_inet")
    G_PIE_INET = VarCfg("pie_inet")
    G_CAB_TABLA = VarCfg("cab_tabla")
    
    G_MINUTOS_POR_CATEG = Val(VarCfg("minutos_por_categ"))
    
    G_CAB_RESULTADOS = VarCfg("cab_resultados")
    G_CAB1_RESULTADOS = VarCfg("cab1_resultados")
    
    G_MARCAR_PUNTOS = IIf(UCase(VarCfg("marcar_puntos")) = "S", True, False)
    G_MOSTRAR_NUM_PUNTOS = IIf(UCase(VarCfg("mostrar_num_puntos")) = "S", True, False)

    G_VELOCIDAD_LIBRETA = Val(VarCfg("velocidad_libreta"))
    G_PUB_RES_DEC_LETRAS = Val(VarCfg("decremento_letras_pub_resultados"))
    G_RETRASO_RESULTADOS = Val(VarCfg("retraso_resultados"))
    G_PAREJAS_AEBDC_UN_APELLIDO = VarCfg("parejas_aebdc_un_apellido", "N", "Carga las parejas AEBDC con un único apellido", "Load the couples AEBDC with only the last name")
    
 

End Sub
Sub EstablecerFase(cbFase As ComboBox, ByVal iFase As Integer)
Dim i As Integer
    For i = 0 To cbFase.ListCount - 1
        If Val(cbFase.List(i)) = iFase Then
            cbFase.ListIndex = i
            Exit Sub
        End If
    Next
End Sub

Function MinFaseCateg(ByVal lCodCateg As Long) As Integer
Dim rs As Recordset

    MinFaseCateg = -1
    'Calcular la menor fase
    Set rs = db.OpenRecordset("SELECT MIN(fase) as min_fase FROM dorsales WHERE cod_categoria = " & lCodCateg, dbOpenSnapshot)
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("min_fase")) Then
            MinFaseCateg = rs.Fields("min_fase")
        End If
    End If
    rs.Close

End Function

Sub CrearDirectorios(sDirFichas As String, Optional sPath As String = "")
    
    On Local Error Resume Next
    MkDir sDirFichas
    MkDir sDirFichas & "\COPIA_PDA\"
    MkDir sDirFichas & "\COPIA_PDA\FORZADA\"
    MkDir sDirFichas & "\Inet\"
    MkDir sDirFichas & "\TMP\"
    MkDir sDirFichas & "\TMP\Errores\"
    MkDir sDirFichas & "\TMP\PuntuacionesManuales\"
    MkDir sDirFichas & "\P1"
    MkDir sDirFichas & "\P2"
    MkDir sDirFichas & "\P3"
    MkDir sDirFichas & "\P4"
    MkDir sDirFichas & "\P5"
    MkDir sDirFichas & "\P6"
    MkDir sDirFichas & "\P7"
    MkDir sDirFichas & "\P8"
    MkDir sDirFichas & "\P9"
    
    If sPath <> "" Then
        MkDir sPath & "\P1"
        MkDir sPath & "\P2"
        MkDir sPath & "\P3"
        MkDir sPath & "\P4"
        MkDir sPath & "\P5"
        MkDir sPath & "\P6"
        MkDir sPath & "\P7"
        MkDir sPath & "\P8"
        MkDir sPath & "\P9"
        MkDir sPath & "\COPIA_BD"
    End If
End Sub

Function TienePuntuaciones(lCodCateg As Long) As Boolean
Dim rs As Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM puntuaciones WHERE cod_categoria = " & lCodCateg, dbOpenSnapshot)
    If rs.Fields(0) > 0 Then
        TienePuntuaciones = True
    Else
        TienePuntuaciones = False
    End If
    rs.Close
End Function

Function PanelJuecesCateg(ByVal lCodCateg As Long) As String
Dim rs As Recordset

    Set rs = db.OpenRecordset("SELECT cod_panel FROM paneles WHERE cod_categoria = " & lCodCateg, dbOpenSnapshot)
    If Not rs.EOF Then
        PanelJuecesCateg = rs.Fields("cod_panel")
    Else
        PanelJuecesCateg = ""
    End If
    rs.Close
End Function

Function PreguntaOperacion() As Boolean

    If MsgBox(G_PREGUNTA_OPERACION, vbYesNo Or vbQuestion, G_MSG_PREGUNTA) = vbYes Then
        PreguntaOperacion = True
    Else
        PreguntaOperacion = False
    End If
    
End Function

Function QuitarAcentos(sCad As String) As String
Dim i As Integer

    QuitarAcentos = ""
    For i = 1 To Len(sCad)
        QuitarAcentos = QuitarAcentos & CarSinAcento(Mid(sCad, i, 1))
    Next

End Function
Function CarSinAcento(sCad As String) As String
    Select Case sCad
        Case "á"
            CarSinAcento = "a"
        Case "é"
            CarSinAcento = "e"
        Case "í"
            CarSinAcento = "i"
        Case "ó"
            CarSinAcento = "o"
        Case "ú"
            CarSinAcento = "u"
        Case "ñ"
            CarSinAcento = "n"
        Case "Á"
            CarSinAcento = "A"
        Case "É"
            CarSinAcento = "E"
        Case "Í"
            CarSinAcento = "I"
        Case "Ó"
            CarSinAcento = "0"
        Case "Ú"
            CarSinAcento = "U"
        Case "Ñ"
            CarSinAcento = "Ñ"
        Case "º"
            CarSinAcento = "o"
        Case "ª"
            CarSinAcento = "a"
        Case "ç"
            CarSinAcento = "c"
        Case "Ç"
            CarSinAcento = "C"
        Case Else
            CarSinAcento = sCad
    End Select
End Function

