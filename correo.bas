Attribute VB_Name = "correos"
Public Type TConfig
   Disco As String
   Directorio As String
   FicheroAnexo As String
   De As String
   Para As String
   Smtp As String
   AnexoVisible As Boolean
   HayAnexo As Boolean
End Type
Public Cfg As TConfig
Public Paso As Integer
Public Cadena As String
Public NFich As String
Public ConexionEstablecida As Boolean
Public Cabecera(100) As String
Public Boundary As String
Public InicioAnexo As Integer
Public InicioBoundaryAnexo As Integer
Public EnviandoMensaje As Boolean
Public ContTotal As Long

Function EnviarDatos(Cadena As String)
   ' Cambio el color del LED de Efectos Especiales.
   correo.Sck.SendData (Cadena)
   ' Este temporizador sirve de "TimeOut"
   ' Si pasan X segundos y aun no hay respuesta se genera un mensaje de error
   correo.TimeOut.Enabled = True
End Function
