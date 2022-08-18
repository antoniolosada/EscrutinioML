Attribute VB_Name = "modImagen"
Option Explicit
'------------------------------------------------------------------
'Código para grabar y leer imagenes en campos de bases ( 9/Abr/98)'
'Adaptado de un par de ejemplos de la ayuda de VB5'
'©Guillermo 'guille' Som, 1998 <guille@costasol.net>
'------------------------------------------------------------------
Dim DataFile As Integer
Dim Chunk() As Byte
Const conChunkSize As Integer = 16384

Private Type PALETTEENTRY
         peRed As Byte
         peGreen As Byte
         peBlue As Byte
         peFlags As Byte
End Type
Private Type LOGPALETTE
         palVersion As Integer
         palNumEntries As Integer
         palPalEntry(255) As PALETTEENTRY
End Type
Private Type GUID
     Data1 As Long
     Data2 As Integer
     Data3 As Integer
     Data4(7) As Byte
End Type
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "USER32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Type PicBmp
     Size As Long
     Type As Long
     hBmp As Long
     hPal As Long
     Reserved As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

     Dim hDCMemory As Long
     Dim hBmp As Long
     Dim hBmpPrev As Long
     Dim r As Long
     Dim hDCSrc As Long
     Dim hPal As Long
     Dim hPalPrev As Long
     Dim RasterCapsScrn As Long
     Dim HasPaletteScrn As Long
     Dim PaletteSizeScrn As Long
     Dim LogPal As LOGPALETTE ' Depending on the value of Client get the proper device context
     If Client Then
         hDCSrc = GetDC(hWndSrc) ' Get device context for client area
     Else
         hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire ' window
     End If
     ' Create a memory device context for the copy process
     hDCMemory = CreateCompatibleDC(hDCSrc)
     ' Create a bitmap and place it in the memory DC
     hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
     hBmpPrev = SelectObject(hDCMemory, hBmp)
     ' Get screen properties
     RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) 'Raster 'capabilities
     HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette 'support
     PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
     ' palette
     'If the screen has a palette make a copy and realize it
     If HasPaletteScrn And (PaletteSizeScrn = 256) Then
         ' Create a copy of the system palette
         LogPal.palVersion = &H300
         LogPal.palNumEntries = 256
         r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
         hPal = CreatePalette(LogPal)
         ' Select the new palette into the memory DC and realize it
         hPalPrev = SelectPalette(hDCMemory, hPal, 0)
         r = RealizePalette(hDCMemory)
     End If
     ' Copy the on-screen image into the memory DC
     r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
     ' Remove the new copy of the on-screen image
     hBmp = SelectObject(hDCMemory, hBmpPrev)
     ' If the screen has a palette get back the palette that was selected in previously
     If HasPaletteScrn And (PaletteSizeScrn = 256) Then
         hPal = SelectPalette(hDCMemory, hPalPrev, 0)
     End If
     ' Release the device context resources back to the system
     r = DeleteDC(hDCMemory)
     r = ReleaseDC(hWndSrc, hDCSrc)
     ' Call CreateBitmapPicture to create a picture object from the
     ' bitmap and palette handles. Then return the resulting picture ' object.
     Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
     Dim r As Long
     Dim Pic As PicBmp ' IPicture requires a reference to "Standard OLE Types"
     Dim IPic As IPicture
     Dim IID_IDispatch As GUID ' Fill in with IDispatch Interface ID
     With IID_IDispatch
         .Data1 = &H20400
         .Data4(0) = &HC0
         .Data4(7) = &H46
     End With
     ' Fill Pic with necessary parts
     With Pic
         .Size = Len(Pic) ' Length of structure
         .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
         .hBmp = hBmp ' Handle to bitmap
         .hPal = hPal ' Handle to palette (may be null)
     End With
     ' Create Picture object
     r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
     ' Return the new Picture object
     Set CreateBitmapPicture = IPic
End Function


Public Sub GuardarBinary(campoBinary As Field, unPicture As PictureBox)
    'Guardar el contenido del Picture en el campo de la base
    Dim i As Integer
    Dim Fragment As Integer, Fl As Long, Chunks As Integer
    'NOTA:
    ' El recordset debe estar preparado para Editar o Añadir
    'Guardar el contenido del picture en un fichero temporal
    SavePicture unPicture.Picture, "pictemp"
    'Leer el fichero y guardarlo en el campo
    DataFile = FreeFile
    Open "pictemp" For Binary Access Read As DataFile
    Fl = LOF(DataFile) ' Longitud de los datos en el archivo
    If Fl = 0 Then Close DataFile: Exit Sub
    Chunks = Fl \ conChunkSize
    Fragment = Fl Mod conChunkSize
    ReDim Chunk(Fragment)
    Get DataFile, , Chunk()
    campoBinary.AppendChunk Chunk()
    ReDim Chunk(conChunkSize)
    For i = 1 To Chunks
        Get DataFile, , Chunk()
        campoBinary.AppendChunk Chunk()
    Next i
    Close DataFile
    'Ya no necesitamos el fichero, así que borrarlo
    On Local Error Resume Next
    If Len(Dir$("pictemp")) Then
        Kill "pictemp"
    End If
    Err = 0
End Sub


Public Sub LeerBinary(campoBinary As Field, unPicture As PictureBox)
    'Leer la imagen del campo de la base y asignarlo al Picture
    Dim lngCompensación As Long
    Dim lngTamañoTotal As Long
    'Se usa un fichero temporal para guardar la imagen
    DataFile = FreeFile
    Open "pictemp" For Binary Access Write As DataFile
    lngTamañoTotal = campoBinary.FieldSize
    Do While lngCompensación < lngTamañoTotal
        Chunk() = campoBinary.GetChunk(lngCompensación, conChunkSize)
        Put DataFile, , Chunk()
        lngCompensación = lngCompensación + conChunkSize
    Loop
    Close DataFile
        'Ahora se carga esa imagen en el control
    unPicture.Picture = LoadPicture("pictemp")
    'Ya no necesitamos el fichero, así que borrarlo
    On Local Error Resume Next
    If Len(Dir$("pictemp")) Then
        Kill "pictemp"
    End If
    Err = 0
End Sub
