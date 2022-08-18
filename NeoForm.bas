Attribute VB_Name = "NeoForm"
' ==============================================================================
' NeoForm Module - Create irregular shaped forms with ease!
' ==============================================================================
' This a compilation (with tweaks and changes here and there) of code I've found
' around the net. Thanks (and credit) goto Unlimited Realities and The Scarms for teaching me
' the concepts. They are really valuable resources for those in need of more
' challenging code and projects. Programming is fun when you understand what
' you're doing! :)
' ------------------------------------------------------------------------------
' Instructions: Use the following format.
'     InitializeSurfaceCapture <Form>
'     |
'     |-> CreateSurfacefromPoints <X1>,<Y1>,<X2>,<Y2>,...,<Xn>,<Yn>,<X1>,<Y1>
'     |-> CreateSurfacefromRect <X1>,<Y1>,<X2>,<Y2>
'     |-> CreateSurfacefromEllipse <X1>,<Y1>,<X2>,<Y2>
'     |-> CreateSurfacefromMask <Object with picture element> , <Optional
'     |       transparency color>
'     |       (if transparency isn't provided, the color at 0,0 will be used.
'     |       this is more accurate since color depths are different per
'     |       bitpitch.)
'     |
'     ReleaseSurfaceCapture <Form>
'
'     Example Code:
'       Private Sub Form_Load()
'           InitializeSurfaceCapture Me
'           CreateSurfacefromPoints 0, 0, 40, 40, 60, 180, 0, 0
'           CreateSurfacefromEllipse 120, 120, 180, 180
'           CreateSurfacefromRect 60, 60, 150, 150
'           CreateSurfacefromMask Me
'           ReleaseSurfaceCapture Me
'       End Sub
'
'     The best way to create a good mask from a picture, is to load it into the
'     form's picture property. It'll give you a visual preview of how the final
'     shape would be.
'
'     NOTE: CreateSurfacefromMask will compute the regions in two ways. The first
'     uses a ptr array, and the second uses API GetPixel. The ptr array is much,
'     much faster but I've only been able to use 256 color bmps (1 planar) on it.
'     But nevertheless, if the bmp <> 256 colors, the routine will shunt the pic
'     to the second routine to compute via GetPixel, which is accurate in all
'     bits per px bmps, but really, really slow. (I'm trying to get the ptr array
'     method to work with all bitpitches.)
'
'     If you need your form to move with mousedown, add :
'          HookSurfaceHwnd <Form>
'     to the form's mousedown method.
'
'     Example Code:
'       Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'           HookSurfaceHwnd Me
'       End Sub
'
' ------------------------------------------------------------------------------
'     Make sure Form.AutoRedraw = True and Form.BorderStyle = None and also
'     recommend taking out the caption and controlbox [IF] you're using the
'     mask to create a surface, else it looks nasty! But you're free to
'     experiment! :)
' ------------------------------------------------------------------------------
' Limitations:
'     1) Only objects that have the picture property can be used to make masks.
'     2) Fast routines only support 256 bit graphics currently.
' ------------------------------------------------------------------------------
' Remember to add exit methods when there exists no close button.
' ==============================================================================
' Freeware as always. Give credit where it's due.
' Proud member of the FreeCode project.
' ------------------------------------------------------------------------------
' Compiled by KoPP3x - [EarthDate|56|29|21|18|12|2000] - nasa_jpl@hotmail.com
' ==============================================================================

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1
Const RGN_OR = 2

Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "USER32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Public lRegion As Long

Public Sub InitializeSurfaceCapture(frm As Form)
    lRegion = CreateRectRgn(0, 0, 0, 0)
    frm.Visible = False
End Sub

Public Sub ReleaseSurfaceCapture(frm As Form)
    ApplySurfaceTo frm
    frm.Visible = True
    Call DeleteObject(lRegion)
End Sub

Public Sub ApplySurfaceTo(frm As Form)
    Call SetWindowRgn(frm.hwnd, lRegion, True)
End Sub

' Create a polygonal region - has to be more than 2 pts (or 4 input values)
Public Sub CreateSurfacefromPoints(ParamArray XY())
    Dim lRegionTemp As Long
    Dim XY2() As POINTAPI
    Dim nIndex As Integer
    Dim nTemp As Integer
    Dim nSize As Integer
    nSize = CInt(UBound(XY) / 2) - 1
    ReDim XY2(nSize + 2)
    nIndex = 0
    For nTemp = 0 To nSize
        XY2(nTemp).X = XY(nIndex)
        nIndex = nIndex + 1
        XY2(nTemp).Y = XY(nIndex)
        nIndex = nIndex + 1
    Next nTemp
    lRegionTemp = CreatePolygonRgn(XY2(0), (UBound(XY2) + 1), 2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)
End Sub

' Create a ciruclar/elliptical region
Public Sub CreateSurfacefromEllipse(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    Dim lRegionTemp As Long
    lRegionTemp = CreateEllipticRgn(X1, Y1, X2, Y2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)
End Sub

' Create a rectangular region
Public Sub CreateSurfacefromRect(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    Dim lRegionTemp As Long
    lRegionTemp = CreateRectRgn(X1, Y1, X2, Y2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)
End Sub

' My best creation (more like tweak) yet! Super fast routines qown j00!
Public Sub CreateSurfacefromMask(obj As Object, Optional lBackColor As Long)
    ' Insight: Down with getpixel!!
    Dim lReturn   As Long
    Dim lRgnTmp   As Long
    Dim lSkinRgn  As Long
    Dim lStart    As Long
    Dim lRow      As Long
    Dim lCol      As Long
    Dim glHeight  As Integer
    Dim glWidth   As Integer
    Dim pict() As Byte
    Dim pict2() As Byte
    Dim sa As SAFEARRAY2D
    Dim bmp As BITMAP
    GetObjectAPI obj.Picture, Len(bmp), bmp
    ' Load the bmp into a safearray ptr
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = bmp.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bmp.bmWidthBytes
        .pvData = bmp.bmBits
    End With
    ' Unfortunately this only supports 256 color bmps (damn high bit graphics!!)
    If bmp.bmBitsPixel <> 8 Then
        CreateSurfacefromMask_GetPixel obj
        Exit Sub
    End If
    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Get the dimensions for future reference
    glHeight = UBound(pict, 2)
    glWidth = UBound(pict, 1)
    ' Create an identity array to flip the damn inversed regions
    ReDim pict2(glWidth, glHeight)
    ' Flip em!
    Dim nTempX As Integer
    Dim nTempY As Integer
    For nTempX = glWidth To 0 Step -1
        For nTempY = glHeight To 0 Step -1
            pict2(nTempX, nTempY) = pict(nTempX, glHeight - nTempY)
        Next nTempY
    Next nTempX
    ' Clear the original array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    ' Let's make our regions!
    lSkinRgn = CreateRectRgn(0, 0, 0, 0)
    With obj
        If lBackColor < 1 Then lBackColor = pict2(0, 0)
        For lRow = 0 To glHeight
            lCol = 0
            Do While lCol < glWidth
                Do While lCol < glWidth
                    If pict2(lCol, lRow) = lBackColor Then
                        lCol = lCol + 1
                    Else
                        Exit Do
                    End If
                Loop
                If lCol < glWidth Then
                    lStart = lCol
                    Do While lCol < glWidth
                        If pict2(lCol, lRow) <> lBackColor Then
                            lCol = lCol + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    If lCol > glWidth Then lCol = glWidth
                    lRgnTmp = CreateRectRgn(lStart, lRow, lCol, (lRow + 1))
                    lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                    Call DeleteObject(lRgnTmp)
                End If
            Loop
        Next
    End With
    ' Clear the identity array
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
    ' Return the f****** fast generated region!
    lReturn = CombineRgn(lRegion, lRegion, lSkinRgn, RGN_OR)
End Sub

' XCopied from The Scarms! Felt like my obligation to leave this code intact w/o
' any changes to variables, etc (cept for the sub's name). Thanks d00d!
Public Sub CreateSurfacefromMask_GetPixel(obj As Object, Optional lBackColor As Long)
    Dim lReturn   As Long
    Dim lRgnTmp   As Long
    Dim lSkinRgn  As Long
    Dim lStart    As Long
    Dim lRow      As Long
    Dim lCol      As Long
    Dim glHeight  As Integer
    Dim glWidth   As Integer
    lSkinRgn = CreateRectRgn(0, 0, 0, 0)
    With obj
        glHeight = .Height / Screen.TwipsPerPixelY
        glWidth = .Width / Screen.TwipsPerPixelX
        If lBackColor < 1 Then lBackColor = GetPixel(.hDC, 0, 0)
        For lRow = 0 To glHeight - 1
            lCol = 0
            Do While lCol < glWidth
                Do While lCol < glWidth And GetPixel(.hDC, lCol, lRow) = lBackColor
                    lCol = lCol + 1
                Loop
                If lCol < glWidth Then
                    lStart = lCol
                    Do While lCol < glWidth And GetPixel(.hDC, lCol, lRow) <> lBackColor
                        lCol = lCol + 1
                    Loop
                    If lCol > glWidth Then lCol = glWidth
                    lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                    lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                    Call DeleteObject(lRgnTmp)
                End If
            Loop
        Next
    End With
    lReturn = CombineRgn(lRegion, lRegion, lSkinRgn, RGN_OR)
End Sub

Public Sub HookSurfaceHwnd(frm As Form)
    Call ReleaseCapture
    Call SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
