Attribute VB_Name = "Main"
   Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
   End Type
   Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
'  Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
   Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
   Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'  Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
   Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'  Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
   Declare Function newGetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
  
   Public Const SRCCOPY = &HCC0020
'  Public Const SRCAND = &H8800C6
'  Public Const SRCPAINT = &HEE0086
'  Public Const NOTSRCCOPY = &H330008
