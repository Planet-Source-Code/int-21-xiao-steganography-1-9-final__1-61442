Attribute VB_Name = "Module"
Option Explicit
Public bPurchase As Boolean
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Public Enum TypeFileEncode
'    BMPFile = 1
'    MP3File
'End Enum

Public Type tBits
    Bits(0 To 7) As Single
End Type

'Public vHex, vBin

'file header, total 14 bytes
Type winBMPFileHeader
     strFileType As String * 2 ' file type always 4D42h or "BM"
     lngFileSize As Long       'size in bytes ussually 0 for uncompressed
     bytReserved1 As Integer   ' always 0
     bytReserved2 As Integer   ' always 0
     lngBitmapOffset As Long   'starting position of image data in bytes
End Type


'image header, total 40 bytes
Type BITMAPINFOHEADER
     biSize As Long          'Size of this header
     biWidth As Long         'width of your image
     biHeight As Long        'height of your image
     biPlanes As Integer     'always 1
     byBitCount As Integer   'number of bits per pixel 1, 4, 8, or 24
     biCompression As Long   '0 data is not compressed
     biSizeImage As Long     'size of bitmap in bytes, typicaly 0 when uncompressed
     biXPelsPerMeter As Long 'preferred resolution in pixels per meter
     biYPelsPerMeter As Long 'preferred resolution in pixels per meter
     biClrUsed As Long       'number of colors that are actually used (can be 0)
     biClrImportant As Long  'which color is most important (0 means all of them)
End Type

'palette, 4 bytes * 256 = 1024
Type BITMAPPalette
     lngBlue As Byte
     lngGreen As Byte
     lngRed As Byte
     lngReserved As Byte
End Type


Private Type POINTAPI
   X As Long
   y As Long
End Type

Private Type MSG
   hwnd     As Long        'window where message occured
   Message  As Long        'message id itself
   wParam   As Long        'further defines message
   lParam   As Long        'further defines message
   time     As Long        'time of message event
   pt       As POINTAPI    'position of mouse
End Type

Public Message As MSG         'holds message recieved from queue
Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Const PM_REMOVE = &H1       'paramater on peekmessage to remove or leave message in queue

'myDoEvents() 43% more faster than Vb-DoEvents
Public Function myDoEvents() As Boolean
      If PeekMessage(Message, 0, 0, 0, PM_REMOVE) Then        'checks for a message in the queue and removes it if there is one
         TranslateMessage Message                             'translates the message(dont need if there is no menu)
         DispatchMessage Message                              'dispatches the message to be handled
         myDoEvents = True
      End If
End Function

Function Binary2String(laData() As tBits)
Dim ArrEnd() As Byte
Dim strEnd$, I&
    
    ReDim ArrEnd(0 To UBound(laData()))
    strEnd$ = ""
    For I = 0 To UBound(laData())
        strEnd = strEnd & Chr(Bin2Asc(laData(I)))
    Next I
strEnd = VBA.Left$(strEnd, Len(strEnd) - 1)
 Binary2String = strEnd
End Function

Function AscChar(ByVal MyMot As String) As String
Dim J As Integer
Dim K
Dim TmpMot As String

    For J = 1 To Len(MyMot)
    K = Mid(MyMot, J, 1)
    If Asc(K) >= 20 Then
    TmpMot = TmpMot & K
    End If
    Next J
    If TmpMot = VBA.Space(Len(TmpMot)) Then TmpMot = ""
    AscChar = TmpMot
End Function

'Function String2Binary(theStr As String, RetArray() As tBits, olbStatus As Label, ProgBar As ProgressBar)
'Dim I&, hexRes$, LenBy&
'Dim arrHexStr() As String, arrHexBy() As Byte, BinRes() As tBits
'
'    For I = 1 To Len(theStr)
'        hexRes = hexRes & Asc(Mid(theStr, I, 1)) & ","
'    Next I
'
'    arrHexStr() = Split(hexRes, ",")
'
'    LenBy = UBound(arrHexStr)
'    ReDim arrHexBy(0 To LenBy)
'
'    For I = 0 To LenBy - 1
'        arrHexBy(I) = CByte(arrHexStr(I))
'        ProgBar.Value = I * 100 / LenBy
'    Next I
'
'    Convert2BinaryArray arrHexBy(), BinRes(), olbStatus, ProgBar
'
'RetArray = BinRes()
'End Function

'Function Bin2Hex(Bin As tBits)
'Dim Nibble1$, Nibble2$, I&
'Dim Res$
'    For I = 0 To 3
'        Nibble1 = Nibble1 & Bin.Bits(I)
'    Next I
'    For I = 4 To 7
'        Nibble2 = Nibble2 & Bin.Bits(I)
'    Next I
'    For I = 0 To UBound(vBin)
'        If (Nibble1 = vBin(I)) And Res = "" Then
'            Res = vHex(I)
'            I = 0
'        End If
'        If (Nibble2 = vBin(I)) And Res <> "" Then
'            Res = Res & vHex(I)
'            Exit For
'        End If
'    Next I
'Bin2Hex = Res
'End Function

Function Bin2Asc(Bin As tBits) As Integer
Dim num As Integer
Dim Fact%, I&
    num = 0
    Fact = 128
    For I = 0 To 7
        num = num + Bin.Bits(I) * Fact
        Fact = Fact / 2
    Next I
    Bin2Asc = num
End Function

Function ByteToBinary(ByVal Data As Byte) As tBits
Dim tmpBit As tBits
    Dim I As Long, J&
    
    I = &H80 '10000000
    
    While I
        tmpBit.Bits(J) = IIf(Data And I, "1", "0")
        I = I \ 2
        J = J + 1
    Wend
ByteToBinary = tmpBit
    
End Function
