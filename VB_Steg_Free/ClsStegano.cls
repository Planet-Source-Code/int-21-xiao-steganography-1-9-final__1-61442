VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsStegano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************
'BMP Header Struct
Private BmpHead As winBMPFileHeader
Private BmpInfo As BITMAPINFOHEADER
Private bmpPalette() As BITMAPPalette
Private DeepColor&
'*************************************
'Private mFile2Encode As TypeFileEncode
'*************************************
'BinaryAttach carried the data for each file added in binary format
'BinaryImg() carried tha data for the main image in binary format
Dim BinaryAttach() As tBits, BinaryImg() As tBits
'bAttachdata carried the data for each file added in byte format
'bImgData() carried tha data for the main image in byte format
Dim bImg2Conv() As Byte, bImgData() As Byte, bAttachData() As Byte
'very hard to figure it out that?
'*************************************

Dim mImageFile$ 'Image Filename
Dim mOutputImageFile$ 'New Image Filename

Dim mFilesAdded& 'Count files added
Dim OutFile& 'Pointer to file
Dim mBytesLimit& 'Bytes limit to be added
Dim mBytesAdded& 'Bytes to attach
Dim mImgSize$ 'Image size(width x height)
Dim mImgRes& 'Image resolution 8,16,24 bit
Dim mEncrypMe As Boolean 'Encryp data
Dim mAreLock As Boolean
Dim mBytesExtra& '4 bytes extra when we use encryp over the files, don't ask me why.. just i'd figure it

Dim colFiles As Collection 'My files's collection

Dim cTAG() As Byte 'the main tag to identify if the file carried any file attached

Dim myEncryp As clsEncryp  'Pointer to Encryp class

Event StatusChanged(prcDone As Long, strStatus As String) 'Raise this event to notify what whe are doing
Event SomeError(strDescription As String) 'Raise this event to notify when some error ocurr

' NewEnum tiene que devolver la interfaz IUnknown del
' enumerador de una colección.
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
   Set NewEnum = colFiles.[_NewEnum]
End Function

Public Property Get ImageFile() As String
    ImageFile = mImageFile
End Property

Public Property Let ImageFile(ByVal vNewValue As String)
Dim tmpFil&
    mImageFile = vNewValue
    mBytesLimit = (FileLen(mImageFile) \ 8) - 1024 '1024bytes reserved, to prevent corrupt the image
    tmpFil = FreeFile
    Open mImageFile For Binary As tmpFil
        ReadHeadImg_ tmpFil
        mImgSize = BmpInfo.biWidth & " x " & BmpInfo.biHeight
        mImgRes = BmpInfo.byBitCount
    Close tmpFil
End Property

Public Property Get FilesAdded() As Long
    FilesAdded = mFilesAdded
End Property
'AddFile
'strFile:the filename will be attach
'strTitle the Shortname fot this file, must be the same name with out the extension and the large path
'Key:the unique identify key for this file
Public Function AddFile(strFile As String, strTitle As String, Key As String) As Boolean
Dim tmpFile As ClsFile
On Local Error GoTo AddErr
    Set tmpFile = New ClsFile
    'fill data
    If FileExist(strFile) Then
        With tmpFile
            .KeyFile = Key
            .Filename = strFile
            .FileTitle = strTitle
            .LenBytes = FileLen(strFile) 'get len in bytes
            .TypeFile = VBA.Right$(strFile, 3) 'get type. (.exe,.txt,.bmp...)
            
            mBytesAdded = mBytesAdded + .LenBytes
            If mBytesAdded > mBytesLimit Then 'if the files to attach is too long, can't be carried
                mBytesAdded = mBytesAdded - .LenBytes
                Err.Raise 9001, "AddFile", "The File can't be add. Too long to be attach!"
            End If
        End With
        
    End If
    colFiles.Add tmpFile, Key
    mFilesAdded = mFilesAdded + 1
    AddFile = True
Exit Function
AddErr:
    RaiseEvent SomeError(Err.Description & " in " & Err.Source)
End Function

Public Function RemoveFile(Key As String) As Boolean
On Local Error GoTo AddErr
Dim tmpFile As ClsFile
    Set tmpFile = colFiles(Key) 'remove form the collection the file added
    mBytesAdded = mBytesAdded - tmpFile.LenBytes 'rest the bytes added too
    Set tmpFile = Nothing 'Free memory
    colFiles.Remove Key 'remove item
    RemoveFile = True
    mFilesAdded = mFilesAdded - 1
Exit Function
AddErr:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
End Function

Public Function GetFile(Key As String) As ClsFile
Attribute GetFile.VB_UserMemId = 0
On Local Error GoTo GetErr
    Set GetFile = colFiles(Key) 'return info about any file added
Exit Function
GetErr:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
End Function

Private Sub Class_Initialize()
    Set colFiles = New Collection
    
    Set myEncryp = New clsEncryp
    myEncryp.EncryptionAlgorithm = RC2
    myEncryp.HashAlgorithm = MD5
    myEncryp.Speed = [1KB]
    
    cTAG() = StrConv("TAG:Int21", vbFromUnicode)
End Sub

Public Function Encodeit() As Boolean
Dim strFile$
Dim It As ClsFile
On Local Error GoTo EncodeErr
    
    If FileExist(mImageFile) Then 'Validate filename exist
        
        
        Dim tmpPalette As BITMAPPalette ' To calculate len of struct
        
        If mEncrypMe Then EncrypFiles
        
        'Process data Image
        Call ReadImg_
        'convert image data to binary
        Call Convert2BinaryArray_(bImg2Conv(), BinaryImg())
        
        RaiseEvent StatusChanged(0, "Preparing data to be write...")
        
        OutFile = FreeFile 'The Main Buffer file
        'in this files we going to put all the data, TAG, and each file added
        Open "c:\tmp_C23F41AA.dat" For Binary As #OutFile
            
            Put #OutFile, , cTAG() 'TAG identifer
            Put #OutFile, , mFilesAdded 'count files added
            Put #OutFile, , mBytesAdded 'count bytes added
            Put #OutFile, , CLng(mAreLock) 'was encryp??
            
            RaiseEvent StatusChanged(0, "Please Wait...")
            
            ReadAttach_
            
        Close #OutFile
        
        ConvertAttach_
        Join_Img_Files_
        
        Kill "c:\tmp_C23F41AA.dat" 'delete buffer file
        
        RaiseEvent StatusChanged(100, "Encode done!")
        
    Else
        RaiseEvent SomeError("File doesn't exist") 'Dumb !!
    End If
    
Exit Function
EncodeErr:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
    Close
End Function

Public Function Decodeit() As Boolean
    If Not ReadTag_ Then 'Look for tag
        RaiseEvent SomeError("The selected image no contain any data to extract or haven't a Xiao format")
    Else
        ExtractData_
        Decodeit = True 'return successful
    End If
End Function

Public Sub Save2Image()
Dim strDone$
If mOutputImageFile <> "" Then
    
    RaiseEvent StatusChanged(0, "Saving file...")
    
    If Not SaveImg_() Then strDone = "Some error saving to new image" Else strDone = "Files was saved!"
    RaiseEvent StatusChanged(100, strDone)
    
Else
        RaiseEvent StatusChanged(0, "Image to save not was found!")
End If
    
End Sub

Private Function SaveImg_() As Boolean
Dim I&, J&, xFil&, lngCounter&
Dim maxArr&
On Local Error GoTo SaveImgErr
    'save to new file in disc our image with the file added
    maxArr = UBound(bImg2Conv()) 'get max data image
    
    For J = 0 To UBound(BinaryImg()) 'Len image in binary format, must be equal LenImageInBytes * 8

        bImg2Conv(I) = Bin2Asc(BinaryImg(J)) 'Convert the binary data to byte, 11111111 = 255

        I = I + 1

        If I > maxArr Then
            Exit For
        End If

        If myDoEvents Then RaiseEvent StatusChanged(J * 100 / maxArr, "Saving new image...")

    Next J
    
    
    xFil = FreeFile 'prepare our file to be write
    Open mOutputImageFile For Binary As #xFil
        Put #xFil, , BmpHead 'write header 1st
        Put #xFil, , BmpInfo '2th, write info
        'write the image data with the files hiden
        For lngCounter = 1 To DeepColor 'if exist..write palette data
            Put #1, , bmpPalette(lngCounter)
        Next lngCounter
        
        Put #xFil, , bImg2Conv() 'finally write the new data with our hide data
        Put #xFil, , bImgData() 'put rest of data
        
    Close #xFil 'end of the magic....=)
    SaveImg_ = True
Exit Function
SaveImgErr:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
End Function

Private Function EncrypFiles()
 Dim It As ClsFile
 Dim strFile$
    For Each It In colFiles 'read the files added in the image
        strFile = "C:\" & It.FileTitle & ".enc"
        myEncryp.EncryptToDestFile It.Filename, strFile, 23
        It.Filename = strFile
        'when we encryp, the len file changed, we must update that
        mBytesAdded = mBytesAdded - It.LenBytes
        It.LenBytes = FileLen(strFile)
        mBytesAdded = mBytesAdded + It.LenBytes
    Next
    mBytesExtra = 4
End Function

Private Sub ReadAttach_()
Dim xFil&, I&, lenBy&
Dim It As ClsFile
Dim vData() As Byte, strOut() As Byte
Dim Str3 As String * 3, Str10 As String * 10
Dim strShort$
On Local Error GoTo ReadAttachErr
    
xFil = FreeFile
'Read attach file
RaiseEvent StatusChanged(0, "Reading file to attach...")
I = 0
For Each It In colFiles 'read the files added in the image
    
    Open It.Filename For Binary As #xFil  ' for each file added, build a new temp file in disc
        
        vData = InputB(LOF(xFil), #xFil)
                
        Str3 = It.TypeFile 'txt, bmp, jpg, gif, png
        Str10 = It.FileTitle 'the short name
        
        I = I + 1
        If myDoEvents Then RaiseEvent StatusChanged((I * 100 / mFilesAdded), "Reading file to attach..." & Str10)
        
        strOut() = StrConv(Str3, vbFromUnicode)
        Put #OutFile, , strOut()
        Put #OutFile, , It.LenBytes
        strOut() = StrConv(Str10, vbFromUnicode)
        Put #OutFile, , strOut()
        
        Put #OutFile, , vData()
        
    Close #xFil
    
Next
Exit Sub
ReadAttachErr:
RaiseEvent SomeError(Err.Description)
Err.Clear
End Sub

Private Sub ReadHeadImg_(pFile&)
Dim tmpPalette As BITMAPPalette
Dim I&
    'teh 1st step is read al header for the bitmap, and skip it, to going directly to the image data
    Get #pFile, , BmpHead 'fill head struct
    Get #pFile, , BmpInfo 'fill info struct
    
    'calculate deepcolor
    DeepColor = ((BmpHead.lngBitmapOffset - 54) / Len(tmpPalette))
        
    If DeepColor > 0 Then ReDim bmpPalette(1 To DeepColor) 'Rezise
    
    For I = 1 To DeepColor
        Get #pFile, , bmpPalette(I)
     Next I
    
End Sub

Private Sub ReadImg_()
Dim xFil&, Bytes2Hide&, RestBytes&, lngCounter&
Dim bytColor As Byte

'Read the Img File
xFil = FreeFile
Open mImageFile For Binary As #xFil

    RaiseEvent StatusChanged(0, "Reading Header...")
    
    ReadHeadImg_ xFil
    
    'Calculate len image data, without headers
    'Only read the len of bytes we going to hide
    
    'calculate the len data must be read
    '17= len of main header,TAG:Int21(9bytes)+filesadded(4bytes)+bytesadded(4bytes)
    '17=len of file header, type(3bytes)+filelen(4bytes)+filename(10byte)
    Bytes2Hide = (mBytesAdded + 17 + (17 * mFilesAdded) + mBytesExtra) * 8 'LOF(ImgFile) - Loc(ImgFile)
    ReDim bImg2Conv(0 To Bytes2Hide) 'NEW- 1 bytes to hide is equal to 8 bytes of data
    RestBytes = (BmpHead.lngFileSize - BmpHead.lngBitmapOffset) - Bytes2Hide 'New
    ReDim bImgData(0 To RestBytes)
    
    RaiseEvent StatusChanged(0, "Reading Image Data...")
    
    For lngCounter = 0 To Bytes2Hide ' this is the data where we going to hide our files
        If Not EOF(xFil) Then
            Get #xFil, , bytColor 'Read each rgb byte info
            bImg2Conv(lngCounter) = bytColor
        End If
        
        If myDoEvents Then RaiseEvent StatusChanged(lngCounter * 100 / Bytes2Hide, "Reading Image Data...")
                
    Next lngCounter
        
    For lngCounter = 0 To RestBytes ' this is the rest data
        If Not EOF(xFil) Then
            Get #xFil, , bytColor 'Read each rgb byte info
            bImgData(lngCounter) = bytColor
        End If
        
        If myDoEvents Then RaiseEvent StatusChanged(lngCounter * 100 / RestBytes, "Reading Image Data...")
                
    Next lngCounter
    
Close #xFil
    
End Sub
'Look for our tag in the image file, if doesn't exist skip all
Private Function ReadTag_() As Boolean
Dim binData() As tBits, binTag() As tBits
Dim I&, J&, Cur&, bytColor As Byte, Ret&
Dim strMyTag As String * 9
Dim lenStruct&, xFil&
Dim bBytes() As Byte
    
    RaiseEvent StatusChanged(0, "Searching header...")
    
    lenStruct = 17 'the len for the tag is always 17 bytes
    ReDim binTag(0 To lenStruct)
    
    xFil = FreeFile
    Open mImageFile For Binary As #xFil
        
        ReadHeadImg_ xFil  'Read header for bitmap
        
        lenStruct = 168 '8 bytes = 1 extra-byte, TAG= 21 bytes * 8 bytes = 168 bytes
        
        ReDim bImgData(0 To lenStruct)
        
        For I = 0 To lenStruct
            If Not EOF(xFil) Then
                Get #xFil, , bytColor
                bImgData(I) = bytColor
            End If
        Next I
    
    Close #xFil
    
    Call Convert2BinaryArray_(bImgData(), binData())
    
    lenStruct = UBound(binTag()) 'len data in binary
    
    Cur = 0
    lenStruct = 9 'the len tag is alway 9bytes
    For I = 0 To lenStruct
        For J = 0 To 7
            binTag(I).Bits(J) = binData(Cur).Bits(7)
            Cur = Cur + 1
            
        Next J
        If Cur >= 72 Then Exit For
    Next I
    
    strMyTag = Binary2String(binTag)
    If strMyTag = "TAG:Int21" Then
        ReDim binTag(0 To 4)
        ReDim bBytes(0 To 4)
        Cur = 72
        For I = 0 To 4
            For J = 0 To 7
                binTag(I).Bits(J) = binData(Cur).Bits(7)
                Cur = Cur + 1
            
            Next J
            If Cur >= 104 Then Exit For
            bBytes(I) = Bin2Asc(binTag(I))
        Next I
        CopyMemory mFilesAdded, bBytes(0), Len(mFilesAdded)
        
        Cur = 104
        For I = 0 To 4
            For J = 0 To 7
                binTag(I).Bits(J) = binData(Cur).Bits(7)
                Cur = Cur + 1
            
            Next J
            If Cur >= 136 Then Exit For
            bBytes(I) = Bin2Asc(binTag(I))
        Next I
        CopyMemory mBytesAdded, bBytes(0), Len(mBytesAdded)
        
        
        Cur = 136
        For I = 0 To 4
            For J = 0 To 7
                binTag(I).Bits(J) = binData(Cur).Bits(7)
                Cur = Cur + 1
            
            Next J
            If Cur >= 168 Then Exit For
            bBytes(I) = Bin2Asc(binTag(I))
        Next I
        CopyMemory Ret, bBytes(0), Len(Ret)
        mAreLock = Ret
        If mAreLock Then mBytesExtra = 4
        ReadTag_ = True
    End If
    
    'ReadTag_ = (strMyTag = "TAG:Int21")
    
End Function

Private Sub ExtractData_()
Dim OutFile&, ImgFile&
Dim tmpFile&
Dim dataOut() As Byte
Dim BinOut() As tBits
Dim Bytes2Read&, Cur&, I&, J&
Dim bytColor As Byte
Dim sTAg$, lFA&, lBA&, sTF$, lLF&, sNF$

    ImgFile& = FreeFile
    
    Open mImageFile$ For Binary As #ImgFile 'open the main image
        'skip the bmp header, to get the real image data
        Call ReadHeadImg_(ImgFile)
        'calculate the len data must be read
        '17= len of main header,TAG:Int21(9bytes)+filesadded(4bytes)+bytesadded(4bytes)
        '17=len of file header, type(3bytes)+filelen(4bytes)+filename(10byte)
        Bytes2Read = (mBytesAdded + 17 + (17 * mFilesAdded) + mBytesExtra) * 8 'LOF(ImgFile) - Loc(ImgFile)

        ReDim dataOut(0 To Bytes2Read)
        For I = 0 To (Bytes2Read)
            If Not EOF(ImgFile) Then
                Get #ImgFile, , bytColor
                dataOut(I) = bytColor
            End If
        Next I
    Close #ImgFile

        Call Convert2BinaryArray_(dataOut(), BinaryImg())
        
        Bytes2Read = UBound(BinaryImg()) 'len image in binary
        mBytesAdded = (mBytesAdded + 17 + (17 * mFilesAdded) + mBytesExtra)
        ReDim dataOut(0 To mBytesAdded)
        ReDim BinOut(0 To Bytes2Read)
        
        
        Cur = 0
        'we going to read the bytes 7 for each byte in the image data
        'and put it in other array to extract the hide data
        For I = 0 To mBytesAdded
            For J = 0 To 7
                If Cur >= Bytes2Read Then Exit For
                BinOut(I).Bits(J) = BinaryImg(Cur).Bits(7)
                Cur = Cur + 1
                
            Next J
            dataOut(I) = Bin2Asc(BinOut(I)) 'convert the binary hide in bytes
            myDoEvents
        Next I
        
        OutFile = FreeFile
        Open "c:\tmp_DD2741C.dat" For Binary As #OutFile 'tmp file to read data
            Put #OutFile, , dataOut()
        Close OutFile
        
        OutFile = FreeFile
        
        mBytesAdded = 0
        mFilesAdded = 0
        
        Open "c:\tmp_DD2741C.dat" For Binary As #OutFile 'tmp file to read data
            
            sTAg = ExtractItem_(OutFile, 9, 0, 1) 'Read the main tag
            lFA = ExtractItem_(OutFile, 4, 0, 0) 'read the number of files added
            lBA = ExtractItem_(OutFile, 4, 0, 0) 'read the len of bytes added
            mAreLock = ExtractItem_(OutFile, 4, 0, 0) 'was encryp??
            Dim strFile$
            Dim It As ClsFile
            For I = 1 To lFA
                sTF = ExtractItem_(OutFile, 3, 0, 1) 'Read the type file(txt,bmp,gif,jpg,png)
                lLF = ExtractItem_(OutFile, 4, 0, 0) 'read the len in bytes for this file
                sNF = ExtractItem_(OutFile, 10, 0, 1) 'read the short name for this file
                
                strFile = "c:\" & sNF & "DD2741C." & sTF 'build the buffer filename
                
                tmpFile = FreeFile
                Open strFile For Binary As tmpFile
                    dataOut() = InputB(lLF, OutFile) 'read n-bytes, the len for this file
                    Put tmpFile, , dataOut() 'write in disc
                Close tmpFile
                
                
                AddFile strFile, sNF, CStr("c0" & I) 'add in the class
                'mBytesAdded = mBytesAdded + lLF 'counter the bytes added in the image
                
            Next
        
        Close OutFile
        
        Kill "c:\tmp_DD2741C.dat"
        
    
End Sub

Private Function ExtractItem_(pFile As Long, Bytes2Read As Long, Bytes2Look As Long, RetType As Integer)
Dim Memo() As Byte
Dim lLong&
Dim strEnd$
    Memo() = InputB(Bytes2Read, pFile) 'read n-bytes from disc
    
If RetType = 0 Then 'Numeric
    CopyMemory lLong, Memo(0), Len(lLong)
    ExtractItem_ = lLong
ElseIf RetType = 1 Then 'String
    strEnd = Memo()
    ExtractItem_ = StrConv(strEnd, vbUnicode)
End If

End Function

Private Sub ConvertAttach_()
Dim byt As Byte
Dim LenF&, I&
On Local Error GoTo ErrConvert
'Read all files added and convert to binary
    OutFile = FreeFile
    
    Open "c:\tmp_C23F41AA.dat" For Binary As #OutFile
        LenF = LOF(OutFile) - 1
        ReDim bAttachData(0 To LenF)
        For I = 0 To LenF
            If Not EOF(OutFile) Then
                Get OutFile, , byt
                bAttachData(I) = byt
            End If
        Next
    Close #OutFile
        
    Call Convert2BinaryArray_(bAttachData(), BinaryAttach())
    
Exit Sub
ErrConvert:
    RaiseEvent SomeError(Err.Description)
    Err.Clear
End Sub
'the magic function, joing image and files to attach in only one file
Private Sub Join_Img_Files_()
Dim I&, J&, K&, LenImg&, LenF&
    
    LenImg = UBound(BinaryImg()) 'len in binary of image
    LenF = UBound(BinaryAttach()) 'len in binary for files to attach
    I = 0
    
    For J = 0 To LenF
        For K = 0 To 7
            BinaryImg(I).Bits(7) = BinaryAttach(J).Bits(K) 'put one bit from binary data to hide in the bit 7
            I = I + 1
        Next K
        If I >= LenImg Then Exit For
        
        If myDoEvents Then RaiseEvent StatusChanged((I * 100 / LenImg), "Joining files with image...")
        
    Next J
    
End Sub
'Convert2BinaryArray_
'Source(): the file data in bytes
'retArray(): the Binary data to be return
Private Sub Convert2BinaryArray_(Source() As Byte, RetArray() As tBits)
Dim LenArray&, I&
Dim arrBinary() As tBits
Dim Bits8 As tBits

    LenArray = UBound(Source())
    
     ReDim arrBinary(0 To LenArray)

    For I = 0 To LenArray
    
        Bits8 = ByteToBinary(Source(I)) 'convert 1 byte to binary
        arrBinary(I) = Bits8
        
        If myDoEvents Then RaiseEvent StatusChanged((I * 100 / LenArray), "Convert Hex to Binary...")
                
     Next I
     
RetArray = arrBinary

End Sub

Private Function FileExist(strFile As String) As Boolean
Dim Rs$, Tama As Boolean
Dim tm&
    Rs = Dir(strFile)
    FileExist = (Len(Rs))
End Function

Public Property Get OutputImageFile() As String
    OutputImageFile = mOutputImageFile
End Property

Public Property Let OutputImageFile(ByVal sNewFile As String)
    mOutputImageFile = sNewFile
End Property

Private Sub Class_Terminate()
Dim tmpClass As ClsFile
For Each tmpClass In colFiles
    Set tmpClass = Nothing
Next
Set colFiles = Nothing

If myEncryp.IsHashReady Then myEncryp.DestroyHash
Set myEncryp = Nothing

'free memory
Erase BinaryAttach()
Erase BinaryImg()
Erase bImgData()
Erase bImg2Conv()
Erase bAttachData()
End Sub

Public Property Get BytesAdded() As Long
    BytesAdded = mBytesAdded
End Property


Public Property Get BytesTotal() As Variant
    BytesTotal = mBytesLimit
End Property

'Public Property Get File2Encode() As TypeFileEncode
'    File2Encode = mFile2Encode
'End Property
'
'Public Property Let File2Encode(ByVal tNewType As TypeFileEncode)
'    mFile2Encode = tNewType
'End Property

Public Property Get ImgSize() As String
    ImgSize = mImgSize
End Property

Public Property Get ImgRes() As Long
    ImgRes = mImgRes
End Property

Public Property Get EncrypMe() As Boolean
    EncrypMe = mEncrypMe
End Property

Public Property Let EncrypMe(ByVal bNewEncryp As Boolean)
    mEncrypMe = bNewEncryp
    mAreLock = mEncrypMe
End Property

Public Property Get EncryptionAlgorithm() As EC_CRYPT_ALGO_ID
    EncryptionAlgorithm = myEncryp.EncryptionAlgorithm
End Property

Public Property Let EncryptionAlgorithm(ByVal ecEncryptID As EC_CRYPT_ALGO_ID)
    myEncryp.EncryptionAlgorithm = ecEncryptID
End Property

Public Property Get HashAlgorithm() As EC_HASH_ALG_ID
    HashAlgorithm = myEncryp.HashAlgorithm
End Property

Public Property Let HashAlgorithm(ByVal hAlgoId As EC_HASH_ALG_ID)
    myEncryp.HashAlgorithm = hAlgoId
End Property

Public Property Get Pwd() As String
    Pwd = myEncryp.Password
End Property

Public Property Let Pwd(ByVal sPassword As String)
    myEncryp.Password = sPassword
End Property

Public Property Get areLock() As Boolean
    areLock = mAreLock
End Property

Public Property Let areLock(ByVal bNewLock As Boolean)
    mAreLock = bNewLock
End Property


Public Function UnLockMe(strLockFile As String, DestFile As String)
    If mAreLock Then myEncryp.DecryptToDestFile strLockFile, DestFile, 23
End Function
