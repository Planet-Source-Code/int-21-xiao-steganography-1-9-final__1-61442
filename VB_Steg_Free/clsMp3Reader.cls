VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMp3Reader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MPEG
    Filename As String
    FileLen As Long
    BitRate As Long
    Frequency As Long
    Version As Integer
    Layer As Integer
    Header As Long
    Mode As String
    Emphasis As String
    Original As String
    Copyrighted As String
    Private As String
    CRCs As String
    Seconds As Long
    FrameSize As Integer
    Frames As Long
    Padding As String
    BitString As String
End Type

Private Type ID3v1Data           'This type is standard for ID3v1 tags
  Title       As String * 30    '30 bytes Title
  Artist      As String * 30    '30 bytes Artist
  Album       As String * 30    '30 bytes Album
  Year        As String * 4     '4 bytes Year
  Comments    As String * 28    '28 bytes Comments
  IsTrack     As Byte           '1 byte Istrack / +1 byte comments
  Tracknumber As Byte           '1 byte Tracknumber / +1 byte comments
  Genre       As Byte           '1 byte Genre
End Type


Dim byteArray(4) As Byte
Dim strFilename As String
Dim mLenHeader&
Private MP3 As MPEG
Private ID3v1Info As ID3v1Data     'Declare a variable as the ID3v1Data type

Public Property Get Artista() As String
    Artista = AscChar(ID3v1Info.Artist)
End Property

Public Property Get Tema() As String
    Tema = AscChar(ID3v1Info.Title)
End Property

Public Property Get BitRate() As Integer
    BitRate = Int(MP3.BitRate)
End Property
Public Property Get BitStr() As String
    BitStr = MP3.BitString
End Property

Public Property Get ArchivoMp3() As String
    ArchivoMp3 = MP3.Filename
End Property

Public Property Let ArchivoMp3(ByVal Data As String)
    If FileExists(Data) Then
        strFilename = Data
        GetMPEGInfo
        GetID3 strFilename
    Else
        strFilename = ""
    End If
End Property

Public Property Get TamaMp3() As String
    TamaMp3 = MP3.FileLen
End Property
Public Property Get Frames() As Integer
    Frames = Int(MP3.Frames)
End Property
Public Property Get Frecuencia() As Long
    Frecuencia = MP3.Frequency
End Property
Public Property Get Header() As String
    Header = MP3.Header
End Property
Public Property Get Mode() As String
    Mode = MP3.Mode
End Property
Public Property Get Version() As String
    Version = MP3.Version
End Property
Public Property Get VersionLayer() As String
    If MP3.Version <> 25 Then VersionLayer = "MPEG " & MP3.Version & ".0 Layer " & MP3.Layer
    If MP3.Version = 25 Then VersionLayer = "MPEG Version 2.5 Layer " & MP3.Layer
End Property
Public Property Get Layer() As String
    Layer = MP3.Layer
End Property
Public Property Get Segun2() As Integer
    Segun2 = Round(MP3.Seconds)
End Property

Private Function GetID3(Filename As String) As Boolean
  Dim TaG As String * 3   'We use this variable to make sure the file has an ID3v1 tag

  GetID3 = True
  Open Filename For Binary As #1
  Get #1, FileLen(Filename) - 127, TaG            'Looks after tag 128 bytes from the end of the file
  If TaG = "TAG" Then                             '"TAG" is put at position filesize-127 to show that this file indeed contains an ID3v1 tag
    Get #1, FileLen(Filename) - 124, ID3v1Info    'if the file has a tag, we put it into our earlier declared variable ID3v1Info
  Else
    ID3v1Info.Title = ""                          'if the "TAG" wasnt at position filesize-127
    ID3v1Info.Artist = ""
    ID3v1Info.Album = ""
    ID3v1Info.Year = ""
    ID3v1Info.Comments = ""
    ID3v1Info.IsTrack = 0
    ID3v1Info.Tracknumber = 0
    ID3v1Info.Genre = 255
    GetID3 = False                                'Return False as there was no ID3v1 tag
  End If
  Close #1                                        'close the file

End Function

Private Sub GetMPEGInfo()
On Error Resume Next
Dim Header, TempVar As Byte, BitString As String, I As Long
Dim Version, Layer, Mode, Emphasis, Frequency, BitRate
Dim YesNo, NoYes
Dim File As Integer
    YesNo = Array("Yes", "No")
    NoYes = Array("No", "Yes")
    MP3.Filename = strFilename
    File = FreeFile
    Open strFilename For Binary Access Read As File
        MP3.FileLen = FileLen(MP3.Filename)
        For I = 1 To MP3.FileLen
        Get File, I, TempVar
        If TempVar = 255 Then
            Get File, I + 1, TempVar
            If TempVar > 242 And TempVar < 252 Then
            Header = I
            MP3.Header = Header
            Exit For
            End If
        End If
        Next I
        
        mLenHeader = I
        
        For I = 1 To 4
        Get File, (Header + (I - 1)), byteArray(I)
        Next I
    Close File
    
    mLenHeader = mLenHeader + I - 1
    
    BitString = ByteToBit(byteArray)
    MP3.BitString = BitString
    'Check for the version(either version 2.0 or 1.0)
    Version = Array(2, 1)
    MP3.Version = Version(BinToDec(Mid(BitString, 13, 1)))
    'Check for the layer of the mp3(either 1, 2, 3)
    Layer = Array(0, 3, 2, 1)
    MP3.Layer = Layer(BinToDec(Mid(BitString, 14, 2)))
    'Check for which mode the mp3 is(either Stereo, Joint Stereo, 2 Channel, or 1 Channel)
    Mode = Array("Stereo", "Joint Stereo", "Dual Channel", "Mono")
    MP3.Mode = Mode(BinToDec(Mid(BitString, 25, 2)))
    'Check for the emphasis(I have no idea what the emphasis does)
    Emphasis = Array("None", "50/15", "Reserved", "CCITT J 17")
    MP3.Emphasis = Emphasis(BinToDec(Mid(BitString, 31, 2)))
    'there are different frequency tables for the different versions
    Select Case MP3.Version
        Case 1
            Frequency = Array(44100, 48000, 32000)
        Case 2
            Frequency = Array(22050, 24000, 16000)
        Case Else
            Frequency = Array(0)
    End Select
    'Set the Frequency
    MP3.Frequency = Frequency(BinToDec(Mid(BitString, 21, 2)))
    Select Case MP3.Version
        Case 1
            If MP3.Layer = 1 Then BitRate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
            If MP3.Layer = 2 Then BitRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
            If MP3.Layer = 3 Then BitRate = Array(0, 32, 40, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320)
        Case 2
            If MP3.Layer = 1 Then BitRate = Array(0, 32, 64, 96, 128, 160, 192, 224, 256, 288, 320, 352, 384, 416, 448)
            If MP3.Layer = 2 Then BitRate = Array(0, 32, 48, 56, 64, 80, 96, 112, 128, 160, 192, 224, 256, 320, 384)
            If MP3.Layer = 3 Then BitRate = Array(0, 8, 16, 24, 32, 64, 80, 56, 64, 128, 160, 112, 128, 256, 320)
        Case Else
            BitRate = Array(0)
    End Select
    'Set the BitRate
    MP3.BitRate = BitRate(BinToDec(Mid(BitString, 17, 4)))
    'Check if original, copyrighted, private, padding, and CRCs = yes or no
    MP3.Original = NoYes(Mid(BitString, 30, 1))
    MP3.Copyrighted = NoYes(Mid(BitString, 29, 1))
    MP3.Private = NoYes(Mid(BitString, 24, 1))
    'i am not so sure if the padding thing here is right
    MP3.Padding = NoYes(Mid(BitString, 23, 1))
    MP3.CRCs = YesNo(Mid(BitString, 16, 1))
    'this may not work, because i am not sure if the padding thing that i set up is right
    MP3.FrameSize = (MP3.BitRate * 144000) / (MP3.Frequency)
    If MP3.Padding = "Yes" Then MP3.FrameSize = MP3.FrameSize + 1
    'i am not sure if this is right, but i multiplied by two so it would be the same amount of frames that winamp came up with
    MP3.Frames = Int(MP3.FileLen / MP3.FrameSize)
    MP3.Seconds = (MP3.FileLen * 8) / (MP3.BitRate * 1000)
End Sub

'This function converts 4*1 byte array to 4*8 bits
Private Function ByteToBit(byteArray) As String
Dim z As Integer, I As Integer
  ByteToBit = ""
  For z = 1 To 4
    For I = 7 To 0 Step -1
      If Int(byteArray(z) / (2 ^ I)) = 1 Then
        ByteToBit = ByteToBit & "1"
        byteArray(z) = byteArray(z) - (2 ^ I)
      Else
        If ByteToBit <> "" Then
          ByteToBit = ByteToBit & "0"
        End If
      End If
    Next I
  Next z
End Function
'This function converts Binary string to decimal integer
Private Function BinToDec(BinValue As String) As Long
Dim I As Integer
  BinToDec = 0
  For I = 1 To Len(BinValue)
    If Mid(BinValue, I, 1) = 1 Then
      BinToDec = BinToDec + 2 ^ (Len(BinValue) - I)
    End If
  Next I
End Function

Private Function FileExists(Filename As String) As Boolean
  'Cheks if file exists
  FileExists = Not (Dir(Filename) = "")
End Function

Public Property Get HeaderLen() As Variant
    HeaderLen = mLenHeader
End Property

