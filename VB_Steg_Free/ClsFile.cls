VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mbFileName() As Byte 'The filename
Dim mbFileTitle() As Byte 'The short title for the filename
Dim mbType() As Byte 'what type is: txt, bmp, jpg, gif, png
Dim msKey As String * 3 'the unique key
Dim mlLenFile& 'len in bytes for this file

Public Property Get KeyFile() As String
    KeyFile = msKey
End Property

Public Property Let KeyFile(ByVal sNewKey As String)
    msKey = sNewKey
End Property

Public Property Get TypeFile() As String
    TypeFile = StrConv(mbType(), vbUnicode)
End Property

Public Property Let TypeFile(ByVal sNewType As String)
    mbType() = StrConv(sNewType, vbFromUnicode)
End Property

Public Property Get LenBytes() As Long
    LenBytes = mlLenFile
End Property

Public Property Let LenBytes(ByVal lNewLen As Long)
    mlLenFile = lNewLen
End Property

Public Property Get FileName() As String
    FileName = StrConv(mbFileName(), vbUnicode)
End Property

Public Property Let FileName(ByVal sNewFile As String)
    mbFileName() = StrConv(sNewFile, vbFromUnicode)
End Property

Public Property Get FileTitle() As String
    FileTitle = StrConv(mbFileTitle(), vbUnicode)
End Property

Public Property Let FileTitle(ByVal sNewTitle As String)
    mbFileTitle() = StrConv(sNewTitle, vbFromUnicode)
End Property
