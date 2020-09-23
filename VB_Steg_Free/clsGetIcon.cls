VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGetIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSID
    id((123)) As Byte
End Type
Private Const MAX_PATH = 260
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1

Public Function SetIcon(ByVal aFileName As String) As IPictureDisp
Dim fName As String
Dim Index As Integer
Dim hIcon As Long
Dim item_num As Long
Dim icon_pic As IPictureDisp
Dim sh_info As SHFILEINFO
Dim cls_id As CLSID
Dim hRes As Long
Dim new_icon As TypeIcon
Dim lpUnk As IUnknown

fName = Trim(aFileName)
If fName = "" Then
    Exit Function
    End If
SHGetFileInfo fName, 0, sh_info, Len(sh_info), SHGFI_ICON + SHGFI_SMALLICON    'SHGFI_LARGEICON
hIcon = sh_info.hIcon
With new_icon
    .cbSize = Len(new_icon)
    .picType = vbPicTypeIcon
    .hIcon = hIcon
End With
With cls_id
    .id(8) = &HC0
    .id(15) = &H46
End With
hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
If hRes = 0 Then Set icon_pic = lpUnk
Set SetIcon = icon_pic 'return image
End Function


