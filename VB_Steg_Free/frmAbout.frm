VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   683
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "<<Back"
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbMe 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label lbTips 
      Caption         =   $"frmAbout.frx":0000
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTips$, sAboutme$

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
sAboutme = "Developed by Int21" & vbCrLf
sAboutme = sAboutme & "(c)2005" & vbCrLf
sAboutme = sAboutme & "Software from Venezuela" & vbCrLf
sAboutme = sAboutme & "webSite: http://www26.brinkster.com/blackc"
lbMe = sAboutme

sTips = "Steganography is the art of covered or hidden writing. The purpose of steganography is covert communication to hide a message from a third party. This differs from cryptography, the art of secret writing, which is intended to make a message unreadable by a third party but does not hide the existence of the secret communication. Although steganography is separate and distinct from cryptography, there are many analogies between the two, and some authors categorize steganography as a form of cryptography since hidden communication is a form of secret writing (Bauer 2002)." & vbCrLf
sTips = sTips & "The most common steganography method in audio and image files employs some type of least significant bit substitution or overwriting. The least significant bit term comes from the numeric significance of the bits in a byte. The high-order or most significant bit is the one with the highest arithmetic value (i.e., 27=128), whereas the low-order or least significant bit is the one with the lowest arithmetic value (i.e., 20=1)." & vbCrLf
sTips = sTips & "As a simple example of least significant bit substitution, imagine ""hiding"" the character 'G' across the following eight bytes of a carrier file (the least significant bits are underlined):" & vbCrLf & vbCrLf
sTips = sTips & "10010101  00001101  11001001  10010110" & vbCrLf
sTips = sTips & "00001111  11001011  10011111  00010000" & vbCrLf & vbCrLf
sTips = sTips & "A 'G' is represented in the American Standard Code for Information Interchange (ASCII) as the binary string 01000111. These eight bits can be ""written"" to the least significant bit of each of the eight carrier bytes as follows:" & vbCrLf & vbCrLf
sTips = sTips & "10010100 00001101 11001000 10010110" & vbCrLf
sTips = sTips & "00001110 11001011 10011111 00010001" & vbCrLf & vbCrLf
sTips = sTips & "In the sample above, only half of the least significant bits were actually changed (shown above in italics). This makes some sense when one set of zeros and ones are being substituted with another set of zeros and ones." & vbCrLf & vbCrLf
sTips = sTips & "for more information, go to http://www.fbi.gov/hq/lab/fsc/backissu/july2004/research/2004_03_research01.htm"

lbTips = sTips

End Sub
