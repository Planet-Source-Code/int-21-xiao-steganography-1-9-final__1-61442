VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRead 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Read File from a Image"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRead.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   4560
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRead.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRead.frx":0B24
            Key             =   "imgAttach"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Files 
      Left            =   12120
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frStep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   6735
      Begin VB.PictureBox PicCont 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         ScaleHeight     =   205
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   4
         Top             =   240
         Width           =   3615
         Begin VB.HScrollBar HS 
            Height          =   255
            Left            =   0
            Max             =   1
            Min             =   1
            TabIndex        =   5
            Top             =   2820
            Value           =   1
            Width           =   3300
         End
         Begin VB.VScrollBar VS 
            Height          =   3075
            Left            =   3300
            Max             =   1
            Min             =   1
            TabIndex        =   6
            Top             =   0
            Value           =   1
            Width           =   255
         End
         Begin VB.PictureBox TheImage 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2775
            Left            =   0
            ScaleHeight     =   183
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   215
            TabIndex        =   7
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Label cmdTarget 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Load Image Target"
         Height          =   225
         Left            =   4080
         MouseIcon       =   "frmRead.frx":10BE
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   2850
         Width           =   2115
      End
      Begin VB.Label lbImgName 
         Caption         =   "Name:"
         Height          =   195
         Left            =   3840
         TabIndex        =   21
         Top             =   360
         Width           =   2730
      End
      Begin VB.Label lbImgSize 
         Caption         =   "Size:"
         Height          =   195
         Left            =   3840
         TabIndex        =   20
         Top             =   600
         Width           =   2730
      End
      Begin VB.Label lbImgRes 
         Caption         =   "Resolution:"
         Height          =   195
         Left            =   3840
         TabIndex        =   19
         Top             =   840
         Width           =   2730
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   5880
         Picture         =   "frmRead.frx":1210
         Top             =   2835
         Width           =   240
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   405
         Left            =   3960
         Top             =   2760
         Width           =   2295
      End
   End
   Begin VB.Frame frStep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Label lbWait 
         AutoSize        =   -1  'True
         Caption         =   "Wait while the files are read from the image..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   4980
      End
   End
   Begin VB.Frame frStep 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Frame frLock 
         Height          =   615
         Left            =   1920
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtLock 
            Height          =   285
            Left            =   3000
            TabIndex        =   18
            Top             =   200
            Width           =   1575
         End
         Begin VB.Label lbLock 
            AutoSize        =   -1  'True
            Caption         =   "Put the password to unlock files:"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   2790
         End
      End
      Begin MSComctlLib.ListView lvwFilesAdded 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImgLst"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Type"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size (K)"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Shape ShpExtract 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   120
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label cmdExtract 
         AutoSize        =   -1  'True
         Caption         =   "Extract File"
         Height          =   195
         Left            =   600
         MouseIcon       =   "frmRead.frx":1614
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   2955
         Width           =   1185
      End
      Begin VB.Image ImgExtract 
         Height          =   240
         Left            =   240
         Picture         =   "frmRead.frx":1766
         Top             =   2925
         Width           =   240
      End
      Begin VB.Label lbRema 
         AutoSize        =   -1  'True
         Caption         =   "Bytes in Image:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.Image ImgBack 
      Height          =   165
      Left            =   240
      Picture         =   "frmRead.frx":1CF0
      Top             =   4170
      Width           =   150
   End
   Begin VB.Label cmdBack 
      AutoSize        =   -1  'True
      Caption         =   "Back"
      Enabled         =   0   'False
      Height          =   195
      Left            =   480
      MouseIcon       =   "frmRead.frx":1D45
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4155
      Width           =   660
   End
   Begin VB.Shape ShpBack 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   120
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image ImgNext 
      Height          =   240
      Left            =   1440
      Picture         =   "frmRead.frx":1E97
      Top             =   4155
      Width           =   240
   End
   Begin VB.Label cmdNext 
      AutoSize        =   -1  'True
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   195
      Left            =   1680
      MouseIcon       =   "frmRead.frx":2421
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   4155
      Width           =   675
   End
   Begin VB.Shape ShpNext 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   1320
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Image ImgDone 
      Height          =   240
      Left            =   2640
      Picture         =   "frmRead.frx":2573
      Top             =   4125
      Width           =   240
   End
   Begin VB.Label cmdFinish 
      AutoSize        =   -1  'True
      Caption         =   "Finish"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3000
      MouseIcon       =   "frmRead.frx":2932
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4155
      Width           =   480
   End
   Begin VB.Shape ShpDone 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   2520
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lbStep 
      AutoSize        =   -1  'True
      Caption         =   "Step 1: Select the source image:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3525
   End
End
Attribute VB_Name = "frmRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strImageFile$ 'the main image where the files was attached
Dim m_Millimeter&
Dim vSteps
Const MAX_STEPS = 2
Dim currentStep&
Dim imgIco%
Dim WithEvents clsStegaRead As ClsStegano 'the magic class
Attribute clsStegaRead.VB_VarHelpID = -1

Private Sub clsStegaRead_SomeError(strDescription As String)
    MsgBox strDescription, , "Xiao 1.9"
End Sub

Private Sub clsStegaRead_StatusChanged(prcDone As Long, strStatus As String)
    frmLoad.ProgBar.Value = prcDone
    frmLoad.lbStatus = strStatus
End Sub

Private Sub cmdExtract_Click()
Dim ItSel As ListItem
Dim OutFile&, tmpFile&
Dim dataOut() As Byte
On Error GoTo errExtrac
    
    Set ItSel = lvwFilesAdded.SelectedItem
    
    If Not ItSel Is Nothing Then
        Files.Filename = ""
        Files.Filter = "File|*." & ItSel.Text
        Files.ShowSave
        If Files.Filename <> "" Then
            Dim ItFile As ClsFile
            tmpFile = FreeFile
            Set ItFile = clsStegaRead.GetFile(ItSel.Key)
            If Len(Trim(txtLock)) > 0 Then
                clsStegaRead.Pwd = Trim(txtLock)
                clsStegaRead.UnLockMe ItFile.Filename, Files.Filename
            Else
                Open ItFile.Filename For Binary As tmpFile
                    dataOut() = InputB(LOF(tmpFile), tmpFile)
                Close tmpFile
                
                OutFile = FreeFile
                Open Files.Filename For Binary As OutFile
                    Put OutFile, , dataOut()
                Close OutFile
            
            End If
            
            
            MsgBox "File extract was successful!", , "Xiao 1.9"
            
        End If
        
    End If
    
Exit Sub

errExtrac:
MsgBox Err.Description, , "Xiao 1.9"
End Sub

Private Sub cmdFinish_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    cmdNext.Enabled = False
    cmdBack.Enabled = True
    frStep(currentStep).Visible = False
    currentStep = currentStep + 1
    lbStep = vSteps(currentStep)
    frStep(currentStep).Visible = True
    If currentStep >= MAX_STEPS Then
        cmdBack.Enabled = False
        cmdNext.Enabled = False
        cmdFinish.Enabled = True
    End If
    
    If currentStep = 2 Then 'Extract Data
        If clsStegaRead.areLock Then
            frLock.Visible = True
            imgIco = 1
        End If
    End If

    If currentStep = 1 Then 'Search for tag
        frmLoad.Show vbModeless, Me
        If Not clsStegaRead.Decodeit Then
            MsgBox "The selected image no contain any data to extract or haven't a Xiao format", , "Xiao 1.9"
            currentStep = MAX_STEPS - 1
        Else
            imgIco = 2
            cmdNext_Click
            Unload frmLoad
            
            Dim ItTmp As ClsFile, Itlvw As ListItem
            Dim I&
            I = 1
            For Each ItTmp In clsStegaRead
                Set Itlvw = lvwFilesAdded.ListItems.Add(, ItTmp.KeyFile, ItTmp.TypeFile, , imgIco)
                Itlvw.SubItems(1) = ItTmp.FileTitle
                Itlvw.SubItems(2) = ItTmp.LenBytes
                I = I + 1
            Next
            lbRema = "Bytes attached in Image: " & clsStegaRead.BytesAdded
        End If

    End If
    

End Sub

Private Sub cmdTarget_Click()
Dim ToW&, ToH&, InW&, InH&, mPad&
    mPad = 2


    Files.Filter = "Image File|*.bmp"
    Files.ShowOpen
    strImageFile = Files.Filename
    If strImageFile <> "" And VBA.Right$(strImageFile, 4) = ".bmp" Then
        TheImage.Picture = LoadPicture(strImageFile)
        
        clsStegaRead.ImageFile = strImageFile
        
        lbImgName = "Name: " & Files.FileTitle
        lbImgSize = "Size: " & clsStegaRead.ImgSize
        lbImgRes = "Resolution: " & clsStegaRead.ImgRes & " bits"
        
        ToW = Me.ScaleWidth - mPad - mPad
        ToH = Me.ScaleHeight - mPad - mPad
        InW = TheImage.Picture.Width / m_Millimeter
        InH = TheImage.Picture.Height / m_Millimeter
        
        VS.Max = InH - HS.Top + mPad
        HS.Max = InW - VS.Left + mPad
        VS.LargeChange = ToH - HS.Height
        HS.LargeChange = ToW - VS.Width
        
        cmdNext.Enabled = True
    Else
        MsgBox "Error: trying to add unsupport image", vbCritical, "Bad image"
    End If
End Sub

Private Sub Form_Load()
    currentStep = 0
    'for futture use
    'vHex = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", " E", " F")
    'vBin = Array("0000", "0001", "0010", "0011", "0100", "0101", "0110", "0111", "1000", "1001", "1010", "1011", "1100", "1101", " 1110", " 1111")
    Set clsStegaRead = New ClsStegano

    m_Millimeter = ScaleX(100, vbPixels, vbMillimeters)
    vSteps = Array("Step 1: Select the Source image:", "Step 2: Extracting file into image:", "Step 3: Select the files you want to extract:", "Finished: Files extraction")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim It As ListItem
For Each It In lvwFilesAdded.ListItems
    Kill clsStegaRead(It.Key).Filename
Next
Set clsStegaRead = Nothing
End Sub


Private Sub HS_Change()
    TheImage.Move -HS.Value
End Sub

Private Sub HS_Scroll()
    TheImage.Move -HS.Value
End Sub
Private Sub VS_Change()
    TheImage.Move TheImage.Left, -VS.Value
End Sub

Private Sub VS_Scroll()
    TheImage.Move TheImage.Left, -VS.Value
End Sub
