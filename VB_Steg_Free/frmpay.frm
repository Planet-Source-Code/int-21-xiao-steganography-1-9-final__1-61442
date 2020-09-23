VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmPay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Xiao"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   139
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet 
      Left            =   5520
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "Unlock"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtSN 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   4095
   End
   Begin VB.TextBox txtRegto 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Serial Key:"
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Register to:"
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1110
   End
End
Attribute VB_Name = "frmPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
