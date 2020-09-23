VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoad 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   311
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lbStatus 
      AutoSize        =   -1  'True
      Caption         =   "Status..."
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   585
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

