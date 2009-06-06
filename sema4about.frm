VERSION 5.00
Begin VB.Form sema4About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4170
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3705
   ClipControls    =   0   'False
   Icon            =   "sema4About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2878.208
   ScaleMode       =   0  'User
   ScaleWidth      =   3479.187
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   120
      Picture         =   "sema4About.frx":00EA
      ScaleHeight     =   474.075
      ScaleMode       =   0  'User
      ScaleWidth      =   1053.5
      TabIndex        =   1
      Top             =   3240
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2280
      TabIndex        =   0
      Top             =   3570
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   3267.9
      Y1              =   2101.713
      Y2              =   2101.713
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   1770
      Left            =   90
      TabIndex        =   2
      Top             =   1125
      Width           =   3285
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   90
      TabIndex        =   3
      Top             =   240
      Width           =   3405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   3267.9
      Y1              =   2112.066
      Y2              =   2112.066
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   90
      TabIndex        =   4
      Top             =   720
      Width           =   3405
   End
End
Attribute VB_Name = "sema4About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription = App.Comments & vbCrLf & vbCrLf & _
                     "Original Mike Bolton 2005" & vbCrLf & _
                     "Modified Chris White, Mark Patrick 2008-2009"
End Sub
