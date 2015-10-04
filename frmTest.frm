VERSION 5.00
Begin VB.Form frmTest 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Test"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4620
   ScaleHeight     =   2265
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btShow 
      Caption         =   "Show"
      Height          =   735
      Left            =   1980
      TabIndex        =   1
      Top             =   660
      Width           =   1515
   End
   Begin VB.PictureBox pictureTest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1080
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   780
      Width           =   495
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btShow_Click()

    Dim i As Integer
    
    Call Randomize
    
    i = CInt((Rnd * 128) + 1)
    
    Call ShowIcon(pictureTest, "shell32.dll", i)
    
End Sub

