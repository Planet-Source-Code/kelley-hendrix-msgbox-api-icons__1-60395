VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Msgbox Icon Example"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option5 
      Caption         =   "All of them"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Critical"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Exclamation"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Information"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Question"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AddMsgPicture(PictureType As MsgPicTypes, x As Long, y As Long, Optional ClearScreen As Boolean = True)

Dim hIcon As Long

    If ClearScreen Then Me.Cls
    hIcon = LoadStandardIcon(0&, PictureType)
    Call DrawIcon(Me.hDC, x, y, hIcon)
    
End Sub

Private Sub Form_Load()
    Me.AutoRedraw = True
End Sub

Private Sub Option1_Click()
    Call AddMsgPicture(Question, 25, 25)
End Sub

Private Sub Option2_Click()
    Call AddMsgPicture(Information, 25, 25)
End Sub

Private Sub Option3_Click()
    Call AddMsgPicture(Exclamation, 25, 25)
End Sub

Private Sub Option4_Click()
    Call AddMsgPicture(Critical, 25, 25)
End Sub

Private Sub Option5_Click()
    Call AddMsgPicture(Question, 25, 25)
    Call AddMsgPicture(Information, 60, 25, False)
    Call AddMsgPicture(Exclamation, 95, 25, False)
    Call AddMsgPicture(Critical, 130, 25, False)
End Sub
