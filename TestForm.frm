VERSION 5.00
Begin VB.Form TestForm 
   Caption         =   "Hello"
   ClientHeight    =   3195
   ClientLeft      =   8700
   ClientTop       =   3795
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox (Text1.Text)

End Sub
