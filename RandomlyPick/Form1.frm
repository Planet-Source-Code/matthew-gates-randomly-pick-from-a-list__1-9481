VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Randomly Picker by: Matthew Gates"
   ClientHeight    =   3225
   ClientLeft      =   3060
   ClientTop       =   2955
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6315
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Randomly Pick"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function randomnumber(finished)
Randomize
randomnumber = Int((Val(finished) * Rnd) + 1)
End Function

Private Sub Command1_Click()
On Error Resume Next
x = InputBox("Enter a word:")
If x = "" Then Exit Sub
List1.AddItem (x)
End Sub

Private Sub Command2_Click()
On Error Resume Next
i = randomnumber(List1.ListCount)
List1.ListIndex = i
End Sub

Private Sub Command3_Click()
On Error Resume Next
If List1.ListCount < 0 Then Exit Sub
  List1.RemoveItem List1.ListIndex
End Sub

Private Sub Form_Load()
For i = 0 To 10
List1.AddItem (i)
Next i
End Sub
