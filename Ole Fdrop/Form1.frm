VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   945
      Width           =   4755
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      ItemData        =   "Form1.frx":0006
      Left            =   105
      List            =   "Form1.frx":000D
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   270
      Width           =   2085
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "Form1.frx":002D
      Left            =   90
      List            =   "Form1.frx":0034
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub List2_Click()
Text1.Text = List1.List(List2.ListIndex)
End Sub

Private Sub List2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Data.GetFormat(vbCFFiles) = True) Then
 Dim FN
 Dim FL As String
 For Each FN In Data.Files
    FL = FN
    List1.AddItem FL
    List2.AddItem Filenm(FL)
    Next FN
End If
'List1.Selected(1) = True
List2.Selected(1) = True
End Sub
Private Function Filenm(strx As String) As String
Dim sps As Integer
'sl As Integer,
'sl = Len(strx)
For sps = Len(strx) To 1 Step -1
If Mid(strx, sps, 1) = "\" Then
Filenm = Mid$(strx, sps + 1)
Exit For
End If
Next
End Function
