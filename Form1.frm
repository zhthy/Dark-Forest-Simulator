VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "主界面"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   4305
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "导出史志"
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "剩余文明"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "总文明："
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "友善文明："
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "中立文明："
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "好斗文明："
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
For i = 0 To Form2.List1.ListCount
    Text = Text & vbCrLf & Form2.List1.List(i)
Next i
Open App.Path + "\1.txt" For Output As #1
Print #1, Text
Close #1
Name App.Path + "/1.txt" As App.Path + "/" + Str(Day(Date)) + "日的史志.txt"
MsgBox "导出成功！", , "提示"
End Sub

Private Sub Form_Load()
Form2.Show
Form3.Show
Form3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
