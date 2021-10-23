VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "史志"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550
   LinkTopic       =   "Form2"
   ScaleHeight     =   5310
   ScaleWidth      =   8550
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   5280
      ItemData        =   "Form2.frx":0000
      Left            =   0
      List            =   "Form2.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
