VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "���㴰��"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5325
   LinkTopic       =   "Form3"
   ScaleHeight     =   3675
   ScaleWidth      =   5325
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1920
      Top             =   2280
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(0 To 100)
Dim zong, hao, zhong, you, n, sheng, max, lin
Dim zong1, hao1, zhong1, you1
Dim b(0 To 100)
Private Sub Form_Load()
Randomize
zong = Int(Rnd * 100) + 1
max = Int(Rnd * 100) + 1
For i = 0 To zong
    a(i) = Int(Rnd * 3) + 1
    b(i) = Int(Rnd * (max + 1))
Next
For i = 1 To lin
    If a(i) = 1 Then hao = hao + 1
    If a(i) = 2 Then zhong = zhong + 1
    If a(i) = 3 Then you = you + 1
Next i
Form1.Text1.Text = Str(hao)
    Form2.List1.AddItem "����ö�����"
Form1.Text2.Text = Str(zhong)
    Form2.List1.AddItem "������������"
Form1.Text3.Text = Str(you)
    Form2.List1.AddItem "������������"
Form1.Text4.Text = Str(zong)
    Form2.List1.AddItem "����������"
    Form2.List1.AddItem "���������С:" & max & "���ʿˡ�"
    Form2.List1.AddItem "������ɡ�"
    Form2.List1.AddItem "���ã�����������ӭ����������ģ������"
n = 0
lin = zong
    Form1.Text4.Text = zong
    Form1.Text1.Text = hao
    Form1.Text2.Text = zhong
    Form1.Text3.Text = you
End Sub

Private Sub Timer1_Timer()
n = n + 1
Form2.List1.AddItem "��ʼ��" & n & "��ģ�⡣"
For l = 1 To lin
    For i = 1 To lin - 1
        If b(i + 1) = b(i) And a(i) <> 0 Then
            Randomize
            sheng = Int(Rnd * 100)
            Select Case a(i)
                Case 1
                    If sheng <= 75 Then
                        a(i + 1) = 0
                        b(i + 1) = -1
                        If sheng < 31 Then a(i) = 1
                        Form2.List1.AddItem i & "��������" & (i + 1) & "��������ս��, " & i & "������ ʤ����, " & (i + 1) & "  �����������ˡ�"
                    Else
                        a(i) = 0
                        b(i) = -1
                        If sheng < 31 Then a(i + 1) = 1
                        Form2.List1.AddItem i & "��������" & (i + 1) & "��������ս��, " & (i + 1) & "������ ʤ����, " & i & "  �����������ˡ�"
                    End If
                Case 2
                    If sheng <= 50 Then
                        a(i + 1) = 0
                        b(i + 1) = -1
                        If sheng < 31 Then a(i) = 1
                        Form2.List1.AddItem i & "��������" & (i + 1) & "��������ս��, " & i & "������ ʤ����, " & (i + 1) & "  �����������ˡ�"
                    Else
                        a(i) = 0
                        b(i) = -1
                        If sheng < 31 Then a(i + 1) = 1
                        Form2.List1.AddItem i & "��������" & (i + 1) & "��������ս��, " & (i + 1) & "������ ʤ����, " & i & "  �����������ˡ�"
                    End If
                Case 3
                    If sheng <= 25 Then
                        a(i + 1) = 0
                        b(i + 1) = -1
                        If sheng < 31 Then a(i) = 1
                        Form2.List1.AddItem i & "��������" & (i + 1) & "��������ս��, " & i & "������ ʤ����, " & (i + 1) & "  �����������ˡ�"
                    Else
                        a(i) = 0
                        b(i) = -1
                        If sheng < 31 Then a(i + 1) = 1
                        Form2.List1.AddItem i & "��������" & (i + 1) & "��������ս��, " & (i + 1) & "������ ʤ����, " & i & "  �����������ˡ�"
                    End If
            End Select
        Else
            Form2.List1.AddItem "��������ǳ�������������֮�䲢û�з��ֶԷ���������ʱ������ƣ�������������λ�÷�����һЩ�ı䡣"
            For k = 0 To lin
                b(k) = Int(Rnd * (max + 1))
            Next
        End If
    Next i
Next l

For i = 1 To lin
        If a(i) = 1 Then hao1 = hao1 + 1
        If a(i) = 2 Then zhong1 = zhong1 + 1
        If a(i) = 3 Then you1 = you1 + 1
        If a(i) = 0 Then zong1 = zong1 + 1
Next i
    Form1.Text4.Text = zong - zong1
    Form1.Text1.Text = hao1
    Form1.Text2.Text = zhong1
    Form1.Text3.Text = you1
        Form2.List1.AddItem "ʣ��������" & Form1.Text4.Text & "����"
        Form2.List1.AddItem "ʣ��ö�������" & Form1.Text1.Text & "����"
        Form2.List1.AddItem "ʣ������������" & Form1.Text2.Text & "����"
        Form2.List1.AddItem "ʣ������������" & Form1.Text3.Text & "����"
    hao1 = 0
    zhong1 = 0
    you1 = 0
    zong1 = 0
    If Val(Form1.Text1.Text) < Val(Form1.Text3.Text) / 2 Or Val(Form1.Text1.Text) < Val(Form1.Text2.Text) / 2 Then
        Form2.List1.AddItem "���������ڵ�" & n & "�����ݣ��ع��ƽ������ս����"
        Timer1.Enabled = False
    End If
    If Val(Form1.Text1.Text) = Val(Form1.Text3.Text) / 2 Or Val(Form1.Text1.Text) = Val(Form1.Text2.Text) / 2 Then
        Form2.List1.AddItem "���������ڵ�" & n & "�����ݣ��ع��ƽ������ֻ����ʱ�ģ�������һ��һ�ж�����ͷ��ʼ��ս�����ڻ��ȼ��"
        Timer1.Enabled = False
    End If
    If Val(Form1.Text4.Text) <= 2 And Val(Form1.Text4.Text) >= 1 Then
        Form2.List1.AddItem "���������ڵ�" & n & "�����ݣ������ˣ�ֻ�в���������ǻ𻹴����ˡ�"
        Timer1.Enabled = False
    End If
    If Val(Form1.Text4.Text) < 1 Then
        Form1.Text4.Text = "0"
        Form2.List1.AddItem "���������ڵ�" & n & "�����ݣ��������"
        Timer1.Enabled = False
    End If
    If Val(Form1.Text2.Text) > 0 And Val(Form1.Text3.Text) = 0 And Val(Form1.Text1.Text) = 0 Then
        Form2.List1.AddItem "���������ڵ�" & n & "�����ݣ���ȡ���������档"
        Timer1.Enabled = False
    End If
    If Val(Form1.Text2.Text) = 0 And Val(Form1.Text3.Text) > 0 And Val(Form1.Text1.Text) = 0 Then
        Form2.List1.AddItem "���������ڵ�" & n & "�����ݣ���ȡ���������档"
        Timer1.Enabled = False
    End If
End Sub










