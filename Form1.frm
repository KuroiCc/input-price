VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7575
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rtex$, Fgb%, Sgb%, Ssel%, F As Boolean

'rtex  用于存储键盘按下后输入文本前的text
'fgb   光标位置
'sge   光标与末尾的间隔字符
'ssel  选中文本长度
'F     判定 防止第二次被format(逻辑型)


Private Sub Text1_Change()
    Dim i%, Ndot%, HaveDot As Boolean
    'Ndot     小数点数量
    'HaveDot  是否有小数点
    t = Text1.Text
    For i = 1 To Len(t)
        If Mid(t, i, 1) = "." Then Ndot = Ndot + 1
    Next
    If Ndot <> 0 Then HaveDot = True Else HaveDot = False
    
    '第一位不能为小数点,小数点只能有一个,只能有两位小数
    If Mid(t, 1, 1) = "." Or Ndot > 1 Or (InStr(t, ".") <> 0 And Len(t) - InStr(t, ".") > 2) Then
        Text1.Text = Rtex
        Text1.SelStart = Fgb
    Else
        If F = True And HaveDot = True Then '有小数点的情况
            Text1.Text = Format(t, "#,###.##")
            If Ssel = 0 Then    '删除时没有被多选
                Text1.SelStart = Len(Text1.Text) - Sgb
            Else                '删除时被多选
                Text1.SelStart = Len(Text1.Text) - (Sgb - Ssel)
            End If
            F = False
        ElseIf F = True Then '没小数点的情况
            Text1.Text = Format(t, "#,###")
            If Ssel = 0 Then
                Text1.SelStart = Len(Text1.Text) - Sgb
            Else
                Text1.SelStart = Len(Text1.Text) - (Sgb - Ssel)
            End If
            F = False
        End If
        If Mid(t, 1, 1) = "," Then 'BUG1修复
            Text1.Text = Right(t, Len(t) - 1)
            If Text1.SelStart <> 0 Then Text1.SelStart = Fgb - 1
        End If
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 46 And KeyAscii <> 8 Then KeyAscii = 0
    If KeyAscii <> 46 Then F = True
    Rtex = Text1.Text
    Fgb = Text1.SelStart
    Ssel = Text1.SelLength
    Sgb = Len(Rtex) - Text1.SelStart
End Sub

'Bug1 删除时如果(1,321,654,654)>>(,654)逗号会保留,而且后面再输入format无效
'Bug2 如果输入30个及以上字符会崩,答案样稿里也有这个BUG
' 字符串处理 中间输入!选中删除!光标定位!艹他MB的烦!!!!!!!还好没有复制粘贴,不然更烦!!!!!!!!!!
