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
   StartUpPosition =   3  '����ȱʡ
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

'rtex  ���ڴ洢���̰��º������ı�ǰ��text
'fgb   ���λ��
'sge   �����ĩβ�ļ���ַ�
'ssel  ѡ���ı�����
'F     �ж� ��ֹ�ڶ��α�format(�߼���)


Private Sub Text1_Change()
    Dim i%, Ndot%, HaveDot As Boolean
    'Ndot     С��������
    'HaveDot  �Ƿ���С����
    t = Text1.Text
    For i = 1 To Len(t)
        If Mid(t, i, 1) = "." Then Ndot = Ndot + 1
    Next
    If Ndot <> 0 Then HaveDot = True Else HaveDot = False
    
    '��һλ����ΪС����,С����ֻ����һ��,ֻ������λС��
    If Mid(t, 1, 1) = "." Or Ndot > 1 Or (InStr(t, ".") <> 0 And Len(t) - InStr(t, ".") > 2) Then
        Text1.Text = Rtex
        Text1.SelStart = Fgb
    Else
        If F = True And HaveDot = True Then '��С��������
            Text1.Text = Format(t, "#,###.##")
            If Ssel = 0 Then    'ɾ��ʱû�б���ѡ
                Text1.SelStart = Len(Text1.Text) - Sgb
            Else                'ɾ��ʱ����ѡ
                Text1.SelStart = Len(Text1.Text) - (Sgb - Ssel)
            End If
            F = False
        ElseIf F = True Then 'ûС��������
            Text1.Text = Format(t, "#,###")
            If Ssel = 0 Then
                Text1.SelStart = Len(Text1.Text) - Sgb
            Else
                Text1.SelStart = Len(Text1.Text) - (Sgb - Ssel)
            End If
            F = False
        End If
        If Mid(t, 1, 1) = "," Then 'BUG1�޸�
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

'Bug1 ɾ��ʱ���(1,321,654,654)>>(,654)���Żᱣ��,���Һ���������format��Ч
'Bug2 �������30���������ַ����,��������Ҳ�����BUG
' �ַ������� �м�����!ѡ��ɾ��!��궨λ!ܳ��MB�ķ�!!!!!!!����û�и���ճ��,��Ȼ����!!!!!!!!!!
