Attribute VB_Name = "ģ��2"
' ģ�鹦�ܣ�����������Latex��ʽ���ı�ת��ΪMathType��ʽ��ʽ

' ˵����
' 1. ���д˴������ȹ�ѡ��MathType�⣺
'    ���� - ���� - ��ѡMathTypeCommands������MathTypeѡ� - ȷ��

Sub LatexToMtEq()
    Dim doc As Document
    Dim rng As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim foundText As String
    
    ' ��ȡ��ǰ�ĵ������ݷ�Χ
    Set doc = ActiveDocument
    Set rng = doc.Content
    
    ' ���ò��ҵĲ���
    With rng.Find
        .ClearFormatting
        .Text = "\$*\$"
        .MatchWildcards = True
        
        ' ��������ƥ����ı�
        Do While .Execute
            ' ��ȡƥ����ı���Χ
            startPos = rng.Start
            endPos = rng.End
            
            ' ��ȡƥ����ı�
            foundText = rng.Text
            
            ' ѡ��ƥ����ı�
            rng.Select
            Debug.Print "Selected text: " & foundText
            
            ' ִ���������������緭ת�ı�
            MathTypeCommands.MTCommand_TeXToggle
            
            ' Ϊ�˷�ֹ��˸
            DoEvents
        Loop
    End With
End Sub


