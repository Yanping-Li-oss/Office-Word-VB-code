Attribute VB_Name = "ģ��1"
' ģ�鹦�ܣ�����������MathType��ʽ���ı�ת��ΪLatex��ʽ��ʽ

' ˵����
' 1. ���д˴������ȹ�ѡ��MathType�⣺
'    ���� - ���� - ��ѡMathTypeCommands������MathTypeѡ� - ȷ��
' 2. ���д˴���ǰ����ȫѡ�����ı�

Sub mtEqToLatex()
    Dim fd As Field
    Dim fw As Range
    Dim i As Long
    
    t0 = Now
    Set fw = Selection.Range
    If fw.Start = fw.End Then Exit Sub
    Application.ScreenUpdating = False
    
    For Each fd In fw.Fields
    
        If fd.Code Like "*EMBED Equation.*" Then

            fd.Select
            MathTypeCommands.MTCommand_TeXToggle
            i = i + 1
            
        End If
        
    Next
    
    fw.Select
    Application.ScreenUpdating = True
    
    Set fw = Nothing
    Set fd = Nothing
    
    Debug.Print DateDiff("s", t0, Now)
    
    MsgBox Format(i, "��� ��������0����ʽ")
    
End Sub
