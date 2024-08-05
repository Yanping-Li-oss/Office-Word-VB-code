Attribute VB_Name = "模块1"
' 模块功能：批量将包含MathType公式的文本转换为Latex公式格式

' 说明：
' 1. 运行此代码请先勾选上MathType库：
'    工具 - 引用 - 勾选MathTypeCommands或其他MathType选项卡 - 确定
' 2. 运行此代码前请先全选测试文本

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
    
    MsgBox Format(i, "完成 共处理了0个公式")
    
End Sub
