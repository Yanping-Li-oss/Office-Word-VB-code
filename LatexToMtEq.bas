Attribute VB_Name = "模块2"
' 模块功能：批量将包含Latex公式的文本转换为MathType公式格式

' 说明：
' 1. 运行此代码请先勾选上MathType库：
'    工具 - 引用 - 勾选MathTypeCommands或其他MathType选项卡 - 确定

Sub LatexToMtEq()
    Dim doc As Document
    Dim rng As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim foundText As String
    
    ' 获取当前文档的内容范围
    Set doc = ActiveDocument
    Set rng = doc.Content
    
    ' 设置查找的参数
    With rng.Find
        .ClearFormatting
        .Text = "\$*\$"
        .MatchWildcards = True
        
        ' 查找所有匹配的文本
        Do While .Execute
            ' 获取匹配的文本范围
            startPos = rng.Start
            endPos = rng.End
            
            ' 获取匹配的文本
            foundText = rng.Text
            
            ' 选中匹配的文本
            rng.Select
            Debug.Print "Selected text: " & foundText
            
            ' 执行其他操作，例如翻转文本
            MathTypeCommands.MTCommand_TeXToggle
            
            ' 为了防止闪烁
            DoEvents
        Loop
    End With
End Sub


