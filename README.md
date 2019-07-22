# Translate-plug-in-for-Excel
**Translate plug-in for Excel,Excel的翻译插件，适合一些经常与英文Excel打交道的英文菜鸟**
## 使用方法
使用的网易有道在线翻译API
+ 1、将“Excel文件翻译（备注）.xla”文件放入到 C:\Users\{你的windows用户名}\AppData\Roaming\Microsoft\AddIns 目录下
+ 2、随便打开一个Excel，点击菜单：文件->选项->加载项->下面的“管理”“转到（G）” 弹出“加载项”选项卡
    + 2.1 点击“浏览(B)” 弹出文件选择框
    + 2.2 选择 C:\Users\{你的windows用户名}\AppData\Roaming\Microsoft\AddIns\Excel文件翻译（备注）.xla文件
    + 2.3 确认已勾选后点击“确定”即可
    + 2.4 关闭Excel，再重新打开Excel（如果提示需要宏启用的，点击启用），在文件、...视图..的主菜单位置后面会出现一个“加载项”的菜单，里面就有“翻译到备注”的功能按钮了
+ 3、操作说明
    + 3.1 点击 “① 提取中英文”菜单，会临时在最前面加一个工作簿，名称为fanyi_en2zh，用于临时字典存放于手工处理的地方，之后会遍历所有可见的工作表，对其进行中英文检索，输出到fanyi_en2zh表的A列当中去，之后自动在C列添加相应的有道在线翻译公式，翻译后将把结果值转到B列，并清理C列的公式，避免提示“需要更新链接”之类的，之后就能看到对于的中英文对照关系，一些专业词汇翻译的可能不是很准确，且有特殊符合的问题，需要手工整理下（有道API给的就是这个，我也没辙，如果一些翻译不出来的，可能是意大利语或者法语之类的，毕竟是少数，自己再百度翻译下看看是否提示可能是某某语言）
    + 3.2 点击 “② 将翻译结果添加到备注”菜单，将会把“fanyi_en2zh”整理好的字典，一个个的写入到对应单元格备注里
    + 3.3 点击 “③ 清除翻译备注”菜单，会清理之前写入的字典备注
    + 3.4 点击 “④ 清理临时工作簿”菜单，会提示将“fanyi_en2zh”工作簿进行删除。
    
## 代码分析
```VBA
  Option Explicit
'启动时添加菜单“翻译到备注”
Private Sub Workbook_Open()
    AddMenuItemExample
End Sub
' 添加菜单
Public Sub AddMenuItemExample()
    Dim cbWSMenuBar As CommandBar
    Dim cbc As CommandBarControl

    Set cbWSMenuBar = Application.CommandBars("Worksheet Menu Bar")
    Set cbc = cbWSMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    cbc.Tag = "翻译到备注"
    With cbc
        .Caption = "&翻译到备注"
        With .Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = "① 提取中英文"
            .OnAction = "ThisWorkbook.提取所有英文"
            .Tag = "Item1"
        End With
        With .Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = "② 将翻译结果添加到备注"
            .OnAction = "ThisWorkbook.切换成中文"
            .BeginGroup = True
            .Tag = "Item4"
        End With
        With .Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = "③ 清除翻译备注"
            .OnAction = "ThisWorkbook.清除翻译备注"
            .Tag = "Item5"
            .BeginGroup = True
        End With
        With .Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = "④ 清理临时工作簿"
            .OnAction = "ThisWorkbook.清理临时工作簿"
            .Tag = "Item5"
            .BeginGroup = True
        End With
    End With
End Sub
 
Sub 提取所有英文()
    Dim arr, i&, j&, txt$, en2cnSheet As Worksheet, wb, ws, d, maxR&, maxC&, EN2CN$, titleRange As Range
    
    EN2CN = "fanyi_en2zh"
    Set wb = ActiveWorkbook
    ' 创建一个字典对象：用于数据去重复
    Set d = CreateObject("scripting.dictionary")

    ' 遍历查找是否存在 fanyi_en2zh 工作簿
    For Each ws In wb.Worksheets
        If ws.Name = EN2CN Then
            Set en2cnSheet = ws
        End If
    Next
    
    On Error Resume Next
    ' 如果没有就临时新建一个fanyi_en2zh 工作簿，用于存放中英文对照
    If en2cnSheet Is Nothing Then
        wb.Sheets.Add Before:=Sheets(Sheets.Count)
        wb.Sheets(Sheets.Count).Name = EN2CN
        Set en2cnSheet = wb.Sheets(EN2CN)
        en2cnSheet.Cells(1, 1) = "提取的英文"
        en2cnSheet.Cells(1, 2) = "清理后的中文"
        Set titleRange = Range("A1:B1")
        titleRange.Interior.ColorIndex = 6
        titleRange.Font.Size = 16
        titleRange.Font.Bold = True
        titleRange.HorizontalAlignment = Excel.xlCenter
    End If
    ' 防止之前存在却被隐藏的可能
    en2cnSheet.Visible = True
    ' 读取这个中英文对照表转化为数组存入字典中：可以点击多次也不会报错的
    arr = en2cnSheet.UsedRange
    maxR = UBound(arr)
    maxC = UBound(arr, 2)
    For i = 2 To maxR
        d(arr(i, 1)) = arr(i, 2)
    Next
    
    ' 遍历所有非隐藏的工作簿
    For Each ws In wb.Worksheets
        If ws.Name <> EN2CN And ws.Visible Then
            'MsgBox ("正在查找工作簿：" & ws.Name & "中的所有英文...")
            With ws
                arr = .UsedRange
                maxR = UBound(arr)
                maxC = UBound(arr, 2)
                For i = 1 To maxR
                    For j = 1 To maxC
                        ' 遍历拿到每个单元格数据，测试是否是字符串类型且不为空，且不是数字且不在已有字典中的，将会加入新的字典中
                        txt = arr(i, j)
                        If VarType(txt) = 8 And txt <> "" And Not IsNumeric(txt) And Not d.exists(txt) Then
                            d(txt) = ""
                        End If
                    Next
                Next
             End With
             
         End If
    Next
    ' 将字典结果批量导出到fanyi_en2zh 工作簿中
    en2cnSheet.Range("a2").Resize(d.Count, 1) = Application.Transpose(d.keys)
    en2cnSheet.Select
    ' 设置AB两列自动宽度显示
    en2cnSheet.Columns("A:B").EntireColumn.AutoFit
    
    msg "提取所有文字", "成功！", "即将进行在线翻译（请确保联网！）"
    ' 使用网易有道官方API进行翻译
    写入翻译公式 en2cnSheet
    msg "已在线翻译", "成功！", "即将清理联网公式痕迹"
    ' 清除公式（避免下次打开的时候一直提示需要更新链接）
    清理翻译 en2cnSheet
    en2cnSheet.Select
    msg "清理联网公式痕迹", "成功！", "请先看看翻译结果是否正确，不正确的请自行处理后，再使用菜单【加载项】-【翻译到备注】-【② 将翻译结果添加到备注】执行翻译"
End Sub

Sub 切换成中文()
   Dim arr, i&, j&, txt$, en2cnSheet As Worksheet, wb, ws, d, maxR&, maxC&, EN2CN$, cell As Range
    EN2CN = "fanyi_en2zh"
    Set wb = ActiveWorkbook
    
    For Each ws In wb.Worksheets
        If ws.Name = EN2CN Then
            Set en2cnSheet = ws
        End If
    Next
    On Error Resume Next
    If en2cnSheet Is Nothing Then
        msg "错误", "请先执行第①步：菜单【加载项】-【翻译到备注】-【① 提取英文】", ""
        Exit Sub
    End If

    ' 加载fanyi_en2zh工作簿中已经编辑好的字典
    Set d = CreateObject("scripting.dictionary")
    arr = en2cnSheet.UsedRange
    maxR = UBound(arr)
    maxC = UBound(arr, 2)
    For i = 2 To maxR
        d(arr(i, 1)) = arr(i, 2)
    Next
    
    If d.Count < 2 Then
         msg "错误", "请先执行第①步：菜单【加载项】-【翻译到备注】-【① 提取英文】", ""
         Exit Sub
    End If
    
    ' 遍历工作簿所有需要翻译的单元格，将其翻译结果放到备注中（以免修改原值会影响到一些公式的使用）
    For Each ws In wb.Worksheets
        If ws.Name <> EN2CN And ws.Visible Then
            For Each cell In ws.UsedRange
                txt = cell.Value
                If VarType(txt) = 8 And txt <> "" And Not IsNumeric(txt) And d.exists(txt) Then
                    cell.Select
                    ' 设置备注
                    setActiveCellComments cell, d(txt)
                End If
            Next
        End If
    Next
    
    msg "将翻译结果添加备注", "成功！", "如果您想取消这些备注请使用：菜单【加载项】-【翻译到备注】-【③ 清除翻译备注】；【④ 清理临时工作簿】将删除“fanyi_en2zh”这个临时工作簿"
    en2cnSheet.Select
End Sub

 ' 自动确定包含总行
Function total_rows(tsheet)
    Dim StartRow As Long
    Dim ASh
    On Error Resume Next
    With tsheet.UsedRange
        ASh = .Rows
        StartRow = .Row
        total_rows = StartRow + UBound(ASh, 1) - 1
    End With
End Function

 ' 自动确定包含总列
Function total_cols(tsheet)
    Dim StartColumn As Integer
    Dim ASh
    On Error Resume Next
    With tsheet.UsedRange
        ASh = .Rows
        StartColumn = .Column
        total_cols = StartColumn + UBound(ASh, 2) - 1
    End With
End Function

Sub 写入翻译公式(en2cnSheet)
    'ActiveWorkbook.ActiveSheet
    Dim i&, maxR&
    maxR = total_rows(en2cnSheet)
    
    '此处使用的是XML过滤器函数对网易有道翻译结果进行抽取
    For i = 2 To maxR
        en2cnSheet.Cells(i, 3).Formula = "=FILTERXML(WEBSERVICE(""http://fanyi.youdao.com/translate?&i=""&A" + Trim(Str(i)) + "&""&doctype=xml&version""),""//translation"")"
    Next
    
End Sub

Sub 清理翻译(en2cnSheet)
    Dim i&, maxR&
    maxR = total_rows(en2cnSheet)
    ' 清理翻译公式，避免下次提示需要更新链接
    For i = 2 To maxR
        en2cnSheet.Cells(i, 2).Value = en2cnSheet.Cells(i, 3).Value
        en2cnSheet.Cells(i, 3).Clear
    Next
    
End Sub

Sub 清理临时工作簿()
    Dim wb, ws, en2cnSheet As Worksheet, EN2CN$
    EN2CN = "fanyi_en2zh"
    Set wb = ActiveWorkbook
    
    ' 删除临时工作簿
    For Each ws In wb.Worksheets
        If ws.Name = EN2CN Then
            ws.Select
            ws.Delete
        End If
    Next
End Sub
' 添加备注
Sub setActiveCellComments(cell, info)
        Dim ocm$, reg, ncm$
        ncm = "翻译:" & Chr(10) & "【 " & info & " 】"
        If cell.Comment Is Nothing Then
            ' 添加一个新的备注
            cell.AddComment
            ' 设置不自动显示，需要鼠标滑过才会显示
            cell.Comment.Visible = False
            ' 设置备注内容
            cell.Comment.Text Text:=ncm
        ' 如果原来只有翻译备注的话
        ElseIf cell.Comment.Text Like "翻译*" Then
            cell.Comment.Text Text:=ncm
        ' 如果有翻译也有其他备注的话，请启用正则表达式进行更新替换
        ElseIf cell.Comment.Text Like "*翻译*" Then
            ' 先取出原有备注信息
            ocm = cell.Comment.Text
            ' 创建一个正则表达式对象
            Set reg = CreateObject("vbscript.regexp")
            With reg
                ' 设置为全局匹配
                .Global = True
                ' 设置为区分大小写
                .IgnoreCase = True
                ' 设置正则表达式规则
                .Pattern = "翻译\:\s\【.*? \】"
                ' 正则替换成的内容
                ocm = .Replace(ocm, ncm)
            End With
            cell.Comment.Text Text:=ocm
            Set reg = Nothing
        ' 如果没有翻译备注的话，追加备注
        ElseIf Not cell.Comment.Text Like "翻译*" And Not cell.Comment.Text Like "*翻译*" Then
            ocm = cell.Comment.Text
            ocm = ocm & Chr(10) & ncm
            cell.Comment.Text Text:=ocm
        End If
End Sub

Sub 清除翻译备注()
    Dim txt$, wb, ws, cell As Range, EN2CN$, cm$, reg
    EN2CN = "fanyi_en2zh"
    Set wb = ActiveWorkbook
    Set reg = CreateObject("vbscript.regexp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.Pattern = "翻译\:\s\【.*? \】"
    
    ' 遍历所有可见工作簿的有备注的单元格，进行清理原来翻译的备注
    For Each ws In wb.Worksheets
        If ws.Name <> EN2CN And ws.Visible Then
            For Each cell In ws.UsedRange
                txt = cell.Value
                If Not cell.Comment Is Nothing Then
                    ' 取出原有备注内容
                    cm = cell.Comment.Text
                    ' 如果是以“翻译”打头的，说明是俺做的，直接清理就好了
                    If cell.Comment.Text Like "翻译*" Then
                        cell.ClearComments
                    ' 如果“翻译”两字在中间的话，则表明之前是追加进来的，需要过滤替换掉就行
                    ElseIf cell.Comment.Text Like "*翻译*" Then
                        cm = reg.Replace(cm, "")
                        cell.Comment.Text Text:=cm
                    End If
                End If
            Next
        End If
    Next
    Set wb = Nothing
    Set ws = Nothing
    Set reg = Nothing
    Set cell = Nothing
    
    msg "清除翻译", "成功！", "如果不想留下临时工作簿“fanyi_en2zh” ，请使用：菜单【加载项】-【翻译到备注】-【④ 清理临时工作簿】"
End Sub
'自定义的提示消息过程
Sub msg(title, msg, tip)
    MsgBox title & ":" & msg & Chr(10) & "------------------------------------------------------------------------------------" & Chr(10) & "提示：" & tip, vbOKOnly, title
End Sub
```
