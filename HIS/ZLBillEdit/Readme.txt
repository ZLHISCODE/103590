'缺省属性值:
LocateCol = 1 '定位列(如果不想某列被用户选中，如果用户选择，则控件会自动定位到该列
CmdEnable = False '该控件中的按钮Enable属性
CmdVisible = False '该控件中的按钮Visible属性
CobEnable = False '该控件中的按钮Enable属性
CobVisible = False '该控件中的按钮Visible属性
TxtEnable = False '该控件中的按钮Enable属性
TxtVisible = False '该控件中的按钮Visible属性
MonVisible = False '该控件中的按钮Visible属性 
MonEnable = False '该控件中的按钮Enable属性
    
****设置该控件的列值***** '缺省为0
    '如果列值为1，则显示按钮
    '如果列值为2，则显示按钮，但显示日期控件
    '如果列值为3，则显示下拉框
    '如果列值为4，则显示文本框
    '如果列值为5，则不允许选择,若选择则定位至定位列
    '如果列值为0，则用户可以选择,但不能更改
    '如果列值为其它值，则用户不能选择

相关信息，请仔细阅读------单据控件.Doc

例子源代码：

Private Sub Form_Load()
    msf.Cols = 8

    msf.Clear
    msf.active=true

    msf.AddItem "壹"
    msf.AddItem "贰"
    msf.AddItem "叁"
    msf.AddItem "肆"
    msf.AddItem "伍"
    msf.AddItem "陆"
    msf.AddItem "柒"
    msf.AddItem "捌"
    msf.AddItem "玖"
    msf.AddItem "拾"
    
    msf.TextMatrix(0, 0) = "第1列"
    msf.TextMatrix(0, 1) = "第2列"
    msf.TextMatrix(0, 2) = "第3列"
    msf.TextMatrix(0, 3) = "第4列"
    msf.TextMatrix(0, 4) = "第5列"
    msf.TextMatrix(0, 5) = "第6列"
    msf.TextMatrix(0, 6) = "第7列"
    msf.TextMatrix(0, 7) = "第8列"
    
    msf.ColData(0) = 1
    msf.ColData(1) = 0
    msf.ColData(2) = 2
    msf.ColData(3) = 3
    msf.ColData(4) = 4
    msf.ColData(5) = 5
    msf.ColData(6) = 4
    msf.ColData(7) = 5
    
    msf.ColAlignment(0) = 1
    msf.ColAlignment(1) = 1
    msf.ColAlignment(2) = 1
    msf.ColAlignment(3) = 1
    msf.ColAlignment(4) = 7
    msf.ColAlignment(5) = 4
    msf.ColAlignment(6) = 7
    msf.ColAlignment(7) = 4
    
    msf.TextMatrix(1, 5) = "不能选择"
    msf.TextMatrix(1, 7) = "不能选择"
    
    msf.MaxDate = "9999-12-31"
    msf.MinDate = "1901-01-01"

    Dim Lop As Integer
    msf.Row = 0
    For Lop = 0 To msf.Cols - 1
        msf.Col = Lop
        msf.CellAlignment = 4
    Next
    msf.Row = 1
End Sub

Private Sub msf_cmdselectclick()
    if msf.col=0 then
        MsgBox "谢谢您使用中联财务软件公司的软件！", vbInformation, "中联"
        msf.TextMatrix(msf.Row, 1) = "Thanks！"
        msf.Col = msf.LocateCol
        msf.CmdVisible = False
    endif
End Sub
