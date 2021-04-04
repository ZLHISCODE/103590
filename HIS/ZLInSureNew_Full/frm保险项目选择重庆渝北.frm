VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm保险项目选择重庆渝北 
   AutoRedraw      =   -1  'True
   Caption         =   "医保项目选择"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm保险项目选择重庆渝北.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7845
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1575
      Width           =   45
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   7845
      TabIndex        =   1
      Top             =   4965
      Width           =   7845
      Begin VB.CommandButton cmd新增 
         Caption         =   "新增(&N)"
         Height          =   350
         Left            =   2625
         TabIndex        =   10
         ToolTipText     =   "从中心下载服务项目、病种信息和定点医疗机构"
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdRequery 
         Caption         =   "项目下载"
         Height          =   350
         Left            =   1335
         TabIndex        =   5
         ToolTipText     =   "从中心下载服务项目、病种信息和定点医疗机构"
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印列表"
         Height          =   350
         Left            =   15
         TabIndex        =   4
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   3
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   2
         Top             =   180
         Width           =   1100
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshGrid 
      Height          =   3990
      Left            =   3045
      TabIndex        =   0
      Top             =   390
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   7038
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择重庆渝北.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择重庆渝北.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   7
      Top             =   255
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7144
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目明细(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   9
      Top             =   15
      Width           =   4710
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目大类(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   8
      Top             =   0
      Width           =   2970
   End
End
Attribute VB_Name = "frm保险项目选择重庆渝北"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mstrCode As String
Private mstrName As String
Private mblnOK As Boolean

Private mLocalCode As String '指向编码
Private mblnFirst As Boolean
'服务目录文件导出
Private Declare Function ExportKA02K3 Lib "YHMdcrAsistntSvr.dll" Alias "_ExportKA02K3@12" (ByVal strYab003 As String, ByVal strFileName As String, ByRef tmpStrut As Struct) As Boolean
'病种目录文件导出
Private Declare Function ExportKA06K1 Lib "YHMdcrAsistntSvr.dll" Alias "_ExportKA06K1@12" (ByVal strYab003 As String, ByVal strFileName As String, ByRef tmpStrut As Struct) As Boolean

Private Declare Function ExportKA03K1 Lib "YHMdcrAsistntSvr.dll" Alias "_ExportKA03K1@12" (ByVal strYab003 As String, ByVal strFileName As String, ByRef tmpStrut As Struct) As Boolean

Private Type Struct
    lngAppCode  As Long   '标志服务执行状态代码。等于1时表示服务执行正常结束，小于0时表示服务执行异常或错误。
    strErrMsg  As String  '当服务执行状态代码AppCod小于0时，描述服务执行的异常或错误信息。
End Type
Private mbln诊疗 As Boolean
 

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(mshGrid.TextMatrix(mshGrid.Row, 0)) = "" Then
        MsgBox "没有选择项目！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '返回选择项目编码
    mstrCode = mshGrid.TextMatrix(mshGrid.Row, 0)
    mstrName = mshGrid.TextMatrix(mshGrid.Row, 1)
    mblnOK = True
    Unload Me
End Sub

Private Function Loadtree() As Boolean
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim tmpNode As Node
    mblnOK = False
    
    On Error GoTo ErrHand:
    
    '装载数据
    gstrSQL = "" & _
        "   Select distinct  编码,名称 From 医保项目分类 " & IIf(mbln诊疗, " Where 编码='61'", "")
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        MsgBox "医保前置服务器中没有医保项目分类，无法选择，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Function
    End If
    
    tvwClass.Nodes.Clear
    Do Until rsTemp.EOF
        Set tmpNode = tvwClass.Nodes.Add(, 4, "K" & Nvl(rsTemp!编码), "【" & Nvl(rsTemp("编码")) & "】" & Nvl(rsTemp("名称")), "Detail", "Detail")
        tmpNode.Sorted = True
        rsTemp.MoveNext
    Loop
    tvwClass.Nodes(1).Selected = True
    Call FillList
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    Loadtree = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Loadtree = False
End Function
Public Function GetCode(ByVal frmMain As Form, strCode As String, strName As String, Optional bln诊疗 As Boolean = False) As Boolean
    '功能：获取编码
    '参数：
    '返回：成功返回True
    mLocalCode = strCode
    mbln诊疗 = bln诊疗
    
    frm保险项目选择重庆渝北.Show vbModal, frm保险项目
    '返回值
    If mblnOK = True Then
        strCode = mstrCode
        strName = mstrName
    End If
    GetCode = mblnOK
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetGrdColHead()
    With mshGrid
        .Clear
        .Rows = 2
        .Cols = 35
        .TextMatrix(0, 0) = "商品代码"
        .TextMatrix(0, 1) = "商品名"
        .TextMatrix(0, 2) = "药品通用中文名"
        .TextMatrix(0, 3) = "药品通用英文名"
        .TextMatrix(0, 4) = "商品曾用名"
        .TextMatrix(0, 5) = "别名"
        .TextMatrix(0, 6) = "包装规格"
        .TextMatrix(0, 7) = "医院大类编码"
        .TextMatrix(0, 8) = "医保编码"
        .TextMatrix(0, 9) = "最小包装单位"
        .TextMatrix(0, 10) = "最小计量单位"
        .TextMatrix(0, 11) = "每日最大用量"
        .TextMatrix(0, 12) = "指导价格"
        .TextMatrix(0, 13) = "招标价格"
        .TextMatrix(0, 14) = "基金支付限价1"
        .TextMatrix(0, 15) = "基金支付限价2"
        .TextMatrix(0, 16) = "基金支付限价3"
        .TextMatrix(0, 17) = "实际执行价格"
        .TextMatrix(0, 18) = "自付比例1"
        .TextMatrix(0, 19) = "自付比例2"
        .TextMatrix(0, 20) = "自付比例3"
        .TextMatrix(0, 21) = "自付比例4"
        .TextMatrix(0, 22) = "自付比例5"
        .TextMatrix(0, 23) = "自付比例6"
        .TextMatrix(0, 24) = "自付比例7"
        .TextMatrix(0, 25) = "自付比例8"
        .TextMatrix(0, 26) = "自付比例9"
        .TextMatrix(0, 27) = "自付比例10"
        .TextMatrix(0, 28) = "自付比例11"
        .TextMatrix(0, 29) = "自付比例12"
        .TextMatrix(0, 30) = "标准编号"
        .TextMatrix(0, 31) = "五笔助记码1"
        .TextMatrix(0, 32) = "拼音助记码1"
        .TextMatrix(0, 33) = "备注"
        .TextMatrix(0, 34) = "目录分类"
    End With

End Sub
Private Sub FillList()
    '功能：显示当前类别下的医保明细
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, fld As ADODB.Field
    Dim str类别代码 As String, blnColSet As Boolean
    Dim lngCol  As Long
    Dim varValue As Variant
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo errHandle
    With tvwClass.SelectedItem
        str类别代码 = IIf(.Key = "Root", "", " And 医院大类编码 ='" & Mid(.Key, 2) & "'")
    End With
    
    
    rsTemp.CursorLocation = adUseClient
    
'    gstrSQL = " select  商品代码,  医院大类编码, 医保编码, 药品通用中文名, 药品通用英文名,商品名, 商品曾用名, 服务项目结算方式, 报销标识, 医保标识, 是否处方用药, 药品适应症, 限制医生, 审批权限, 别名, 包装规格, " & _
             "         最小包装单位, 最小计量单位, 每日最大用量, 指导价格, 招标价格, 基金支付限价1, 基金支付限价2, 基金支付限价3, 实际执行价格, 自付比例1, 自付比例2, 自付比例3, 自付比例4, 自付比例5, 自付比例6, 自付比例7, 自付比例8,  " & _
             "         自付比例9, 自付比例10, 自付比例11, 自付比例12, 医院使用状态, 中心使用状态, 标准编号,  " & _
             "         五笔助记码1, 五笔助记码2, 五笔助记码3, 拼音助记码1, 拼音助记码2, 拼音助记码3, 备注, 医保经办机构,机构标准编号, 医疗机构编号, " & _
             "          修改时间, 目录分类  " & _
             "  from 医保服务项目目录" & _
             "  where 1=1 " & str类别代码
    
    gstrSQL = " select  商品代码,商品名, 药品通用中文名, 药品通用英文名, 商品曾用名, 别名, 包装规格,医院大类编码, 医保编码, " & _
             "         最小包装单位, 最小计量单位, 每日最大用量, 指导价格, 招标价格, 基金支付限价1, 基金支付限价2, 基金支付限价3, 实际执行价格, 自付比例1, 自付比例2, 自付比例3, 自付比例4, 自付比例5, 自付比例6, 自付比例7, 自付比例8,  " & _
             "         自付比例9, 自付比例10, 自付比例11, 自付比例12, 标准编号,  " & _
             "         五笔助记码1, 拼音助记码1, 备注, " & _
             "         目录分类  " & _
             "  from 医保服务项目目录" & _
             "  where 1=1 " & str类别代码
    
    rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    
    If rsTemp.RecordCount = 0 Then
        '设置列头
        Call SetGrdColHead
    Else
        Set mshGrid.DataSource = rsTemp
    End If
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 '   LockWindowUpdate 0
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdPrint_Click()
    If gstrUserName = "" Then Call GetUserInfo
    subPrint 1
End Sub

Private Sub subPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwClass.SelectedItem
    Set objPrint.Body = mshGrid
    objPrint.Title.Text = "保险项目"
    
    objRow.Add "医保大类：" & nod.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & gstrUserName
    objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Private Sub cmdRequery_Click()
    Dim strInPut As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln病种 As Boolean
    
    If MsgBox("本操作可能会花比较长的时间，是否继续？" & vbCrLf & vbCrLf & "另外注意，本操作只更新医保项目明细，而不包括对应关系。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox("本次仅下载病种吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        bln病种 = False
    Else
        bln病种 = True
    End If
    
    MousePointer = vbHourglass
    zlCommFun.ShowFlash "医保项目选择（正在读取从文件或网络读取保险项目明细，该过程是一个较长的过程，请等待......）"
    
    If InitInfor_重庆渝北.经办机构代码 = "" Then
        ShowMsgbox "经办机构代码不能为空!"
        Exit Sub
    End If
    
    picCmd.Enabled = False
    tvwClass.Enabled = False
    '检查本次是全量更新还是增量更新(修改量大)
    If Not bln病种 Then
        Call Get服务目录
    End If
    
    '门诊病种下载:
    Call Get病种目录
    
    '住院病种
    Call Get住院病种目录
    
    Me.Caption = "正在填充明细数据，请稍后..."
    '重新装入明细
    Call FillList
    zlCommFun.StopFlash
    MousePointer = vbDefault
    Me.Caption = "医保项目选择"
    picCmd.Enabled = True
    tvwClass.Enabled = True
End Sub

Private Function Get服务目录() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:导出服务项目目录
    '返回:导出成功True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim tmpStrut As Struct
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim StrSQL As String
    Dim strHead As String
    Dim strArr
    Dim strText As String
    Dim strTemp As String
    Dim i As Long
    Dim lngRow As Long
    
    tmpStrut.strErrMsg = Space(5000)
    strFile = App.Path & "\医保"
    Get服务目录 = False

    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(strFile) Then
        '不存在文件夹，需创建
        objFile.CreateFolder strFile
    End If
    
    strFile = strFile & "\医疗服务目录.Txt"
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile
    End If
    
    
    ExportKA02K3 InitInfor_重庆渝北.经办机构代码, strFile, tmpStrut
    
    If tmpStrut.lngAppCode < 0 Then
        ShowMsgbox tmpStrut.strErrMsg
        Exit Function
    End If
    
    strHead = "ZL_医保服务项目目录_UPDATE("
    Set objText = objFile.OpenTextFile(strFile)
    lngRow = 1
    Do While Not objText.AtEndOfStream
        strTemp = Trim(objText.ReadLine)
        strArr = Split(strTemp, vbTab)
        StrSQL = ""
        For i = 0 To UBound(strArr)
            If InStr(1, strArr(i), "'") <> 0 Then
                strTemp = Replace(strArr(i), "'", "‘")
                strArr(i) = strTemp
            End If
            If Trim(strArr(i)) = "" Then
                    StrSQL = StrSQL & ",null"
            Else
                If i < 22 Or i >= 41 And i <= 50 Or i >= 52 Then
                    If i >= 43 And i <= 48 Then
                        StrSQL = StrSQL & ",'" & UCase(strArr(i)) & "'"
                    Else
                        StrSQL = StrSQL & ",'" & strArr(i) & "'"
                    End If
                ElseIf i = 51 Then
                    '修改时间
                    StrSQL = StrSQL & ",to_date('" & Format(strArr(i), "yyyy-mm-dd") & "','yyyy-mm-dd')"
                Else
                    StrSQL = StrSQL & "," & Val(strArr(i))
                End If
            End If
        Next
        If StrSQL <> "" Then
            StrSQL = strHead & Mid(StrSQL, 2) & ")"
            gcnOracle_CQYB.Execute StrSQL, , adCmdStoredProc
            DoEvents
        End If
        
        Me.Caption = "医保服务项目下载:已经更新了 " & lngRow & "条记录"
        lngRow = lngRow + 1
    Loop
    objText.Close
    Get服务目录 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function Get病种目录() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:导出病种目录
    '返回:导出成功True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim tmpStrut As Struct
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim StrSQL As String
    Dim strSQL1 As String
    Dim strHead As String
    Dim strArr
    Dim strText As String
    Dim strTemp As String
    Dim i As Long
    Dim lngRow As Long
    
    tmpStrut.strErrMsg = Space(5000)
    strFile = App.Path & "\医保"
    Get病种目录 = False
    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(strFile) Then
        '不存在文件夹，需创建
        objFile.CreateFolder strFile
    End If
    
    strFile = strFile & "\医疗病种目录.Txt"
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile
    End If
    
    
    ExportKA06K1 InitInfor_重庆渝北.经办机构代码, strFile, tmpStrut
    
    If tmpStrut.lngAppCode < 0 Then
        ShowMsgbox tmpStrut.strErrMsg
        Exit Function
    End If
    
    strHead = "ZL_医保病种目录_UPDATE("
    Set objText = objFile.OpenTextFile(strFile)
    lngRow = 1
    Do While Not objText.AtEndOfStream
        strTemp = Trim(objText.ReadLine)
        strArr = Split(strTemp, vbTab)
        StrSQL = ""
        For i = 0 To UBound(strArr)
            If Trim(strArr(i)) = "" Then
                    StrSQL = StrSQL & ",null"
            Else
                    StrSQL = StrSQL & ",'" & strArr(i) & "'"
            End If
            If i >= 5 Then Exit For
        Next
        If StrSQL <> "" Then
            StrSQL = strHead & "1," & Mid(StrSQL, 2) & ")"
            gcnOracle_CQYB.Execute StrSQL, , adCmdStoredProc
        End If
        Me.Caption = "门诊病种目录下载:已经更新了 " & lngRow & "条记录"
        lngRow = lngRow + 1
    Loop
    
    objText.Close
    Get病种目录 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Function Get住院病种目录() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:导出住院病种目录
    '返回:导出成功True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    Dim tmpStrut As Struct
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    Dim StrSQL As String
    Dim strSQL1 As String
    Dim strHead As String
    Dim strArr
    Dim strText As String
    Dim strTemp As String
    Dim i As Long
    Dim lngRow As Long
    
    tmpStrut.strErrMsg = Space(5000)
    strFile = App.Path & "\医保"
    Get住院病种目录 = False
    Err = 0
    On Error GoTo ErrHand:
    If Not objFile.FolderExists(strFile) Then
        '不存在文件夹，需创建
        objFile.CreateFolder strFile
    End If
    
    strFile = strFile & "\住院病种目录.Txt"
    If Not objFile.FileExists(strFile) Then
        objFile.CreateTextFile strFile
    End If
    
    
    ExportKA03K1 InitInfor_重庆渝北.经办机构代码, strFile, tmpStrut
    
    If tmpStrut.lngAppCode < 0 Then
        ShowMsgbox tmpStrut.strErrMsg
        Exit Function
    End If
    
    strHead = "ZL_医保病种目录_UPDATE("
    Set objText = objFile.OpenTextFile(strFile)
    lngRow = 1
    
    '1   string  20      病种编码
    '2   string  100     病种名称
    '3   string  6       明细标识，见代码表
    '4   string  10      助记码
    
    Do While Not objText.AtEndOfStream
        strTemp = Trim(objText.ReadLine)
        strArr = Split(strTemp, vbTab)
        
        If strTemp <> "" Then
            'ZL_医保病种目录_UPDATE
            StrSQL = "ZL_医保病种目录_UPDATE("
            '  性质_In         In 医保病种目录.性质%Type,
            StrSQL = StrSQL & "" & 2 & ","
            '  编码_In         In 医保病种目录.编码%Type,
            StrSQL = StrSQL & "'" & strArr(0) & "',"
            '  名称_In         In 医保病种目录.名称%Type,
            StrSQL = StrSQL & "'" & strArr(1) & "',"
            '  支付类别_In     In 医保病种目录.支付类别%Type,
            StrSQL = StrSQL & "'" & strArr(2) & "',"
            '  助记码_In       In 医保病种目录.助记码%Type,
            StrSQL = StrSQL & "'" & strArr(3) & "',"
            '  病种结算办法_In In 医保病种目录.病种结算办法%Type,
            StrSQL = StrSQL & "NULL,"
            '  经办构构代码_In In 医保病种目录.经办构构代码%Type
            StrSQL = StrSQL & "NULL)"
            gcnOracle_CQYB.Execute StrSQL, , adCmdStoredProc
        End If
        Me.Caption = "住院病种目录下载:已经更新了 " & lngRow & "条记录"
        lngRow = lngRow + 1
    Loop
    
    objText.Close
    Get住院病种目录 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function
Private Sub cmd新增_Click()
    Dim blnReturn As Boolean
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    
    blnReturn = frm保险项目编辑.EditCard(Me, "", Mid(Me.tvwClass.SelectedItem.Key, 2))
    If blnReturn = False Then Exit Sub
     Call FillList
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If Loadtree = False Then
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = tvwClass.Width
    
    On Error Resume Next
    
    tvwClass.Left = 0: tvwClass.Top = lblClass.Top + lblClass.Height
    tvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = tvwClass.Top
    picSplit.Left = tvwClass.Left + tvwClass.Width
    picSplit.Height = tvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If tvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
        
    With mshGrid
        .Top = tvwClass.Top
        .Left = lblDetail.Left
        .Width = lblDetail.Width
        .Height = tvwClass.Height
    End With
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
    cmdRequery.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshgrid_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvwClass.Width + x < 1000 Or mshGrid.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        mshGrid.Left = mshGrid.Left + x
        mshGrid.Width = mshGrid.Width - x
    End If
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
End Sub





