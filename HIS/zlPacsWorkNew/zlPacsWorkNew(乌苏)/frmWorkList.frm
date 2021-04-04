VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmWorklist 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraWorkList 
      Height          =   6330
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   11310
      Begin VB.Frame frmResultFilter 
         Caption         =   "查询结束条件"
         Height          =   1215
         Left            =   3120
         TabIndex        =   15
         ToolTipText     =   "检查执行到哪一步之后，Worklist中不再能提取到检查信息"
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton optResultFilter 
            Caption         =   "检查完成"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   17
            ToolTipText     =   "检查报告完成后，Worklist查询不再返回该检查"
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton optResultFilter 
            Caption         =   "图像采集（默认）"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   16
            ToolTipText     =   "接收到设备发回的图像后，Worklist查询不再返回该检查"
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "部位对码：多部位设置"
         Height          =   1215
         Left            =   5760
         TabIndex        =   9
         Top             =   245
         Width           =   4215
         Begin VB.OptionButton optMultiParts 
            Caption         =   "多序列"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   14
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtPatsSpliter 
            Height          =   300
            Left            =   960
            TabIndex        =   13
            Top             =   817
            Width           =   855
         End
         Begin VB.OptionButton optMultiParts 
            Caption         =   "分隔符"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optMultiParts 
            Caption         =   "多记录"
            Height          =   255
            Index           =   2
            Left            =   2160
            TabIndex        =   11
            Top             =   840
            Width           =   855
         End
         Begin VB.OptionButton optMultiParts 
            Caption         =   "无"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdBodypartCode 
         Caption         =   "部位对码"
         Height          =   350
         Left            =   10080
         TabIndex        =   8
         Top             =   630
         Width           =   1100
      End
      Begin VB.ComboBox cboMatchOther 
         Height          =   300
         ItemData        =   "frmWorkList.frx":0000
         Left            =   1140
         List            =   "frmWorkList.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   247
         Width           =   1665
      End
      Begin VB.TextBox txtSearch 
         Height          =   300
         Left            =   900
         MaxLength       =   4
         TabIndex        =   3
         Top             =   680
         Width           =   435
      End
      Begin VB.CheckBox chkForceResult 
         Caption         =   "使用强制结果"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1380
      End
      Begin VB.CommandButton cmdResetWLResult 
         Caption         =   "恢复默认值"
         Height          =   350
         Left            =   10080
         TabIndex        =   1
         Top             =   247
         Width           =   1100
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   4560
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   11085
         _cx             =   19553
         _cy             =   8043
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "辅助匹配(&A)"
         Height          =   180
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "该参数针对[数据库项目]按""病人标识号""/""检查号""匹配有效"
         Top             =   307
         Width           =   990
      End
      Begin VB.Label LblSe 
         Caption         =   "检索最近      天的申请"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   733
         Width           =   1980
      End
   End
End
Attribute VB_Name = "frmWorklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ColReturn
    ColID = 0
    Col服务ID
    Col标记
    Col上级ID
    Col中文标题
    Col英文标题
    Col数据值
    Col是否嵌套数据
    Col是否递增
    Col值类型
    Col选中
    Col元素类型
    Col强制结果值
    Col默认值
    Col默认选中
    Col默认强制结果值
End Enum
Private mlngSrvID As Long
Private Const mstrDBItem As String = "|[CallingAET]|[首次日期]|[首次时间]|[影像类别]|[执行间]|[执行过程]|[医嘱ID]|[发送号]|[医嘱ID]_[发送号]|[检查号]|[标识号]|[英文名]|[性别]|[年龄]|[出生日期]|[中文名]|[检查设备]|[检查部位]|[体重]|[附加主述]|[对码部位名称]|[对码部位代码]"

Public Sub ShowRefresh(ByVal SrvID As Long)
    mlngSrvID = SrvID
    If mlngSrvID = 0 Then
        fraWorkList.Caption = "上方列表中所选服务尚未保存，不能进行设置！"
        fraWorkList.Enabled = False
    Else
        fraWorkList.Caption = ""
        fraWorkList.Enabled = True
    End If
    Call RefreshPara
    Call ReFreshReturnData
End Sub

Private Sub cmdBodypartCode_Click()
    '打开部位对码设置窗口
    frmMWLBodypartCode.zlSohwMe Me, mlngSrvID
End Sub

Private Sub cmdResetWLResult_Click()
Dim i As Integer
    With vfgList
        For i = 1 To .Rows - 1
            .TextMatrix(i, Col数据值) = .TextMatrix(i, Col默认值)
            .TextMatrix(i, Col强制结果值) = .TextMatrix(i, Col默认强制结果值)
            .TextMatrix(i, Col是否递增) = ""
            .TextMatrix(i, Col选中) = .TextMatrix(i, Col默认选中)
        Next
    End With
End Sub

Public Sub SavePara()
    Dim i As Long
    Dim iMatch As Integer
    
    On Error GoTo errHandle
    zlCommFun.ShowFlash "正在保存数据", Me
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'WorkList过滤方式','" & NeedNo(cboMatchOther.Text) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存WorkList按设备过滤")
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'WorkList使用强制结果','" & chkForceResult.value & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存WorkList使用强制结果")
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'WorkList检索天数','" & txtSearch.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存WorkList检索天数")
    
    If optMultiParts(1).value = True Then
        iMatch = 1
    ElseIf optMultiParts(2).value = True Then
        iMatch = 2
    ElseIf optMultiParts(3).value = True Then
        iMatch = 3
    Else
        iMatch = 0
    End If
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'Worklist多部位方式','" & iMatch & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存Worklist多部位方式")
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'Worklist多部位分隔符','" & txtPatsSpliter.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存Worklist多部位分隔符")
    
    If optResultFilter(1).value = True Then
        iMatch = 1
    Else
        iMatch = 0
    End If
    gstrSQL = "Zl_影像DICOM服务参数_SAVE(" & mlngSrvID & ",'Worklist查询结束条件','" & iMatch & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存Worklist查询结束条件")
    
    With vfgList
        For i = 1 To .Rows - 1
            gstrSQL = "Zl_影像MWL结果集_UPDATE(" & .TextMatrix(i, ColID) & ",'" & .TextMatrix(i, Col数据值) & "'," & IIf(.TextMatrix(i, Col是否递增) = "", 0, 1) & "," & IIf(.TextMatrix(i, Col选中) = "", 0, 1) & ",'" & .TextMatrix(i, Col强制结果值) & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新MWL结果集")
        Next
    End With
    zlCommFun.StopFlash
   Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Call InitvfgList
End Sub
Private Sub RefreshPara()
    Dim rsTemp As New ADODB.Recordset, i As Integer
    
    On Error GoTo err
    gstrSQL = "select 服务ID,参数名称 ,参数值 from 影像DICOM服务参数 where 服务ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", mlngSrvID)
    chkForceResult.value = False
    txtSearch.Text = 3
    cboMatchOther.ListIndex = 0
    optMultiParts(0).value = True
    txtPatsSpliter.Text = ""
    optResultFilter(0).value = True '默认使用图像采集作为Worklist的查询结束条件
    
    Do Until rsTemp.EOF
        Select Case rsTemp!参数名称
            Case "WorkList过滤方式"
                Call SeekIndexWithNo(cboMatchOther, Nvl(rsTemp!参数值, 0), True)
            Case "WorkList使用强制结果"
                chkForceResult.value = Nvl(rsTemp!参数值)
            Case "WorkList检索天数"
                txtSearch.Text = Nvl(rsTemp!参数值)
            Case "Worklist多部位方式"   '0-无，1-分隔符，2-多记录，3-多序列
                If Nvl(rsTemp!参数值, 0) = 1 Then
                    optMultiParts(1).value = True
                    txtPatsSpliter.Enabled = True
                    txtPatsSpliter.BackColor = &H80000005
                ElseIf Nvl(rsTemp!参数值, 0) = 2 Then
                    optMultiParts(2).value = True
                    txtPatsSpliter.Enabled = False
                    txtPatsSpliter.BackColor = &H8000000B
                ElseIf Nvl(rsTemp!参数值, 0) = 3 Then
                    optMultiParts(3).value = True
                    txtPatsSpliter.Enabled = False
                    txtPatsSpliter.BackColor = &H8000000B
                Else
                    optMultiParts(0).value = True
                    txtPatsSpliter.Enabled = False
                    txtPatsSpliter.BackColor = &H8000000B
                End If
            Case "Worklist多部位分隔符"
                txtPatsSpliter.Text = Nvl(rsTemp!参数值)
            Case "Worklist查询结束条件"
                If Nvl(rsTemp!参数值, 0) = 1 Then
                    optResultFilter(1).value = True
                Else
                    optResultFilter(0).value = True
                End If
        End Select
        rsTemp.MoveNext
    Loop
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ReFreshReturnData()
'刷新Worklist的结果数据集

    Dim rsTemp As New ADODB.Recordset
    Dim rsQuery As New ADODB.Recordset  '复制一个rsTemp的数据集，用来判断是否有上级ID。
    
    On Error GoTo err
    
    InitvfgList
    gstrSQL = "select ID,服务ID,组号,元素号,上级ID,中文标题,英文标题,数据值," & _
                    "是否嵌套数据,是否递增,值类型,选中,元素类型,强制结果值,默认值,默认选中,默认强制结果值" & _
                    " from 影像MWL结果集 WHERE 服务ID=[1] Order by 组号,元素号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取参数", mlngSrvID)
    
    Set rsQuery = rsTemp.Clone
    
    rsTemp.Filter = "上级ID = NULL"
    Call AddMWLDataset(rsTemp, rsQuery, "")
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AddMWLDataset(rsTemp As ADODB.Recordset, rsQuery As ADODB.Recordset, ByVal strPrefix As String)
    
    On Error GoTo err
    
    If rsTemp.EOF = False Then
        If Not IsNull(rsTemp!上级ID) Then strPrefix = strPrefix & ">"
    End If
    
    While rsTemp.EOF = False
        '添加结果集
        With vfgList
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, ColID) = rsTemp!ID
            .TextMatrix(.Rows - 1, Col服务ID) = rsTemp!服务ID
            .TextMatrix(.Rows - 1, Col标记) = strPrefix & rsTemp!组号 & "," & rsTemp!元素号
            .TextMatrix(.Rows - 1, Col上级ID) = Nvl(rsTemp!上级ID)
            .TextMatrix(.Rows - 1, Col中文标题) = strPrefix & rsTemp!中文标题
            .TextMatrix(.Rows - 1, Col英文标题) = strPrefix & rsTemp!英文标题
            .TextMatrix(.Rows - 1, Col数据值) = Nvl(rsTemp!数据值)
            .TextMatrix(.Rows - 1, Col是否嵌套数据) = IIf(Nvl(rsTemp!是否递增, 0) = 1, "√", "")
            .TextMatrix(.Rows - 1, Col是否递增) = IIf(rsTemp!是否递增 = 1, "√", "")
            .TextMatrix(.Rows - 1, Col值类型) = rsTemp!值类型
            .TextMatrix(.Rows - 1, Col选中) = IIf(rsTemp!选中 = 1, "√", "")
            .TextMatrix(.Rows - 1, Col元素类型) = rsTemp!元素类型
            .TextMatrix(.Rows - 1, Col强制结果值) = Nvl(rsTemp!强制结果值)
            .TextMatrix(.Rows - 1, Col默认值) = Nvl(rsTemp!默认值)
            .TextMatrix(.Rows - 1, Col默认选中) = IIf(Nvl(rsTemp!默认选中, 0) = 1, "√", "")
            .TextMatrix(.Rows - 1, Col默认强制结果值) = Nvl(rsTemp!默认强制结果值)
        End With
        
        '查找是否有其他数据集的上级ID=当前id,如果有，则添加这些数据集
        rsQuery.Filter = "上级ID=" & rsTemp!ID
        If rsQuery.RecordCount > 0 Then
            '查到数据集，需要处理这些嵌套的数据集
            Dim rsClone As New ADODB.Recordset
            Set rsClone = rsTemp.Clone
            rsClone.Filter = "上级ID=" & rsTemp!ID
            Call AddMWLDataset(rsClone, rsQuery, strPrefix)
        End If
        rsTemp.MoveNext
    Wend
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitvfgList()
    
    On Error GoTo err
    
    With vfgList
        .Clear
        .FixedRows = 1
        .Rows = 1
        .Cols = 16

        
        .ColWidth(ColID) = 0
        .ColWidth(Col服务ID) = 0
        .ColWidth(Col标记) = 1400
        .ColWidth(Col上级ID) = 0
        .ColWidth(Col中文标题) = 2000
        .ColWidth(Col英文标题) = 3100
        .ColWidth(Col数据值) = 1700
        .ColWidth(Col是否嵌套数据) = 0
        .ColWidth(Col是否递增) = 600
        .ColWidth(Col值类型) = 0
        .ColWidth(Col选中) = 600
        .ColWidth(Col元素类型) = 0
        .ColWidth(Col强制结果值) = 1200
        .ColWidth(Col默认值) = 0
        .ColWidth(Col默认选中) = 0
        .ColWidth(Col默认强制结果值) = 0

        .TextMatrix(0, ColID) = "ID"
        .TextMatrix(0, Col服务ID) = "服务ID"
        .TextMatrix(0, Col标记) = "标记"
        .TextMatrix(0, Col上级ID) = "上级ID"
        .TextMatrix(0, Col中文标题) = "中文标题"
        .TextMatrix(0, Col英文标题) = "英文标题"
        .TextMatrix(0, Col数据值) = "数据值"
        .TextMatrix(0, Col是否嵌套数据) = "嵌套"
        .TextMatrix(0, Col是否递增) = "递增"
        .TextMatrix(0, Col值类型) = "值类型"
        .TextMatrix(0, Col选中) = "使用"
        .TextMatrix(0, Col元素类型) = "元素类型"
        .TextMatrix(0, Col强制结果值) = "强制结果"
        .TextMatrix(0, Col默认值) = "默认值"
        .TextMatrix(0, Col默认选中) = "默认选中"
        .TextMatrix(0, Col默认强制结果值) = "默认强制结果值"
        
        .ColAlignment(ColID) = flexAlignLeftCenter
        .ColAlignment(Col服务ID) = flexAlignLeftCenter
        .ColAlignment(Col标记) = flexAlignLeftCenter
        .ColAlignment(Col上级ID) = flexAlignLeftCenter
        .ColAlignment(Col中文标题) = flexAlignLeftCenter
        .ColAlignment(Col英文标题) = flexAlignLeftCenter
        .ColAlignment(Col数据值) = flexAlignLeftCenter
        .ColAlignment(Col是否嵌套数据) = flexAlignLeftCenter
        .ColAlignment(Col是否递增) = flexAlignLeftCenter
        .ColAlignment(Col值类型) = flexAlignLeftCenter
        .ColAlignment(Col选中) = flexAlignLeftCenter
        .ColAlignment(Col元素类型) = flexAlignLeftCenter
        .ColAlignment(Col强制结果值) = flexAlignLeftCenter
        .ColAlignment(Col默认值) = flexAlignLeftCenter
        .ColAlignment(Col默认选中) = flexAlignLeftCenter
        .ColAlignment(Col默认强制结果值) = flexAlignLeftCenter
        
        .Editable = flexEDKbdMouse
        .ComboSearch = flexCmbSearchNone
        .ColComboList(Col数据值) = mstrDBItem
    End With
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optMultiParts_Click(Index As Integer)
    If Index = 1 Then
        txtPatsSpliter.Enabled = True
        txtPatsSpliter.BackColor = &H80000005
    Else
        txtPatsSpliter.Enabled = False
        txtPatsSpliter.BackColor = &H8000000B
    End If
End Sub

Private Sub vfgList_Click()
    With vfgList
        If .Col = Col是否递增 Or .Col = Col选中 Then
            .Editable = flexEDNone
        ElseIf .Col = Col标记 Or .Col = Col中文标题 Or .Col = Col英文标题 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .Col = Col是否递增 Or .Col = Col选中 Then
            If .TextMatrix(.Row, .Col) = "" Then
                .TextMatrix(.Row, .Col) = "√"
            Else
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
End Sub
Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not (Col = Col强制结果值 Or Col = Col数据值) Then
        KeyAscii = 0
    End If
End Sub

Private Sub vfgList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    On Error GoTo err
    
    '验证数据值的输入
    If Col = Col数据值 Then
        '在字符串中，不允许出现'号；[]数量要匹配；[]中的内容是数据库字段
        Dim strTmp As String, strValue As String
        Dim intResult As Integer
        Dim strDBItems() As String
        Dim i As Integer
        Dim blnDBMatch As Boolean
        
        strTmp = vfgList.EditText
        strDBItems = Split(mstrDBItem, "|")
        
        If InStr(strTmp, "'") > 0 Then
            Cancel = True
            intResult = 1
        ElseIf InStr(strTmp, "[") <> 0 Then
            Do Until InStr(strTmp, "[") = 0
                If InStr(strTmp, "]") = 0 Or InStr(strTmp, "]") < InStr(strTmp, "[") Then
                    Cancel = True
                    intResult = 2
                    Exit Do
                End If
                blnDBMatch = False
                strValue = Mid(strTmp, InStr(strTmp, "["), InStr(strTmp, "]") - InStr(strTmp, "[") + 1)
                For i = 1 To UBound(strDBItems)
                    If strDBItems(i) = strValue Then
                        blnDBMatch = True
                        Exit For
                    End If
                Next i
                
                If blnDBMatch = False Then
                    Cancel = True
                    intResult = 3
                    Exit Do
                End If
                strTmp = Mid(strTmp, InStr(strTmp, "]") + 1)
            Loop
        End If
        
        If Cancel Then
            If intResult = 1 Then
                MsgBoxD Me, "输入参数不合法,不能使用符号“'”作为连接符。", vbInformation, gstrSysName
            ElseIf intResult = 2 Then
                MsgBoxD Me, "输入参数不合法,“[”和“]”的数目不匹配。", vbInformation, gstrSysName
            Else
                MsgBoxD Me, "输入参数不合法,“[]”中的内容不是数据库项目。", vbInformation, gstrSysName
            End If
        End If
    End If
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
