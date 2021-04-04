Attribute VB_Name = "mdlPubMethods"
'---------------------------------------------------------------------------------------
'创    建:蔡青松
'创建时间:2019-12-06
'模块功能:公共方法
'---------------------------------------------------------------------------------------

Option Explicit

Public Enum EKey
    vbKeyUp = 38
    vbKeyDown = 40
End Enum

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-04-08
'功    能:  截断长度超过4000的字符串，并返回数组。为防止字符串长度超过4000出入数据库报错
'入    参:
'           strIn       需要处理的字符串
'           strSplit    字符串分隔符
'出    参:
'返    回:  截断之后的数组
'调整影响:
'---------------------------------------------------------------------------------------
Public Function TruncatedExtraLongStr(ByVal strIn As String, ByVal strSplit As String) As Variant
          Dim strReturn() As String
          Dim blnRdm As Boolean
          Dim lngLen As Long
          Dim lngS As Long

1         On Error GoTo TruncatedExtraLongStr_Error

2         lngLen = 1500
3         Do While Len(strIn) > lngLen  '由于中午字符占用两个字节，所以2000个中文字符传入到数据库就是4000，所以使用1500
4             lngS = InStr(lngLen, strIn, strSplit)
5             If lngS = 0 Then
6                 lngS = InStrRev(strIn, strSplit)
7             End If
8             If blnRdm Then
9                 ReDim Preserve strReturn(UBound(strReturn) + 1)
10            Else
11                ReDim strReturn(0)
12                blnRdm = True
13            End If
14            If lngS > 0 Then
15                strReturn(UBound(strReturn)) = Left(strIn, lngS - 1)
16                strIn = Mid(strIn, lngS + Len(strSplit))
17            End If
18        Loop
19        If blnRdm Then
20            ReDim Preserve strReturn(UBound(strReturn) + 1)
21        Else
22            ReDim strReturn(0)
23        End If
24        strReturn(UBound(strReturn)) = strIn
25        TruncatedExtraLongStr = strReturn



26        Exit Function
TruncatedExtraLongStr_Error:
27        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "执行(TruncatedExtraLongStr)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
28        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/7/16
'功    能:将记录集绑定到VSF中
'入    参:
'           objVSF          需要绑定数据的VSF
'           rsTmp           绑定的记录集
'           intBdingType    绑定类型，0=一次性绑定所有数据，1=逐行遍历绑定，2=只绑定一行记录，用于单行刷新，3=按照分类进行绑定
'           [objImgList     需要绑定图标的图标对象]
'           [strTitle       标题，格式:列名1,列宽,对其方式;列名2,列宽,对其方式]
'           [strID          当intBdingType=2时，该参数有效，并且记录集中必须有ID字段，否则会出错]
'           [strGroup       当intBdingType=3时，该参数有效，按照strGroup的值(记录集中的字段,通常该字段前需要一个能显示的列)，进行树形绑定]
'           [blnSort        当intBdingType=3事，该参数有效，是否按照strGroup参数进行排序]
'           [lngFontSize    VSF控件字体，若不指定，则使用全局字体]
'出    参:
'           [lngRowFind     当intBdingType=2时，该参数有效，返回被刷新的行的行号]
'返    回:
'---------------------------------------------------------------------------------------
Public Function SetDataToVSF(ByVal objVSF As Object, ByVal rsTmp As ADODB.Recordset, _
                                Optional ByVal intBdingType As Integer, Optional objImgList As Object, _
                                Optional ByVal strTitle As String, Optional ByVal strID As String, _
                                Optional ByRef lngRowFind As Long, Optional ByVal strGroup As String, _
                                Optional ByVal blnSort As Boolean = True, Optional ByVal lngFontSize As Long) As Boolean
          Dim CollIco As New Collection
          Dim lngRow As Long
          Dim lngCol As Long
          Dim strFL As String
          Dim i As Integer

          '如果没有传入列明，则默认从记录集中获取
1         On Error GoTo SetDataToVSF_Error

2         If intBdingType = 1 Or intBdingType = 3 Then  '当intBdingType=0时，使用Set objVSF.DataSource = rsTmp方式，会丢失之前设置的VSF风格，所以当intBdingType=0时，不做设置
3             If strTitle = "" Then
4                 For i = 0 To rsTmp.Fields.Count - 1
5                     strTitle = strTitle & ";" & rsTmp.Fields(i).Name & "," & Len(rsTmp.Fields(i).Name) * 400 & "," & flexAlignLeftCenter
6                 Next
7                 If strTitle <> "" Then strTitle = Mid(strTitle, 2)
8             End If
9             Call vfgSetting(0, objVSF, strTitle, , lngFontSize)  '设置VSF风格
10        End If

          '先将图标对象加载到集合中
11        If Not objImgList Is Nothing Then
12            For i = 1 To objImgList.ListImages.Count
13                CollIco.Add objImgList.ListImages(i).ExtractIcon, objImgList.ListImages(i).Key
14            Next
15        End If


          '绑定数据
16        If intBdingType = 0 Then
17            Call vfgSetting(0, objVSF, , , lngFontSize) '设置VSF风格
              '整体绑定
18            With objVSF
                  '数据
19                Set .DataSource = rsTmp
                  '固定行文字居中
20                If .FixedRows > 0 And .Cols > 0 Then
21                    .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
                      '                .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignLeftCenter
22                End If
                  '内容行靠左居中对齐
23                If .Rows > 1 And .Cols > 0 Then
24                    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
25                End If
                  '设置各列Key值
26                For lngCol = 0 To .Cols - 1
27                    .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
28                Next

                  '设置行列图标
29                If Not objImgList Is Nothing Then
30                    For lngRow = 0 To .Rows - 1
31                        For lngCol = 0 To .Cols - 1
32                            Call SetVSFIco(objVSF, lngRow, lngCol, CollIco, Trim(.TextMatrix(lngRow, lngCol)))
33                        Next
34                    Next
35                End If
36            End With

37        ElseIf intBdingType = 1 Then
              '逐行绑定
38            With objVSF
                  '设置固定行列图标
39                If Not objImgList Is Nothing Then
40                    For lngRow = 0 To .FixedRows - 1
41                        For lngCol = 0 To .Cols - 1
42                            Call SetVSFIco(objVSF, lngRow, lngCol, CollIco, Trim(.TextMatrix(lngRow, lngCol)))
43                        Next
44                    Next
45                End If

46                Do Until rsTmp.EOF
47                    For i = 0 To rsTmp.Fields.Count - 1
48                        If .ColIndex(rsTmp.Fields(i).Name) >= 0 Then
49                            .TextMatrix(.Rows - 1, .ColIndex(rsTmp.Fields(i).Name)) = CStr("" & rsTmp.Fields(i).value)
50                            If Not objImgList Is Nothing Then
51                                Call SetVSFIco(objVSF, .Rows - 1, .ColIndex(rsTmp.Fields(i).Name), CollIco, Trim(CStr("" & rsTmp.Fields(i).value)))  '设置行列图标
52                            End If
53                        End If
54                    Next
55                    .Rows = .Rows + 1
56                    rsTmp.MoveNext
57                Loop
58                If .Rows > .FixedRows + 1 Then .Rows = .Rows - 1
59            End With
60        ElseIf intBdingType = 2 Then
              '单行绑定
61            rsTmp.Filter = "ID='" & strID & "'"
62            If rsTmp.RecordCount > 0 Then
63                With objVSF
64                    For lngRow = .FixedRows To .Rows - 1
65                        If .TextMatrix(lngRow, .ColIndex("ID")) = strID Then
66                            For i = 0 To rsTmp.Fields.Count - 1
67                                If .ColIndex(rsTmp.Fields(i).Name) >= 0 Then
68                                    .TextMatrix(lngRow, .ColIndex(rsTmp.Fields(i).Name)) = CStr("" & rsTmp.Fields(i).value)
69                                    If Not objImgList Is Nothing Then
70                                        Call SetVSFIco(objVSF, lngRow, .ColIndex(rsTmp.Fields(i).Name), CollIco, Trim(CStr("" & rsTmp.Fields(i).value)))  '设置行列图标
71                                    End If
72                                End If
73                            Next
74                            lngRowFind = lngRow
75                            Exit For
76                        End If

77                    Next
78                    If lngRowFind < .FixedRows Then
79                        .Rows = .Rows + 1
80                        For i = 0 To rsTmp.Fields.Count - 1
81                            If .ColIndex(rsTmp.Fields(i).Name) >= 0 Then
82                                .TextMatrix(lngRow, .ColIndex(rsTmp.Fields(i).Name)) = CStr("" & rsTmp.Fields(i).value)
83                                If Not objImgList Is Nothing Then
84                                    Call SetVSFIco(objVSF, lngRow, .ColIndex(rsTmp.Fields(i).Name), CollIco, Trim(CStr("" & rsTmp.Fields(i).value)))  '设置行列图标
85                                End If
86                            End If
87                        Next
88                        lngRowFind = .Rows - 1
89                    End If
90                End With
91            End If
92        ElseIf intBdingType = 3 Then
              '树形绑定
93            If blnSort Then rsTmp.Sort = strGroup
94            With objVSF
95                .OutlineBar = flexOutlineBarComplete
96                .OutlineCol = 0
97                .SubtotalPosition = flexSTAbove
98                .Rows = 1

                  '设置固定行列图标
99                For lngRow = 0 To .FixedRows - 1
100                   For lngCol = 0 To .Cols - 1
101                       Call SetVSFIco(objVSF, lngRow, lngCol, CollIco, Trim(.TextMatrix(lngRow, lngCol)))
102                   Next
103               Next

104               Do While Not rsTmp.EOF
105                   If InStr("<SP>" & strFL & "<SP>", "<SP>" & rsTmp(strGroup) & "<SP>") <= 0 Then
106                       .Rows = .Rows + 1
107                       .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = rsTmp(strGroup) & ""

                          '合并
108                       .MergeRow(.Rows - 1) = True
109                       .MergeCellsFixed = flexMergeRestrictRows

                          '加粗
110                       .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True

                          '缩进
111                       .IsSubtotal(.Rows - 1) = True

112                       strFL = strFL & "<SP>" & rsTmp(strGroup)
113                   End If
114                   .Rows = .Rows + 1
115                   .RowOutlineLevel(.Rows - 1) = 1
116                   For i = 0 To rsTmp.Fields.Count - 1
117                       If .ColIndex(rsTmp.Fields(i).Name) >= 0 Then
118                           .TextMatrix(.Rows - 1, .ColIndex(rsTmp.Fields(i).Name)) = CStr("" & rsTmp.Fields(i).value)
119                           If Not objImgList Is Nothing Then
120                               Call SetVSFIco(objVSF, .Rows - 1, .ColIndex(rsTmp.Fields(i).Name), CollIco, Trim(CStr("" & rsTmp.Fields(i).value)))  '设置行列图标
121                           End If
122                       End If
123                   Next

124                   rsTmp.MoveNext
125               Loop
126               .ColHidden(.ColIndex(strGroup)) = True
127           End With
128       End If

129       With objVSF
              '序号列设置不同的背景色
130           If .Rows - 1 > 0 Then
131               If .ColIndex("序号") >= 0 Then
132                   .Cell(flexcpBackColor, .FixedRows, .ColIndex("序号"), .Rows - 1, .ColIndex("序号")) = 16772055
133               End If
134           End If
135       End With

136       SetDataToVSF = True

137       Set CollIco = Nothing



138       Exit Function
SetDataToVSF_Error:
139       Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "执行(SetDataToVSF)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
140       Err.Clear

End Function

Private Sub SetVSFIco(ByVal objVSF As Object, ByVal lngRow As Long, ByVal lngCol As Long, ByVal CollIco As Collection, ByVal strKey As String)
    On Error Resume Next
    With objVSF
        Set .Cell(flexcpPicture, lngRow, lngCol, lngRow, lngCol) = Nothing
        Set .Cell(flexcpPicture, lngRow, lngCol, lngRow, lngCol) = CollIco(strKey)
    End With
    
End Sub



'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/9/29
'功    能:删除非法字符
'入    参:
'       strChar         要处理的字符
'       strInvalidChar  非法字符串，如果为空，则为~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>,否则按传入的字符处理
'出    参:
'返    回:处理之后的字符串
'调整影响:
'---------------------------------------------------------------------------------------
Public Function StringDelInvalidWord(ByVal strChar As String, Optional ByVal strInvalidChar As String) As String
          Dim StrBit As String, i As Integer, strWord As String

1         On Error GoTo StringDelInvalidWord_Error

2         strWord = "`#@$%&|\{}[]?;""'" & Chr(&HD) & Chr(&HA)
          
3         If strInvalidChar <> "" Then strWord = strInvalidChar
4         For i = 1 To Len(strWord)
5             StrBit = StrBit & "," & Asc(Mid(strWord, i, 1))
6         Next
7         strWord = StrBit & ","
          
8         If Len(strChar) > 0 Then
9             For i = 1 To Len(strChar)
10                StrBit = "," & Asc(Mid$(strChar, i, 1)) & ","
11                If InStr(strWord, StrBit) <= 0 Then
12                    StringDelInvalidWord = StringDelInvalidWord & Mid$(strChar, i, 1)
13                End If
14            Next
15        End If


16        Exit Function
StringDelInvalidWord_Error:
17        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "执行(StringDelInvalidWord)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
18        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-03-07
'功    能:  判断当前标本中的项目是否为耐受试验项目
'入    参:
'           lngSampleID     标本ID
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function IsTre(ByVal lngSampleID As Long) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo IsTre_Error
          
2         If VerCompare(gSysInfo.VersionLIS, "10.35.130") = -1 Then
3             Exit Function
4         End If
          
5         strSQL = "Select b.id From 检验报告明细 A, 检验组合项目 B Where a.组合ID = b.id And a.标本ID = [1] And b.是否耐受项目 = 1"
6         Set rsTmp = OpenSQLRecord(Sel_Lis_DB, strSQL, "耐受项目", lngSampleID)
7         If rsTmp.EOF Then
8             IsTre = False
9         Else
10            IsTre = True
11        End If


12        Exit Function
IsTre_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "执行(IsTre)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
14        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/9/4
'功    能:通过上下键改变选择行
'入    参:
'           objVSF      需要改变选择行的VSF表格
'           KeyAscii    按键码，vbkeyUp=方向键上键，vbkeyDown=方向键下键
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Sub UpOrDown(objVSF As VSFlexGrid, ByVal KeyAscii As EKey)
          Dim i As Integer


1         On Error GoTo UpOrDown_Error

2         With objVSF
3             .Tag = .Row
4             If KeyAscii = vbKeyUp Then
                  '按方向键上键
5                 If .RowSel > 1 Then
6                     For i = Val(.Tag) - 1 To 1 Step -1
7                         If .RowHidden(i) = False Then
8                             .Row = i
9                             Exit Sub
10                        End If
11                    Next
12                End If
13            End If
14            If KeyAscii = vbKeyDown Then
                  '按方向键下键
15                If .RowSel < .Rows - 1 Then
16                    For i = Val(.Tag) + 1 To .Rows - 1
17                        If .RowHidden(i) = False Then
18                            .Row = i
19                            Exit Sub
20                        End If
21                    Next
22                End If
23            End If
24        End With




25        Exit Sub
UpOrDown_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "执行(UpOrDown)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
27        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/7/12
'功    能:创建工具栏
'入    参:
'           cbsMain         工具栏对象
'           Buttons         菜单集合,每个元素的格式为 菜单id,标题,是否分组
'           blnLargeIcons   是否大图标
'           Position        菜单位置
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub CbsButtonInit(ByRef cbsMain As CommandBars, Buttons As Collection, _
                         Optional blnLargeIcons As Boolean = False, _
                         Optional Position As XTPBarPosition)
          
          Dim objBar As CommandBar
          Dim objControl As CommandBarControl
          
          Dim strButton As Variant
          Dim varButton As Variant

1         On Error GoTo CbsButtonInit_Error

2         Call CbsSetting(cbsMain)
          '工具栏定义:包括公共部份
          '-----------------------------------------------------
3         cbsMain.ActiveMenuBar.Visible = False
4         Set objBar = cbsMain.Add("工具栏", Position)
5         objBar.ContextMenuPresent = False
6         objBar.ShowTextBelowIcons = False
7         objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
8         cbsMain.Options.LargeIcons = blnLargeIcons  '小图标
9         objBar.Position = Position   '工具栏在顶部

10        For Each strButton In Buttons
11            varButton = Split(strButton, ",")
12            With objBar.Controls
13                Set objControl = .Add(xtpControlButton, Val(varButton(0)), varButton(1))     '固有
14                objControl.Style = xtpButtonIconAndCaption
15                If UCase(varButton(2)) = "TRUE" Then objControl.BeginGroup = True '固有
16            End With
17        Next
18        cbsMain.RecalcLayout



19        Exit Sub
CbsButtonInit_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "执行(CbsButtonInit)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/7/12
'功    能:创建菜单
'入    参:
'           cbsMain         菜单对象
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Function CbsSetting(ByRef cbsMain As CommandBars)


1         On Error GoTo CbsSetting_Error

2         CommandBarsGlobalSettings.App = App
3         CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
4         CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
5         cbsMain.VisualTheme = xtpThemeOffice2003
6         With cbsMain.Options
7             .ShowExpandButtonAlways = False
8             .ToolBarAccelTips = True
9             .AlwaysShowFullMenus = False
              '.UseFadedIcons = True '放在VisualTheme后有效
10            .IconsWithShadow = True '放在VisualTheme后有效
11            .UseDisabledIcons = True
12            .LargeIcons = True
13            .SetIconSize True, 24, 24
14            .SetIconSize False, 16, 16
              
15        End With
16        cbsMain.EnableCustomization False
      '17        cbsMain.Icons = frmPublicIco.imgPublic.Icons
17        cbsMain.ActiveMenuBar.ContextMenuPresent = False    '禁止右键选择工具栏来取消
18        cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap  '禁止移动工具栏





19        Exit Function
CbsSetting_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "执行(CbsSetting)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/3/21
'功    能:文本框全选
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Public Sub selAllText(objTxt As TextBox)
    On Error Resume Next
    With objTxt
        If .Tag = "T" Then
            .ForeColor = &H80000008
            .Text = ""
            .Tag = ""
        End If
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
