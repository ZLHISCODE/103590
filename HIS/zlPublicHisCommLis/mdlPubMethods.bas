Attribute VB_Name = "mdlPubMethods"
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-12-06
'ģ�鹦��:��������
'---------------------------------------------------------------------------------------

Option Explicit

Public Enum EKey
    vbKeyUp = 38
    vbKeyDown = 40
End Enum

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-04-08
'��    ��:  �ضϳ��ȳ���4000���ַ��������������顣Ϊ��ֹ�ַ������ȳ���4000�������ݿⱨ��
'��    ��:
'           strIn       ��Ҫ������ַ���
'           strSplit    �ַ����ָ���
'��    ��:
'��    ��:  �ض�֮�������
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function TruncatedExtraLongStr(ByVal strIn As String, ByVal strSplit As String) As Variant
          Dim strReturn() As String
          Dim blnRdm As Boolean
          Dim lngLen As Long
          Dim lngS As Long

1         On Error GoTo TruncatedExtraLongStr_Error

2         lngLen = 1500
3         Do While Len(strIn) > lngLen  '���������ַ�ռ�������ֽڣ�����2000�������ַ����뵽���ݿ����4000������ʹ��1500
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
27        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "ִ��(TruncatedExtraLongStr)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
28        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/7/16
'��    ��:����¼���󶨵�VSF��
'��    ��:
'           objVSF          ��Ҫ�����ݵ�VSF
'           rsTmp           �󶨵ļ�¼��
'           intBdingType    �����ͣ�0=һ���԰��������ݣ�1=���б����󶨣�2=ֻ��һ�м�¼�����ڵ���ˢ�£�3=���շ�����а�
'           [objImgList     ��Ҫ��ͼ���ͼ�����]
'           [strTitle       ���⣬��ʽ:����1,�п�,���䷽ʽ;����2,�п�,���䷽ʽ]
'           [strID          ��intBdingType=2ʱ���ò�����Ч�����Ҽ�¼���б�����ID�ֶΣ���������]
'           [strGroup       ��intBdingType=3ʱ���ò�����Ч������strGroup��ֵ(��¼���е��ֶ�,ͨ�����ֶ�ǰ��Ҫһ������ʾ����)���������ΰ�]
'           [blnSort        ��intBdingType=3�£��ò�����Ч���Ƿ���strGroup������������]
'           [lngFontSize    VSF�ؼ����壬����ָ������ʹ��ȫ������]
'��    ��:
'           [lngRowFind     ��intBdingType=2ʱ���ò�����Ч�����ر�ˢ�µ��е��к�]
'��    ��:
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

          '���û�д�����������Ĭ�ϴӼ�¼���л�ȡ
1         On Error GoTo SetDataToVSF_Error

2         If intBdingType = 1 Or intBdingType = 3 Then  '��intBdingType=0ʱ��ʹ��Set objVSF.DataSource = rsTmp��ʽ���ᶪʧ֮ǰ���õ�VSF������Ե�intBdingType=0ʱ����������
3             If strTitle = "" Then
4                 For i = 0 To rsTmp.Fields.Count - 1
5                     strTitle = strTitle & ";" & rsTmp.Fields(i).Name & "," & Len(rsTmp.Fields(i).Name) * 400 & "," & flexAlignLeftCenter
6                 Next
7                 If strTitle <> "" Then strTitle = Mid(strTitle, 2)
8             End If
9             Call vfgSetting(0, objVSF, strTitle, , lngFontSize)  '����VSF���
10        End If

          '�Ƚ�ͼ�������ص�������
11        If Not objImgList Is Nothing Then
12            For i = 1 To objImgList.ListImages.Count
13                CollIco.Add objImgList.ListImages(i).ExtractIcon, objImgList.ListImages(i).Key
14            Next
15        End If


          '������
16        If intBdingType = 0 Then
17            Call vfgSetting(0, objVSF, , , lngFontSize) '����VSF���
              '�����
18            With objVSF
                  '����
19                Set .DataSource = rsTmp
                  '�̶������־���
20                If .FixedRows > 0 And .Cols > 0 Then
21                    .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
                      '                .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignLeftCenter
22                End If
                  '�����п�����ж���
23                If .Rows > 1 And .Cols > 0 Then
24                    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
25                End If
                  '���ø���Keyֵ
26                For lngCol = 0 To .Cols - 1
27                    .ColKey(lngCol) = Trim(.TextMatrix(0, lngCol))
28                Next

                  '��������ͼ��
29                If Not objImgList Is Nothing Then
30                    For lngRow = 0 To .Rows - 1
31                        For lngCol = 0 To .Cols - 1
32                            Call SetVSFIco(objVSF, lngRow, lngCol, CollIco, Trim(.TextMatrix(lngRow, lngCol)))
33                        Next
34                    Next
35                End If
36            End With

37        ElseIf intBdingType = 1 Then
              '���а�
38            With objVSF
                  '���ù̶�����ͼ��
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
51                                Call SetVSFIco(objVSF, .Rows - 1, .ColIndex(rsTmp.Fields(i).Name), CollIco, Trim(CStr("" & rsTmp.Fields(i).value)))  '��������ͼ��
52                            End If
53                        End If
54                    Next
55                    .Rows = .Rows + 1
56                    rsTmp.MoveNext
57                Loop
58                If .Rows > .FixedRows + 1 Then .Rows = .Rows - 1
59            End With
60        ElseIf intBdingType = 2 Then
              '���а�
61            rsTmp.Filter = "ID='" & strID & "'"
62            If rsTmp.RecordCount > 0 Then
63                With objVSF
64                    For lngRow = .FixedRows To .Rows - 1
65                        If .TextMatrix(lngRow, .ColIndex("ID")) = strID Then
66                            For i = 0 To rsTmp.Fields.Count - 1
67                                If .ColIndex(rsTmp.Fields(i).Name) >= 0 Then
68                                    .TextMatrix(lngRow, .ColIndex(rsTmp.Fields(i).Name)) = CStr("" & rsTmp.Fields(i).value)
69                                    If Not objImgList Is Nothing Then
70                                        Call SetVSFIco(objVSF, lngRow, .ColIndex(rsTmp.Fields(i).Name), CollIco, Trim(CStr("" & rsTmp.Fields(i).value)))  '��������ͼ��
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
84                                    Call SetVSFIco(objVSF, lngRow, .ColIndex(rsTmp.Fields(i).Name), CollIco, Trim(CStr("" & rsTmp.Fields(i).value)))  '��������ͼ��
85                                End If
86                            End If
87                        Next
88                        lngRowFind = .Rows - 1
89                    End If
90                End With
91            End If
92        ElseIf intBdingType = 3 Then
              '���ΰ�
93            If blnSort Then rsTmp.Sort = strGroup
94            With objVSF
95                .OutlineBar = flexOutlineBarComplete
96                .OutlineCol = 0
97                .SubtotalPosition = flexSTAbove
98                .Rows = 1

                  '���ù̶�����ͼ��
99                For lngRow = 0 To .FixedRows - 1
100                   For lngCol = 0 To .Cols - 1
101                       Call SetVSFIco(objVSF, lngRow, lngCol, CollIco, Trim(.TextMatrix(lngRow, lngCol)))
102                   Next
103               Next

104               Do While Not rsTmp.EOF
105                   If InStr("<SP>" & strFL & "<SP>", "<SP>" & rsTmp(strGroup) & "<SP>") <= 0 Then
106                       .Rows = .Rows + 1
107                       .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = rsTmp(strGroup) & ""

                          '�ϲ�
108                       .MergeRow(.Rows - 1) = True
109                       .MergeCellsFixed = flexMergeRestrictRows

                          '�Ӵ�
110                       .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True

                          '����
111                       .IsSubtotal(.Rows - 1) = True

112                       strFL = strFL & "<SP>" & rsTmp(strGroup)
113                   End If
114                   .Rows = .Rows + 1
115                   .RowOutlineLevel(.Rows - 1) = 1
116                   For i = 0 To rsTmp.Fields.Count - 1
117                       If .ColIndex(rsTmp.Fields(i).Name) >= 0 Then
118                           .TextMatrix(.Rows - 1, .ColIndex(rsTmp.Fields(i).Name)) = CStr("" & rsTmp.Fields(i).value)
119                           If Not objImgList Is Nothing Then
120                               Call SetVSFIco(objVSF, .Rows - 1, .ColIndex(rsTmp.Fields(i).Name), CollIco, Trim(CStr("" & rsTmp.Fields(i).value)))  '��������ͼ��
121                           End If
122                       End If
123                   Next

124                   rsTmp.MoveNext
125               Loop
126               .ColHidden(.ColIndex(strGroup)) = True
127           End With
128       End If

129       With objVSF
              '��������ò�ͬ�ı���ɫ
130           If .Rows - 1 > 0 Then
131               If .ColIndex("���") >= 0 Then
132                   .Cell(flexcpBackColor, .FixedRows, .ColIndex("���"), .Rows - 1, .ColIndex("���")) = 16772055
133               End If
134           End If
135       End With

136       SetDataToVSF = True

137       Set CollIco = Nothing



138       Exit Function
SetDataToVSF_Error:
139       Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "ִ��(SetDataToVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
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
'��    ��:������
'����ʱ��:2018/9/29
'��    ��:ɾ���Ƿ��ַ�
'��    ��:
'       strChar         Ҫ������ַ�
'       strInvalidChar  �Ƿ��ַ��������Ϊ�գ���Ϊ~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>,���򰴴�����ַ�����
'��    ��:
'��    ��:����֮����ַ���
'����Ӱ��:
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
17        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "ִ��(StringDelInvalidWord)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
18        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-03-07
'��    ��:  �жϵ�ǰ�걾�е���Ŀ�Ƿ�Ϊ����������Ŀ
'��    ��:
'           lngSampleID     �걾ID
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Function IsTre(ByVal lngSampleID As Long) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo IsTre_Error
          
2         If VerCompare(gSysInfo.VersionLIS, "10.35.130") = -1 Then
3             Exit Function
4         End If
          
5         strSQL = "Select b.id From ���鱨����ϸ A, ���������Ŀ B Where a.���ID = b.id And a.�걾ID = [1] And b.�Ƿ�������Ŀ = 1"
6         Set rsTmp = OpenSQLRecord(Sel_Lis_DB, strSQL, "������Ŀ", lngSampleID)
7         If rsTmp.EOF Then
8             IsTre = False
9         Else
10            IsTre = True
11        End If


12        Exit Function
IsTre_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "ִ��(IsTre)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
14        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/4
'��    ��:ͨ�����¼��ı�ѡ����
'��    ��:
'           objVSF      ��Ҫ�ı�ѡ���е�VSF���
'           KeyAscii    �����룬vbkeyUp=������ϼ���vbkeyDown=������¼�
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Public Sub UpOrDown(objVSF As VSFlexGrid, ByVal KeyAscii As EKey)
          Dim i As Integer


1         On Error GoTo UpOrDown_Error

2         With objVSF
3             .Tag = .Row
4             If KeyAscii = vbKeyUp Then
                  '��������ϼ�
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
                  '��������¼�
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
26        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "ִ��(UpOrDown)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/7/12
'��    ��:����������
'��    ��:
'           cbsMain         ����������
'           Buttons         �˵�����,ÿ��Ԫ�صĸ�ʽΪ �˵�id,����,�Ƿ����
'           blnLargeIcons   �Ƿ��ͼ��
'           Position        �˵�λ��
'��    ��:
'��    ��:
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
          '����������:������������
          '-----------------------------------------------------
3         cbsMain.ActiveMenuBar.Visible = False
4         Set objBar = cbsMain.Add("������", Position)
5         objBar.ContextMenuPresent = False
6         objBar.ShowTextBelowIcons = False
7         objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
8         cbsMain.Options.LargeIcons = blnLargeIcons  'Сͼ��
9         objBar.Position = Position   '�������ڶ���

10        For Each strButton In Buttons
11            varButton = Split(strButton, ",")
12            With objBar.Controls
13                Set objControl = .Add(xtpControlButton, Val(varButton(0)), varButton(1))     '����
14                objControl.Style = xtpButtonIconAndCaption
15                If UCase(varButton(2)) = "TRUE" Then objControl.BeginGroup = True '����
16            End With
17        Next
18        cbsMain.RecalcLayout



19        Exit Sub
CbsButtonInit_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "ִ��(CbsButtonInit)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/7/12
'��    ��:�����˵�
'��    ��:
'           cbsMain         �˵�����
'��    ��:
'��    ��:
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
              '.UseFadedIcons = True '����VisualTheme����Ч
10            .IconsWithShadow = True '����VisualTheme����Ч
11            .UseDisabledIcons = True
12            .LargeIcons = True
13            .SetIconSize True, 24, 24
14            .SetIconSize False, 16, 16
              
15        End With
16        cbsMain.EnableCustomization False
      '17        cbsMain.Icons = frmPublicIco.imgPublic.Icons
17        cbsMain.ActiveMenuBar.ContextMenuPresent = False    '��ֹ�Ҽ�ѡ�񹤾�����ȡ��
18        cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap  '��ֹ�ƶ�������





19        Exit Function
CbsSetting_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlPubMethods", "ִ��(CbsSetting)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/21
'��    ��:�ı���ȫѡ
'��    ��:
'��    ��:
'��    ��:
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
