VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPubSel 
   AutoRedraw      =   -1  'True
   Caption         =   "选择器"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "frmPubSel.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   6840
      TabIndex        =   9
      Top             =   0
      Width           =   6840
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找"
         Height          =   300
         Left            =   5880
         TabIndex        =   13
         Top             =   97
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   4320
         TabIndex        =   12
         Top             =   97
         Width           =   1455
      End
      Begin VB.CheckBox chkShowChild 
         Caption         =   "包含下级项目"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择一个项目,然后点击确定"
         Height          =   180
         Left            =   180
         TabIndex        =   10
         Top             =   157
         Width           =   2430
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   6840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3840
      Width           =   6840
      Begin VB.CommandButton cmdSelALL 
         Caption         =   "全选(&A)"
         Height          =   360
         Left            =   300
         TabIndex        =   4
         ToolTipText     =   "Ctrl+A"
         Top             =   100
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdClear 
         Cancel          =   -1  'True
         Caption         =   "全清(&R)"
         Height          =   360
         Left            =   1320
         TabIndex        =   5
         ToolTipText     =   "Ctrl+R"
         Top             =   100
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5295
         TabIndex        =   3
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4170
         TabIndex        =   2
         Top             =   105
         Width           =   1100
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3240
      Left            =   2205
      TabIndex        =   1
      Top             =   555
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   5715
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   3240
      Left            =   15
      TabIndex        =   0
      Top             =   540
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   5715
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2145
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3210
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   540
      Width           =   45
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   4725
      Top             =   1425
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   2400
      ScaleHeight     =   1110
      ScaleWidth      =   2220
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   2280
   End
End
Attribute VB_Name = "frmPubSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrKey As String
Private mblnMulti As Boolean
Private mstrFind As String
Private mlngFindIndex As Long

'入口参数
Private mstrTitle As String
Private mstrNote As String
Private mbytStyle As Byte
Private mstrSeek As String
Private mbln末级 As Boolean
Private mblnShowSub As Boolean
Private mblnShowRoot As Boolean
Private mblnMultiOne As Boolean
Private mstrColWith As String '列宽设置参数
Private mstrTipCol As String   '悬浮提示的列
Private mbytSize As Byte '字体大小
Private mcnOracle As ADODB.Connection

Private mstrSaveTag As String '注册表区分键
Private mstrSQL As String
Private marrInput() As Variant
Private marrHideCols()  As Variant '可以隐藏的列的名字
Private mblnSearch As Boolean '是否通过输入行号检索
Private mblnNotShowNon As Boolean '不显示没有子项的分类，bytStyle=2
Private mblnNoneWin As Boolean
Private mlngX As Long, mlngY As Long, mlngTxtH As Long
Private mstrCheck As String
Private mstrCheckMult As String
'出口参数
Private mrsSel As ADODB.Recordset

'程序变量
Private mblnOk As Boolean

'Public Function ShowSelect(frmParent As Object, ByVal cnOracle As ADODB.Connection, ByVal strSQL As String, bytStyle As Byte, _
'    ByVal strTitle As String, bln末级 As Boolean, _
'    ByVal strSeek As String, ByVal strNote As String, _
'    ByVal blnShowSub As Boolean, blnShowRoot As Boolean, _
'    ByVal blnNoneWin As Boolean, ByVal X As Long, _
'    ByVal y As Long, ByVal txtH As Long, _
'    ByRef Cancel As Boolean, ByVal blnMultiOne As Boolean, _
'    ByVal blnSearch As Boolean, ByVal blnMulti As Boolean, _
'    Optional arrInput As Variant) As ADODB.Recordset
''功能：多功能选择器
''参数：
''     frmParent=显示的父窗体
''     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
''     bytStyle=选择器风格
''       为0时:列表风格:ID,…
''       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
''       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
''             双表风格如果列名存在含Check结尾的字段，则该字段作为是否勾选的值存储字段。=1为勾选，0=不勾选。
''             双表风格如果列名存在*名称，*简码，*编码的，则显示右上角的查询功能，以供查询项目，
''                    编码列必须整个匹配，匹配成功后定位到该分类的该项目上，按F3支持查找下一个。
''     strTitle=选择器功能命名,也用于个性化区分
''     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
''     strSeek=当bytStyle<>2时有效,缺省定位的项目。
''             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
''             bytStyle=1时,可以是编码或名称
''     strNote=选择器的说明文字
''     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
''     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
''     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
''     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
''     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
''     blnSearch=是否显示行号,并可以输入行号定位
''     blnMulti=是否允许多选
''     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
''     arrInput第一个参数如果是“不显示没有子项的分类”，则不显示没有子项的分类
''     arrInput中，
''           格式为："bytSize=?"表示设置字体大小(0-小字体,1-大字体;小字体为9号字,大字体为12号字),默认小字体。
''           格式为：ColSet:...时表示列宽设置,ColSet格式:列宽设置|列名1,宽度1;列名2,宽度2.....|悬浮提示|列名。
''返回：取消=Nothing,选择=SQL源的单行记录集
''说明：
''     1.ID和上级ID可以为字符型数据
''     2.末级等字段不要带空值
''应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
'
'    Dim blnHaveColSet As Boolean
'    Dim i As Long, j As Integer, arrTmp As Variant
'    Dim strColSet As String
'    Dim blnFontSize As Boolean
'
'    mstrSQL = strSQL
'    mstrColWith = "": mstrTipCol = ""
'    mblnNotShowNon = False
'    mbytSize = 0
'    marrInput = Array()
'    '从参数数组中解析特殊设置
'    If TypeName(arrInput) <> "Error" Then
'        '从可变参数中分各种参数
'        If UBound(arrInput) >= 0 Then
'            For i = LBound(arrInput) To UBound(arrInput)
'                If TypeName(arrInput(i)) = "Error" Then arrInput(i) = "" '将没传的参数，转换为空串，不然使用会出错
'                If UCase(arrInput(i)) Like "BYTSIZE*=*" Then
'                    mbytSize = Val(Split(arrInput(i), "=")(1))
'                ElseIf i = UBound(arrInput) And UCase(arrInput(i)) Like "COLSET:*" Then 'COLSET放在最后一位
'                    arrTmp = Split(arrInput(i), ":")
'                    arrTmp = Split(arrTmp(1), "|")
'                    For j = LBound(arrTmp) To UBound(arrTmp) Step 2
'                        If arrTmp(j) = "列宽设置" Then
'                            mstrColWith = arrTmp(j + 1)
'                        ElseIf arrTmp(j) = "悬浮提示" Then
'                            mstrTipCol = arrTmp(j + 1)
'                        End If
'                    Next
'                ElseIf bytStyle = 2 And i = 0 Then '不显示没有子项的分类放在第一位
'                    If arrInput(i) = "不显示没有子项的分类" Then
'                        mblnNotShowNon = True
'                    Else
'                        ReDim Preserve marrInput(UBound(marrInput) + 1)
'                        marrInput(UBound(marrInput)) = arrInput(i)
'                    End If
'                Else
'                    ReDim Preserve marrInput(UBound(marrInput) + 1)
'                    marrInput(UBound(marrInput)) = arrInput(i)
'                End If
'            Next
'        End If
'    End If
'
'    marrHideCols = Array()
'    Call GetHideCols '获取可隐藏列
'    mstrTitle = strTitle
'    mstrNote = strNote
'    mbytStyle = bytStyle
'    mbln末级 = bln末级
'    mstrSeek = strSeek
'    mblnShowSub = blnShowSub
'    mblnShowRoot = blnShowRoot
'    mblnMultiOne = blnMultiOne
'    mblnNoneWin = blnNoneWin
'    mlngX = X: mlngY = y: mlngTxtH = txtH
'    mblnSearch = blnSearch
'    mblnMulti = blnMulti
'
'    If Not frmParent Is Nothing Then
'        mstrSaveTag = frmParent.name & "_" & strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
'    Else
'        mstrSaveTag = strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
'    End If
'    On Error Resume Next
'    Me.Show 1, frmParent
'    On Error GoTo 0
'    If mblnOK Then
'        Cancel = False
'        Set ShowSelect = mrsSel
'    Else
'        Cancel = True
'        Set ShowSelect = Nothing
'    End If
'End Function

Public Function ShowSelectV2(ByVal cnOracle As ADODB.Connection, frmParent As Object, ByVal objControl As Object, ByVal strSQL As String, bytStyle As Byte, _
                                                ByVal strTitle As String, ByVal bln末级 As Boolean, ByVal strSeek As String, ByVal strNote As String, _
                                                ByVal blnShowSub As Boolean, ByVal blnShowRoot As Boolean, ByVal blnNoneWin As Boolean, ByRef Cancel As Boolean, _
                                                Optional ByVal blnMultiOne As Boolean, Optional ByVal blnSearch As Boolean, Optional ByVal blnMulti As Boolean, _
                                                Optional ByVal strOtherInfo As String, Optional arrInput As Variant) As ADODB.Recordset
'功能：多功能选择器
'参数：
'     frmParent=显示的父窗体
'     objControl=调用界面输入框
'     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
'     bytStyle=选择器风格
'       为0时:列表风格:ID,…
'       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
'       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
'             双表风格如果列名存在含Check结尾的字段，则该字段作为是否勾选的值存储字段。=1为勾选，0=不勾选。
'             双表风格如果列名存在*名称，*简码，*编码的，则显示右上角的查询功能，以供查询项目，
'                    编码列必须整个匹配，匹配成功后定位到该分类的该项目上，按F3支持查找下一个。
'     strTitle=选择器功能命名,也用于个性化区分
'     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
'     strSeek=当bytStyle<>2时有效,缺省定位的项目。
'             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
'             bytStyle=1时,可以是编码或名称
'     strNote=选择器的说明文字
'     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
'     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
'     blnNoneWin=处理成非窗体风格
'     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
'     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
'     blnSearch=是否显示行号,并可以输入行号定位
'     blnMulti=是否允许多选
'     strOtherInfo=格式为：项目名称1=内容1#项目2=内容2#......
'                           当前项目有：bytSize=0,1;字体大小(0-小字体,1-大字体;小字体为9号字,大字体为12号字),默认小字体
'                                              ColSet=列宽设置|列名1,宽度1;列名2,宽度2.....|悬浮提示|列名。
'                                              NotShowNon=0,1;0-默认处理，显示没有子项的分类，1-不显示没有子项的分类;bytStyle=2有作用
'     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
'返回：取消=Nothing,选择=SQL源的单行记录集
'说明：
'     1.ID和上级ID可以为字符型数据
'     2.末级等字段不要带空值
'应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    Dim arrInfo As Variant, arrTmp As Variant, arrTmp2 As Variant
    Dim i As Long, j As Long
    Dim lngH As Long, lngW As Long, vRect As RECT, sngX As Single, sngY As Single
    Dim vPoint As POINTAPI
    
    Set mcnOracle = cnOracle
    mstrSQL = strSQL
    mstrColWith = ""
    mstrTipCol = ""
    mblnNotShowNon = False
    mbytSize = 0
    '解析strOtherInfoInfo
    arrInfo = Split(strOtherInfo, "#")
    For i = LBound(arrInfo) To UBound(arrInfo)
        If Trim(arrInfo(i)) <> "" Then
            arrTmp = Split(Trim(arrInfo(i)), "=")
            If UBound(arrTmp) = 1 Then
                Select Case UCase(arrTmp(0))
                    Case "BYTSIZE" '字体
                        mbytSize = Val(arrTmp(1))
                    Case "COLSET" '列宽与悬浮列设置
                        arrTmp2 = Split(arrTmp(1), "|")
                        For j = LBound(arrTmp) To UBound(arrTmp) Step 2
                            If arrTmp2(j) = "列宽设置" Then
                                mstrColWith = arrTmp2(j + 1)
                            ElseIf arrTmp2(j) = "悬浮提示" Then
                                mstrTipCol = arrTmp2(j + 1)
                            End If
                        Next
                    Case "NOTSHOWNON" '不显示没有子项的分类
                        If bytStyle = 2 Then mblnNotShowNon = Val(arrTmp(1))
                End Select
            End If
        End If
    Next
    '通过Api计算出控件的相关坐标信息
    If Not objControl Is Nothing Then
        Select Case UCase(TypeName(objControl))
            Case UCase("VSFlexGrid")
                vPoint = GetClientPoint(objControl.hwnd)
                sngX = vPoint.X
                sngY = vPoint.Y + objControl.Height
                lngH = objControl.CellHeight
                lngW = objControl.CellWidth
                sngY = sngY - lngH
            Case UCase("BILLEDIT")
                vPoint = GetClientPoint(objControl.MsfObj.hwnd)
                sngX = vPoint.X
                sngY = vPoint.Y + objControl.MsfObj.Height
                lngH = objControl.MsfObj.CellHeight
                lngW = objControl.MsfObj.CellWidth
            Case Else
                vRect = GetControlRect(objControl.hwnd)
                sngX = vRect.Left - 15
                sngY = vRect.Top
                lngH = objControl.Height
                lngW = objControl.Width
        End Select
    End If
    mlngX = sngX: mlngY = sngY: mlngTxtH = lngH
    marrInput = arrInput
    marrHideCols = Array()
    Call GetHideCols '获取可隐藏列
    mstrTitle = strTitle
    mstrNote = strNote
    mbytStyle = bytStyle
    mbln末级 = bln末级
    mstrSeek = strSeek
    mblnShowSub = blnShowSub
    mblnShowRoot = blnShowRoot
    mblnMultiOne = blnMultiOne
    mblnNoneWin = blnNoneWin
    mblnSearch = blnSearch
    mblnMulti = blnMulti
    If Not frmParent Is Nothing Then
        mstrSaveTag = frmParent.name & "_" & strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
    Else
        mstrSaveTag = strTitle & "_" & bytStyle & IIf(blnNoneWin, 0, 1)
    End If
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    If mblnOk Then
        Cancel = False
        Set ShowSelectV2 = mrsSel
    Else
        Cancel = True
        Set ShowSelectV2 = Nothing
    End If
End Function


Private Sub chkShowChild_Click()
    mblnShowSub = chkShowChild.value = 1
    If Not tvw_s.SelectedItem Is Nothing Then mstrKey = "": Call tvw_s_NodeClick(tvw_s.SelectedItem)
End Sub

Private Sub cmdCancel_Click()
    Set mrsSel = Nothing
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    Dim strFilter As String
    
    strFilter = mrsSel.Filter
    For i = 1 To lvw.ListItems.Count
        lvw.ListItems(i).Checked = False
        If mstrCheck <> "" Then
            mrsSel.Filter = "ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "' And 末级=1"
            If mrsSel.RecordCount > 0 Then
                mrsSel.Fields(mstrCheck).value = "0"
                mrsSel.Update
            End If
        End If
    Next
    mrsSel.Filter = strFilter
    
    If Not lvw.SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim int性质 As Integer, i As Long, j As Long, K As Long
    Dim strFilter As String
    Dim lngTmp As Long, strTemp As String
    
    If txtFind.Text <> "" And mlngFindIndex > 0 Then
        With mrsSel
            strFilter = .Filter
            .Filter = "末级=1"
            If .RecordCount > 0 Then .AbsolutePosition = mlngFindIndex
            strFind = UCase(Trim(txtFind.Text))
            If IsCharChinese(txtFind.Text) Then
                '中文的只查名称
                int性质 = 1
            ElseIf IsCharAlpha(txtFind.Text) Then
                '英文查名称和简码
                int性质 = 2
            Else
                '否则查名称简码和编码
                int性质 = 3
            End If
            For i = mlngFindIndex To .RecordCount
                If int性质 = 1 Then
                    For j = 0 To UBound(Split(mstrFind, ","))
                        If Split(mstrFind, ",")(j) Like "*名称" Then
                            If .Fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                lngTmp = !Id
                                '75926,冉俊明,2014-7-28
                                strTemp = "_" & Nvl(!上级id)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "直接调用"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For K = 1 To lvw.ListItems.Count
                                        If lvw.ListItems.Item(K).Key Like "*_" & lngTmp Then
                                            lvw.ListItems.Item(K).Selected = True
                                            lvw.ListItems.Item(K).EnsureVisible
                                            Exit For
                                        End If
                                    Next
                                Else
                                    lvw.ListItems("_" & lngTmp).Selected = True
                                    lvw.ListItems("_" & lngTmp).EnsureVisible
                                End If
                                .Filter = strFilter
                                Exit Sub
                            End If
                        End If
                    Next
                ElseIf int性质 = 2 Then
                    For j = 0 To UBound(Split(mstrFind, ","))
                        If Split(mstrFind, ",")(j) Like "*名称" Or Split(mstrFind, ",")(j) Like "*简码" Then
                            If .Fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                lngTmp = !Id
                                '75926,冉俊明,2014-7-28
                                strTemp = "_" & Nvl(!上级id)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "直接调用"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For K = 1 To lvw.ListItems.Count
                                        If lvw.ListItems.Item(K).Key Like "*_" & lngTmp Then
                                            lvw.ListItems.Item(K).Selected = True
                                            lvw.ListItems.Item(K).EnsureVisible
                                            Exit For
                                        End If
                                    Next
                                Else
                                    lvw.ListItems("_" & lngTmp).Selected = True
                                    lvw.ListItems("_" & lngTmp).EnsureVisible
                                End If
                                .Filter = strFilter
                                Exit Sub
                            End If
                        End If
                    Next
                Else
                    For j = 0 To UBound(Split(mstrFind, ","))
                        If Split(mstrFind, ",")(j) Like "*名称" Or Split(mstrFind, ",")(j) Like "*简码" Then
                            If .Fields(Split(mstrFind, ",")(j)).value & "" Like "*" & strFind & "*" Then
                                mlngFindIndex = i + 1
                                lngTmp = !Id
                                '75926,冉俊明,2014-7-28
                                strTemp = "_" & Nvl(!上级id)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "直接调用"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For K = 1 To lvw.ListItems.Count
                                        If lvw.ListItems.Item(K).Key Like "*_" & lngTmp Then
                                            lvw.ListItems.Item(K).Selected = True
                                            lvw.ListItems.Item(K).EnsureVisible
                                            Exit For
                                        End If
                                    Next
                                Else
                                    lvw.ListItems("_" & lngTmp).Selected = True
                                    lvw.ListItems("_" & lngTmp).EnsureVisible
                                End If
                                .Filter = strFilter
                                Exit Sub
                            End If
                        ElseIf Split(mstrFind, ",")(j) Like "*编码" Then
                            If .Fields(Split(mstrFind, ",")(j)).value & "" = strFind Then
                                mlngFindIndex = i + 1
                                lngTmp = !Id
                                '75926,冉俊明,2014-7-28
                                strTemp = "_" & Nvl(!上级id)
                                If strTemp = "_" Then strTemp = "Root"
                                tvw_s.Nodes(strTemp).Selected = True
                                tvw_s.Nodes(strTemp).EnsureVisible
                                tvw_s.Nodes(strTemp).Tag = "直接调用"
                                Call tvw_s_NodeClick(tvw_s.Nodes(strTemp))
                                If mblnSearch = True Then
                                    For K = 1 To lvw.ListItems.Count
                                        If lvw.ListItems.Item(K).Key Like "*_" & lngTmp Then
                                            lvw.ListItems.Item(K).Selected = True
                                            lvw.ListItems.Item(K).EnsureVisible
                                            Exit For
                                        End If
                                    Next
                                Else
                                    lvw.ListItems("_" & lngTmp).Selected = True
                                    lvw.ListItems("_" & lngTmp).EnsureVisible
                                End If
                                .Filter = strFilter
                                Exit Sub
                            End If
                        End If
                    Next
                End If
                .MoveNext
            Next
            If mlngFindIndex = 1 Then
                MsgBox "未找到您查询的项目。", vbInformation, Me.Caption
            ElseIf mlngFindIndex <> 1 Then
                MsgBox "已经查找完最后一个项目了。", vbInformation, Me.Caption
                mlngFindIndex = 1
            End If
            .Filter = strFilter
        End With
    End If
End Sub

Private Sub cmdOK_Click()
    If mrsSel Is Nothing Then Exit Sub
    If mrsSel.RecordCount = 0 Then Exit Sub
    
    If mbln末级 And mbytStyle = 1 Then
        If mrsSel!末级 <> 1 Then Exit Sub
    End If
    
    If mblnMulti And mbytStyle = 2 Then
        mrsSel.Filter = GetListFilter
    End If
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    Dim strFilter As String
    
    strFilter = mrsSel.Filter
    For i = 1 To lvw.ListItems.Count
        lvw.ListItems(i).Checked = True
        If mstrCheck <> "" Then
            mrsSel.Filter = "ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "' And 末级=1"
            If mrsSel.RecordCount > 0 Then
                mrsSel.Fields(mstrCheck).value = "1"
                mrsSel.Update
            End If
        End If
    Next
    mrsSel.Filter = strFilter
    
    If Not lvw.SelectedItem Is Nothing Then
        Call lvw_ItemClick(lvw.SelectedItem)
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If lvw.Visible Then
        If lvw.ListItems.Count = 0 And tvw_s.Visible = True Then
            tvw_s.SetFocus
        Else
            lvw.SetFocus
        End If
    Else
        tvw_s.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And cmdOK.Enabled And Me.ActiveControl.name <> "txtFind" And Me.ActiveControl.name <> "cmdFind" Then
        cmdOK_Click
    ElseIf KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        cmdCancel_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If lvw.Checkboxes Then Call cmdSelAll_Click
    ElseIf (KeyCode = vbKeyR Or KeyCode = vbKeyC) And Shift = vbCtrlMask Then
        If lvw.Checkboxes Then Call cmdClear_Click
    ElseIf KeyCode = vbKeyF3 Then
        cmdFind_Click
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long
    Dim lngColW As Long, i As Integer
    Dim blnSame As Boolean, strItemID As String
    Dim strCode As String, strName As String
    Dim objNode As Node
    Dim lngIndex As Long
    Dim arrCols As Variant, arrTmp As Variant
    Dim blnLike As Boolean '是否是输入匹配
    
    Screen.MousePointer = 11
    
    On Error GoTo ErrH
    
    mblnOk = False
    mstrKey = ""
    mlngFindIndex = 1
    
    '设置控件字体大小
    Call SetFontSize(Me, mbytSize)
    '打开SQL语句
    Set mrsSel = gclsBase.OpenSQLRecordByArray(mcnOracle, mstrSQL, Me.Caption, marrInput)
    If mrsSel.RecordCount > 0 Then mrsSel.MoveFirst
    '没有数据则返回
    If mrsSel.EOF Then
        Screen.MousePointer = 0
        Set mrsSel = Nothing
        mblnOk = True: Unload Me: Exit Sub
    End If
    
    If mstrSQL Like "*%*" Then
        blnLike = True
    Else
        For i = LBound(marrInput) To UBound(marrInput)
            If marrInput(i) Like "*%*" Then
                blnLike = True: Exit For
            End If
        Next
    End If
    '输入匹配时自动返回的情况
    If blnLike Then
        If mrsSel.RecordCount = 1 Then '只有一行数据
            Screen.MousePointer = 0
            mblnOk = True: Unload Me: Exit Sub
        ElseIf mblnMultiOne And mbytStyle = 0 Then '多行相同数据
            blnSame = True
            For i = 1 To mrsSel.RecordCount
                If i = 1 Then
                    strItemID = mrsSel!Id
                Else
                    If mrsSel!Id <> strItemID Then blnSame = False: Exit For
                End If
                mrsSel.MoveNext
            Next
            mrsSel.MoveFirst
            If blnSame Then
                Screen.MousePointer = 0
                mblnOk = True: Unload Me: Exit Sub
            End If
        End If
    End If
    
    '后面会对 mrsSel 记录集进行操作，要变为动态的
    mrsSel.Filter = 0
    Set mrsSel = CopyNewRec(mrsSel)
     '删除没有子项的分类
    If mblnNotShowNon Then Call DeleteNotHave
    
    If mstrNote <> "" And mbytStyle = 2 Then
        If InStr(1, UCase(mstrNote), "[COUNT]") > 0 Then
            mrsSel.Filter = "末级=1"
            mstrNote = Replace(UCase(mstrNote), "[COUNT]", "[" & mrsSel.RecordCount & "]")
        End If
        For i = 0 To mrsSel.Fields.Count - 1
            If InStr(1, mstrNote, "[" & mrsSel.Fields(i).name & "=") > 0 Then
                lngIndex = InStr(1, mstrNote, "[" & mrsSel.Fields(i).name & "=") + Len(mrsSel.Fields(i).name) + 1
                strCode = Mid(mstrNote, lngIndex)
                strCode = Mid(strCode, 1, InStr(1, strCode, "]") - 1)
                mrsSel.Filter = "末级=1 And " & mrsSel.Fields(i).name & strCode
                mstrNote = Replace(mstrNote, "[" & mrsSel.Fields(i).name & strCode & "]", "[" & mrsSel.RecordCount & "]")
            End If
        Next i
        mrsSel.Filter = ""
    End If
    
    '确定名称字段
    strCode = "": strName = ""
    For i = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(i).name = "编码" Then strCode = "编码"
        If mrsSel.Fields(i).name = "名称" Then
            strName = mrsSel.Fields(i).name
        ElseIf mrsSel.Fields(i).name = "姓名" And strName = "" Then
            strName = mrsSel.Fields(i).name
        End If
    Next
    If strName = "" Then strName = "名称"
    
    '在填充数据之前设置CheckBox样式
    If mbytStyle <> 1 And mblnMulti Then
        lvw.Checkboxes = True
        cmdSelALl.Visible = True
        cmdClear.Visible = True
    End If
    
    '填充数据
    Select Case mbytStyle
        Case 0
            '构造列头
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.Fields.Count - 1
                If (Not mrsSel.Fields(i).name Like "*ID" Or mrsSel.Fields(i).name = "病人ID") And mrsSel.Fields(i).name <> "末级" Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.Fields(i).name, mrsSel.Fields(i).name
                    If mrsSel.Fields(i).name Like "*价*" Or mrsSel.Fields(i).name Like "*额*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Alignment = lvwColumnRight
                    End If
                Else
                    If UCase(mrsSel.Fields(i).name) = "CHECKID" Then
                        mstrCheckMult = mrsSel.Fields(i).name
                    End If
                End If
            Next
            If mblnSearch Then lvw.ColumnHeaders.Add , "_行", "行", , 2
            lvw.Sorted = False
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Position = 1
            
            lvw.ListItems.Clear
            Call FillList
        Case 1
            '所有树形数据
            Set objNode = tvw_s.Nodes.Add(, , "Root", "所有" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            objNode.Sorted = True
            
            If Not mrsSel.EOF Then
                For i = 1 To mrsSel.RecordCount
                    If strCode <> "" Then
                        If IsNull(mrsSel!上级id) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!Id, IIf(IsNull(mrsSel!编码), "", "[" & mrsSel!编码 & "]") & mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!上级id, 4, "_" & mrsSel!Id, IIf(IsNull(mrsSel!编码), "", "[" & mrsSel!编码 & "]") & mrsSel.Fields(strName).value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!上级id) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!Id, mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!上级id, 4, "_" & mrsSel!Id, mrsSel.Fields(strName).value, 1)
                        End If
                    End If
                    If objNode.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then
                        objNode.Selected = True
                        objNode.Parent.Expanded = True
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If tvw_s.SelectedItem.Index = 1 Then tvw_s.Nodes(1).Child.Selected = True
            End If
            tvw_s.SelectedItem.EnsureVisible
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
        Case 2
            '非末级树形数据
            Set objNode = tvw_s.Nodes.Add(, , "Root", "所有" & mstrTitle, 1)
            objNode.Expanded = True
            objNode.Selected = True
            objNode.Sorted = True
            
            If Not mrsSel.EOF Then
                mrsSel.Filter = "末级=0"
                For i = 1 To mrsSel.RecordCount
                    If strCode <> "" Then
                        If IsNull(mrsSel!上级id) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!Id, IIf(IsNull(mrsSel!编码), "", "[" & mrsSel!编码 & "]") & mrsSel.Fields(strName).value, 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!上级id, 4, "_" & mrsSel!Id, IIf(IsNull(mrsSel!编码), "", "[" & mrsSel!编码 & "]") & mrsSel.Fields(strName).value, 1)
                        End If
                    Else
                        If IsNull(mrsSel!上级id) Then
                            Set objNode = tvw_s.Nodes.Add("Root", 4, "_" & mrsSel!Id, mrsSel.Fields(strName).value & "", 1)
                        Else
                            Set objNode = tvw_s.Nodes.Add("_" & mrsSel!上级id, 4, "_" & mrsSel!Id, mrsSel.Fields(strName).value & "", 1)
                        End If
                    End If
                    objNode.Sorted = True
                    mrsSel.MoveNext
                Next
                If Not tvw_s.Nodes(1).Child Is Nothing Then tvw_s.Nodes(1).Child.Selected = True
            End If
            
            '构造列头
            lvw.ColumnHeaders.Clear
            For i = 0 To mrsSel.Fields.Count - 1
                If (Not mrsSel.Fields(i).name Like "*ID" Or mrsSel.Fields(i).name = "病人ID") And mrsSel.Fields(i).name <> "末级" Then
                    lvw.ColumnHeaders.Add , "_" & mrsSel.Fields(i).name, mrsSel.Fields(i).name
                    If mrsSel.Fields(i).name Like "*价*" Or mrsSel.Fields(i).name Like "*额*" Then
                        lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Alignment = lvwColumnRight
                    End If
                    If UCase(mrsSel.Fields(i).name) Like "*CHECK" Then
                        mstrCheck = mrsSel.Fields(i).name
                    End If
                    If mrsSel.Fields(i).name Like "*名称" Or mrsSel.Fields(i).name Like "*简码" Or mrsSel.Fields(i).name Like "*编码" Then
                        mstrFind = mstrFind & "," & mrsSel.Fields(i).name
                    End If
                End If
            Next
            mstrFind = Mid(mstrFind, 2)
            If mblnSearch Then lvw.ColumnHeaders.Add , "_行", "行", , 2
            lvw.Sorted = False
            If mblnSearch Then lvw.ColumnHeaders(lvw.ColumnHeaders.Count).Position = 1
            
            Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End Select
    
    '设置控件可见性
    '---------------------------------------------------------------
    
    If mstrTitle <> "" Then
        Me.Caption = mstrTitle & "选择"
    End If
    If mstrNote <> "" Then
        lblInfo.Caption = mstrNote
    End If
    If mblnNoneWin Then
        pic.Width = 30
        pic.BackColor = vbBlack
        pic.ZOrder
        If Not (mbytStyle = 2 And mstrNote <> "") Then picInfo.Visible = False
        picCmd.Visible = False
        lvw.Appearance = ccFlat
        lvw.BorderStyle = ccFixedSingle
        tvw_s.Appearance = ccFlat
        tvw_s.BorderStyle = ccFixedSingle
    Else
        '字体变大时，调整控件位置
        If mbytSize = 1 Then
            picInfo.Height = picInfo.Height + 60
            lblInfo.Top = lblInfo.Top + 15
            
            chkShowChild.Top = chkShowChild.Top + 30
            chkShowChild.Left = lblInfo.Left + lblInfo.Width + 200
            
            txtFind.Height = 360: txtFind.Left = chkShowChild.Left + chkShowChild.Width + 200
            
            cmdFind.Height = 420: cmdFind.Width = 1300
            cmdFind.Top = cmdFind.Top - 50: cmdFind.Left = txtFind.Left + txtFind.Width + 50
            
            picCmd.Height = picCmd.Height + 30
            cmdSelALl.Height = 420: cmdSelALl.Width = 1500
            cmdSelALl.Top = cmdSelALl.Top - 30
            
            cmdClear.Height = 420: cmdClear.Width = 1500
            cmdClear.Top = cmdClear.Top - 30: cmdClear.Left = cmdSelALl.Left + cmdSelALl.Width + 20
            
            cmdOK.Height = 420: cmdOK.Width = 1500
            cmdOK.Top = cmdOK.Top - 30:
            
            cmdCancel.Height = 420: cmdCancel.Width = 1500
            cmdCancel.Top = cmdCancel.Top - 30
        End If
    End If
    Me.Left = mlngX
    arrCols = Split(mstrColWith, ";")
    For i = LBound(arrCols) To UBound(arrCols)
        arrTmp = Split(arrCols(i), ",")
        lvw.ColumnHeaders("_" & arrTmp(0)).Width = Val(arrTmp(1))
    Next
    If mbytStyle = 1 Then
        Me.Width = 5400
        If mbytSize = 1 Then Me.Width = Me.Width + 500
    Else
        If mbytStyle <> 2 Then Me.Width = 7200: Me.Height = 4800
        If mbytStyle = 2 Then Me.Width = 8400: Me.Height = 5600
        If Me.Left + Me.Width > Screen.Width Then
            If mbytSize = 1 Then tvw_s.Width = tvw_s.Width + 500
            lngScrW = GetSystemMetrics(SM_CXVSCROLL) * 15 + 75
            For i = 1 To lvw.ColumnHeaders.Count
                lngColW = lngColW + lvw.ColumnHeaders(i).Width
            Next
            If mbytStyle = 2 And lngColW < 1.5 * tvw_s.Width Then lngColW = 1.5 * tvw_s.Width
            If mbytStyle = 2 Then lngColW = lngColW + tvw_s.Width
            
            If Me.Left + lngColW + lngScrW > Screen.Width Then
                lngColW = 0
                For i = 1 To lvw.ColumnHeaders.Count
                    If InStr(";" & mstrColWith, ";" & lvw.ColumnHeaders(i).Text & ",") = 0 Then
                        If lvw.ColumnHeaders(i).Width > IIf(mbytSize = 1, 2400, 1800) Then lvw.ColumnHeaders(i).Width = IIf(mbytSize = 1, 2400, 1800)
                    End If
                    lngColW = lngColW + lvw.ColumnHeaders(i).Width
                Next
                If Me.Left + lngColW + lngScrW > Screen.Width Then
                    Me.Width = Screen.Width - Me.Left
                Else
                    Me.Width = lngColW + lngScrW
                End If
            Else
                Me.Width = lngColW + lngScrW
            End If
        End If
    End If
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '屏幕可用高度
    If mlngY + mlngTxtH + Me.Height > lngScrH Then
        Me.Top = mlngY - Me.Height
    Else
        Me.Top = mlngY + mlngTxtH
    End If
    Select Case mbytStyle
        Case 0
            lvw.Visible = True
            tvw_s.Visible = False
            pic.Visible = False
            cmdFind.Visible = False
            txtFind.Visible = False
        Case 1
            lvw.Visible = False
            tvw_s.Visible = True
            pic.Visible = False
            cmdFind.Visible = False
            txtFind.Visible = False
        Case 2
            lvw.Visible = True
            tvw_s.Visible = True
            pic.Visible = True
            chkShowChild.Visible = True
            If mstrFind <> "" Then
                cmdFind.Visible = True
                txtFind.Visible = True
            End If
    End Select
    
    '调整窗体尺寸
    '---------------------------------------------------------------
    If mblnNoneWin Then
        Call FormSetCaption(Me, False, False)
    End If
    Call Form_Resize
    Screen.MousePointer = 0
    Exit Sub
ErrH:
    Screen.MousePointer = 0
    If 0 = 1 Then
        Resume
    End If
    Unload Me
End Sub

Private Sub DeleteNotHave()
'功能：删除没有子项的分类
    Dim i As Long
    Dim strFilter As String
    Dim rsTmp As Recordset
    Dim rstmp1 As Recordset
    
    strFilter = mrsSel.Filter
    mrsSel.Filter = "末级=1"
    Set rsTmp = CopyNewRec(mrsSel)
    mrsSel.Filter = "末级=0"
    Set rstmp1 = CopyNewRec(mrsSel)
    If mrsSel.RecordCount > 0 Then mrsSel.MoveFirst
    For i = mrsSel.RecordCount To 1 Step -1
        mrsSel.AbsolutePosition = i
        rstmp1.Filter = "上级ID=" & mrsSel!Id & " And ID<>-1"
        rsTmp.Filter = "上级ID=" & mrsSel!Id
        If rstmp1.RecordCount = 0 And rsTmp.RecordCount = 0 Then
            rstmp1.Filter = "ID=" & mrsSel!Id
            rstmp1!Id = "-1"
            mrsSel!Id = "-1"
        End If
    Next
    mrsSel.Filter = "ID=-1"
    Do While Not mrsSel.EOF
        mrsSel.Delete
        If mrsSel.RecordCount >= 0 Then mrsSel.MoveNext
    Loop
    mrsSel.Filter = IIf(strFilter = "0", 0, strFilter)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    Select Case mbytStyle
        Case 0 'ListView
            lvw.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            lvw.Left = 0
            lvw.Width = Me.ScaleWidth
            lvw.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
        Case 1
            tvw_s.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Left = 0
            tvw_s.Width = Me.ScaleWidth
            tvw_s.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
        Case 2
            tvw_s.Left = 0
            tvw_s.Top = IIf(picInfo.Visible, picInfo.Height, 0)
            tvw_s.Height = Me.ScaleHeight - IIf(picInfo.Visible, picInfo.Height, 0) - IIf(picCmd.Visible, picCmd.Height, 0)
            
            pic.Top = tvw_s.Top
            pic.Height = tvw_s.Height
            lvw.Top = tvw_s.Top
            lvw.Height = tvw_s.Height
            
            If mblnNoneWin Then
                pic.Left = tvw_s.Width - pic.Width / 2
                lvw.Left = tvw_s.Width
                lvw.Width = Me.ScaleWidth - tvw_s.Width
            Else
                pic.Left = tvw_s.Width
                lvw.Left = tvw_s.Width + pic.Width
                lvw.Width = Me.ScaleWidth - tvw_s.Width - pic.Width
            End If
    End Select
    
    picBack.Left = lvw.Left
    picBack.Top = lvw.Top
    picBack.Width = lvw.Width
    picBack.Height = lvw.Height
    
    If Me.ScaleWidth - cmdCancel.Width * 1.3 - cmdOK.Width >= cmdClear.Left + cmdClear.Width * 1.3 Then
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.3
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 20
    End If
    
    Me.Refresh
End Sub

Private Sub lvw_DblClick()
    If cmdOK.Enabled And Not lvw.SelectedItem Is Nothing Then cmdOK_Click
End Sub

Private Function GetListFilter() As String
    Dim strFilter As String
    Dim i As Long
    
    '如果有勾选，以勾选的为准
    If mblnMulti Then
        For i = 1 To lvw.ListItems.Count
            If lvw.ListItems(i).Checked Then
                If mrsSel.Fields("ID").Type = adVarChar Or mrsSel.Fields("ID").Type = adChar Then
                    If mbytStyle = 2 Then
                        strFilter = strFilter & " Or (ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "' And 末级=1)"
                    Else
                        strFilter = strFilter & " Or ID='" & Split(lvw.ListItems(i).Key, "_")(1) & "'"
                    End If
                Else
                    If mbytStyle = 2 Then
                        strFilter = strFilter & " Or (ID=" & Split(lvw.ListItems(i).Key, "_")(1) & " And 末级=1)"
                    Else
                        strFilter = strFilter & " Or ID=" & Split(lvw.ListItems(i).Key, "_")(1)
                    End If
                End If
            End If
        Next
        strFilter = Mid(strFilter, 4)
    End If
    
    '如果没有勾选，则以当前项为准
    If strFilter = "" Then
        If mrsSel.Fields("ID").Type = adVarChar Or mrsSel.Fields("ID").Type = adChar Then
            strFilter = "ID='" & Split(lvw.SelectedItem.Key, "_")(1) & "'"
        Else
            strFilter = "ID=" & Split(lvw.SelectedItem.Key, "_")(1)
        End If
        If mbytStyle = 2 Then strFilter = strFilter & " And 末级=1"
    End If
    
    GetListFilter = strFilter
End Function

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim strFilter As String

    If mstrCheck <> "" Then
        strFilter = mrsSel.Filter
        mrsSel.Filter = "ID='" & Split(Item.Key, "_")(1) & "' And 末级=1"
        If mrsSel.RecordCount > 0 Then
            mrsSel.Fields(mstrCheck).value = IIf(Item.Checked, "1", "0")
            mrsSel.Update
        End If
        mrsSel.Filter = strFilter
    End If
    mrsSel.Filter = GetListFilter
    cmdOK.Enabled = mrsSel.RecordCount > 0
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mrsSel.Filter = GetListFilter
    cmdOK.Enabled = mrsSel.RecordCount > 0
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If mblnSearch Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If lvw.ListItems.Count >= CInt(strIdx) And CInt(strIdx) > 0 Then
                lvw.ListItems(CInt(strIdx)).Selected = True
                lvw.SelectedItem.EnsureVisible
                Call lvw_ItemClick(lvw.SelectedItem)
            End If
        End If
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or lvw.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        lvw.Left = lvw.Left + X
        lvw.Width = lvw.Width - X
        picBack.Left = picBack.Left + X
        picBack.Width = picBack.Width - X
        Me.Refresh
    End If
End Sub

Private Sub FillList()
'功能：装入ListView数据
    Dim i As Long, j As Long, K As Long
    Dim objItem As ListItem
    Dim arrCols As Variant
    Dim arrTmp As Variant
    
    On Error Resume Next
    
    lvw.Visible = False
    Screen.MousePointer = 11
    For i = 1 To mrsSel.RecordCount
        For j = 0 To mrsSel.Fields.Count - 1
            If (Not mrsSel.Fields(j).name Like "*ID" Or mrsSel.Fields(j).name = "病人ID") And mrsSel.Fields(j).name <> "末级" Then
                If lvw.ColumnHeaders("_" & mrsSel.Fields(j).name).Index = 1 Then
                    If mblnSearch Then '关键字加入行号
                        Set objItem = lvw.ListItems.Add(, i & "_" & mrsSel!Id, IIf(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value), , 1)
                    Else
                        Set objItem = lvw.ListItems.Add(, "_" & mrsSel!Id, IIf(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value), , 1)
                    End If
                    If objItem.Text Like "*" & mstrSeek & "*" And mstrSeek <> "" Then objItem.Selected = True
                Else
                    objItem.SubItems(lvw.ColumnHeaders("_" & mrsSel.Fields(j).name).Index - 1) = IIf(IsNull(mrsSel.Fields(j).value), "", mrsSel.Fields(j).value)
                End If
            End If
        Next
        If mblnSearch Then objItem.SubItems(lvw.ColumnHeaders("_行").Index - 1) = i
        If mstrCheckMult = "CHECKID" Then
            objItem.Checked = Val(mrsSel.Fields(mstrCheckMult).value & "")
        End If
        If mstrCheck <> "" Then objItem.Checked = objItem.SubItems(lvw.ColumnHeaders("_" & mstrCheck).Index - 1)
        mrsSel.MoveNext
    Next
    
    Call LvwSetColWidth(lvw, , mbytSize)
    '20031013:限制最大宽度
    If lvw.Width > Screen.Width / 2 Then
        For i = 1 To lvw.ColumnHeaders.Count
            If InStr(";" & mstrColWith, ";" & lvw.ColumnHeaders(i).Text & ",") = 0 Then
                If lvw.ColumnHeaders(i).Width > 1800 Then lvw.ColumnHeaders(i).Width = 1800
            End If
        Next
    End If
    arrCols = Split(mstrColWith, ";")
    For i = LBound(arrCols) To UBound(arrCols)
        arrTmp = Split(arrCols(i), ",")
        lvw.ColumnHeaders("_" & arrTmp(0)).Width = Val(arrTmp(1))
    Next
    K = LBound(marrHideCols)
    If K <> -1 Then
        For i = 1 To lvw.ColumnHeaders.Count
            For j = K To IIf(i > UBound(marrHideCols), UBound(marrHideCols), i)
                If lvw.ColumnHeaders(i).Text = marrHideCols(j) Then
                    lvw.ColumnHeaders(i).Width = 0: K = j: Exit For
                End If
            Next
        Next
    End If
    If Not lvw.SelectedItem Is Nothing Then
        cmdOK.Enabled = True
        
        lvw.SelectedItem.EnsureVisible
        Call lvw_ItemClick(lvw.SelectedItem)
    Else
        cmdOK.Enabled = False
    End If
    lvw.Refresh
    lvw.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub tvw_s_DblClick()
    If cmdOK.Enabled And mbytStyle = 1 Then cmdOK_Click
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim mstrKeys As String, i As Integer
    Dim strFilter As String
    
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    If Node.Tag = "直接调用" Then
        Node.Tag = ""
    Else
        mlngFindIndex = 1
    End If
    
    If mbytStyle = 1 Then
        If Node.Key <> "Root" Then
            If mrsSel.Fields("ID").Type = adVarChar Then
                mrsSel.Filter = "ID='" & Mid(Node.Key, 2) & "'"
            Else
                mrsSel.Filter = "ID=" & Mid(Node.Key, 2)
            End If
            If mbln末级 Then
                cmdOK.Enabled = (mrsSel!末级 = 1)
            Else
                cmdOK.Enabled = True
            End If
        Else
            cmdOK.Enabled = False
        End If
    ElseIf mbytStyle = 2 Then
        lvw.ListItems.Clear
        If Node.Key = "Root" Then
            If mblnShowRoot Then
                mrsSel.Filter = "末级=1" '数据量大时很慢
            Else
                mrsSel.Filter = "末级=-1"
            End If
            If Visible Then lvw.SetFocus
        Else
            If mblnShowSub Then
                mstrKeys = GetSubTree(Node) '数据量大时很慢
            Else
                mstrKeys = Mid(Node.Key, 2)
            End If
            For i = 0 To UBound(Split(mstrKeys, ","))
                If mrsSel.Fields("上级ID").Type = adVarChar Then
                    strFilter = strFilter & " Or (末级=1 And 上级ID='" & Split(mstrKeys, ",")(i) & "')"
                Else
                    strFilter = strFilter & " Or (末级=1 And 上级ID=" & Split(mstrKeys, ",")(i) & ")"
                End If
            Next
            strFilter = Mid(strFilter, 5)
            mrsSel.Filter = strFilter
            
'            If mrsSel.Fields("上级ID").Type = adVarChar Then
'                mrsSel.Filter = "末级=1 And 上级ID='" & Mid(Node.Key, 2) & "'"
'            Else
'                mrsSel.Filter = "末级=1 And 上级ID=" & Mid(Node.Key, 2)
'            End If
        End If
        If Not mrsSel.EOF Then Call FillList
    End If
End Sub

Private Function GetSubTree(ByVal objNode As Node) As String
'功能：返回一个结点的子树结点的Key(含该结点)
    Dim mstrKeys As String
    Dim objTmp As Node
    
    mstrKeys = "," & Mid(objNode.Key, 2) & mstrKeys
    Set objTmp = objNode.Child
    Do While Not objTmp Is Nothing
        If objTmp.Children > 0 Then
            mstrKeys = "," & GetSubTree(objTmp) & mstrKeys
        Else
            mstrKeys = "," & Mid(objTmp.Key, 2) & mstrKeys
        End If
        Set objTmp = objTmp.Next
    Loop
    GetSubTree = Mid(mstrKeys, 2)
End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Static blnDesc As Boolean
    Static intIdx As Integer
    
    If mblnSearch And ColumnHeader.Key = "_行" Then Exit Sub
    
    If intIdx = ColumnHeader.Index Then
        blnDesc = Not blnDesc
    Else
        blnDesc = False
    End If
    lvw.SortKey = ColumnHeader.Index - 1
    If blnDesc Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
    lvw.Sorted = True
        
    If mblnSearch Then
        For intIdx = 1 To lvw.ListItems.Count
            lvw.ListItems(intIdx).SubItems(lvw.ColumnHeaders("_行").Index - 1) = intIdx
        Next
    End If
    intIdx = ColumnHeader.Index
    If Not lvw.SelectedItem Is Nothing Then lvw.SelectedItem.EnsureVisible
End Sub

Private Sub txtFind_Change()
    mlngFindIndex = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdFind_Click
    End If
End Sub

Private Sub GetHideCols()
'功能：解析SQL，获取可隐藏的列
'           NUll 列名 或 NULL AS 列名 才可以隐藏
    Dim arrFileds As Variant
    Dim i As Long
    Dim strSQLTmp As String
    Dim arrTmp As Variant
    
    strSQLTmp = Replace(mstrSQL, vbCrLf, " ")
    strSQLTmp = Replace(strSQLTmp, vbLf, " ")
    strSQLTmp = Replace(strSQLTmp, vbCr, " ")
    strSQLTmp = Trim(Replace(strSQLTmp, vbTab, " "))
    '去除空格
    i = 5
    Do While i > 1
        strSQLTmp = Replace(strSQLTmp, String(i, " "), " ")
        If InStr(strSQLTmp, String(i, " ")) = 0 Then i = i - 1
    Loop
    strSQLTmp = UCase(strSQLTmp)
    arrFileds = Split(strSQLTmp, ",")
    '解析空值列
    For i = LBound(arrFileds) To UBound(arrFileds)
        '发现可隐藏列
        If Trim(arrFileds(i)) Like "NULL ?*" Or Trim(arrFileds(i)) Like "NULL AS ?*" Then
            arrTmp = Split(Trim(arrFileds(i)), " ")
            If arrTmp(UBound(arrTmp)) <> "" Then
                ReDim Preserve marrHideCols(UBound(marrHideCols) + 1)
                marrHideCols(UBound(marrHideCols)) = arrTmp(UBound(arrTmp))
            End If
        End If
    Next
End Sub

Private Sub SetFontSize(ByVal objForm As Object, ByVal bytSize As Byte)
'功能：设置界面控件字体大小
'入参：objForm-窗体对象
'      bytSize-字体大小: 0-小字体,1-大字体;小字体为9号字,大字体为12号字
    Dim objCtl As Control
    
    On Error Resume Next
    For Each objCtl In objForm.Controls
        '0-小字体,1-大字体;小字体为9号字,大字体为12号字
        objCtl.Font.Size = IIf(bytSize = 1, 12, 9)
    Next
End Sub

Private Sub LvwSetColWidth(objLvw As Object, Optional blnHideNullCol As Boolean, Optional ByVal bytSize As Byte = 0)
'功能：根据ListView中当前的内容自动调整列为最小匹配宽度,并保持至少可以显示列头文字的宽度
'参数：objLvw=要调整的ListView对象
'      blnHideNullCol=是否隐藏没有任何数据的列
'      bytSize=字体大小：0-小字体(9号) 1-大字体(12号)
    Dim i As Integer, lngW As Long, lngAvgW As Long
    
    lngAvgW = IIf(bytSize = 1, 115, 90)
    For i = 1 To objLvw.ColumnHeaders.Count
        SendMessage objLvw.hwnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If blnHideNullCol Then If objLvw.ColumnHeaders(i).Width < 200 Then objLvw.ColumnHeaders(i).Width = 0
        If objLvw.ColumnHeaders(i).Width < (ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * lngAvgW And objLvw.ColumnHeaders(i).Width <> 0 Then
            objLvw.ColumnHeaders(i).Width = (ActualLen(objLvw.ColumnHeaders(i).Text) + 2) * lngAvgW
        End If
    Next
End Sub



