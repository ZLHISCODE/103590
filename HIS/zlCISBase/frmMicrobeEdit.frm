VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMicrobeEdit 
   BorderStyle     =   0  'None
   Caption         =   "细菌信息"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox cbo标本类型 
      Height          =   300
      ItemData        =   "frmMicrobeEdit.frx":0000
      Left            =   1035
      List            =   "frmMicrobeEdit.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   2925
      Width           =   1620
   End
   Begin VB.ComboBox cboEdit 
      Height          =   300
      Index           =   1
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Tag             =   "革兰氏染色分类"
      Top             =   2100
      Width           =   1425
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&P"
      Height          =   285
      Index           =   0
      Left            =   4980
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   0
      Left            =   1020
      TabIndex        =   17
      Top             =   2505
      Width           =   4230
   End
   Begin VB.TextBox txt结果 
      Height          =   300
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   26
      Top             =   2925
      Width           =   1485
   End
   Begin VB.ComboBox cbo默认方法 
      Height          =   300
      ItemData        =   "frmMicrobeEdit.frx":0004
      Left            =   1215
      List            =   "frmMicrobeEdit.frx":0006
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3330
      Width           =   1440
   End
   Begin VB.TextBox txtWHONET码 
      Height          =   300
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1695
      Width           =   1545
   End
   Begin VB.ComboBox cbo默认药敏 
      Height          =   300
      ItemData        =   "frmMicrobeEdit.frx":0008
      Left            =   4095
      List            =   "frmMicrobeEdit.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3330
      Width           =   1155
   End
   Begin VB.TextBox txt缩写 
      Height          =   300
      Left            =   1020
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1695
      Width           =   1425
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   1020
      MaxLength       =   10
      TabIndex        =   3
      Top             =   495
      Width           =   1425
   End
   Begin VB.TextBox txt中文 
      Height          =   300
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   5
      Top             =   885
      Width           =   4245
   End
   Begin VB.TextBox txt英文 
      Height          =   300
      Left            =   1020
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1290
      Width           =   4245
   End
   Begin VB.ComboBox cbo细菌类型 
      Height          =   300
      ItemData        =   "frmMicrobeEdit.frx":000C
      Left            =   1020
      List            =   "frmMicrobeEdit.frx":0013
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   105
      Width           =   4290
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   1800
      Left            =   105
      TabIndex        =   24
      Top             =   3975
      Width           =   5085
      _cx             =   8969
      _cy             =   3175
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
      BackColorFixed  =   15790320
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
      Cols            =   3
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
      AutoResize      =   -1  'True
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.ComboBox cboEdit 
      Height          =   300
      Index           =   0
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Tag             =   "细菌类别"
      Top             =   2100
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标本类型"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   2985
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "革兰氏染色"
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "菌属分类"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   2580
      Width           =   720
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "细菌类别"
      Height          =   180
      Index           =   0
      Left            =   2910
      TabIndex        =   14
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lbl结果 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认结果"
      Height          =   180
      Left            =   2910
      TabIndex        =   25
      Top             =   2985
      Width           =   720
   End
   Begin VB.Label lbl默认方法 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认检测方法"
      Height          =   180
      Left            =   90
      TabIndex        =   19
      Top             =   3390
      Width           =   1080
   End
   Begin VB.Label lblWHONET码 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WHONET码"
      Height          =   180
      Left            =   2910
      TabIndex        =   10
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lbl抗生素组 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对应进行药敏实验的抗生素组:"
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   3720
      Width           =   2430
   End
   Begin VB.Label lbl默认药敏 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认药敏结果"
      Height          =   180
      Left            =   2910
      TabIndex        =   21
      Top             =   3390
      Width           =   1080
   End
   Begin VB.Label lbl缩写 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "英文缩写"
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "细菌编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   570
      Width           =   720
   End
   Begin VB.Label lbl中文 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "中文名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl英文 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "英文名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lbl细菌类型 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前类型"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   165
      Width           =   720
   End
End
Attribute VB_Name = "frmMicrobeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngGermId As Long          '当前显示的类型id

Private Enum mCol
    ID = 0: 选择: 编码: 名称: 备注
End Enum

'刘兴宏:加入
Private Enum mcboIndex
    idx_细菌类别 = 0
    idx_革兰氏染色 = 1
End Enum
Private Enum mTxtIndex
    idx_菌属分类 = 0
End Enum

Dim lngCount As Long
Private Function SelectItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strkey As String, ByVal strTable As String, ByVal strTittle As String) As Boolean
    '------------------------------------------------------------------------------
    '功能:多功能选择器
    '参数:objCtl-文本框控件
    '     strKey-要搜索的值
    '     strTable-表名
    '     strTittle-选择器名称
    '返回:
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single, int匹配方式 As Integer
    
    Dim rsTemp  As ADODB.Recordset
    int匹配方式 = Val(zlDatabase.GetPara("输入匹配", , , True))
    
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    gstrSql = "Select rownum as ID,a.* From " & strTable & " a"
    
    If strkey <> "" Then
        gstrSql = gstrSql & _
        "   Where ((名称) like [1] or  编码  like [1] or  简码  like  upper([1]))  " & _
        "    "
    End If
    gstrSql = gstrSql & _
    "   order by 编码"
    
    strkey = IIf(int匹配方式 = 0, "%", "") & strkey & "%"
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hWnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSql, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strkey)
    frmMain.SetFocus
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        MsgBox "没有找到满足条件的内容,请检查!", vbDefaultButton1 + vbInformation, gstrSysName
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If

    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
            .Cell(flexcpData, .Row, .Col) = Nvl(rsTemp!名称)
        End With
    Else
        If objCtl.Enabled Then objCtl.SetFocus
        objCtl.Text = Nvl(rsTemp!名称)
        objCtl.Tag = Nvl(rsTemp!名称)
        zlCommFun.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置参考值列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 5: .FixedCols = 0
            .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.编码) = "编码"
            .TextMatrix(0, mCol.名称) = "名称": .TextMatrix(0, mCol.备注) = "备注"
            .ColWidth(mCol.编码) = 900: .ColWidth(mCol.名称) = 2600: .ColWidth(mCol.备注) = 1000
        End If
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.选择) = 280: .TextMatrix(0, mCol.选择) = ""
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.选择) = 1 Then
                .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked
            Else
                .Cell(flexcpChecked, lngCount, mCol.选择) = flexUnchecked
            End If
            .TextMatrix(lngCount, mCol.选择) = ""
        Next
        .Redraw = flexRDDirect
    End With
End Sub
Private Sub InitData()
    '------------------------------------------------------------------------------
    '功能:初始化相应combox数据，及值
    '返回:
    '编制:刘兴宏
    '日期:2008/03/18
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select 编码,名称,简码,缺省标志 From 检验细菌类别 order by 编码"
    zlDatabase.OpenRecordset rsTemp, gstrSql, Me.Caption
    With cboEdit(mcboIndex.idx_细菌类别)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!编码) & "." & Nvl(rsTemp!名称)
            If Val(Nvl(rsTemp!缺省标志)) = 1 Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
    End With
    gstrSql = "Select 编码,名称,简码,缺省标志 From 革兰染色分类 order by 编码"
    zlDatabase.OpenRecordset rsTemp, gstrSql, Me.Caption
    With cboEdit(mcboIndex.idx_革兰氏染色)
        .Clear
        Do While Not rsTemp.EOF
            .AddItem Nvl(rsTemp!编码) & "." & Nvl(rsTemp!名称)
            If Val(Nvl(rsTemp!缺省标志)) = 1 Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function zlRefresh(lngGermId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    mlngGermId = lngGermId
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select ID, 编码, 中文名称 From 检验细菌类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cbo细菌类型.Clear
        Do While Not .EOF
            Me.cbo细菌类型.AddItem !编码 & "-" & !中文名称
            Me.cbo细菌类型.ItemData(Me.cbo细菌类型.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    '清除此前项目的显示
    
    Me.txt编码.Text = "": Me.txt中文.Text = "": Me.txt英文.Text = "": Me.txt缩写.Text = ""
    Me.cbo细菌类型.ListIndex = -1: Me.cbo默认方法.ListIndex = -1: Me.cbo默认药敏.ListIndex = -1: Me.cbo标本类型.ListIndex = -1
    '刘兴宏加入:2008/03/17
    Me.cboEdit(mcboIndex.idx_革兰氏染色).ListIndex = -1: Me.cboEdit(mcboIndex.idx_细菌类别).ListIndex = -1
    Me.txtEdit(mTxtIndex.idx_菌属分类).Text = ""
    If lngGermId = 0 Then setListFormat: zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    gstrSql = "Select 编码, 中文名, 英文名, 简码, 类型id, 默认药敏, 默认方法, Whonet码,默认结果, 细菌类别,细菌菌属 ,革兰氏分类" & vbCrLf & _
              " From 检验细菌 Where ID = [1]  order by 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngGermId)
    With rsTemp
        Me.txt编码.MaxLength = .Fields("编码").DefinedSize
        Me.txt中文.MaxLength = .Fields("中文名").DefinedSize
        Me.txt英文.MaxLength = .Fields("英文名").DefinedSize
        Me.txt缩写.MaxLength = .Fields("简码").DefinedSize
        Me.txtWHONET码.MaxLength = .Fields("WHONET码").DefinedSize
        Me.txt结果.MaxLength = .Fields("默认结果").DefinedSize
        If .RecordCount > 0 Then
            Me.txt编码.Text = "" & !编码
            Me.txt中文.Text = "" & !中文名
            Me.txt英文.Text = "" & !英文名
            Me.txt缩写.Text = "" & !简码
            Me.txtWHONET码.Text = "" & !WHONET码
            Me.txt结果.Tag = "" & !默认结果
            
            For lngCount = 0 To Me.cbo标本类型.ListCount - 1
                If InStr(Me.txt结果.Tag, Mid(Me.cbo标本类型.List(lngCount), InStr(Me.cbo标本类型.List(lngCount), "-") + 1)) > 0 Then
                    Me.cbo标本类型.ListIndex = lngCount
                    Exit For
                End If
            Next
            If Me.cbo标本类型.ListIndex = -1 And Me.cbo标本类型.ListCount > 0 Then Me.cbo标本类型.ListIndex = 0
            
            For lngCount = 0 To Me.cbo细菌类型.ListCount - 1
                If Me.cbo细菌类型.ItemData(lngCount) = Val("" & !类型id) Then Me.cbo细菌类型.ListIndex = lngCount: Exit For
            Next
            
            For lngCount = 0 To Me.cbo默认药敏.ListCount - 1
                If Left(Me.cbo默认药敏.List(lngCount), 1) = "" & !默认药敏 Then Me.cbo默认药敏.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo默认方法.ListCount - 1
                If Mid(Me.cbo默认方法.List(lngCount), 4) = "" & !默认方法 Then Me.cbo默认方法.ListIndex = lngCount: Exit For
            Next
            '刘兴宏:2008/03/18加入
            Me.txtEdit(mTxtIndex.idx_菌属分类).Text = Nvl(rsTemp!细菌菌属)
            Me.txtEdit(mTxtIndex.idx_菌属分类).Tag = Nvl(rsTemp!细菌菌属)
            With Me.cboEdit(mcboIndex.idx_革兰氏染色)
                For lngCount = 0 To .ListCount - 1
                    strTemp = .List(lngCount): strTemp = Mid(strTemp, InStr(1, strTemp, ".") + 1)
                    
                    If strTemp = Nvl(rsTemp!革兰氏分类) Then .ListIndex = lngCount: Exit For
                Next
                If .ListIndex < 0 And Trim(Nvl(rsTemp!革兰氏分类)) <> "" Then
                        .AddItem Nvl(rsTemp!革兰氏分类)
                        .ListIndex = .NewIndex
                End If
            End With
            With Me.cboEdit(mcboIndex.idx_细菌类别)
                For lngCount = 0 To .ListCount - 1
                    strTemp = .List(lngCount): strTemp = Mid(strTemp, InStr(1, strTemp, ".") + 1)
                    
                    If strTemp = Nvl(rsTemp!细菌类别) Then .ListIndex = lngCount: Exit For
                Next
                If .ListIndex < 0 And Trim(Nvl(rsTemp!细菌类别)) <> "" Then
                        .AddItem Nvl(rsTemp!细菌类别)
                        .ListIndex = .NewIndex
                End If
            End With
            
        End If
    End With
    
    gstrSql = "Select I.ID, 1 As 选择, I.编码, I.名称, Decode(D.缺省标志, 1, '←默认试验组', '') As 备注" & vbNewLine & _
            "From 检验细菌抗生素 D, 检验抗生素组 I" & vbNewLine & _
            "Where D.抗生素分组id = I.ID And 细菌id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngGermId)
    Set Me.vfgList.DataSource = rsTemp
    Call setListFormat(True)
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngGermId As Long) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngGermId-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
    Dim str编码 As String
    Dim intLoop As Integer
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(编码), 0) As 编码, Nvl(Max(Length(编码)), 0) As 长度 From 检验细菌"

'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
'            Call SQLTest
        With rsTemp
            If !长度 <> 0 And !长度 <= Me.txt编码.MaxLength Then
                If Val(!编码) = 0 Then
                    For intLoop = 1 To Len(!编码)
                        If IsNumeric(Mid(!编码, intLoop, 1)) Then
                            str编码 = str编码 & Format((Val(Mid(!编码, intLoop)) + 1), String(Len(Mid(!编码, intLoop)), "0"))
                            Exit For
                        Else
                            str编码 = str编码 & Mid(!编码, intLoop, 1)
                        End If
                    Next
                    Me.txt编码.Text = str编码
                Else
                    Me.txt编码.Text = Format(Val(!编码) + 1, String(!长度, "0"))
                End If
            Else
                Me.txt编码.Text = Format(Val(!编码) + 1, String(Me.txt编码.MaxLength, "0"))
            End If
        End With
        
        '清除并设置备注值
        Me.txt中文.Text = "": Me.txt英文.Text = "": Me.txt缩写.Text = "": Me.txt结果.Tag = ""
    End If
    If Me.cbo细菌类型.ListIndex = -1 And Me.cbo细菌类型.ListCount > 0 Then Me.cbo细菌类型.ListIndex = 0
    If Me.cbo默认药敏.ListIndex = -1 And Me.cbo默认药敏.ListCount > 0 Then Me.cbo默认药敏.ListIndex = 0
    If Me.cbo默认方法.ListIndex = -1 And Me.cbo默认方法.ListCount > 0 Then Me.cbo默认方法.ListIndex = 0
    
    gstrSql = "Select I.ID, Decode(D.细菌id, Null, 0, 1) As 选择, I.编码, I.名称, Decode(D.缺省标志, 1, '←默认试验组', '') As 备注" & vbNewLine & _
            "From (Select * From 检验细菌抗生素 Where 细菌id = [1]) D, 检验抗生素组 I" & vbNewLine & _
            "Where D.抗生素分组id(+) = I.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngGermId)
    Set Me.vfgList.DataSource = rsTemp
    Call setListFormat(True)
    
    mlngGermId = lngGermId
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "增加", "修改"): Call Form_Resize
    Me.txt编码.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Enabled = False: Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngGermId)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long
    Dim strLists As String
    
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked Then
                strLists = strLists & "|" & .TextMatrix(lngCount, mCol.ID) & _
                    ";" & IIf(Trim(.TextMatrix(lngCount, mCol.备注)) <> "", 1, 0)
            End If
        Next
    End With
'    If strLists = "" Then MsgBox "没有选择药敏试验抗生素！", vbInformation, gstrSysName: zlEditSave = False: Exit Function
    strLists = Mid(strLists, 2)
    
    '一般特性检查
    If Me.cbo细菌类型.ListIndex = -1 Then
        MsgBox "请选择当前细菌类型！", vbInformation, gstrSysName
        Me.cbo细菌类型.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt编码.Text) > Val(String(Me.txt编码.MaxLength, "9")) Then
        MsgBox "编码太大！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt中文.Text) = "" Then
        MsgBox "请输入中文名称！", vbInformation, gstrSysName
        Me.txt中文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt中文.Text), vbFromUnicode)) > Me.txt中文.MaxLength Then
        MsgBox "中文名称超长（最多" & Me.txt中文.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt中文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt英文.Text), vbFromUnicode)) > Me.txt英文.MaxLength Then
        MsgBox "英文名称超长（最多" & Me.txt英文.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt英文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt缩写.Text), vbFromUnicode)) > Me.txt缩写.MaxLength Then
        MsgBox "缩写超长（最多" & Me.txt缩写.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt缩写.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtWHONET码.Text), vbFromUnicode)) > Me.txtWHONET码.MaxLength Then
        MsgBox "WHONET码超长（最多" & Me.txtWHONET码.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txtWHONET码.SetFocus: zlEditSave = 0: Exit Function
    End If
    '刘兴宏2008/03/17加入
    If txtEdit(mTxtIndex.idx_菌属分类).Tag = "" And Trim(txtEdit(mTxtIndex.idx_菌属分类).Text) <> "" Then
        MsgBox "菌属分类选择错误,请检查！", vbInformation, gstrSysName
        Me.txtEdit(mTxtIndex.idx_菌属分类).SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '并发操作检查
    If zlExistItem("检验细菌类型", "ID", Val(Me.cbo细菌类型.ItemData(Me.cbo细菌类型.ListIndex)), _
                                   Me.cbo细菌类型.List(Me.cbo细菌类型.ListIndex)) = False Then
        Me.zlRefresh (mlngGermId)
        zlEditSave = 0: Exit Function
    End If
    
    If zlExistItem("细菌检测方法", "名称", Mid(Me.cbo默认方法.Text, 4), Me.cbo默认方法.Text) = False Then
        Me.zlRefresh (mlngGermId)
        zlEditSave = 0: Exit Function
    End If
    
    
    '数据保存语句组织
    If Me.Tag = "增加" Then
        lngNewId = zlDatabase.GetNextId("检验细菌")
    Else
        If zlExistItem("检验细菌", "ID", mlngGermId, Trim(Me.txt中文.Text)) = False Then
            zlEditSave = 0: Exit Function
        End If
        lngNewId = mlngGermId
        
    End If
    gstrSql = "'" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt中文.Text) & "','" & Trim(Me.txt英文.Text) & "','" & Trim(Me.txt缩写.Text) & "'"
    gstrSql = gstrSql & "," & Me.cbo细菌类型.ItemData(Me.cbo细菌类型.ListIndex) & ",'" & Left(Me.cbo默认药敏.Text, 1) & "'"
    gstrSql = gstrSql & ",'" & Mid(Me.cbo默认方法.Text, 4) & "','" & Trim(Me.txtWHONET码.Text) & "'"
    
    '刘兴宏加入:
    Dim strTemp As String
    '  细菌类别_In   In 检验细菌.细菌类别%Type,
    strTemp = Me.cboEdit(mcboIndex.idx_细菌类别).Text: strTemp = Mid(strTemp, InStr(1, strTemp, ".") + 1)
    gstrSql = gstrSql & ",'" & strTemp & "'"
    '  细菌菌属_In   In 检验细菌.细菌菌属%Type,
    gstrSql = gstrSql & ",'" & Trim(txtEdit(mTxtIndex.idx_菌属分类).Tag) & "'"
    '  革兰氏分类_In In 检验细菌.革兰氏分类%Type,
    strTemp = Me.cboEdit(mcboIndex.idx_革兰氏染色).Text: strTemp = Mid(strTemp, InStr(1, strTemp, ".") + 1)
    gstrSql = gstrSql & ",'" & strTemp & "'"
            
    '更新最后的记录
    Call txt结果_LostFocus
    
    If Me.Tag = "增加" Then
        gstrSql = "Zl_检验细菌_Edit(1," & lngNewId & "," & gstrSql & ",'" & strLists & "','" & Me.txt结果.Tag & "')"
    Else
        gstrSql = "Zl_检验细菌_Edit(2," & lngNewId & "," & gstrSql & ",'" & strLists & "','" & Me.txt结果.Tag & "')"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "增加" Then mlngGermId = lngNewId
    Me.Enabled = False: Me.Tag = "": Call Form_Resize
    zlEditSave = mlngGermId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Sub cboEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    If KeyCode = vbKeyDelete Then cboEdit(Index).ListIndex = -1
End Sub

Private Sub cbo标本类型_Click()
    Dim intLoop As Integer
    Dim strItem() As String
    Dim strTmp As String
    
    Me.txt结果.Text = ""
    strTmp = Mid(cbo标本类型, InStr(cbo标本类型, "-") + 1)
    
    strItem = Split(Me.txt结果.Tag, ";")
    
    For intLoop = 1 To UBound(strItem)
        If Split(strItem(intLoop), ",")(0) = strTmp Then
            Me.txt结果.Text = Split(strItem(intLoop), ",")(1)
        End If
    Next
    
    If Me.txt结果.Tag <> "" And InStr(Me.txt结果.Tag, ",") = 0 And InStr(Me.txt结果.Tag, ";") = 0 Then
        Me.txt结果.Text = Me.txt结果.Tag
    End If
End Sub

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------
Private Sub cbo默认方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo默认药敏_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo细菌类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
 

Private Sub cmdEdit_Click(Index As Integer)
    Select Case Index
    Case mTxtIndex.idx_菌属分类
        If SelectItem(Me, txtEdit(mTxtIndex.idx_菌属分类), "", "检验细菌菌属", "检验细菌菌属选择器") = False Then Exit Sub
    End Select
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mlngGermId = 0
    
    With Me.cbo默认药敏
        .AddItem "R-耐药": .AddItem "I-中介": .AddItem "S-敏感"
    End With
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select 编码, 名称, 简码 From 细菌检测方法"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cbo默认方法.Clear
        Do While Not .EOF
            Me.cbo默认方法.AddItem !编码 & "-" & !名称
            .MoveNext
        Loop
    End With
    
    gstrSql = "SELECT 编码,名称 FROM 诊疗检验标本 order by 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cbo标本类型.Clear
        Do While Not .EOF
            Me.cbo标本类型.AddItem !编码 & "-" & !名称
            .MoveNext
        Loop
        If Me.cbo标本类型.ListCount > 0 Then
            Me.cbo标本类型.ListIndex = 1
        End If
    End With
    
    '------------------------------------------------------
    '刘兴宏:2008/03/18加入
    Call InitData
    '------------------------------------------------------
    Call setListFormat
    Exit Sub
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.vfgList.Height = Me.ScaleHeight - Me.vfgList.Top - 180
    If Me.Tag <> "" Then
        Me.BackColor = RGB(250, 250, 250)
        Me.vfgList.FocusRect = flexFocusHeavy
    Else
        Me.BackColor = &H8000000F
        Me.vfgList.FocusRect = flexFocusNone
    End If
End Sub
 

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
    
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlCommFun.OpenIme (True)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case mTxtIndex.idx_菌属分类
        If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If txtEdit(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If SelectItem(Me, txtEdit(Index), DelInvalidChar(Trim(txtEdit(Index).Text)), "检验细菌菌属", "检验细菌菌属选择器") = False Then Exit Sub
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub txtWHONET码_GotFocus()
    Me.txtWHONET码.SelStart = 0: Me.txtWHONET码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtWHONET码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt结果_LostFocus()
    Dim strTmp As String
    Dim intLoop As Integer
    Dim strItem() As String
    strTmp = Mid(cbo标本类型, InStr(cbo标本类型, "-") + 1)
    If Trim(txt结果.Text) <> "" Then
        '增加
        If InStr(txt结果.Tag, ";" & strTmp & ",") > 0 Then
            strItem = Split(txt结果.Tag, ";")
            Me.txt结果.Tag = ""
            For intLoop = 1 To UBound(strItem)
                If Split(strItem(intLoop), ",")(0) = strTmp Then
                    Me.txt结果.Tag = Me.txt结果.Tag & ";" & strTmp & "," & Me.txt结果.Text
                Else
                    Me.txt结果.Tag = Me.txt结果.Tag & ";" & strItem(intLoop)
                End If
            Next
        Else
            Me.txt结果.Tag = Me.txt结果.Tag & ";" & strTmp & "," & Me.txt结果
        End If
    Else
        '删除
        If InStr(txt结果.Tag, ";" & strTmp & ",") > 0 Then
            strItem = Split(txt结果.Tag, ";")
            Me.txt结果.Tag = ""
            For intLoop = 1 To UBound(strItem)
                If Split(strItem(intLoop), ",")(0) = strTmp Then
                    Me.txt结果.Tag = Me.txt结果.Tag & ""
                Else
                    Me.txt结果.Tag = Me.txt结果.Tag & ";" & strItem(intLoop)
                End If
            Next
        End If
    End If
End Sub

Private Sub txt缩写_GotFocus()
    Me.txt缩写.SelStart = 0: Me.txt缩写.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt缩写_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt英文_GotFocus()
    Me.txt英文.SelStart = 0: Me.txt英文.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt英文_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt中文_GotFocus()
    Me.txt中文.SelStart = 0: Me.txt中文.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt中文_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag = "" Then Exit Sub
    With Me.vfgList
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Col = mCol.备注 Then
            If .Cell(flexcpChecked, .Row, mCol.选择) <> flexChecked Then Exit Sub
            For lngCount = .FixedRows To .Rows - 1
                .TextMatrix(lngCount, mCol.备注) = IIf(lngCount = .Row, "←默认试验组", "")
            Next
        Else
            If .Cell(flexcpChecked, .Row, mCol.选择) = flexChecked Then
                .Cell(flexcpChecked, .Row, mCol.选择) = flexUnchecked
                .TextMatrix(.Row, mCol.备注) = ""
            Else
                .Cell(flexcpChecked, .Row, mCol.选择) = flexChecked
            End If
        End If
    End With
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfgList_DblClick
End Sub

