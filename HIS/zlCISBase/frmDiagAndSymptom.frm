VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiagAndSymptom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊断与病种对应"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8865
   Icon            =   "frmDiagAndSymptom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   2640
      Picture         =   "frmDiagAndSymptom.frx":000C
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1290
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&C)"
      Height          =   350
      Left            =   1350
      Picture         =   "frmDiagAndSymptom.frx":0156
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1290
   End
   Begin VB.CommandButton cmdDiag 
      Caption         =   "&P"
      Height          =   300
      Left            =   8460
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   870
      Width           =   285
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      Picture         =   "frmDiagAndSymptom.frx":02A0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "关闭(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7665
      TabIndex        =   10
      Top             =   5340
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -405
      Top             =   6090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagAndSymptom.frx":03EA
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagAndSymptom.frx":0984
            Key             =   "ItemStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagAndSymptom.frx":0F1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDiag 
      Height          =   300
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   2
      Top             =   885
      Width           =   7200
   End
   Begin VB.Frame fra应用于 
      Caption         =   "应用于(&B)"
      Height          =   1050
      Left            =   120
      TabIndex        =   5
      Top             =   4155
      Width           =   8640
      Begin VB.ComboBox cbo分类 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   570
         Width           =   2100
      End
      Begin VB.OptionButton opt病种应用范围 
         Caption         =   "仅应用于当前诊断(&1)"
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton opt病种应用范围 
         Caption         =   "应用于                        的所有诊断(&2)"
         Height          =   240
         Index           =   1
         Left            =   330
         TabIndex        =   8
         Top             =   630
         Width           =   5145
      End
      Begin VB.OptionButton opt病种应用范围 
         Caption         =   "应用于所属分类的所有诊断(&3)"
         Height          =   240
         Index           =   2
         Left            =   5595
         TabIndex        =   7
         Top             =   300
         Width           =   2730
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vs病种 
      Height          =   2820
      Left            =   105
      TabIndex        =   4
      Top             =   1245
      Width           =   8670
      _cx             =   15293
      _cy             =   4974
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDiagAndSymptom.frx":2C28
      ScrollTrack     =   0   'False
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   6570
      TabIndex        =   9
      Top             =   5340
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   330
      Picture         =   "frmDiagAndSymptom.frx":2C78
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "指定诊断(&Z)"
      Height          =   180
      Left            =   195
      TabIndex        =   1
      Top             =   930
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请选择指定的疾病诊断后，设置该疾病诊断所对应的保险病种。"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   390
      Width           =   5400
   End
End
Attribute VB_Name = "frmDiagAndSymptom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mbln编辑 As Boolean
Private mlng诊断ID As Long                  '诊疗项目ID
Private mlng分类id As Long
Private mblnChange As Boolean
Private mbln中医  As Boolean
Private Sub cmdClear_Click()
    Dim lngRow As Long
    With vs病种
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, .ColIndex("病种")) = ""
            .Cell(flexcpData, lngRow, .ColIndex("病种")) = ""
        Next
    End With
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int(glngSys / 100))
End Sub
Private Function MulitSelect诊断(strkey As String) As Boolean
    '------------------------------------------------------------------------------------------
    '功能:多选诊断信息
    '参数:strKey-条件索引值
    '返回:True选择了诊断信息,否则:选择失败!
    '编制:刘兴宏
    '日期:2007/06/17
    '------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSearchKey As String, strTittle As String, lngH As Long
    Dim vRect As RECT
    Dim rsTemp As New ADODB.Recordset
 
    On Error GoTo errHandle
    gstrSql = "" & _
        "   Select Distinct decode(a.类别,1,'西医','中医') as 类别, a.Id ,a.编码,a.名称,a.说明,a.编者 " & _
        "   From 疾病诊断目录 a,疾病诊断别名 b " & _
        "   Where a.id=b.诊断id and b.性质=1  and a.类别=" & IIf(mbln中医, 2, 1) & " and (a.编码 like [1] or a.名称 like [1] Or b.简码 like [2])" & _
        " and (a.撤档时间 Is Null Or a.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')) " & _
        "   Order by  类别,编码"
    
    
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
    
    
    
    vRect = zlControl.GetControlRect(txtDiag.hWnd)
    lngH = txtDiag.Height
    strSearchKey = gstrMatch & strkey & "%"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strSearchKey, CStr(UCase(strSearchKey)))
    If blnCancel = True Then
        If txtDiag.Enabled Then txtDiag.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "不存在指定的疾病诊断,请检查!"
        If txtDiag.Enabled Then txtDiag.SetFocus
        Exit Function
    End If
    txtDiag.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
    txtDiag.Tag = Nvl(rsTemp!ID)
    
    
    Call init所属分类(Val(txtDiag.Tag), 0)
    Call Init缺省病种(Val(txtDiag.Tag))
    MulitSelect诊断 = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub cmdDiag_Click()

    Dim rsTemp As New ADODB.Recordset
    Dim blnCancel As Boolean
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "" & _
    "   Select Id||'_' As Id ,上级id||decode(上级id,Null,Null,'_') As 上级id,编码,名称,0 As 末级,'' 说明,'' 编者 " & _
    "   From 疾病诊断分类 where 类别=" & IIf(mbln中医, 2, 1) & " Start With 上级id Is Null Connect By Prior Id=上级id " & _
    "   Union All " & _
    "   Select A.Id|| '_'||b.分类ID As Id,b.分类id||'_' As 上级id,a.编码,a.名称,1 As 末级,a.说明,a.编者 " & _
    "   From 疾病诊断目录 A,疾病诊断属类 b" & _
    "   Where a.ID = b.诊断ID and a.类别=" & IIf(mbln中医, 2, 1) & "" & _
    " and (a.撤档时间 Is Null Or a.撤档时间 >= To_Date('3000-01-01', 'yyyy-mm-dd')) "
    
    '-------------------------------------------------------------------------------------------------------------------------------
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
    '-------------------------------------------------------------------------------------------------------------------------------
    Set rsTemp = zlDatabase.ShowSelect(Me, gstrSql, 1, IIf(mbln中医, "中医", "西医") & "疾病诊断", True, , "选择指定的疾病诊断", , , , , , , blnCancel, False)
    If rsTemp Is Nothing Then Exit Sub
    If blnCancel = True Then Exit Sub
    txtDiag.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
    txtDiag.Tag = Split(Nvl(rsTemp!ID), "_")(0)
    
    Call init所属分类(Val(txtDiag.Tag), Val(Split(Nvl(rsTemp!ID), "_")(1)))
    Call Init缺省病种(Val(txtDiag.Tag))
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function Init缺省病种(ByVal lng诊断ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------
    '功能:根据诊断,加载已经设置好的对应病种
    '参数:lng诊断id-诊断id
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/17
    '---------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    Dim strSql As String, rsTemp As New ADODB.Recordset
    strSql = "Select a.险类,a.病种ID,c.编码||'-'||c.名称 as 病种 From 诊断病种对应 a ,保险病种 C Where  a.诊断id=[1] And a.病种id=c.Id"
    
    Err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng诊断ID)
    With vs病种
        For i = 1 To .Rows - 1
            rsTemp.Filter = "险类=" & .RowData(i)
            If rsTemp.EOF = True Then
                .TextMatrix(i, .ColIndex("病种")) = ""
                .Cell(flexcpData, i, .ColIndex("病种")) = ""
            Else
                .TextMatrix(i, .ColIndex("病种")) = Nvl(rsTemp!病种)
                .Cell(flexcpData, i, .ColIndex("病种")) = Nvl(rsTemp!病种ID)
            End If
        Next
    End With
    rsTemp.Close
    Init缺省病种 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function init所属分类(ByVal lng诊断ID As Long, ByVal lng分类id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------
    '功能:根据指定的诊断,将所属分类填加了combox控件中
    '参数:lng分类id-分类id
    '     lng诊断id-诊断id
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/08/17
    '---------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo ErrHand:
    strSql = " Select a.Id,a.编码,a.名称 From 疾病诊断分类 a,疾病诊断属类 b Where a.Id=b.分类id And b.诊断ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng诊断ID)
    With rsTemp
        cbo分类.Clear
        Do While Not .EOF
            cbo分类.AddItem Nvl(!编码) & "-" & Nvl(!名称)
            cbo分类.ItemData(cbo分类.NewIndex) = Val(Nvl(!ID))
            If Val(Nvl(!ID)) = lng分类id Then
                cbo分类.ListIndex = cbo分类.NewIndex
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 And cbo分类.ListIndex < 0 Then cbo分类.ListIndex = 0
    End With
    rsTemp.Close
    init所属分类 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub Load诊断信息()
    '------------------------------------------------------------------------------
    '功能:加载指定的诊断信息
    '参数:
    '返回:
    '编制:刘兴宏
    '日期:2007/08/17
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    If mlng诊断ID = 0 Then GoTo Init:
    On Error GoTo errHandle
    gstrSql = "Select id,编码,名称 From 疾病诊断目录 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlng诊断ID)
    If rsTemp.EOF Then
        GoTo Init:
    End If
    txtDiag.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
    txtDiag.Tag = Nvl(rsTemp!ID)
    
Init:
    Call init所属分类(Val(txtDiag.Tag), mlng分类id)
    Call Init缺省病种(Val(txtDiag.Tag))
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub cmdRestore_Click()
    '恢复相关的设置
    Call init所属分类(Val(txtDiag.Tag), mlng分类id)
    Call Init缺省病种(Val(txtDiag.Tag))
End Sub

Private Function IsValid() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:分析输入的病种是否有效
    '参数:
    '返回值:有效返回True,否则为False
    '------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTemp As String
    
    '检查
    With vs病种
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("病种"))) <> "" And Val(.Cell(flexcpData, i, .ColIndex("病种"))) = 0 Then
                MsgBox "所输入的病种不正确,请重新输入!", vbInformation + vbDefaultButton1, gstrSysName
                .Row = i
                If .RowIsVisible(i) = False Then
                    .TopRow = i
                End If
                .SetFocus
                Exit Function
            End If
        Next
    End With
    If opt病种应用范围(1).Value = True Then
        If cbo分类.ListIndex < 0 Then
            MsgBox "未选择指定的分类,请检查!", vbInformation + vbDefaultButton1, gstrSysName
            If cbo分类.Enabled Then cbo分类.SetFocus
            Exit Function
        End If
    End If
    If Val(txtDiag.Tag) = 0 Then
        MsgBox "未选择指定的分类,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        If txtDiag.Enabled Then txtDiag.SetFocus
        Exit Function
    End If
    IsValid = True
End Function


Private Sub cmdSave_Click()
    '功能:保证诊断与病种的对应关系
    Dim n As Long, str病种 As String, int病种应用范围 As Integer
   
    If IsValid() = False Then Exit Sub
    
    For n = 0 To opt病种应用范围.UBound
         If opt病种应用范围(n).Value = True Then
             int病种应用范围 = n
             Exit For
         End If
    Next
    str病种 = ""
    With vs病种
        For n = 1 To .Rows - 1
            If .RowData(n) <> 0 Then
                If Val(.Cell(flexcpData, n, .ColIndex("病种"))) <> 0 Then
                    str病种 = str病种 & "," & .RowData(n) & "|" & Val(.Cell(flexcpData, n, .ColIndex("病种")))
                End If
            End If
        Next
    End With
    If str病种 <> "" Then str病种 = Mid(str病种, 2)
    
    'Zl_诊断病种对应_Update
    gstrSql = "Zl_诊断病种对应_Update("
    '  诊断ID_In     In 疾病诊断目录.ID%Type,
    gstrSql = gstrSql & "" & Val(txtDiag.Tag) & ","
    '  分类id_In In 疾病诊断属类.分类id%Type := 0, --指定的分类ID
    If int病种应用范围 = 1 Then
        gstrSql = gstrSql & "" & cbo分类.ItemData(cbo分类.ListIndex) & ","
    Else
        gstrSql = gstrSql & "" & 0 & ","
    End If
    '  病种_In   In Varchar2 := Null, --病种id串,险类1|病种id1,险类2|病种id2.....
    gstrSql = gstrSql & "'" & str病种 & "',"
    '  应用_In   In Number := 0 --病种的应用范围:0-应用于当前项目;1-应用于指定分类;2-应用于所属分类
    gstrSql = gstrSql & "" & int病种应用范围 & ")"
    
    Err = 0: On Error GoTo ErrHand:
    Me.Enabled = False
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Me.Enabled = True
    MsgBox "保存成功!", vbInformation + vbDefaultButton1, gstrSysName
    Exit Sub
ErrHand:
    Me.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    If Init险类() = False Then Unload Me: Exit Sub
    Call Load诊断信息
    Call CtlEnableSet
    mblnChange = False
End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub opt病种应用范围_Click(Index As Integer)
    cbo分类.Enabled = opt病种应用范围(1).Value
End Sub

Private Sub txtDiag_Change()
    txtDiag.Tag = ""
    mblnChange = True
End Sub

Private Sub txtDiag_GotFocus()
    zlControl.TxtSelAll txtDiag
End Sub

Private Sub txtDiag_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub

    
    strTemp = UCase(Trim(Me.txtDiag.Text))
    If strTemp = "" Then mlng诊断ID = 0: Me.txtDiag.Tag = "": Me.txtDiag.Text = "": Exit Sub

    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    If MulitSelect诊断(strTemp) = False Then
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Function Init险类() As Boolean
    '--------------------------------------------------------------------------------------------------------------------
    '功能:初始险类
    '参数:
    '返回:设置成功,返回true,否则返回false
    '编制:刘兴宏
    '日期:2007/08/17
    '--------------------------------------------------------------------------------------------------------------------
    Dim rs险类 As New ADODB.Recordset, i As Long
    Err = 0: On Error GoTo ErrHand:
    
    gstrSql = "Select 序号,名称 From 保险类别 where 医院编码 is not null"
    Call zlDatabase.OpenRecordset(rs险类, gstrSql, Me.Caption)
    If rs险类.RecordCount = 0 Then
        MsgBox "未安装相关的医保,请找系统管理员!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    With vs病种
        If rs险类.RecordCount = 0 Then
            .Rows = 2
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .Cell(flexcpData, 1, i) = ""
                .RowData(1) = 0
            Next
            .Editable = flexEDNone
            Exit Function
        End If
        .Rows = rs险类.RecordCount + 1
        i = 1
        Do While Not rs险类.EOF
            .RowData(i) = Val(Nvl(rs险类!序号))
            .TextMatrix(i, .ColIndex("险类")) = Nvl(rs险类!名称)
            .TextMatrix(i, .ColIndex("病种")) = ""
            .Cell(flexcpData, i, .ColIndex("病种")) = ""
            i = i + 1
             rs险类.MoveNext
        Loop
        If mbln编辑 = False Then
            .Editable = flexEDKbdMouse
            .ColComboList(.ColIndex("病种")) = ""
        Else
            .Editable = flexEDKbdMouse
            .ColComboList(.ColIndex("病种")) = "..."
        End If
    End With
    Init险类 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Sub ShowEdit(ByVal frmMain As Object, ByVal lng诊断ID As Long, ByVal lng分类id As Long, ByVal bln编辑 As Boolean, ByVal bln中医 As Boolean)
    '-------------------------------------------------------------------------------------------------
    '功能:显示编辑窗口,程序入口
    '参数:frmMain-父窗口
    '     lng诊断id=诊断id
    '     lng分类id=默认的分类id(主要应用在应用于指定分类上)
    '     bln编辑=是否可以编辑
    '编制:刘兴宏
    '日期:2007/08/17
    '-------------------------------------------------------------------------------------------------
    On Error Resume Next
    mblnFirst = True
    mlng诊断ID = lng诊断ID
    mlng分类id = lng分类id
    mbln编辑 = bln编辑
    mbln中医 = bln中医
    
    Me.Show 1, frmMain
End Sub
Private Sub CtlEnableSet()
    '---------------------------------------------------------------------------------------------------------------------
    '功能:设置相关控件的Enable
    '参数:
    '编制:刘兴宏
    '日期:2007/08/17
    '---------------------------------------------------------------------------------------------------------------------
    txtDiag.Enabled = mbln编辑
    cmdDiag.Enabled = mbln编辑
    vs病种.Editable = flexEDKbdMouse
    If mbln编辑 Then
        vs病种.Editable = flexEDKbdMouse
    Else
        vs病种.Editable = flexEDNone
    End If
    cmdClear.Visible = mbln编辑
    cmdRestore.Visible = mbln编辑
    cmdSave.Visible = mbln编辑
    fra应用于.Enabled = mbln编辑
    opt病种应用范围(0).Enabled = mbln编辑
    opt病种应用范围(1).Enabled = mbln编辑
    opt病种应用范围(2).Enabled = mbln编辑
    cbo分类.Enabled = mbln编辑
End Sub


Private Sub vs病种_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vs病种
        Select Case Col
        Case .ColIndex("病种")
             .ColComboList(0) = "..."
        End Select
    End With
End Sub

Private Sub vs病种_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs病种
        Select Case Col
        Case .ColIndex("险类")
             Cancel = True
        End Select
    End With

End Sub

Private Sub vs病种_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case vs病种.ColIndex("病种")
        '选择病种
        Call Select病种(vs病种.RowData(Row), "")
    Case Else
    End Select
End Sub
Private Function Select病种(ByVal lng险类 As Long, ByVal strkey As String)
    '---------------------------------------------------------------------------------
    '功能:选择指定险类的病种
    '参数:lng险类-险类
    '返回:选择成功,返回ture,否则返回False
    '编制:刘兴宏
    '日期:2007/08/15
    '---------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strLeft As String
    Dim blnCancel As Boolean
    
    Dim vRect As RECT

    strLeft = gstrMatch
    
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
    Err = 0: On Error GoTo ErrHand:
    
    Dim sngX As Single, sngY As Single
    Call CalcPosition(sngX, sngY, vs病种)
     
     If strkey <> "" Then
        strkey = strLeft & strkey & "%"
        gstrSql = "" & _
          "   Select Id, 编码, 名称, 简码, decode('0','普通病','1','慢性病','2','特种病','') As 类别, 特殊封顶线, 封顶线金额 " & _
          "    From 保险病种 " & _
          "    Where 险类 = [1] And (编码 Like [2] Or 名称 Like [2] Or 简码 Like [3]) " & _
          "    Order by 编码"
    Else
        gstrSql = "" & _
          "   Select Id, 编码, 名称, 简码, decode('0','普通病','1','慢性病','2','特种病','') As 类别, 特殊封顶线, 封顶线金额 " & _
          "    From 保险病种 " & _
          "    Where 险类 = [1]" & _
          "    Order by 编码"
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "保险病种选择", False, "", "", False, False, True, sngX, sngY - vs病种.CellHeight, vs病种.CellHeight, blnCancel, False, False, lng险类, strkey, CStr(UCase(strkey)))
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then
        MsgBox "不存在指定的病种,请检查!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    With vs病种
        .TextMatrix(.Row, .ColIndex("病种")) = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
        If strkey <> "" Then
            .EditText = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
        End If
        .Cell(flexcpData, .Row, .ColIndex("病种")) = Nvl(rsTemp!ID)
    End With
    Select病种 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vs病种_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vs病种_DblClick()
   Select Case vs病种.Col
    Case vs病种.ColIndex("病种")
        '选择病种
        
        vs病种.ColComboList(vs病种.ColIndex("病种")) = ""
        
        
    Case Else
    End Select
End Sub

Private Sub vs病种_EnterCell()
    vs病种.ColComboList(vs病种.ColIndex("病种")) = "..."
End Sub

Private Sub vs病种_GotFocus()
    With vs病种
        .BackColorSel = &H8000000D
'        .GridColor = &H0&
'        .GridColorFixed = &H0&
    End With
End Sub

Private Sub vs病种_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    If vs病种.Col = vs病种.ColIndex("病种") And KeyCode <> vbKeyReturn Then
       vs病种.ColComboList(vs病种.ColIndex("病种")) = ""
    End If
End Sub

Private Sub vs病种_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
      Dim strkey As String
        
        With vs病种
        
            Select Case Col
            Case .ColIndex("病种")
                strkey = Trim(vs病种.EditText)
                strkey = Replace(strkey, Chr(vbKeyReturn), "")
                strkey = Replace(strkey, Chr(10), "")
                If strkey = "" Then Exit Sub
                .Cell(flexcpData, Row, Col) = ""
                If KeyCode <> vbKeyReturn Then Exit Sub
               If Select病种(.RowData(.Row), strkey) = False Then
                    '选择失败
                    
                End If
                .ColComboList(.ColIndex("病种")) = "..."
                .Col = 1
                .SetFocus
            End Select
        End With
End Sub

Private Sub vs病种_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vs病种_LostFocus()
    With vs病种
        .BackColorSel = &H8000000C
'        .GridColor = &H808080
'        .GridColorFixed = &H808080
    End With
End Sub
Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        x = objPoint.x * 15 'objBill.Left +
        y = objPoint.y * 15 + objBill.Height '+ objBill.Top
    Else
        x = objPoint.x * 15 + objBill.CellLeft
        y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub



