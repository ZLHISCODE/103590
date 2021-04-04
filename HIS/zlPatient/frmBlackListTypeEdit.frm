VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBlackListTypeEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "不良行为分类"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7440
   Icon            =   "frmBlackListTypeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7440
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "基本信息"
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   195
      Width           =   4890
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   3540
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "有效期间"
         Top             =   1095
         Width           =   600
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "简码"
         Top             =   1095
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   840
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "名称"
         Top             =   705
         Width           =   3675
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "编码"
         Top             =   345
         Width           =   1500
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "个月"
         Height          =   180
         Left            =   4200
         TabIndex        =   8
         Top             =   1155
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   765
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编号(&U)"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   0
         Top             =   405
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "有效期间(&Q)"
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   6
         Top             =   1155
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5235
      TabIndex        =   10
      Top             =   330
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5235
      TabIndex        =   11
      Top             =   765
      Width           =   1100
   End
   Begin VB.Frame fra规则 
      Height          =   120
      Left            =   1155
      TabIndex        =   12
      Top             =   1935
      Width           =   6420
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGridRule 
      Height          =   2835
      Left            =   60
      TabIndex        =   9
      Top             =   2235
      Width           =   7305
      _cx             =   12885
      _cy             =   5001
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBlackListTypeEdit.frx":06EA
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
   Begin VB.Label Label1 
      Caption         =   "不良行为控制"
      Height          =   270
      Left            =   45
      TabIndex        =   14
      Top             =   1935
      Width           =   1170
   End
End
Attribute VB_Name = "frmBlackListTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gTypeEdit
    EM_Ty_增加 = 0
    EM_Ty_修改
    EM_Ty_删除
    EM_Ty_规则调整
    EM_Ty_查看
End Enum
Private mbytEditType As gTypeEdit
Private mfrmMain As Object
Private mstrCode As String
Private mblnChange As Boolean     '是否改变了
Private mintSuccess As Integer
Private mblnFirst As Boolean
Private mblnSys As Boolean
Private mblnUnLoad As Boolean

Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytEditType As gTypeEdit, Optional strCode As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:编辑不良行为类别
    '入参:frmMain-调用的主窗体
    '    bytEditType-编辑类别:0-新增;1-修改;2-仅修改控制方式;3-查看;
    '     strCode-编码,新增时传入空
    '返回:编辑成功返回True,否则为False
    '编制:刘兴洪
    '日期:2018-11-08 17:01:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytEditType = bytEditType: mintSuccess = 0
    Set mfrmMain = frmMain: mstrCode = strCode: mblnFirst = True
    mblnUnLoad = False
    
    Me.Show 1, frmMain
    zlShowEdit = mintSuccess > 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetInputDefineSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关控件输入长度（得到数据库的表字段的长度）
    '编制:刘兴洪
    '日期:2018-11-09 17:06:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "SELECT 编码,名称,简码,有效期限 FROM 不良行为分类 Where Rownum<0 "
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "不良行为分类")
    
    txtEdit(1).MaxLength = rsTemp.Fields("编码").DefinedSize
    txtEdit(2).MaxLength = rsTemp.Fields("名称").DefinedSize
    txtEdit(3).MaxLength = rsTemp.Fields("简码").DefinedSize
    txtEdit(4).MaxLength = rsTemp.Fields("有效期限").DefinedSize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub SetCtrolEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控制的enable属性
    '编制:刘兴洪
    '日期:2018-11-13 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, i As Long
    On Error GoTo errHandle
    blnEdit = (mbytEditType = EM_Ty_增加 Or mbytEditType = EM_Ty_修改) And mblnSys = False
    txtEdit(1).Enabled = mbytEditType = EM_Ty_增加
    txtEdit(2).Enabled = blnEdit
    txtEdit(3).Enabled = blnEdit
    txtEdit(4).Enabled = mbytEditType = EM_Ty_增加 Or mbytEditType = EM_Ty_修改
    
    For i = 1 To txtEdit.UBound
        txtEdit(i).BackColor = IIf(txtEdit(i).Enabled, &H80000005, &H8000000F)
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub

 
Private Function ReadData(ByVal strCode As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据编码读取数据
    '入参:strCode-当前编码
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-09 17:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String

    On Error GoTo errHandle
   
    mblnSys = False
    If mbytEditType = 0 Then
        '增加
        txtEdit(1).Text = zlDatabase.GetMax("不良行为分类", "编码", txtEdit(1).MaxLength)
        Call LoadRuleData("")
        Call SetCtrolEnabled
        ReadData = True
        Exit Function
    End If
     
    strSQL = "" & _
    "   SELECT 编码,名称,简码,有效期限 ,是否固定 FROM 不良行为分类  Where 编码=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode)
            
    If rsTemp.EOF Then
        MsgBox "未找到编码为“" & strCode & "”的不良行为原因数据，请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    txtEdit(1).Text = Nvl(rsTemp!编码)
    txtEdit(2).Text = Nvl(rsTemp!名称)
    txtEdit(3).Text = Nvl(rsTemp!简码)
    txtEdit(4).Text = IIf(Val(Nvl(Nvl(rsTemp!有效期限))) = 0, "", Val(Nvl(Nvl(rsTemp!有效期限))))
    
    mblnSys = Val(Nvl(rsTemp!是否固定)) = 1
    
    '加载控制规则
    Call LoadRuleData(Nvl(rsTemp!名称))
    Call SetCtrolEnabled
     
    ReadData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function GetSplitRuleValue(ByVal str控制规则 As String, Optional str控制符_out As String, Optional str数次_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据控制规则值,返回指定的行
    '入参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-12 15:22:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str控制符 As String, str数次 As String, varTemp As Variant
    
    On Error GoTo errHandle
    If InStr(1, str控制规则, ">=") > 0 Then
        varTemp = Split(str控制规则, ">=")
        str控制符 = ">="
        str数次 = Val(varTemp(1))
    ElseIf InStr(1, str控制规则, ">") > 0 Then
        varTemp = Split(str控制规则, ">")
        str控制符 = ">"
        str数次 = Val(varTemp(1))
    ElseIf InStr(1, str控制规则, "<=") > 0 Then
        varTemp = Split(str控制规则, "<=")
        str控制符 = "<="
        str数次 = Val(varTemp(1))
    ElseIf InStr(1, str控制规则, "<") > 0 Then
        varTemp = Split(str控制规则, "<")
        str控制符 = "<"
        str数次 = Val(varTemp(1))
    ElseIf InStr(1, str控制规则, "=") > 0 Then
        varTemp = Split(str控制规则, "=")
        str控制符 = "="
        str数次 = Val(varTemp(1))
    ElseIf InStr(1, str控制规则, "-") > 0 Then
        str控制符 = "数次范围"
        str数次 = str控制规则
    Else
        Exit Function
    End If
    str控制符_out = str控制符
    str数次_Out = str数次
    GetSplitRuleValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function GetRuleRowFromRuleValue(ByVal str控制符 As String, ByVal str数次 As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据规则，获取指定的行
    '入参:str控制符-控制符:>=;<=等
    '     str数次
    '返回:找到返回指定的行，否则返回-1
    '编制:刘兴洪
    '日期:2018-11-12 15:27:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    On Error GoTo errHandle
    With vsGridRule
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("控制符")) = str控制符 And .TextMatrix(i, .ColIndex("数次")) = str数次 Then
                GetRuleRowFromRuleValue = i: Exit Function
            End If
        Next
    End With
    GetRuleRowFromRuleValue = -1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadRuleData(ByVal str行为类别 As String) As Boolean  '
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载行为类别所对应的控制规则
    '编制:刘兴洪
    '日期:2018-11-12 14:35:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, cllType As Collection, blnFind As Boolean
    Dim varTemp As Variant, strTemp As String, lngRow As Long
    Dim str控制符 As String, str数次 As String
    Dim rsTemp As ADODB.Recordset, rs预约方式 As ADODB.Recordset
    On Error GoTo errHandle
    
     
    strSQL = "Select 编码,名称 From 预约方式 order by 编码"
    Set rs预约方式 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set cllType = New Collection
    With rs预约方式
        Do While Not .EOF
            cllType.Add Nvl(!名称)
            .MoveNext
        Loop
    End With
    
    
    strSQL = "" & _
    "   Select a.应用场合,a.行为类别,a.预约方式,a.序号,a.控制规则,a.控制方式  " & _
    "   From 不良行为控制 A" & _
    "   where 行为类别=[1]" & _
    "   Order by 序号 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str行为类别)
   
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Nvl(!预约方式) <> "" Then
                blnFind = False
                rs预约方式.Filter = "名称='" & !预约方式 & "'"
                If rs预约方式.EOF Then
                    cllType.Add Nvl(!预约方式)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    vsGridRule.Redraw = flexRDNone
    Call InitRuleGridColumHead(cllType)
    
    rsTemp.Sort = "序号"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With vsGridRule
        .Clear 1
        Do While Not rsTemp.EOF
            strTemp = Nvl(rsTemp!控制规则)
            If GetSplitRuleValue(strTemp, str控制符, str数次) Then
                lngRow = GetRuleRowFromRuleValue(str控制符, str数次)
                If lngRow = -1 Then
                    If .TextMatrix(.Rows - 1, .ColIndex("控制符")) <> "" Then .Rows = .Rows + 1
                    lngRow = .Rows - 1
                    
                    .TextMatrix(lngRow, .ColIndex("控制符")) = str控制符
                    .TextMatrix(lngRow, .ColIndex("数次")) = str数次
                End If
                If Nvl(rsTemp!应用场合) = "预约" Then
                    If Trim(Nvl(rsTemp!预约方式)) = "" Then
                        .TextMatrix(lngRow, .ColIndex("所有预约")) = decode(Val(Nvl(rsTemp!控制方式)), 1, "禁止", 2, "提示", "")
                    Else
                        .TextMatrix(lngRow, .ColIndex(Trim(Nvl(rsTemp!预约方式)))) = decode(Val(Nvl(rsTemp!控制方式)), 1, "禁止", 2, "提示", "")
                    End If
                Else
                     .TextMatrix(lngRow, .ColIndex(Trim(Nvl(rsTemp!应用场合)))) = decode(Val(Nvl(rsTemp!控制方式)), 1, "禁止", 2, "提示", "")
                End If
            End If
           rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    LoadRuleData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
 
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    If mbytEditType <> 0 Then
        mblnChange = False: Unload Me
        Exit Sub
    End If
    
    mstrCode = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(1).Text = zlDatabase.GetMax("不良行为分类", "编码", txtEdit(1).MaxLength)
    '规则保留上次的不变
    
    mblnChange = False
    txtEdit(1).SetFocus
End Sub

Private Function IsValid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:分析输入有关内容是否有效
    '返回:有效返回True,否则为False
    '编制:刘兴洪
    '日期:2018-11-09 17:22:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, varTemp As Variant, varData As Variant, blnHaveData As Boolean
    Dim strTemp As String
    
    On Error GoTo errHandle
    For i = 1 To 3
        txtEdit(i).Text = Trim(txtEdit(i).Text)
        strTemp = txtEdit(i).Text
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox txtEdit(i).Tag & "不能超过" & Int(txtEdit(i).MaxLength / 2) & "个汉字" & "或" & txtEdit(i).MaxLength & "个字母。", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(i)
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox txtEdit(i).Tag & "中含有非法字符。", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(i)
            Exit Function
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)
    If IsNumeric(txtEdit(4).Text) = False And txtEdit(4).Text <> "" Then
        MsgBox "有效期限输入必须是数字型。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(4)
        Exit Function
    End If
    If Val(txtEdit(4).Text) > 99999 Then
        MsgBox "有效期限最大只能输入99999个月。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(4)
        Exit Function
    End If
    
    If Val(txtEdit(4).Text) < 0 Then
        MsgBox "有效期限输入必须大于等于0个月。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(4)
        Exit Function
    End If
    
    If Len(txtEdit(1).Text) = 0 Or Trim(txtEdit(1).Text) = "" Then
        MsgBox "编码不能为空。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(1)
        Exit Function
    End If
    
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(2)
        Exit Function
    End If
    With vsGridRule
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("控制符"))) <> "" Then
                If Trim(.TextMatrix(i, .ColIndex("数次"))) = "" Then
                    MsgBox "第" & i & "行未输入数次!", vbInformation + vbOKOnly, gstrSysName
                    .Row = i: .Col = .ColIndex("数次")
                    .SetFocus
                    Exit Function
                End If
                If .TextMatrix(i, .ColIndex("控制符")) = "数次范围" Then
                    If InStr(.TextMatrix(i, .ColIndex("数次")), "-") = 0 Then
                        MsgBox "第" & i & "行的数次格式不正确，应该输入的格式为:000-999!", vbInformation + vbOKOnly, gstrSysName
                        .Row = i: .Col = .ColIndex("数次")
                        .SetFocus
                        Exit Function
                    End If
                End If
                blnHaveData = False
                For j = .ColIndex("数次") + 1 To .Cols - 1
                    If .TextMatrix(i, .ColIndex("所有预约")) = "禁止" Then
                       If InStr(";所有预约;挂号;结帐;入院;出院;数次;", ";" & .ColKey(j) & ";") = 0 And .TextMatrix(0, j) = "预约方式" Then
                            If .TextMatrix(i, j) <> "禁止" And j <> .ColIndex("所有预约") And Trim(.TextMatrix(i, j)) <> "" Then
                                MsgBox "第" & i & "行的所有预约业务都是禁止状态，其他预约也应该为禁止或不设置!", vbInformation + vbOKOnly, gstrSysName
                                .Row = i: .Col = j
                                .SetFocus
                                Exit Function
                            End If
                       End If
                    End If
                                    
                    If .TextMatrix(i, j) <> "" Then blnHaveData = True
                Next
                If Not blnHaveData Then
                    MsgBox "第" & i & "行未设置相关的控制方式!", vbInformation + vbOKOnly, gstrSysName
                    .Row = i: .Col = .ColIndex("数次")
                    .SetFocus
                    Exit Function
                End If
           
                For j = i + 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("控制符")) = .TextMatrix(j, .ColIndex("控制符")) Then
                        If .TextMatrix(i, .ColIndex("控制符")) <> "数次范围" Then
                            If Val(.TextMatrix(i, .ColIndex("数次"))) = Val(.TextMatrix(j, .ColIndex("数次"))) Then
                                MsgBox "第" & i & "行制定的规则与第" & j & "行的规则相同，请合并!", vbInformation + vbOKOnly, gstrSysName
                                .Row = i: .Col = .ColIndex("控制符")
                                .SetFocus
                                Exit Function
                            End If
                        Else
                            If .TextMatrix(i, .ColIndex("数次")) = .TextMatrix(j, .ColIndex("数次")) Then
                                If Val(.TextMatrix(i, .ColIndex("数次"))) = Val(.TextMatrix(j, .ColIndex("数次"))) Then
                                    MsgBox "第" & i & "行制定的规则与第" & j & "行的规则相同，请合并!", vbInformation + vbOKOnly, gstrSysName
                                    .Row = i: .Col = .ColIndex("控制符")
                                    .SetFocus
                                    Exit Function
                                End If
                            End If
                            varData = Split(.TextMatrix(i, .ColIndex("数次")) & "-", "-")
                            varTemp = Split(.TextMatrix(j, .ColIndex("数次")) & "-", "-")
                            If Val(varData(0)) = Val(varTemp(0)) And Val(varData(1)) = Val(varTemp(1)) Then
                                MsgBox "第" & i & "行制定的规则与第" & j & "行的规则相同，请合并!", vbInformation + vbOKOnly, gstrSysName
                                .Row = i: .Col = .ColIndex("控制符")
                                .SetFocus
                                Exit Function
                            End If
                        End If
                    End If
                Next
                If CheckRuleDataValid(i) = False Then
                    MsgBox "第" & i & "行未输入控制场合，请检查!", vbInformation + vbOKOnly, gstrSysName
                    .Row = i: .Col = 2
                    .SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
    IsValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-09 17:23:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Long, cllPro As Collection, strSQL As String, blnDel As Boolean
    Dim blnTran As Boolean, strTemp As String, str规则 As String, strRule As String
    On Error GoTo errHandle
    Set cllPro = New Collection
    
    If mbytEditType <> EM_Ty_规则调整 Then
        
        '    Zl_不良行为分类_Update
        strSQL = "Zl_不良行为分类_Update("
        '  操作_In     Number, 0-增加;1-修改
        strSQL = strSQL & "" & IIf(mbytEditType = 0, 0, 1) & ","
        '  编码_In     不良行为分类.编码%Type,
        strSQL = strSQL & "'" & txtEdit(1).Text & "',"
        '  名称_In     不良行为分类.名称%Type,
        strSQL = strSQL & "'" & txtEdit(2).Text & "',"
        '  简码_In     不良行为分类.简码%Type,
        strSQL = strSQL & "'" & txtEdit(3).Text & "',"
        '  是否固定_In 不良行为分类.是否固定%Type := 0,
        strSQL = strSQL & "" & IIf(mblnSys, "1", "0") & ","
        '  有效期限_In 不良行为分类.有效期限%Type := Null
        strSQL = strSQL & "" & IIf(Val(txtEdit(4).Text) = 0, "NULL", Val(txtEdit(4).Text)) & ")"
        zlAddArray cllPro, strSQL
        
    End If
    '取规则保存
    blnDel = True
    With vsGridRule
        strTemp = "": str规则 = ""
        For i = .FixedRows To .Rows - 1
            
            '规则1|控制方式1||规则2|控制方式1||....
            '  --                 规则:如:>=10;<10等
            '  --                 控制方式:格式为应用场合:控制标志(0-不控制(不控制的，即始传入，也不保存);1-禁止;2-提醒):预约方式
            
            If Trim(.TextMatrix(i, .ColIndex("控制符"))) <> "" Then
                strTemp = Trim(.TextMatrix(i, .ColIndex("控制符")))
                If strTemp = "数次范围" Then
                    strTemp = .TextMatrix(i, .ColIndex("数次"))
                Else
                    strTemp = strTemp & Val(.TextMatrix(i, .ColIndex("数次")))
                End If
                  
                strRule = strTemp
                For j = .ColIndex("数次") + 1 To .Cols - 1
                   
                   If Trim(.TextMatrix(i, j)) <> "" Then
                        'strRule = strTemp
                        Select Case .ColKey(j)
                        Case "所有预约"
                            strRule = strRule & "|" & "预约:" & IIf(.TextMatrix(i, j) = "禁止", 1, 2)
                        Case "挂号", "入院", "出院", "结帐"
                            strRule = strRule & "|" & .ColKey(j) & ":" & IIf(.TextMatrix(i, j) = "禁止", 1, 2)
                        Case Else
                            strRule = strRule & "|" & "预约:" & IIf(.TextMatrix(i, j) = "禁止", 1, 2) & ":" & .ColKey(j)
                        End Select
                    End If
                Next
                If strRule <> "" Then
                    If zlCommFun.ActualLen(str规则 & "||" & strRule) > 4000 Then
                        str规则 = Mid(str规则, 3)
                        'Zl_不良行为控制规则_Update
                        strSQL = "Zl_不良行为控制规则_Update("
                        '  行为类别_In 不良行为控制.行为类别%Type,
                        strSQL = strSQL & "'" & txtEdit(2).Text & "',"
                         '  控制规则_In Varchar2,
                        strSQL = strSQL & "'" & str规则 & "',"
                        '  是否删除_In Number:=1
                        strSQL = strSQL & "" & IIf(blnDel, 1, 0) & ")"
                        zlAddArray cllPro, strSQL
                        blnDel = False
                        str规则 = ""
                    End If
                    str规则 = str规则 & "||" & strRule
                End If
            End If
       Next
    End With
    
    str规则 = Mid(str规则, 3)
    'Zl_不良行为控制规则_Update
    strSQL = "Zl_不良行为控制规则_Update("
    '  行为类别_In 不良行为控制.行为类别%Type,
    strSQL = strSQL & "'" & txtEdit(2).Text & "',"
     '  控制规则_In Varchar2,
    strSQL = strSQL & "'" & str规则 & "',"
    '  是否删除_In Number:=1
    strSQL = strSQL & "" & IIf(blnDel, 1, 0) & ")"
    zlAddArray cllPro, strSQL
    
    blnTran = True
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnLoad Then Unload Me: Exit Sub
    If txtEdit(2).Enabled And txtEdit(2).Visible Then txtEdit(2).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl Is vsGridRule Then Exit Sub
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    
    Call SetInputDefineSize '设置缺省的输入长度
    mblnUnLoad = Not ReadData(mstrCode) '读取数据
    
    mblnChange = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub
     

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Then zlCommFun.OpenIme True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("'}|,""/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
    If Index = 4 Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m数字式
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub InitRuleGridColumHead(ByVal cllType As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化控制规则网格列头
    '编制:刘兴洪
    '日期:2018-11-08 18:03:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, strSQL As String
    
    On Error GoTo errHandle
    With vsGridRule
        .Clear:
        .Rows = .FixedRows + 1
        .Cols = cllType.Count + 7
        i = 0
        
        .TextMatrix(0, i) = "控制规则"
        .TextMatrix(1, i) = "控制符": .ColWidth(i) = 800: i = i + 1
        
        .TextMatrix(0, i) = "控制规则"
        .TextMatrix(1, i) = "数次": .ColWidth(i) = 800: i = i + 1
        
        .TextMatrix(0, i) = "所有预约"
        .TextMatrix(1, i) = "所有预约": .ColWidth(i) = 800: i = i + 1
        For j = 1 To cllType.Count
            .TextMatrix(0, i) = "预约方式"
            .TextMatrix(1, i) = cllType(j): .ColWidth(i) = 800
            If Me.TextWidth(" " & cllType(j)) > 800 Then
                 .ColWidth(i) = Me.TextWidth(" " & cllType(j))
            End If
            i = i + 1
        Next
        
        .TextMatrix(0, i) = "挂号":
        .TextMatrix(1, i) = "挂号": .ColWidth(i) = 800: i = i + 1
        .TextMatrix(0, i) = "入院":
        .TextMatrix(1, i) = "入院": .ColWidth(i) = 800: i = i + 1
        .TextMatrix(0, i) = "出院"
        .TextMatrix(1, i) = "出院": .ColWidth(i) = 800: i = i + 1
        .TextMatrix(0, i) = "结帐"
        .TextMatrix(1, i) = "结帐": .ColWidth(i) = 800: i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(1, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            Select Case .ColKey(i)
            Case "数次"
                .ColAlignment(i) = flexAlignCenterCenter
            Case Else
                .ColAlignment(i) = flexAlignLeftCenter
            End Select
            .MergeCol(i) = True
        Next
        .ColComboList(.ColIndex("控制符")) = " |>=|>|=|<=|<|数次范围"
        
        .MergeCells = flexMergeRestrictAll
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeRow(0) = True
        .MergeRow(1) = True
        .Editable = IIf(mbytEditType = EM_Ty_查看 Or mbytEditType = EM_Ty_删除, flexEDNone, flexEDKbdMouse) '0-新增;1-修改;2-仅修改控制方式;3-查看;
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub DeleteRuleRow(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除指定的规则行
    '入参:lngRow-指定的行
    '编制:刘兴洪
    '日期:2018-11-12 16:54:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsGridRule
    
        If lngRow > .Rows - 1 Or lngRow < .FixedRows Then Exit Sub
        If lngRow = .FixedRows And lngRow = .Rows - 1 Then
            .Clear 1
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = ""
            Exit Sub
        End If
        If lngRow < .Rows - 1 Then
            .RemoveItem lngRow
            .Row = lngRow
            Exit Sub
        End If
        .RemoveItem lngRow
        .Row = .Rows - 1
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub vsGridRule_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim varData As Variant
    
    If mbytEditType = gTypeEdit.EM_Ty_查看 Then Exit Sub
    
    With vsGridRule
        Select Case Col
        Case .ColIndex("删除")
             Call DeleteRuleRow(Row)
        Case Else
        End Select
    End With
End Sub

 
Private Sub vsGridRule_ChangeEdit()
    If mbytEditType = EM_Ty_查看 Then Exit Sub
    With vsGridRule
       Select Case .Col
       Case .ColIndex("数次")
       Case Else
       End Select
    End With
End Sub

Private Sub vsGridRule_DblClick()
    
    If mbytEditType = EM_Ty_查看 Then Exit Sub
    With vsGridRule
        If .Row < 0 Then Exit Sub
        Select Case .Col
        Case .ColIndex("数次")
            .EditCell
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
        Case .ColIndex("控制符")
        Case .ColIndex("删除")
        
        Case Else
            If .TextMatrix(.Row, .Col) = "禁止" Then
                .TextMatrix(.Row, .Col) = "提示"
            ElseIf .TextMatrix(.Row, .Col) = "提示" Then
                .TextMatrix(.Row, .Col) = ""
            Else
                .TextMatrix(.Row, .Col) = "禁止"
            End If
        End Select
    End With
End Sub
 

Private Sub vsGridRule_EnterCell()
    
    If mbytEditType = EM_Ty_查看 Then Exit Sub
   
    With vsGridRule
        If .Row < 0 Then Exit Sub
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        Select Case .Col
        Case .ColIndex("数次")
        End Select
    End With
End Sub

Private Sub vsGridRule_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim varTemp As Variant, lngRow As Long
    
    If mbytEditType = EM_Ty_查看 Then Exit Sub
    
    With vsGridRule
        If .Row > .Rows - 1 Or .Row < 1 Then Exit Sub
    
         If KeyCode <> vbKeyReturn And (KeyCode <> Asc("*")) And KeyCode <> vbKeySpace _
            And KeyCode <> vbKeyShift Then
            If Shift = 1 And (KeyCode = 56 Or KeyCode <> Asc("*")) Then
                Call vsGridRule_CellButtonClick(.Row, .Col)
            Else
                If .Col = .ColIndex("控制符") Then .ColComboList(.Col) = ""
            End If
        End If
        '删除
        If KeyCode = vbKeyDelete Then
            Call vsGridRule_CellButtonClick(.Row, .Col)
            Exit Sub
        End If
    End With
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
     
    With vsGridRule
        Select Case .Col
        Case .ColIndex("数次")
            If (Trim(.TextMatrix(.Row, .ColIndex("数次"))) = "" And Trim(.TextMatrix(.Row, .ColIndex("控制符"))) = "" <> "") And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        Case Else
            If (Trim(.TextMatrix(.Row, .ColIndex("数次"))) = "" And Trim(.TextMatrix(.Row, .ColIndex("控制符"))) = "" <> "") And .Row = .Rows - 1 Then OS.PressKey vbKeyTab: Exit Sub
        End Select
            Call zlVsMoveGridCell(vsGridRule, .ColIndex("控制符"), , IIf(mbytEditType = EM_Ty_查看, False, True), lngRow)
    End With
    Call vsGridRule_EnterCell
End Sub
 


Private Sub vsGridRule_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

    Dim strKey As String, lngRow As Long
    
    If mbytEditType = EM_Ty_查看 Then Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsGridRule
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '暂不处理输入
        Select Case Col
        Case .ColIndex("控制符")
        Case .ColIndex("数次")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsGridRule, .ColIndex("控制符"), -1, True, lngRow)
    End With
End Sub

Private Sub vsGridRule_KeyPress(KeyAscii As Integer)
 
    If mbytEditType = EM_Ty_查看 Then Exit Sub
   
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    
    With vsGridRule
        If .Col <> .ColIndex("数次") Then KeyAscii = 0: Exit Sub
    End With
    Call VsFlxGridCheckKeyPress(vsGridRule, vsGridRule.Row, vsGridRule.Col, KeyAscii, m数字式)
End Sub

Private Sub vsGridRule_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngCashRow As Long
    
    If mbytEditType = EM_Ty_查看 Then Exit Sub
     
    With vsGridRule
        Select Case .Col
        Case .ColIndex("数交")
            Call VsFlxGridCheckKeyPress(vsGridRule, Row, Col, KeyAscii, m数字式)
        End Select
    End With
End Sub


Private Sub vsGridRule_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String, lngRow As Long
    If mbytEditType = EM_Ty_查看 Then Exit Sub
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsGridRule
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "") '暂不处理输入
        Select Case Col
        Case .ColIndex("控制符")
            .Col = .ColIndex("数次")
        Case Else
            'Call zlVsMoveGridCell(vsGridRule, .ColIndex("数次"), -1, True, Row)
        End Select

    End With
   
End Sub
Private Sub vsGridRule_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error GoTo Errhand:
    With vsGridRule
        If .MouseRow < 1 Or .MouseRow > .Rows - 1 Then Exit Sub
        If .MouseCol < 0 Or .MouseCol > .Cols - 1 Then Exit Sub
       If .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol)) Then Exit Sub
       .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol))
    End With
Errhand:
    Exit Sub
End Sub

Private Sub vsGridRule_LeaveCell()
    If mbytEditType = EM_Ty_查看 Then Exit Sub
    OS.OpenIme False
End Sub



Private Sub vsGridRule_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 
    
    If mbytEditType = EM_Ty_查看 Then Exit Sub
    
    '设置单元格的编辑长度
    With vsGridRule
       Select Case .Col
           Case .ColIndex("数次")
               .EditMaxLength = 50
       End Select
    End With
End Sub

Private Sub vsGridRule_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strInput As String, varTemp As Variant
    
    With vsGridRule
        If Row <= 0 Then Exit Sub
        
        strInput = Trim(.EditText): strInput = Replace(strInput, Chr(vbKeyReturn), ""): strInput = Replace(strInput, Chr(10), "")
        
        Select Case Col
        Case .ColIndex("控制符")
        Case .ColIndex("数次")
            If .TextMatrix(Row, .ColIndex("控制符")) <> "数次范围" And .TextMatrix(Row, .ColIndex("控制符")) <> "" Then
                If Not IsNumeric(strInput) And strInput <> "" Then
                    MsgBox "输入的数次必须为数字！", vbInformation, gstrSysName
                    .EditCell: .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                If zlDblIsValid(strInput, 5, False, False, 0, .ColKey(Col)) = False Then
                    Cancel = True: Exit Sub
                End If
            ElseIf .TextMatrix(Row, .ColIndex("控制符")) <> "" Then
                If InStr(strInput, "-") = 0 Then
                     MsgBox "输入的数次必须符合格式(XXXXX-XXXX)的范围格式,比如：1-5！", vbInformation, gstrSysName
                     Cancel = True: Exit Sub
                End If
                varTemp = Split(strInput, "-")
                If Val(varTemp(0)) > Val(varTemp(1)) Then
                     MsgBox "输入的数次范围下线大于了上线！", vbInformation, gstrSysName
                     Cancel = True: Exit Sub
                End If
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vsGridRule_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGridRule
        Select Case Col
        Case .ColIndex("控制符")
            If .ComboIndex < 0 Then .TextMatrix(Row, Col) = ""
        Case Else
        End Select
    End With
End Sub
Private Sub vsGridRule_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mbytEditType = EM_Ty_查看 Then Cancel = True: Exit Sub
    With vsGridRule
         .ComboList = ""
        Select Case Col
        Case .ColIndex("删除")
            .ComboList = "..."
           ' .CellButtonPicture = imgDel
            Exit Sub
        Case .ColIndex("控制符")
            .ComboList = ">=|>|=|<=|<|数次范围"
        Case .ColIndex("数次")
            
        Case Else
              .ComboList = " |禁止|提示"
        End Select
    End With
End Sub

Private Function CheckRuleDataValid(ByVal intRow As Integer) As Boolean
    '功能：检查不良行为控制规则表格输入数据的合法性
    '入参：intRow-不良行为控制规则表格的行
    Dim i As Integer, strTemp As String
    
    With vsGridRule
        For i = 2 To .Cols - 1
            If Trim(.TextMatrix(intRow, i)) <> "" Then
                strTemp = strTemp & .TextMatrix(intRow, i)
            End If
        Next
    End With
    CheckRuleDataValid = strTemp <> ""
End Function
