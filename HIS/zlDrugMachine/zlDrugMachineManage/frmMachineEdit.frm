VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmMachineEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品设备编辑"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   Icon            =   "frmMachineEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   6840
      TabIndex        =   15
      Top             =   5430
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      Height          =   345
      Left            =   5640
      TabIndex        =   14
      Top             =   5430
      Width           =   1095
   End
   Begin VB.CheckBox chkContine 
      Appearance      =   0  'Flat
      Caption         =   "连续新增设备接口(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Frame fraInfo 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdLink 
         Caption         =   "…"
         Height          =   300
         Index           =   1
         Left            =   7440
         Picture         =   "frmMachineEdit.frx":038A
         TabIndex        =   16
         ToolTipText     =   "生成连接串"
         Top             =   1200
         Width           =   330
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfInfo 
         Height          =   2415
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   7095
         _cx             =   12515
         _cy             =   4260
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      Begin VB.TextBox txtRemark 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1680
         Width           =   6015
      End
      Begin VB.CommandButton cmdLink 
         Height          =   300
         Index           =   0
         Left            =   7125
         Picture         =   "frmMachineEdit.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "测试连接"
         Top             =   1200
         Width           =   330
      End
      Begin VB.TextBox txtLink 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   5670
      End
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
         Height          =   300
         ItemData        =   "frmMachineEdit.frx":240E
         Left            =   1440
         List            =   "frmMachineEdit.frx":2410
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   750
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   4
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.同库房不同接口的药品剂型不可重复； 2.药品剂型不填，默认为所有药品剂型；"
         Height          =   180
         Index           =   5
         Left            =   360
         TabIndex        =   17
         Top             =   4680
         Width           =   6570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注(&E)"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   10
         Top             =   1710
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*地址(&I)"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*接口类型(&Y)"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*名称(&M)"
         Height          =   180
         Index           =   1
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*编号(&N)"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMachineEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_BILL As String = "药品库房,,3,2000|库房ID,,0,0|药品剂型,,3,4500|剂型编码,,0,0"

Private mblnShow As Boolean                     '显示状态（Load事件后的过程处理）
Private mblnReturn As Boolean                   '返回值； True确认；False取消
Private mblnEdited As Boolean                   '是否已经编辑；True是；False否
Private mbytState As Byte                       '窗体状态；0-查看；1-新增；2-修改
Private mlngID As Long                          '药品设备接口的ID
Private WithEvents mclsVSF As clsVSFlexGridEx
Attribute mclsVSF.VB_VarHelpID = -1

Public Function ShowMe(ByVal frmOwner As Form, ByVal bytState As Byte, Optional ByVal lngID As Long) As Boolean
'功能：
'参数：
'  frmOwner：主调窗体对象
'  bytState：窗体状态
'  lngID：药品设备接口的ID
'返回：True确认；False取消

    If lngID = 0 And bytState <> Val("1-新增") Then
        MsgBox "请传入接口ID！", vbInformation, GSTR_MSG
        Exit Function
    End If

    mbytState = bytState
    mlngID = lngID
    
    Me.Show vbModal, frmOwner
    ShowMe = mblnReturn

End Function

Private Sub cboType_Click()
    If Me.Visible = False Then Exit Sub
        
    Select Case Val(cboType.Text)
    Case Val("2-TOSHO"), Val("5-YUYAMA"), Val("6-高园")
        lbl(3).Caption = "*连接串(&I)"
        txtLink.Locked = True
        txtLink.BackColor = &H8000000F
    Case Else
        lbl(3).Caption = "*地址(&I)"
        txtLink.Locked = False
        txtLink.BackColor = &H80000005
    End Select
    cmdLink(0).Visible = Not (Val(cboType.Text) = 2 Or Val(cboType.Text) = 5 Or Val(cboType.Text) = 6)
    cmdLink(1).Visible = (Val(cboType.Text) = 2 Or Val(cboType.Text) = 5 Or Val(cboType.Text) = 6)
    
    txtLink.Text = ""
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdLink_Click(Index As Integer)
    If Index = 0 Then
        'WebService
        Dim objSOAP As Object
        
        Call CreateSOAP(objSOAP)
        
        With objSOAP
            If Trim(txtLink.Text) = "" Then
                MsgBox "请填写“" & cboType.Text & "”信息！", vbInformation, GSTR_MSG
                If txtLink.Enabled Then txtLink.SetFocus
                Exit Sub
            End If
            
            On Error Resume Next
            .MSSoapInit txtLink.Text
            If Err.Number = 0 Then
                txtLink.Tag = "1"           '标记连接成功
                MsgBox "连接成功！", vbInformation, GSTR_MSG
            Else
                txtLink.Tag = ""            '标记连接失败
                If objSOAP Is Nothing Then
                    MsgBox "“SoapClient”未安装，请联系技术人员！" & vbCrLf & _
                           "注意：SoapClient在WinXP下安装2.0版本。", _
                           vbInformation, GSTR_MSG
                Else
                    MsgBox "连接失败！", vbCritical, GSTR_MSG
                End If
            End If
            On Error GoTo 0
        End With
        Set objSOAP = Nothing
        
    Else
        'OLEDB连接串
        Dim msdLink As New MSDASC.DataLinks
        Dim cnTest As New ADODB.Connection
        
        If msdLink.PromptEdit(cnTest) Then
            On Error Resume Next
            Call cnTest.Open
            If Err.Number <> 0 Then
                txtLink.Text = ""
                txtLink.Tag = ""
                MsgBox "OLEDB连接串不正确，请检查！", vbInformation, GSTR_MSG
            Else
                txtLink.Text = cnTest.ConnectionString
                txtLink.Tag = "1"
            End If
            cnTest.Close
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    '检查
    If Verify() = False Then Exit Sub
    
    '保存
    If Save() = False Then Exit Sub
    
    If mbytState = enuEditState.新增 And chkContine.Value Then
        '连续新增
        txtCode.Text = ""
        txtName.Text = ""
        txtRemark.Text = ""
        
        With vsfInfo
            .Redraw = False
            .Clear 1
            .Rows = 1
            .Redraw = True
        End With
        
        txtCode.SetFocus
    Else
        Unload Me
    End If
    
    mblnReturn = True
End Sub

Private Sub Form_Activate()
    If mblnShow Then
        Screen.MousePointer = vbHourglass
        
        Call InitControls
        If mbytState <> enuEditState.新增 Then
            chkContine.Visible = False
            Call FillData
        End If
        
        mblnShow = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    mblnReturn = False
    mblnEdited = False
    
    Set mclsVSF = New clsVSFlexGridEx
    
    mblnShow = True         '本行放最后
End Sub

Private Sub InitControls()
    Dim arrTemp As Variant
    Dim i As Integer
    
    '控件位置
    With cmdLink(1)
        .Left = cmdLink(0).Left
        .Top = cmdLink(0).Top
        .Width = cmdLink(0).Width
        .Height = cmdLink(0).Height
    End With

    '控件最大字符数
    mdlMain.SetTextMaxLen txtCode, "药品设备接口.编号"
    mdlMain.SetTextMaxLen txtName, "药品设备接口.名称"
    mdlMain.SetTextMaxLen txtLink, "药品设备接口.连接信息"
    mdlMain.SetTextMaxLen txtRemark, "药品设备接口.备注"
    
    '绑定VSF
    With mclsVSF
        .Bunding = vsfInfo
        .Init
        .Head = MSTR_BILL
        .ColsReadonly = ""
        .Editable = EM_Modify
        .Repaint RT_Columns
    End With
    With vsfInfo
        .RowHeight(0) = 350
        .Rows = 2
        .ColComboList(.ColIndex("药品库房")) = "..."
        .ColComboList(.ColIndex("药品剂型")) = "..."
    End With
    
    '初始化类型控件的项
    With cboType
        .Clear
        arrTemp = Split(GSTR_TYPE, "|")
        For i = LBound(arrTemp) To UBound(arrTemp)
            If arrTemp(i) <> "" Then
                .AddItem arrTemp(i)
            End If
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    Call cboType_Click
    
    '遍历设置控件Enabled属性
    If mbytState = enuEditState.查看 Then
        For i = 0 To Me.Controls.Count - 1
            Select Case UCase(TypeName(Me.Controls(i)))
            Case "LABEL"
            Case Else
                Me.Controls(i).Enabled = False
            End Select
        Next
        cmdCancel.Enabled = True
    End If
    
End Sub

Private Sub FillData()
    Dim rsSQL As ADODB.Recordset, rsInfo As ADODB.Recordset
    Dim strInfo As String
    
    gstrSQL = "Select ID, 编号, 名称, 类型, 连接信息, 备注 From 药品设备接口 Where ID = [1] "
    
    On Error GoTo hErr
    Set rsSQL = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "药品设备接口", mlngID)
    If rsSQL.EOF = False Then
        '编号
        txtCode.Text = rsSQL!编号
        '名称
        txtName.Text = rsSQL!名称
        '类型
        cboType.ListIndex = rsSQL!类型 - 1
        '连接串或URL
        If IsNull(rsSQL!连接信息) Then
            txtLink.Text = ""
        Else
            txtLink.Text = mdlMain.Base64Decode(rsSQL!连接信息)
        End If
        '备注
        txtRemark.Text = gobjComLib.zlCommFun.NVL(rsSQL!备注)
        
        '库房与剂型
        gstrSQL = _
            "Select 库房ID, '【' || 库房编码 || '】' || 库房名称 as 药品库房 " & vbNewLine & _
            "    , f_List2str(Cast(Collect(剂型名称 Order By 剂型编码) As t_Strlist), '；') 药品剂型" & vbNewLine & _
            "    , f_List2str(Cast(Collect(剂型编码 Order By 剂型编码) As t_Strlist), '；') 剂型编码" & vbNewLine & _
            "From (Select a.编码 剂型编码, a.名称 剂型名称, d.库房id, b.编码 库房编码, b.名称 库房名称" & vbNewLine & _
            "      From 药品剂型 A, 部门表 B, 药品设备接口 C," & vbNewLine & _
            "         Xmltable('//root/bm' Passing c.扩展信息 Columns 库房id Number(18) Path 'id', 剂型编码 Varchar2(20) Path 'jxbm') D" & vbNewLine & _
            "      Where d.库房id = b.Id(+) And d.剂型编码 = a.编码(+) And c.Id = [1] )" & vbNewLine & _
            "Group By 库房id, 库房编码, 库房名称" & vbNewLine & _
            "Union All " & vbNewLine & _
            "Select 0, '', '', '' From Dual "
        Set rsInfo = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "库房与剂型", mlngID)
        mclsVSF.Recordset = rsInfo
        mclsVSF.Repaint RT_Rows
        rsInfo.Close
        
        If vsfInfo.Rows <= 1 Then vsfInfo.Rows = 2
    End If
    rsSQL.Close
    
    Exit Sub
    
hErr:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVSF = Nothing
'    Set mfrmOwner = Nothing
End Sub

Private Sub txtCode_GotFocus()
    Call gobjComLib.zlControl.TXTSelAll(txtCode)
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    '转大写
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = KeyAscii - 32
    '限制录入
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtLink_Change()
    txtLink.Tag = ""
End Sub

Private Sub txtLink_GotFocus()
    Call gobjComLib.zlControl.TXTSelAll(txtLink)
End Sub

Private Sub txtLink_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtName_GotFocus()
    Call gobjComLib.zlControl.TXTSelAll(txtName)
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    '限制录入
    If InStr("`~!@#$%^&*()+={}|[]\:"";'<>?,./", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtRemark_GotFocus()
    Call gobjComLib.zlControl.TXTSelAll(txtRemark)
End Sub

Private Sub txtRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Function Verify() As Boolean
    Dim l As Long, lngID As Long, lngCount As Long
    Dim blnFind As Boolean
    
    '编号
    If Trim(txtCode.Text) = "" Then
        MsgBox "“编号”内容未填写！", vbInformation, GSTR_MSG
        txtCode.SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtCode.Text, vbFromUnicode)) > txtCode.MaxLength Then
        MsgBox mdlMain.FormatString("“编号”内容超长（最多[1]字符）！", txtCode.MaxLength), vbInformation, GSTR_MSG
        txtCode.SetFocus
        Exit Function
    End If
    If VerifyString(txtCode.Text, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_") = False Then
        MsgBox "“编号”内容存在非法字符！", vbInformation, GSTR_MSG
        txtCode.SetFocus
        Exit Function
    End If
    
    '名称
    If Trim(txtName.Text) = "" Then
        MsgBox "“名称”内容未填写！", vbInformation, GSTR_MSG
        txtName.SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtName.Text, vbFromUnicode)) > txtName.MaxLength Then
        MsgBox mdlMain.FormatString("“名称”内容超长（最多[1]字符）！", txtName.MaxLength), vbInformation, GSTR_MSG
        txtName.SetFocus
        Exit Function
    End If
    If VerifyString(txtName.Text, "`~!@#$%^&*()+={}|[]\:"";'<>?,./", False) = False Then
        MsgBox "“名称”内容存在非法字符！", vbInformation, GSTR_MSG
        txtName.SetFocus
        Exit Function
    End If
    
    '地址
    If Trim(txtLink.Text) = "" Then
        If Val(cboType.Text) = Val("2-TOSHO") Then
            MsgBox mdlMain.FormatString("“[1]”内容未设置！", Split(lbl(3).Caption, "(")(0)), vbInformation, GSTR_MSG
            If cmdLink(1).Enabled And cmdLink(1).Visible Then cmdLink(1).SetFocus
        Else
            MsgBox mdlMain.FormatString("“[1]”内容未填写！", Split(lbl(3).Caption, "(")(0)), vbInformation, GSTR_MSG
            txtLink.SetFocus
        End If
        Exit Function
    End If
    If LenB(StrConv(txtLink.Text, vbFromUnicode)) > txtLink.MaxLength Then
        MsgBox mdlMain.FormatString("“[1]”内容超长（最多[2]字符）！", Split(lbl(3).Caption, "(")(0), txtLink.MaxLength), _
                vbInformation, _
                GSTR_MSG
        txtLink.SetFocus
        Exit Function
    End If
    
    '备注
    If Trim(txtRemark.Text) <> "" Then
        If LenB(StrConv(txtRemark.Text, vbFromUnicode)) > txtRemark.MaxLength Then
            MsgBox mdlMain.FormatString("“备注”内容超长（最多[1]字符）！", txtRemark.MaxLength), vbInformation, GSTR_MSG
            txtRemark.SetFocus
            Exit Function
        End If
        If VerifyString(txtRemark.Text, "`~!@#$%^&*()+={}|[]\:"";'<>?,./", False) = False Then
            MsgBox "“备注”内容存在非法字符！", vbInformation, GSTR_MSG
            txtRemark.SetFocus
            Exit Function
        End If
    End If
    
    '库房与剂型
    With vsfInfo
        For l = 1 To .Rows - 1
            lngID = Val(.TextMatrix(l, .ColIndex("库房ID")))
            If lngID > 0 Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = False Then
            MsgBox "请填写药品库房信息！", vbInformation, GSTR_MSG
            Exit Function
        End If
    End With
        
    '检查库房的药品剂型
    With vsfInfo
        lngCount = 0
        blnFind = False
        For l = 1 To .Rows - 1
            lngID = Val(.TextMatrix(l, .ColIndex("库房ID")))
'            '检查当前接口多库房的药品剂型，不允许出现默认空（即全选）的设置
'            If lngID <> 0 Then
'                lngCount = lngCount + 1
'                If Trim(.TextMatrix(l, .ColIndex("剂型编码"))) = "" Then
'                    blnFind = True
'                End If
'            End If
'            If lngCount > 0 And blnFind Then
'                MsgBox mdlMain.FormatString("“[1]”的当前接口已存在全选的药品剂型（默认空），请检查！", .TextMatrix(l, .ColIndex("药品库房")))
'                Exit Function
'            End If

            If Trim(.TextMatrix(l, .ColIndex("剂型编码"))) = "" And lngID > 0 Then
                '检查其他注册的接口有无设置药品剂型为空的
                If CheckJiXing(lngID, Trim(txtCode.Text)) Then
                    MsgBox mdlMain.FormatString("“[1]”的其他接口已全选药品剂型，请检查！", .TextMatrix(l, .ColIndex("药品库房")))
                    Exit Function
                End If
            End If
        Next
        
    End With
    
    Verify = True
    
End Function

Private Function Save() As Boolean
    Dim strCode As String, strName As String, strType As String, strLink As String
    Dim strInfo As String, strRemark As String, strXml As String
    Dim arrJX As Variant
    Dim i As Long, j As Long
    Dim objXML As New clsXML
    
    strCode = "'" & Trim(txtCode.Text) & "'"
    strName = "'" & Trim(txtName.Text) & "'"
    strType = CStr(Val(cboType.Text))
    strLink = "'" & mdlMain.Base64Encode(Trim(txtLink.Text)) & "'"    '加密
    strRemark = "'" & Trim(txtRemark.Text) & "'"
    
    '库房与剂型
    With vsfInfo
        objXML.ClearXmlText
        objXML.AppendNode "root", False
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("库房ID"))) > 0 Then
                If Trim(.TextMatrix(i, .ColIndex("剂型编码"))) = "" Then
                    objXML.AppendNode "bm", False
                    objXML.AppendData "id", Trim(.TextMatrix(i, .ColIndex("库房ID")))
                    objXML.AppendData "jxbm", ""
                    objXML.AppendNode "bm", True
                Else
                    arrJX = Split(.TextMatrix(i, .ColIndex("剂型编码")), "；")
                    For j = LBound(arrJX) To UBound(arrJX)
                        If arrJX(j) <> "" Then
                            objXML.AppendNode "bm", False
                            objXML.AppendData "id", Trim(.TextMatrix(i, .ColIndex("库房ID")))
                            objXML.AppendData "jxbm", arrJX(j)
                            objXML.AppendNode "bm", True
                        End If
                    Next
                End If
            End If
        Next
        objXML.AppendNode "root", True
    End With
    strInfo = "'" & Replace(Replace(objXML.XmlText, vbNewLine, ""), " ", "") & "'"
    
    Set objXML = Nothing
    
    Select Case mbytState
    Case enuEditState.新增
        gstrSQL = mdlMain.FormatString("ZL_药品设备接口_UPDATE([1], [2], [3], [4], [5], [6], [7])", _
                                        strCode, _
                                        strName, _
                                        strType, _
                                        strLink, _
                                        strInfo, _
                                        "Null", _
                                        strRemark)
    Case enuEditState.修改
        gstrSQL = mdlMain.FormatString("ZL_药品设备接口_UPDATE([1], [2], [3], [4], [5], [6], [7])", _
                                        strCode, _
                                        strName, _
                                        strType, _
                                        strLink, _
                                        strInfo, _
                                        mlngID, _
                                        strRemark)
    End Select
    
    On Error GoTo hErr
    Call gobjComLib.zlDatabase.ExecuteProcedure(gstrSQL, "")
    
    Save = True
    Exit Function
    
hErr:
    Call gobjComLib.ErrCenter
End Function

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()+={}|[]\:"";'<>?,./", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfInfo_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Row <= 0 Then Exit Sub
    
    Select Case Col
    Case vsfInfo.ColIndex("药品库房")
        Call Selector(1, Row)
        vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("药品剂型")) = ""
        vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("剂型编码")) = ""
        
        If CheckDept() = False Then
            vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("药品库房")) = ""
            vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("库房ID")) = ""
        Else
            Call AppendSpaceLine
        End If
        
    Case vsfInfo.ColIndex("药品剂型")
        Call Selector(2, Row)
        
    End Select
End Sub

Private Sub AppendSpaceLine()
    Dim l As Long
    
    With vsfInfo
        l = .Rows - 1
        If Val(.TextMatrix(l, .ColIndex("库房ID"))) <> 0 Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
    End With
End Sub

Private Function CheckDept() As Boolean
    Dim l As Long, lngID As Long
    Dim blnFind As Boolean
    Dim strCode As String
    
    
    With vsfInfo
        '检查库房重复
        lngID = Val(.TextMatrix(.Row, .ColIndex("库房ID")))
        If lngID = 0 Then
            CheckDept = True
            Exit Function
        End If
        
        For l = 1 To .Rows - 1
            If lngID = Val(.TextMatrix(l, .ColIndex("库房ID"))) And l <> .Row Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind Then
            MsgBox "当前填写的“药品库房”已重复，请检查！", vbInformation, GSTR_MSG
            Exit Function
        End If
        
    End With
    
    CheckDept = True
        
End Function

Private Function CheckJiXing(ByVal lngStoreID As Long, ByVal strInf As String) As Boolean
'功能：检查同库房其他接口的药品剂型
'参数：
'  lngStoreID：库房ID
'  strInf：接口编号
'返回：True存在；False不存在

    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo hErr
    
    gstrSQL = _
        "Select Count(1) Rec " & vbNewLine & _
        "From 药品设备接口 A, Xmltable('//root/bm' Passing a.扩展信息 Columns 库房id Number(18) Path 'id', 剂型编码 Varchar2(20) Path 'jxbm') B " & vbNewLine & _
        "Where a.编号 <> [1] And b.库房id = [2] And b.剂型编码 Is Null And Rownum < 2 "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "", strInf, lngStoreID)
    CheckJiXing = rsTmp!Rec > 0
    rsTmp.Close

    Exit Function
    
hErr:
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Private Sub Selector(ByVal intType As Integer, ByVal Row As Long)
'功能：选择器
'参数：
'   intType：1-药品库房；2-药品剂型
'   Row：选择器选中的值要写本指定行

    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String, strCode As String
    Dim vRect As mdlDefine.RECT
    Dim sngTop As Single, sngLeft As Single
    Dim blnCancel As Boolean
    Dim lngDeptID As Long
    
    vRect = mdlMain.GetControlRect(vsfInfo.hwnd)
    sngTop = vRect.Top + vsfInfo.CellTop + vsfInfo.CellHeight
    sngLeft = vRect.Left + vsfInfo.CellLeft
    
    If intType = 1 Then
        '库房（单选）
        gstrSQL = _
            "Select Distinct a.Id, a.编码, a.名称 " & vbNewLine & _
            "From 部门表 A, 部门性质说明 B " & vbNewLine & _
            "Where a.Id = b.部门id And To_Char(Nvl(a.撤档时间, To_Date('3000-1-1', 'yyyy-mm-dd')), 'yyyy') = '3000' " & vbNewLine & _
            "   And b.工作性质 In ('中药库', '西药库', '成药库', '中药房', '西药房', '成药房') " & vbNewLine & _
            "Order By a.名称 "
            
        Set rsTemp = gobjComLib.zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "药品库房", False, "", "选择药品库房", False, False, True, _
                        sngLeft, sngTop, 0, blnCancel, False, False)
    Else
        '剂型（多选）
        lngDeptID = Val(vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("库房ID")))
        
'        '当前的药品剂型
'        strCode = BuildCode(vsfInfo.Row)
'        If strCode = "ALL" Then
'            MsgBox "同一库房的其他行默认全选药品剂型，请调整！", vbInformation, GSTR_MSG
'            Exit Sub
'        End If
        
        '取出药品剂型，过滤掉指定库房已选的药品剂型，供用户选择
        gstrSQL = _
            "Select Rownum ID, 编码, 名称 " & vbNewLine & _
            "From 药品剂型 " & vbNewLine & _
            "Where Not 编码 In (Select b.剂型 " & vbNewLine & _
            "                   From 药品设备接口 A," & vbNewLine & _
            "                      Xmltable('//root/bm' Passing a.扩展信息 Columns 库房id Number(18) Path 'id', 剂型 Varchar2(20) Path 'jxbm') B " & vbNewLine & _
            "                   Where a.ID <> [2] And b.库房id = [1] " & vbNewLine & _
            ") "

'        gstrSQL = _
'            "Select Rownum ID, 编码, 名称 " & vbNewLine & _
'            "From 药品剂型 " & vbNewLine & _
'            "Order By 编码 "

        Set rsTemp = gobjComLib.zlDatabase.ShowSQLMultiSelect(Me, gstrSQL, 0, "药品剂型", False, "编码", "选择药品剂型", False, False, True, _
                        sngLeft, sngTop, 0, blnCancel, False, False, lngDeptID, mlngID)
    End If
    
    If blnCancel = False Then
        If Not rsTemp Is Nothing Then
            With vsfInfo
                If intType = 1 Then
                    '库房
                    .TextMatrix(Row, .ColIndex("库房ID")) = rsTemp!ID
                    .TextMatrix(Row, .ColIndex("药品库房")) = "【" & rsTemp!编码 & "】" & rsTemp!名称
                Else
                    '剂型
                    strTemp = ""
                    strCode = ""
                    Do While rsTemp.EOF = False
                        strCode = strCode & "；" & rsTemp!编码
                        strTemp = strTemp & "；" & rsTemp!名称
                        rsTemp.MoveNext
                    Loop
                    If Left(strTemp, 1) = "；" Then strTemp = Mid(strTemp, 2)
                    If Left(strCode, 1) = "；" Then strCode = Mid(strCode, 2)
                    .TextMatrix(Row, .ColIndex("药品剂型")) = strTemp
                    .TextMatrix(Row, .ColIndex("剂型编码")) = strCode
                End If
            End With
            rsTemp.Close
        Else
            If intType = 2 Then
                MsgBox "同一库房的其他接口默认全选药品剂型，请调整！", vbInformation, GSTR_MSG
            End If
        End If
    End If
End Sub

'Private Function BuildCode(ByVal lngCol As Long) As String
''功能：获取当前网格中同库房ID记录的剂型，除开当前行
''参数：
''  lngCol：当前行
'
'    Dim l As Long, lngDeptID As Long
'    Dim strCode As String, strTmp As String
'
'    With vsfInfo
'        lngDeptID = Val(.TextMatrix(lngCol, .ColIndex("库房ID")))
'
'        For l = 1 To .Rows - 1
'            '同库房ID
'            If l <> lngCol And Val(.TextMatrix(l, .ColIndex("库房ID"))) = lngDeptID Then
'                strCode = Replace(Trim(.TextMatrix(l, .ColIndex("剂型编码"))), "；", ",")
'                If strCode = "" Then
'                    '所有剂型
'                    BuildCode = "ALL"
'                    Exit Function
'                End If
'                strTmp = strTmp & "," & strCode
'            End If
'        Next
'        If Left(strTmp, 1) = "," Then strTmp = Mid(strTmp, 2)
'    End With
'
'    BuildCode = strTmp
'
'End Function

