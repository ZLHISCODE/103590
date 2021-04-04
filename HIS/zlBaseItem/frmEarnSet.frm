VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEarnSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "收入项目设置"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmEarnSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsfBill 
      Height          =   930
      Left            =   1200
      TabIndex        =   16
      Top             =   2760
      Width           =   2830
      _cx             =   4992
      _cy             =   1640
      Appearance      =   2
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4530
      TabIndex        =   21
      Top             =   3360
      Width           =   1100
   End
   Begin VB.ComboBox cmb病案 
      Height          =   300
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1920
      Width           =   2025
   End
   Begin VB.ComboBox cmb收据 
      Height          =   300
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2320
      Width           =   2025
   End
   Begin VB.CheckBox chk公费 
      Alignment       =   1  'Right Justify
      Caption         =   "公费(&G)"
      Height          =   255
      Left            =   450
      TabIndex        =   10
      Top             =   1605
      Width           =   945
   End
   Begin VB.CommandButton cmd上级 
      Caption         =   "…"
      Height          =   240
      Left            =   2970
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   4290
      TabIndex        =   20
      Top             =   -150
      Width           =   30
   End
   Begin VB.CheckBox chk末级 
      Caption         =   "末级(&M)"
      Height          =   225
      Left            =   480
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   5
      Tag             =   "编码"
      Text            =   "111111"
      Top             =   555
      Width           =   1125
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   7
      Tag             =   "名称"
      Top             =   870
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4530
      TabIndex        =   18
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4530
      TabIndex        =   17
      Top             =   150
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   9
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   2025
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "简码"
      Top             =   1260
      Width           =   1305
   End
   Begin VB.TextBox txtTemp 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "编码"
      Text            =   "11"
      Top             =   510
      Width           =   1305
   End
   Begin VB.Label lblEdit 
      Caption         =   "不同场合  收据费目(&F)"
      Height          =   900
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "缺省病案费目(&B)"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   1965
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "收据费目(&T)"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "编码(&U)"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   570
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   930
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "简码(&S)"
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   1290
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "上级(&D)"
      Height          =   180
      Index           =   9
      Left            =   480
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
End
Attribute VB_Name = "frmEarnSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr上级项目ID As String     '当前编辑的上级项目ID
Dim mstrID As String         '当前编辑的项目ID

Dim mstr上级编码 As String    '原始的上级编码的值
Dim mstr编码 As String        '原始的本级编码的值
Dim mint编码 As Integer       '修改前包括下级在内的编码最长的长度
Dim mintSuccess As Integer
Dim mblnChange As Boolean  '已修改
Dim mbln药店 As Boolean
Dim mstr收据费目 As String

Private Sub cmb病案_KeyPress(KeyAscii As Integer)
    '-----------------------------------------------------------------------------------
    '键盘定位
    '-----------------------------------------------------------------------------------
    Dim intI As Integer
    intI = zlControl.CboMatchIndex(cmb病案.hwnd, KeyAscii)
    '根据公共函数CboSetIndex定位到指定索引
    Call zlControl.CboSetIndex(cmb病案.hwnd, intI)
End Sub

Private Sub cmb收据_KeyPress(KeyAscii As Integer)
    '-----------------------------------------------------------------------------------
    '键盘定位
    '-----------------------------------------------------------------------------------
    Dim intI As Integer
    intI = zlControl.CboMatchIndex(cmb收据.hwnd, KeyAscii)
    '根据公共函数CboSetIndex定位到指定索引
    Call zlControl.CboSetIndex(cmb收据.hwnd, intI)
End Sub

Private Sub cmb收据_Validate(Cancel As Boolean)
    With vsfBill
        If .TextMatrix(1, 1) = "" Then .TextMatrix(1, 1) = cmb收据.Text
        If .TextMatrix(2, 1) = "" Then .TextMatrix(2, 1) = cmb收据.Text
        If .TextMatrix(3, 1) = "" Then .TextMatrix(3, 1) = cmb收据.Text
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    If chk末级.Value = 1 Then
        ShowHelp App.ProductName, Me.hwnd, "frm收入项目设置2", Int((glngSys) / 100)
    Else
        ShowHelp App.ProductName, Me.hwnd, "frm收入项目设置1", Int((glngSys) / 100)
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save项目() = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstrID = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(1).Text = GetMaxLocalCode(mstr上级项目ID, "收入项目")
    cmdOK.Enabled = False
    frmEarnManage.FillList frmEarnManage.tvwMain_S.SelectedItem.Key
    txtEdit(1).SetFocus
    txtTemp.MaxLength = GetLocalCodeLength(mstr上级项目ID, "收入项目")
    txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8, txtTemp.MaxLength) - Len(txtTemp.Text)
    mblnChange = False
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能:分析输入收入项目的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = 1 To 3
        strTemp = Trim(txtEdit(i).Text)
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox "所输入内容不能超过" & Int(txtEdit(i).MaxLength / 2) & "个汉字" & "或" & txtEdit(i).MaxLength & "个字母。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(1).Text) = 0 Then
            MsgBox "编码不能为空。", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            Exit Function
        End If
    Else
        If Len(txtEdit(1).Text) < txtEdit(1).MaxLength Then
            MsgBox "编码的长度不够。", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(1).Text) Or InStr(txtEdit(1).Text, ",") > 0 Or InStr(txtEdit(1).Text, ".") > 0 Then
        MsgBox "编码应由数字组成。", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txtEdit(2).Text = ""
        txtEdit(2).SetFocus
        Exit Function
    End If
    If chk末级.Value And cmb收据.ListIndex < 1 Then
        MsgBox "收据费目不能为空。", vbExclamation, gstrSysName
        cmb收据.SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save项目() As Boolean
'功能:保存编辑的内容到收入项目表中
'参数:
'返回值:成功返回True,否则为False
    Dim lng收入项目ID As Long
    Dim str收据费目场合 As String
    
    On Error GoTo ErrHandle
    
    With vsfBill
        str收据费目场合 = .TextMatrix(1, 1) & "|" & .TextMatrix(2, 1) & "|" & .TextMatrix(3, 1)
    End With
    
    If mstrID = "" Then       '新增一条记录
        lng收入项目ID = zlDatabase.GetNextId("收入项目")
        gstrSQL = "zl_收入项目_insert(" & lng收入项目ID & "," & IIF(mstr上级项目ID = "", "null", mstr上级项目ID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & txtEdit(3).Text & "'," & chk公费.Value & ",'" & cmb收据.Text & _
            "','" & cmb病案.Text & "'," & chk末级.Value & ",'" & str收据费目场合 & "')"
    Else    '修改
        gstrSQL = "zl_收入项目_update(" & mstrID & "," & IIF(mstr上级项目ID = "", "null", mstr上级项目ID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & txtEdit(3).Text & "'," & chk公费.Value & ",'" & cmb收据.Text & _
            "','" & cmb病案.Text & "'," & Len(mstr编码) + 1 & "," & chk末级.Value & ",'" & str收据费目场合 & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Save项目 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 编辑项目(ByVal str上级项目 As String, ByVal str上级项目ID As String, ByVal str上级编码 As String, _
    Optional strID As String = "", Optional ByVal bln末级项目 As Boolean) As Boolean
'功能:用来与调用的收入项目管理窗口进行通讯的程序
'参数:str上级项目     上级收入项目的名字
'     str上级项目ID   上级收入项目的ID
'     str上级编码     上级收入项目的编码
'     strID           本收入项目的的ID
'     bln末级项目     本收入项目是否末级
'返回值:编辑成功返回True,否则为False
    
    Dim rs收入项目 As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rs收据费目对应 As ADODB.Recordset
    
    Dim i As Integer
    
    mintSuccess = 0
    mstrID = strID
    
    Load frmEarnSet
    
    mbln药店 = (glngSys \ 100 = 8)
    
    On Error GoTo ErrHandle
    chk末级.Value = 0
    If strID <> "" Then
        rs收入项目.CursorLocation = adUseClient
        gstrSQL = "select A.ID,A.编码,A.名称 from 收入项目 A,收入项目 B " & _
                " where A.ID(+)=B.上级ID and B.ID=[1]"
        Set rs收入项目 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
                
        mstr上级项目ID = IIF(IsNull(rs收入项目("ID")), "", rs收入项目("ID"))
        mstr上级编码 = IIF(IsNull(rs收入项目("编码")), "", rs收入项目("编码"))
        
        txtTemp.Text = mstr上级编码
        txtEdit(9).Text = IIF(IsNull(rs收入项目("名称")), "无", rs收入项目("名称"))
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength(mstr上级项目ID, "收入项目")
        'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
        
        gstrSQL = "select ID,上级ID,编码,名称,简码,末级,公费,收据费目,病案费目 from 收入项目  " & _
            "where ID =[1]"
        Set rs收入项目 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
        txtEdit(1).Text = Mid(rs收入项目("编码"), Len(txtTemp.Text) + 1)
        mstr编码 = rs收入项目("编码")
        '求出包括子节点在内的最长编码
        mint编码 = GetDownCodeLength(mstrID, "收入项目")
        ' 8 - (mint编码 - Len(mstr编码))这个公式的意思是要为它的孩子的编码留有余地
        txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8 - (mint编码 - Len(mstr编码)), txtTemp.MaxLength) - Len(mstr上级编码)
        txtEdit(2).Text = rs收入项目("名称")
        txtEdit(3).Text = IIF(IsNull(rs收入项目("简码")), "", rs收入项目("简码"))
        chk公费.Value = IIF(rs收入项目("公费") = 1, 1, 0)
        chk末级.Value = IIF(rs收入项目("末级") = 1, 1, 0)
        chk末级.Enabled = False
    Else
        mstr上级项目ID = str上级项目ID
        mstr上级编码 = str上级编码
        
        txtTemp.Text = str上级编码
        txtEdit(9).Text = str上级项目
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength(str上级项目ID, "收入项目")
        'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
        '判断编码是否满了
        If Len(mstr上级编码) = 8 Then
            MsgBox "不能再增加子级了，编码长度已经用尽。", vbExclamation, gstrSysName
            mblnChange = False
            Unload frmEarnSet
            Exit Function
        End If
        txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8, txtTemp.MaxLength) - Len(mstr上级编码)
        txtEdit(1).Text = GetMaxLocalCode(str上级项目ID, "收入项目")
        mstr编码 = mstr上级编码 & txtEdit(1).Text
        If bln末级项目 Then chk末级.Value = 1
        
    End If
    If chk末级.Value = 1 Then
        rsTemp.CursorLocation = adUseClient
        
        gstrSQL = "select 名称 from 收据费目 order By 名称"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        If rsTemp.RecordCount = 0 Then
            MsgBox "请先完成“收据费目”的设置。", vbExclamation, gstrSysName
            编辑项目 = False
            mblnChange = False
            Unload frmEarnSet
            Exit Function
        End If
        cmb收据.Clear
        cmb收据.AddItem ""
        Do Until rsTemp.EOF
            cmb收据.AddItem rsTemp("名称")
            mstr收据费目 = IIF(mstr收据费目 = "", rsTemp("名称"), mstr收据费目 & "|" & rsTemp("名称"))
            rsTemp.MoveNext
        Loop
        cmb收据.ListIndex = 0
        rsTemp.Close
        
        If mbln药店 = True Then
            lblEdit(6).Visible = False
            cmb病案.Visible = False
            txtEdit(9).Top = cmb病案.Top
            lblEdit(9).Top = lblEdit(6).Top
            cmd上级.Top = txtEdit(9).Top + 30
            frmEarnSet.Height = 3000
        Else
            '药店系统不处理病案费目
            '刘兴宏:2007/05/17:由于病案中的病案费目存在上下级关系,因此统一调整了病案费目,收入项目的病案费目只能统计末级为1的记录.
            gstrSQL = "select 名称 from 病案费目 where 末级=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

            cmb病案.Clear
            Do Until rsTemp.EOF
                cmb病案.AddItem rsTemp("名称")
                rsTemp.MoveNext
            Loop
            If rsTemp.RecordCount = 0 Then
                mblnChange = False
                MsgBox "请先完成“病案费目”的设置。", vbExclamation, gstrSysName
                mblnChange = False
                Unload frmEarnSet
                编辑项目 = False
                Exit Function
            End If
            cmb病案.ListIndex = 0
            rsTemp.Close
        End If
        
        If mstrID <> "" Then
            On Error Resume Next
            cmb收据.Text = rs收入项目("收据费目")
            If Err <> 0 Then
                cmb收据.AddItem rs收入项目("收据费目")
                cmb收据.Text = rs收入项目("收据费目")
                Err.Clear
            End If
            cmb病案.Text = rs收入项目("病案费目")
            If Err <> 0 Then
                cmb病案.AddItem rs收入项目("病案费目")
                cmb病案.Text = rs收入项目("病案费目")
                Err.Clear
            End If
        End If
        
        '收据费目对应
        With vsfBill
            .Rows = 4
            .Cols = 2
            .Editable = flexEDNone
            
            .FixedAlignment(-1) = flexAlignCenterCenter
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColWidth(0) = 1200
            .ColWidth(1) = 1600
            
            .TextMatrix(0, 0) = "场合"
            .TextMatrix(0, 1) = "收据费目"
            
            .TextMatrix(1, 0) = "0-门诊"
            .TextMatrix(2, 0) = "1-住院"
            .TextMatrix(3, 0) = "2-门诊及住院"
        End With
        
        gstrSQL = "Select 收入项目id, 场合, 收据费目 From 收据费目对应 Where 收入项目ID = [1]"
        Set rs收据费目对应 = zlDatabase.OpenSQLRecord(gstrSQL, "收据费目对应", Val(strID))
        
        With rs收据费目对应
            Do While Not .EOF
                If !场合 = 0 Then
                    vsfBill.TextMatrix(1, 1) = !收据费目
                ElseIf !场合 = 1 Then
                    vsfBill.TextMatrix(2, 1) = !收据费目
                Else
                    vsfBill.TextMatrix(3, 1) = !收据费目
                End If
                .MoveNext
            Loop
        End With
        
    Else
        lblEdit(5).Visible = False
        lblEdit(6).Visible = False
        lblEdit(7).Visible = False
        chk公费.Visible = False
        cmb病案.Visible = False
        cmb收据.Visible = False
        vsfBill.Visible = False
'        txtEdit(9).Top = chk公费.Top
'        lblEdit(9).Top = txtEdit(9).Top + 75
'        cmd上级.Top = txtEdit(9).Top + 30
        frmEarnSet.Height = 2300
        cmdHelp.Top = txtEdit(3).Top
    End If
    
    frmEarnSet.Caption = IIF(chk末级.Value = 1, "收入项目设置", "收入分类设置")
    
    mblnChange = False
    frmEarnSet.Show vbModal
    编辑项目 = mintSuccess > 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmd上级_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim str编码 As String
    Dim int编码  As Integer
    
    strSQL = "select ID,上级ID,名称,编码 from 收入项目  " & _
        "where 末级 <> 1 start with 上级ID is null connect by prior ID =上级ID"
    strID = mstr上级项目ID
    str名称 = txtEdit(9).Text
    str编码 = txtTemp.Text
    blnRe = frmTreeSel.ShowTree(strSQL, strID, str名称, str编码, mstrID, "收入项目", "所有收入项目", , mstr编码)
    '成功返回
    If blnRe Then
        '判断是否合适
        If Len(str编码) >= Len(mstr编码) Then
            If Mid(str编码, 1, Len(str编码)) = mstr编码 Then
                MsgBox "这个上级不合适，因为选择了它自身或其下级。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '新的本级的宽度
        int编码 = GetLocalCodeLength(strID, "收入项目")
        '只有修改才有必要审核
        If mstrID <> "" Then
            '其纯下级编码+新的本级编码<=8
            If mint编码 - Len(mstr编码) + IIF(int编码 = 0, Len(str编码) + 1, int编码) > 8 Then
                MsgBox "这个上级不合适，因为它的编码太长了。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        mstr上级项目ID = strID
        txtEdit(9).Text = str名称
        txtTemp.MaxLength = int编码
        txtTemp.Text = str编码
        If mstrID <> "" Then
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8 - (mint编码 - Len(mstr编码)), txtTemp.MaxLength) - Len(str编码)
            txtEdit(1).Text = GetMaxLocalCode(mstr上级项目ID, "收入项目")
        Else
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8, txtTemp.MaxLength) - Len(str编码)
            txtEdit(1).Text = GetMaxLocalCode(mstr上级项目ID, "收入项目")
        End If
        mblnChange = True
        'txtEdit(1).Text = Mid(txtEdit(1).Text, Len(txtTemp.Text) + 1)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Activate()
    txtEdit(2).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    cmdOK.Enabled = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 2 Then
        zlCommFun.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        '要作报表名称，所以不能有怪字符
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Then
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(1).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(1).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

Private Sub vsfBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfBill
        If Row = 0 Then Exit Sub
        If Col <> 1 Then Exit Sub
        
        .ColComboList(1) = mstr收据费目
    End With
End Sub

Private Sub vsfBill_EnterCell()
    With vsfBill
        .Editable = flexEDKbd
        If .Row = 0 Then Exit Sub
        If .Col = 0 Then Exit Sub
        
        .Editable = flexEDKbdMouse
    End With
End Sub


Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    For i = 0 To vsfBill.ComboCount - 1
        If zlStr.GetCodeByVB(vsfBill.ComboItem(i)) Like UCase(Chr(KeyAscii)) & "*" Then
            vsfBill.ComboIndex = i: Exit For
        End If
    Next
End Sub
