VERSION 5.00
Begin VB.Form frmUnitEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "合约单位设置"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmUnitEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboInfo 
      Height          =   300
      Left            =   3480
      TabIndex        =   26
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5100
      TabIndex        =   25
      Top             =   1380
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "编码"
      Text            =   "111111"
      Top             =   195
      Width           =   885
   End
   Begin VB.CommandButton cmd上级 
      Caption         =   "…"
      Height          =   270
      Left            =   4510
      TabIndex        =   21
      Top             =   3120
      Width           =   285
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   4920
      TabIndex        =   23
      Top             =   -150
      Width           =   30
   End
   Begin VB.CheckBox chk末级 
      Caption         =   "末级(&M)"
      Height          =   225
      Left            =   480
      TabIndex        =   22
      Top             =   3510
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   8
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   15
      Tag             =   "联系人"
      Top             =   2760
      Width           =   1500
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   13
      Tag             =   "帐号"
      Top             =   2370
      Width           =   2400
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   11
      Tag             =   "开户银行"
      Top             =   1980
      Width           =   3620
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   1200
      MaxLength       =   16
      TabIndex        =   9
      Tag             =   "电话"
      Top             =   1590
      Width           =   2400
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   7
      Tag             =   "地址"
      Top             =   1200
      Width           =   3620
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "名称"
      Top             =   510
      Width           =   3620
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   18
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5100
      TabIndex        =   17
      Top             =   150
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   10
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   20
      Top             =   3120
      Width           =   3310
   End
   Begin VB.TextBox txtTemp 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "编码"
      Text            =   "11"
      Top             =   150
      Width           =   1155
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "简码"
      Top             =   870
      Width           =   1155
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "院区(&P)"
      Height          =   180
      Index           =   9
      Left            =   2805
      TabIndex        =   16
      Top             =   2820
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "联系人(&L)"
      Height          =   180
      Index           =   8
      Left            =   300
      TabIndex        =   14
      Top             =   2820
      Width           =   810
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "帐号(&Z)"
      Height          =   180
      Index           =   7
      Left            =   480
      TabIndex        =   12
      Top             =   2430
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "开户银行(&B)"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "电话(&T)"
      Height          =   180
      Index           =   5
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "地址(&A)"
      Height          =   180
      Index           =   4
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "编码(&U)"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   210
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   570
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "简码(&S)"
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   930
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "上级(&H)"
      Height          =   180
      Index           =   10
      Left            =   480
      TabIndex        =   19
      Top             =   3180
      Width           =   630
   End
End
Attribute VB_Name = "frmUnitEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Dim mstr上级单位ID As String     '当前编辑的上级单位ID
Dim mstrID As String         '当前编辑的单位ID

Dim mstr上级编码 As String    '原始的上级编码的值
Dim mstr编码 As String        '原始的本级编码的值
Dim mint编码 As Integer       '修改前包括下级在内的编码最长的长度
Dim mintSuccess As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save单位() = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    If mstrID <> "" Then
        Unload Me
        Exit Sub
    End If
    mstrID = ""
    For i = 2 To 8
        txtEdit(i).Text = ""
    Next
    txtEdit(1).Text = GetMaxLocalCode(mstr上级单位ID, "合约单位")
    cmdOK.Enabled = False
    frmUnit.FillList frmUnit.tvwMain_S.SelectedItem.Key
    txtEdit(1).SetFocus
    txtTemp.MaxLength = GetLocalCodeLength(mstr上级单位ID, "合约单位")
    txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(mstr上级编码)
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能:分析输入有关合约单位的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = 1 To 8
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
    IsValid = True
End Function

Private Function Save单位() As Boolean
'功能:保存编辑的内容到合约单位表中
'参数:
'返回值:成功返回True,否则为False
    Dim lngID As Long
    Dim str站点 As String
    
    On Error GoTo errHandle
    
    If cboInfo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = Mid(cboInfo.Text, 1, InStr(1, cboInfo.Text, "-") - 1)
    End If
    
    If mstrID = "" Then       '新增一条记录
        lngID = zlDatabase.GetNextId("合约单位")
        gstrSQL = "zl_合约单位_insert(" & lngID & "," & IIf(mstr上级单位ID = "", "null", mstr上级单位ID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & UCase(txtEdit(3).Text) & "','" & txtEdit(4).Text & "','" & txtEdit(5).Text & _
            "','" & txtEdit(6).Text & "','" & txtEdit(7).Text & _
            "','" & txtEdit(8).Text & "'," & chk末级.Value & ",null,null,'" & IIf(cboInfo.Text = "", "Null", str站点) & "')"
    Else    '修改
        gstrSQL = "zl_合约单位_update(" & mstrID & "," & IIf(mstr上级单位ID = "", "null", mstr上级单位ID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & UCase(txtEdit(3).Text) & "','" & txtEdit(4).Text & "','" & txtEdit(5).Text & _
            "','" & txtEdit(6).Text & "','" & txtEdit(7).Text & _
            "','" & txtEdit(8).Text & "'," & Len(mstr编码) + 1 & ",null,null,'" & IIf(cboInfo.Text = "", "Null", str站点) & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Save单位 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function 编辑单位(ByVal str上级单位 As String, ByVal str上级单位ID As String, ByVal str上级编码 As String, _
    Optional strID As String = "", Optional ByVal bln末级单位 As Boolean) As Boolean
'功能:用来与调用的合约单位管理窗口进行通讯的程序,用来增加或修改某个合约单位信息
'参数:str上级单位     上级合约单位的名字
'     str上级单位ID   上级合约单位的ID
'     str上级编码     上级合约单位的编码
'     strID           本合约单位的的ID
'     bln末级项目     本收入项目是否末级
'返回值:编辑成功返回True,否则为False
    
    Dim rs合约单位 As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    
    
    mintSuccess = 0
    
    mstrID = strID
    
    On Error GoTo errHandle
    
    strSQL = "Select 编号, 名称 From Zlnodelist"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "站点位查询")
    If Not rsTmp Is Nothing Then
        cboInfo.AddItem ""
        For i = 0 To rsTmp.RecordCount - 1
            cboInfo.AddItem rsTmp!编号 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Next
    End If
    
    If strID <> "" Then
        rs合约单位.CursorLocation = adUseClient
        gstrSQL = "select A.ID,A.编码,A.名称 from 合约单位 A,合约单位 B " & _
                " where A.ID(+)=B.上级ID and B.ID=[1]"
'        Call SQLTest(App.ProductName, Me.Caption, gstrSQL)
'        rs合约单位.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'        Call SQLTest
        Set rs合约单位 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(strID = "", Null, CLng(strID)))
        
        mstr上级单位ID = IIf(IsNull(rs合约单位("ID")), "", rs合约单位("ID"))
        mstr上级编码 = IIf(IsNull(rs合约单位("编码")), "", rs合约单位("编码"))
        
        txtTemp.Text = mstr上级编码
        txtEdit(10).Text = IIf(IsNull(rs合约单位("名称")), "无", rs合约单位("名称"))
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength(mstr上级单位ID, "合约单位")
        'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
        
        rs合约单位.Close
        
        strSQL = "select ID,上级ID,编码,名称,简码,末级,地址,电话,开户银行,帐号,联系人,撤档时间,站点 from 合约单位  " & _
            "where ID =[1]"
'        Call SQLTest(App.ProductName, Me.Caption, strSQL)
'        rs合约单位.Open strSQL, gcnOracle, adOpenStatic, adLockReadOnly
'        Call SQLTest
        Set rs合约单位 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(strID = "", Null, CLng(strID)))

        txtEdit(1).Text = Mid(rs合约单位("编码"), Len(txtTemp.Text) + 1)
        mstr编码 = rs合约单位("编码")
        '求出包括子节点在内的最长编码
        mint编码 = GetDownCodeLength(mstrID, "合约单位")
        ' 8 - (mint编码 - Len(mstr编码))这个公式的意思是要为它的孩子的编码留有余地
        txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10 - (mint编码 - Len(mstr编码)), txtTemp.MaxLength) - Len(mstr上级编码)
        
        txtEdit(2).Text = rs合约单位("名称")
        txtEdit(3).Text = IIf(IsNull(rs合约单位("简码")), "", rs合约单位("简码"))
        txtEdit(4).Text = IIf(IsNull(rs合约单位("地址")), "", rs合约单位("地址"))
        txtEdit(5).Text = IIf(IsNull(rs合约单位("电话")), "", rs合约单位("电话"))
        txtEdit(6).Text = IIf(IsNull(rs合约单位("开户银行")), "", rs合约单位("开户银行"))
        txtEdit(7).Text = IIf(IsNull(rs合约单位("帐号")), "", rs合约单位("帐号"))
        txtEdit(8).Text = IIf(IsNull(rs合约单位("联系人")), "", rs合约单位("联系人"))
        cboInfo.ListIndex = cbo.FindIndex(cboInfo, IIf(IsNull(rs合约单位("站点")), "", rs合约单位("站点")))
        If rs合约单位("末级") Then chk末级.Value = 1
        chk末级.Enabled = False
    Else
        mstr上级单位ID = str上级单位ID
        txtEdit(10).Text = str上级单位
        mstr上级编码 = str上级编码
        
        txtTemp.Text = str上级编码
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength(str上级单位ID, "合约单位")
        'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
        txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(mstr上级编码)
        txtEdit(1).Text = GetMaxLocalCode(str上级单位ID, "合约单位")
        mstr编码 = mstr上级编码 & txtEdit(1).Text
        If bln末级单位 Then chk末级.Value = 1
    End If
    If chk末级.Value <> 1 Then
        For i = 4 To 8
            txtEdit(i).Visible = False
            lblEdit(i).Visible = False
        Next
        
        cboInfo.Top = txtEdit(3).Top
        lblEdit(9).Top = lblEdit(3).Top
        txtEdit(10).Top = txtEdit(4).Top
        lblEdit(10).Top = lblEdit(4).Top
        cmd上级.Top = txtEdit(10).Top
        frmUnitEdit.Height = 2300
    End If
    
'    If gstrNodeNo = "-" Then
'        txtEdit(9).Visible = False
'        lblEdit(9).Visible = False
'    End If
    frmUnitEdit.Show vbModal
    编辑单位 = mintSuccess > 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmd上级_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim str编码 As String
    Dim int编码  As Integer
    
    strSQL = "select ID,上级ID,名称,编码 from 合约单位  " & _
        "where 末级 <> 1 start with 上级ID is null connect by prior ID =上级ID"
    strID = mstr上级单位ID
    str名称 = txtEdit(10).Text
    str编码 = txtTemp.Text
    blnRe = frm树型选择.ShowTree(strSQL, strID, str名称, str编码, mstrID, "合约单位", "所有合约单位", , mstr编码)
    '成功返回
    If blnRe Then
        '新的本级的宽度
        int编码 = GetLocalCodeLength(strID, "合约单位")
        '只有修改才有必要审核
        If mstrID <> "" Then
            If mint编码 - Len(mstr编码) + IIf(int编码 = 0, Len(str编码) + 1, int编码) > 10 Then
                MsgBox "这个上级不合适，因为它的编码太长了。", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        mstr上级单位ID = strID
        txtEdit(10).Text = str名称
        txtTemp.MaxLength = int编码
        txtTemp.Text = str编码
        If mstrID <> "" Then
            txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10 - (mint编码 - Len(mstr编码)), txtTemp.MaxLength) - Len(str编码)
        Else
            txtEdit(1).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(str编码)
        End If
        txtEdit(1).Text = GetMaxLocalCode(mstr上级单位ID, "合约单位")
        'txtEdit(1).Text = Mid(txtEdit(1).Text, Len(txtTemp.Text) + 1)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    cmdOK.Enabled = True
    If Index = 2 Then
        txtEdit(3).Text = zlCommFun.SpellCode(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Or Index = 4 Or Index = 6 Or Index = 8 Then
        OpenIme gstrIme
    ElseIf Index = 1 Or Index = 3 Or Index = 7 Then
        OpenIme
    End If
End Sub


Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Or Index = 4 Or Index = 6 Or Index = 8 Then
        OpenIme
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(1).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(1).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

