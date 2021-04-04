VERSION 5.00
Begin VB.Form frmEInvoicePointSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "电子票据开票点设置"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmEInvoicePointSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5220
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Height          =   3400
      Left            =   3800
      TabIndex        =   25
      Top             =   -90
      Width           =   10
   End
   Begin VB.CommandButton cmd收费员 
      Caption         =   "…"
      Height          =   250
      Left            =   3360
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   840
      MaxLength       =   50
      TabIndex        =   23
      Top             =   3720
      Width           =   2475
   End
   Begin VB.CommandButton cmd部门 
      Caption         =   "…"
      Height          =   250
      Left            =   3360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2085
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   840
      TabIndex        =   6
      Top             =   2070
      Width           =   2475
   End
   Begin VB.CommandButton cmd客户端 
      Caption         =   "…"
      Height          =   250
      Left            =   3360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1725
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   970
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "编码"
      Text            =   "111111"
      Top             =   615
      Width           =   2535
   End
   Begin VB.TextBox txtTemp 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   840
      MaxLength       =   10
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "编码"
      Text            =   "1111111111"
      Top             =   570
      Width           =   2775
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   180
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   1695
      Width           =   2475
   End
   Begin VB.ComboBox cmbStationNo 
      Height          =   300
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2835
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   840
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   840
      MaxLength       =   50
      TabIndex        =   2
      Top             =   945
      Width           =   2775
   End
   Begin VB.CommandButton cmd上级 
      Caption         =   "…"
      Height          =   250
      Left            =   3350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   195
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   840
      MaxLength       =   100
      TabIndex        =   8
      Top             =   2445
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   12
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   13
      Top             =   720
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "收费员"
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   22
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "部门"
      Height          =   180
      Index           =   6
      Left            =   360
      TabIndex        =   21
      Top             =   2130
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "客户端"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   20
      Top             =   1740
      Width           =   540
   End
   Begin VB.Label lblStationNo 
      AutoSize        =   -1  'True
      Caption         =   "院区"
      Height          =   180
      Left            =   360
      TabIndex        =   19
      Top             =   2895
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "上级"
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "简码"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   1365
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "名称"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "编码"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   615
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "位置"
      Height          =   180
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   2505
      Width           =   360
   End
   Begin VB.Menu mnuShort 
      Caption         =   "快捷菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPatient 
         Caption         =   "门诊病人(&O)"
         Index           =   0
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "住院病人(&I)"
         Index           =   1
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "门诊和住院病人(&B)"
         Index           =   2
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "不服务于病人(&N)"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmEInvoicePointSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const mlng编码长度 As Long = 10
Private mstr上级ID As String     '当前编辑的上级发票点ID
Private mstrID As String            '当前编辑的发票点ID
Private mbln末级 As Boolean     '当前编辑的发票点是否为末级
Private mstr编码 As String         '原始的本级编码的值
Private mstr上级编码 As String   '原始的上级编码的值
Private mint编码 As Integer       '修改前包括下级在内的编码最长的长度
Private mblnChange As Boolean     '是否改变了
Private Enum mEdit
    Edit_客户端 = 0
    Edit_编码 = 1
    Edit_名称 = 2
    Edit_简码 = 3
    Edit_上级 = 4
    Edit_位置 = 5
    Edit_部门 = 6
    Edit_收费员 = 7
End Enum
Private mbln开票点对码 As Boolean
Private mintMode As Integer  '对码方式0-按客户端对,1-按收费员对;2-按收费员+客户端对
Private mlng对照ID As Long
Private mblnOK  As Boolean

Private Sub IniStationNo()
    Dim rsRecord As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    lblStationNo.Visible = True
    cmbStationNo.Visible = True
    
    strSQL = "select 编号,名称 from zlnodelist"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "站点查询")
    
    If rsRecord.RecordCount = 0 Then
        lblStationNo.Visible = False
        cmbStationNo.Visible = False
    Else
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!编号 & "-" & rsRecord!名称
                rsRecord.MoveNext
            Loop
        End With
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStationNo(ByVal strNO As String)
    Dim n As Integer
    If cmbStationNo.ListCount = 0 Then Exit Sub
    
    If strNO = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNO Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL  As String
    
    If mbln开票点对码 Then
        If Save开票点对照() = False Then Exit Sub
        mblnChange = False: mbln开票点对码 = False
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If IsValid() = False Then Exit Sub
        
    '检查开票点是否与已有开票点名称相同
    If CheckSame(txtEdit(mEdit.Edit_名称).Text, Val(mstrID)) Then
        If MsgBox("当前录入的开票点名称与已有开票点名称相同!", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    If Save开票点() = False Then Exit Sub
    mblnOK = True
    '改变主窗口的显示
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    Else
    
    End If
    '连续增加
    mstrID = ""
    txtEdit(mEdit.Edit_名称).Text = ""
    txtEdit(mEdit.Edit_简码).Text = ""
    txtEdit(mEdit.Edit_位置).Text = ""
    txtEdit(mEdit.Edit_编码).Text = GetMaxLocalCode(mstr上级ID, "电子票据开票点")
    
    txtTemp.MaxLength = GetLocalCodeLength(mstr上级ID, "电子票据开票点")
    txtEdit(mEdit.Edit_编码).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(txtTemp.Text)

    zlControl.ControlSetFocus txtEdit(mEdit.Edit_名称)
    
    mblnChange = False
End Sub

Private Function IsValid() As Boolean
    Dim i As Long
    Dim blnTmp As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim str部门性质 As String
    Dim strMsg As String
    
    On Error GoTo errHandle
    
    For i = 1 To 5
        If i <> 4 Then
            If zlCommFun.StrIsValid(Trim(txtEdit(mEdit.Edit_编码).Text), txtEdit(i).MaxLength) = False Then
                zlControl.ControlSetFocus txtEdit(i)
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        End If
    Next
    txtEdit(mEdit.Edit_编码).Text = Trim(txtEdit(mEdit.Edit_编码).Text)

    If Len(Trim(txtEdit(mEdit.Edit_上级).Text)) = 0 And Me.Tag = "恢复" Then
        MsgBox "上级不能为空。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_上级)
        Exit Function
    End If
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(mEdit.Edit_编码).Text) = 0 Then
            MsgBox "编码不能为空。", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_编码)
            Exit Function
        End If
    Else
        If Len(txtEdit(mEdit.Edit_编码).Text) < txtEdit(mEdit.Edit_编码).MaxLength Then
            MsgBox "编码的长度不够。", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_编码)
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(mEdit.Edit_编码).Text) Or InStr(txtEdit(mEdit.Edit_编码).Text, ",") > 0 Or InStr(txtEdit(mEdit.Edit_编码).Text, ".") > 0 Then
        MsgBox "编码应由数字组成。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_编码)
        Exit Function
    End If
    If Len(Trim(txtEdit(mEdit.Edit_名称).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txtEdit(mEdit.Edit_名称).Text = ""
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_名称)
        Exit Function
    End If
    If LenB(StrConv(txtEdit(mEdit.Edit_名称).Text, vbFromUnicode)) > 20 Then
        MsgBox "名称长度不能超过10个汉字或者20个字符，请重新录入！", vbInformation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_名称)
        Exit Function
    End If
    If LenB(StrConv(txtEdit(mEdit.Edit_简码).Text, vbFromUnicode)) > 20 Then
        MsgBox "编码长度不能超过20个字符，请重新录入！", vbInformation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_简码)
        Exit Function
    End If

    IsValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function reMoveSpe(ByVal strChar As String) As String
'129884,去掉名称里的回车，换行，及首尾空格
    reMoveSpe = Trim(Replace(Replace(Replace(strChar, vbCrLf, ""), vbCr, ""), vbLf, ""))
End Function

Private Function Save开票点() As Boolean
    'blnDelete-是否删除电子票据开票点
    Dim i As Integer, strSQL As String
    Dim nod As Node
    Dim lst As ListItem
    Dim str站点 As String
    Dim lngID As Long
    
    On Error GoTo errHandle
    
    txtEdit(mEdit.Edit_名称).Text = reMoveSpe(txtEdit(mEdit.Edit_名称).Text)
    If cmbStationNo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = "'" & Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1) & "'"
    End If
        
    If mstrID = "" Then       '新增一条记录
        If Check重复部门(mstr上级ID, Trim(txtEdit(mEdit.Edit_名称).Text)) = True Then
            MsgBox "该级下面已有该部门，不能添加相同部门！", vbInformation, gstrSysName
            Exit Function
        End If
        lngID = zlDatabase.GetNextId("电子票据开票点")
        '  Zl_电子票据开票点_Insert
        strSQL = "Zl_电子票据开票点_Insert("
        '  Id_In       In 电子票据开票点.Id%Type,
        strSQL = strSQL & lngID & ","
        '  上级id_In   In 电子票据开票点.上级id%Type,
        strSQL = strSQL & ZVal(Val(txtEdit(mEdit.Edit_上级).Tag)) & ","
        '  编码_In     In 电子票据开票点.编码%Type,
        strSQL = strSQL & "'" & txtTemp.Text & txtEdit(mEdit.Edit_编码).Text & "',"
        '  名称_In     In 电子票据开票点.名称%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_名称).Text & "',"
        '  简码_In     In 电子票据开票点.简码%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_简码).Text & "',"
        '  院区_In     In 电子票据开票点.院区%Type,
        strSQL = strSQL & str站点 & ","
        '  客户端_In   In 电子票据开票点.客户端%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_客户端).Text & "',"
        '  部门id_In   In 电子票据开票点.部门id%Type,
        strSQL = strSQL & IIf(Val(txtEdit(mEdit.Edit_部门).Tag) = 0, "NULL", "'" & Val(txtEdit(mEdit.Edit_部门).Tag) & "'") & ","
        '  位置_In     In 电子票据开票点.位置%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_位置).Text & "',"
        '  末级_In     In 电子票据开票点.末级%Type := Null,
        strSQL = strSQL & "" & ZVal(IIf(mbln末级, 1, 0)) & ")"
    Else
        '修改
        lngID = Val(mstrID)
        '  Zl_电子票据开票点_Update
        strSQL = "Zl_电子票据开票点_Update("
        '  Id_In       In 电子票据开票点.Id%Type,
        strSQL = strSQL & lngID & ","
        '  上级id_In   In 电子票据开票点.上级id%Type,
        strSQL = strSQL & ZVal(Val(txtEdit(mEdit.Edit_上级).Tag)) & ","
        '  编码_In     In 电子票据开票点.编码%Type,
        strSQL = strSQL & "'" & txtTemp.Text & txtEdit(mEdit.Edit_编码).Text & "',"
        '  名称_In     In 电子票据开票点.名称%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_名称).Text & "',"
        '  简码_In     In 电子票据开票点.简码%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_简码).Text & "',"
        '  院区_In     In 电子票据开票点.院区%Type,
        strSQL = strSQL & str站点 & ","
        '  客户端_In   In 电子票据开票点.客户端%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_客户端).Text & "',"
        '  部门id_In   In 电子票据开票点.部门id%Type,
        strSQL = strSQL & IIf(Val(txtEdit(mEdit.Edit_部门).Tag) = 0, "NULL", "'" & Val(txtEdit(mEdit.Edit_部门).Tag) & "'") & ","
        '  位置_In     In 电子票据开票点.位置%Type
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_位置).Text & "')"
    End If

    Call zlDatabase.ExecuteProcedure(strSQL, "电子票据开票点")
    
    Save开票点 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save开票点对照() As Boolean
    Dim strSQL As String
    Dim lng对照ID As Long
    
    On Error GoTo errHandle
    
    If mlng对照ID = 0 Then       '新增一条记录
        lng对照ID = zlDatabase.GetNextId("电子票据开票点")
        '  Zl_票据开票点对照_Update
        strSQL = "Zl_票据开票点对照_Update("
        '  操作_In     In Number,
        strSQL = strSQL & 0 & ","
        '  Id_In       In 票据开票点对照.Id%Type := Null,
        strSQL = strSQL & lng对照ID & ","
        '  开票点id_In In 电子票据开票点.Id%Type := Null,
        strSQL = strSQL & Val(mstrID) & ","
        '  人员id_In   In 票据开票点对照.人员id%Type := Null,
        strSQL = strSQL & ZVal(txtEdit(Edit_收费员).Tag) & ","
        '  客户端_In   In 票据开票点对照.客户端%Type := Null
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_客户端).Text & "')"
    Else
        '修改
        '  Zl_票据开票点对照_Update
        strSQL = "Zl_票据开票点对照_Update("
        '  操作_In     In Number,
        strSQL = strSQL & 1 & ","
        '  Id_In       In 票据开票点对照.Id%Type := Null,
        strSQL = strSQL & mlng对照ID & ","
        '  开票点id_In In 电子票据开票点.Id%Type := Null,
        strSQL = strSQL & Val(mstrID) & ","
        '  人员id_In   In 票据开票点对照.人员id%Type := Null,
        strSQL = strSQL & ZVal(txtEdit(Edit_收费员).Tag) & ","
        '  客户端_In   In 票据开票点对照.客户端%Type := Null
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_客户端).Text & "')"
    End If

    Call zlDatabase.ExecuteProcedure(strSQL, "电子票据开票点")
    
    Save开票点对照 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Init开票点设置(ByVal strID As String, Optional ByVal str上级ID As String, Optional ByVal bln末级 As Boolean, Optional blnRefresh As Boolean)
    On Error GoTo errHandle
    'strID-电子票据开票点.id
    'str上级ID:上级id
    'bln末级:true-是末级,false-不是末级
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    mstrID = strID
    mbln末级 = bln末级
    mblnOK = False
    Call IniStationNo
    If Not mbln末级 Then Call SetControlVisib
    If strID <> "" Then
        strSQL = "Select a.Id, a.上级id, a.编码, b.编码 as 上级编码,a.名称, a.简码, a.院区, a.客户端, a.位置,a.部门id, b.名称 As 上级名称,c.名称 As 部门 " & _
        "   From 电子票据开票点 A, 电子票据开票点 B,部门表 C " & _
        "   Where a.上级id = b.Id(+) And a.部门id=c.id(+) And a.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
        mstr上级ID = IIf(IsNull(rsTemp("上级ID")), "", rsTemp("上级ID"))
        mstr上级编码 = IIf(IsNull(rsTemp("上级编码")), "", rsTemp("上级编码"))
        mstr编码 = rsTemp("编码")
        txtEdit(mEdit.Edit_上级).Text = IIf(IsNull(rsTemp("上级名称")), "无", rsTemp("上级名称"))
        txtEdit(mEdit.Edit_上级).Tag = IIf(IsNull(rsTemp("上级id")), "0", rsTemp("上级id"))
        txtTemp.Text = mstr上级编码
        '取得上级编码，本级编码长度等值
        txtTemp.MaxLength = GetLocalCodeLength(mstr上级ID, "电子票据开票点")
        'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
        txtEdit(mEdit.Edit_编码).Text = Mid(rsTemp("编码"), Len(txtTemp.Text) + 1)
        '求出包括子节点在内的最长编码
        mint编码 = GetDownCodeLength(mstrID, "电子票据开票点")
        '10 - (mint编码 - Len(mstr编码))这个公式的意思是要为它的孩子的编码留有余地
        txtEdit(mEdit.Edit_编码).MaxLength = IIf(txtTemp.MaxLength = 0, 10 - (mint编码 - Len(mstr编码)), txtTemp.MaxLength) - Len(mstr上级编码)
        txtEdit(mEdit.Edit_名称).Text = rsTemp("名称")
        txtEdit(mEdit.Edit_简码).Text = IIf(IsNull(rsTemp("简码")), "", rsTemp("简码"))
        txtEdit(mEdit.Edit_位置).Text = IIf(IsNull(rsTemp("位置")), "", rsTemp("位置"))
        txtEdit(mEdit.Edit_客户端).Text = IIf(IsNull(rsTemp("客户端")), "", rsTemp("客户端"))
        txtEdit(mEdit.Edit_部门).Text = IIf(IsNull(rsTemp("部门")), "", rsTemp("部门"))
        txtEdit(mEdit.Edit_部门).Tag = IIf(IsNull(rsTemp("部门id")), "", rsTemp("部门id"))
        SetStationNo (IIf(IsNull(rsTemp("院区")), "", rsTemp("院区")))
    Else
        If str上级ID = "oot" Then
            mstr上级ID = ""
            mstr上级编码 = ""
            txtTemp.Text = ""
            txtEdit(mEdit.Edit_上级).Text = "无"
            '取得上级编码，本级编码长度等值
            txtTemp.MaxLength = GetLocalCodeLength("", "电子票据开票点")
        Else
            strSQL = "select 编码 as 上级编码,名称 as 上级名称,ID as 上级ID from 电子票据开票点 where ID=[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str上级ID))
                        
            mstr上级ID = IIf(IsNull(rsTemp("上级ID")), "", rsTemp("上级ID"))
            mstr上级编码 = IIf(IsNull(rsTemp("上级编码")), "", rsTemp("上级编码"))
            txtEdit(mEdit.Edit_上级).Text = IIf(IsNull(rsTemp("上级名称")), "无", rsTemp("上级名称"))
            txtEdit(mEdit.Edit_上级).Tag = IIf(IsNull(rsTemp("上级id")), "0", rsTemp("上级id"))
            txtTemp.Text = mstr上级编码
            '判断编码是否满了
            If Len(mstr上级编码) = mlng编码长度 Then
                MsgBox "不能再增加子部门了，编码长度已经用尽。", vbExclamation, gstrSysName
                Exit Sub
            End If
            '取得上级编码，本级编码长度等值
            txtTemp.MaxLength = GetLocalCodeLength(mstr上级ID, "电子票据开票点")
            'txtTemp.MaxLength为0表示该父节点还没有子节点，要设多长都随便
        End If
        txtEdit(mEdit.Edit_编码).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(mstr上级编码)
        txtEdit(mEdit.Edit_编码).Text = GetMaxLocalCode(mstr上级ID, "电子票据开票点")
        mstr编码 = mstr上级编码 & txtEdit(1).Text
    End If

    mblnChange = False
    Me.Show vbModal
    blnRefresh = mblnOK
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub Init开票点对码(ByVal intMode As Integer, ByVal lng开票点id As Long, Optional ByVal lng对照ID As Long, Optional blnRefresh As Boolean)
    On Error GoTo errHandle
    'intMode:0-按客户端对,1-按收费员对;2-按收费员+客户端对
    'lng开票点id:电子票据开票点.id
    'lng对照ID-修改开票点对码时传入
    
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strFilter As String
    
    If lng开票点id = 0 Then Exit Sub
    mbln开票点对码 = True
    mblnOK = False
    mstrID = lng开票点id
    mintMode = intMode
    mlng对照ID = lng对照ID
    If lng对照ID > 0 Then strFilter = " And b.id=[2] "
    strSQL = "Select a.Id, a.名称, nvl(b.客户端,a.客户端)As 客户端,b.人员id As 收费员id,c.姓名 As 收费员 " & _
    "   From 电子票据开票点 A,票据开票点对照 B,人员表 C " & _
    "   Where a.id=b.开票点id(+) And b.人员id=c.id(+)  And a.ID=[1]" & strFilter
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng开票点id, lng对照ID)
    If rsTemp.EOF Then Exit Sub
    txtEdit(mEdit.Edit_名称).Text = IIf(IsNull(rsTemp("名称")), "", rsTemp("名称"))
    txtEdit(mEdit.Edit_名称).Enabled = False
    If lng对照ID <> 0 Then
        If mintMode = 0 Then
            txtEdit(mEdit.Edit_客户端).Text = IIf(IsNull(rsTemp("客户端")), "", rsTemp("客户端"))
        ElseIf mintMode = 1 Then
            txtEdit(mEdit.Edit_收费员).Text = IIf(IsNull(rsTemp("收费员")), "", rsTemp("收费员"))
            txtEdit(mEdit.Edit_收费员).Tag = IIf(IsNull(rsTemp("收费员id")), "0", rsTemp("收费员id"))
        Else
            txtEdit(mEdit.Edit_客户端).Text = IIf(IsNull(rsTemp("客户端")), "", rsTemp("客户端"))
            txtEdit(mEdit.Edit_收费员).Text = IIf(IsNull(rsTemp("收费员")), "", rsTemp("收费员"))
            txtEdit(mEdit.Edit_收费员).Tag = IIf(IsNull(rsTemp("收费员id")), "0", rsTemp("收费员id"))
        End If
    End If
    mblnChange = False
    Call SetControlStation
    Me.Show vbModal
    blnRefresh = mblnOK
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd部门_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    strSQL = "Select Distinct a.id,a.编码,a.名称, a.位置 From 部门表 A, 部门性质说明 B Where a.Id = b.部门id And Nvl(b.服务对象, 0) <> 0 Order  By a.名称 "
    vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_部门).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取部门", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_部门).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(mEdit.Edit_部门).Text = rsTemp("名称")
        txtEdit(mEdit.Edit_部门).Tag = rsTemp("id")
        txtEdit(mEdit.Edit_位置).Text = NVL(rsTemp("位置"))
    End If
    zlControl.ControlSetFocus txtEdit(mEdit.Edit_部门)
End Sub

Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：本级ID，表名
    '输出参数：成功返回 下级最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with 上级ID is null " & strWhere & " connect by prior id=上级id"
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with 上级ID=" & strID & strWhere & " connect by prior id=上级id"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDownCodeLength")
    
    If rsTemp.RecordCount = 0 Then
        GetDownCodeLength = 0
    Else
        GetDownCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetDownCodeLength = 0
End Function

Private Sub cmd客户端_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    strSQL = "Select Rownum As id,Upper(工作站) as 工作站, Upper(用途) as 用途,Upper(部门) as 部门 From zlclients Order  By 工作站 "
    vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_客户端).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取客户端", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_客户端).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(mEdit.Edit_客户端).Text = rsTemp("工作站")
    End If
    zlControl.ControlSetFocus txtEdit(mEdit.Edit_客户端)
End Sub

Private Sub cmd上级_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim strID As String
    Dim str编码 As String
    Dim int编码  As Integer
    Dim vRect As RECT, rsTemp As ADODB.Recordset
    
    If mstrID <> "" Then
        strSQL = "select id,上级id,名称,编码,简码 from 电子票据开票点 where 撤档时间=to_date('3000-01-01','YYYY-MM-DD') and Nvl(末级, 0) = 0 and id<>" & mstrID & " start with 上级id is null connect by prior id =上级id And 上级id<>" & mstrID
    Else
        strSQL = "select id,上级id,名称,编码,简码 from 电子票据开票点 where 撤档时间=to_date('3000-01-01','YYYY-MM-DD') and Nvl(末级, 0) = 0 start with 上级id is null connect by prior id =上级id "
    End If
    strID = mstr上级ID
    str名称 = txtEdit(mEdit.Edit_上级).Text
    str编码 = txtTemp.Text
    vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_上级).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取上级", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_位置).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(mEdit.Edit_上级).Text = rsTemp("名称")
        txtEdit(mEdit.Edit_上级).Tag = rsTemp("id")
        int编码 = GetLocalCodeLength(txtEdit(mEdit.Edit_上级).Tag, "电子票据开票点")
        strID = rsTemp("id")
        str名称 = rsTemp("名称")
        str编码 = rsTemp("编码")
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_上级)
        '只有修改才有必要审核
        If mstrID <> "" Then
            If mint编码 - Len(mstr编码) + IIf(int编码 = 0, Len(str编码) + 1, int编码) > 10 Then
                MsgBox "这个上级不合适，因为它的编码太长了。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        mstr上级ID = strID
        txtEdit(mEdit.Edit_上级).Text = str名称
        txtTemp.MaxLength = int编码
        txtTemp.Text = str编码
        If mstrID <> "" Then
            txtEdit(mEdit.Edit_编码).MaxLength = IIf(txtTemp.MaxLength = 0, 10 - (mint编码 - Len(mstr编码)), txtTemp.MaxLength) - Len(str编码)
        Else
            txtEdit(mEdit.Edit_编码).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(str编码)
        End If
        txtEdit(mEdit.Edit_编码).Text = GetMaxLocalCode(mstr上级ID, "电子票据开票点")
    End If

    mblnChange = True
End Sub

Private Sub cmd收费员_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    strSQL = "Select Distinct a.Id, a.姓名, a.性别, a.出生日期" & vbNewLine & _
                    "From 人员表 A, 人员性质说明 B" & vbNewLine & _
                    "Where a.Id = b.人员id And b.人员性质 In ('门诊挂号员', '门诊收费员', '预交收款员', '住院结帐员')"
    vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_收费员).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取收费员", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_收费员).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(mEdit.Edit_收费员).Text = rsTemp("姓名")
        txtEdit(mEdit.Edit_收费员).Tag = rsTemp("ID")
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_收费员)
    Else
        MsgBox "没有找到有效的收费员！", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Activate()
    If Not mbln开票点对码 Then
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_名称)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Edit_名称 Then
        txtEdit(mEdit.Edit_简码).Text = zlStr.GetCodeByVB(txtEdit(mEdit.Edit_名称).Text)
    ElseIf Index = Edit_收费员 Then
        If txtEdit(Index) = "" Then txtEdit(Index).Tag = "0"
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = mEdit.Edit_名称 Or Index = mEdit.Edit_位置 Then
        OS.OpenIme True
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    If Index = mEdit.Edit_编码 Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    ElseIf Index = mEdit.Edit_名称 Or Index = mEdit.Edit_简码 Then
        If LenB(StrConv(txtEdit(mEdit.Edit_名称).Text & Chr(KeyAscii), vbFromUnicode)) > 100 And (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack) Then
            KeyAscii = 0
        End If
    ElseIf Index = mEdit.Edit_位置 Then
'        If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
    ElseIf Index = mEdit.Edit_客户端 Then
        If KeyAscii <> vbKeyReturn Then Exit Sub
        strSQL = "Select Rownum As id,Upper(工作站) as 工作站, Upper(用途) as 用途,Upper(部门) as 部门  From zlClients " & _
                  "Where 工作站 Like Upper([1]) Or 用途 Like Upper([1]) Or 部门 Like Upper([1]) " & _
                  "   Or Upper(zlPinYinCode(工作站)) Like Upper([1]) Or Upper(zlPinYinCode(用途)) Like Upper([1]) Or Upper(zlPinYinCode(部门)) Like Upper([1]) Order By 工作站 "
                  
        vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_客户端).hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取客户端", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_客户端).Height, True, False, False, "%" & txtEdit(mEdit.Edit_客户端).Text & "%")
         If Not rsTemp Is Nothing Then
            txtEdit(mEdit.Edit_客户端).Text = rsTemp("工作站")
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_客户端)
        Else
            MsgBox "根据输入的信息未找到有效的客户端，请重试！", vbInformation, gstrSysName
            txtEdit(mEdit.Edit_客户端).Text = ""
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_客户端)
        End If
    ElseIf Index = mEdit.Edit_部门 Then
        If KeyAscii <> vbKeyReturn Then Exit Sub
        strSQL = "Select Distinct a.ID,a.编码,a.名称,a.位置 From 部门表 A, 部门性质说明 B Where a.Id = b.部门id And Nvl(b.服务对象, 0) <> 0 " & _
                  " And A.名称 Like Upper([1]) Or A.简码 Like Upper([1]) Or A.编码 Like Upper([1]) " & _
                  "   Or Upper(zlPinYinCode(A.名称)) Like Upper([1]) Or Upper(zlPinYinCode(A.简码)) Like Upper([1]) Order  By a.名称 "
                  
        vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_部门).hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取客户端", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_部门).Height, True, False, False, "%" & txtEdit(mEdit.Edit_部门).Text & "%")
         If Not rsTemp Is Nothing Then
            txtEdit(mEdit.Edit_部门).Text = rsTemp("名称")
            txtEdit(mEdit.Edit_部门).Tag = rsTemp("ID")
            txtEdit(mEdit.Edit_位置).Tag = NVL(rsTemp("位置"))
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_部门)
        Else
            MsgBox "根据输入的信息未找到有效的部门，请重试！", vbInformation, gstrSysName
            txtEdit(mEdit.Edit_部门).Text = ""
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_部门)
        End If
    ElseIf Index = mEdit.Edit_收费员 Then
        If KeyAscii <> vbKeyReturn Then Exit Sub
        strSQL = "Select Distinct a.Id, a.姓名, a.性别, a.出生日期" & vbNewLine & _
                        "From 人员表 A, 人员性质说明 B" & vbNewLine & _
                        "Where a.Id = b.人员id And b.人员性质 In ('门诊挂号员', '门诊收费员', '预交收款员', '住院结帐员')" & vbNewLine & _
                        " And a.姓名 Like Upper([1]) Or A.简码 Like Upper([1]) Or A.编号 Like Upper([1])"
        vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_收费员).hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "获取收费员", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_收费员).Height, True, False, False, "%" & txtEdit(mEdit.Edit_收费员).Text & "%")
         If Not rsTemp Is Nothing Then
            txtEdit(mEdit.Edit_收费员).Text = rsTemp("姓名")
            txtEdit(mEdit.Edit_收费员).Tag = rsTemp("ID")
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_收费员)
        Else
            MsgBox "没有找到有效的收费员！", vbInformation, gstrSysName
            txtEdit(mEdit.Edit_收费员).Text = ""
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_收费员)
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Or Index = 5 Then
        OS.OpenIme False
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(1).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(1).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

Private Function CheckSame(ByVal strName As String, Optional ByVal lngID As Long) As Boolean
'----------------------------------------------
'功能：检查开票点是否与已有开票点名称相同
'----------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    If lngID = 0 Then
        strSQL = "Select 1 From 电子票据开票点 " & _
              "Where  名称 = [1] "
    Else
      strSQL = "Select 1 From 电子票据开票点 " & _
              "Where  名称 = [1]  and id<> [2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查开票点是否与已有开票点名称相同", strName, lngID)
    CheckSame = Not rsTemp.EOF

    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check重复部门(ByVal str上级ID As String, ByVal str名称 As String) As Boolean
    '功能：用来检查是否已具有该部门
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "select 名称 from 电子票据开票点 where 上级id=[1] and 名称=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询是否有重复部门", str上级ID, str名称)
    If rsTemp.EOF Then
        Check重复部门 = False
    Else
        Check重复部门 = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMaxLocalCode(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '功能描述：根据指定表的上级ID 读取本级的最大编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 最大编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strAllCode As String
    Dim intLength   As Integer
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select nvl(max(to_number(编码)),0)+1 as MaxCode from " & strTableName & " where 上级ID is null" & strWhere
        
        '如果是部门表，则要排除"已删除部门"分类的ID
        If strTableName = "部门表" Then
            strSQL = strSQL & " And 编码 <> '-'"
        End If
    Else
        strSQL = "select nvl(max(to_number(编码)),0)+1 as MaxCode from " & strTableName & " where 上级ID=" & str上级ID & strWhere
    End If
    intCode = GetLocalCodeLength(str上级ID, strTableName, strWhere)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetMaxLocalCode")
    
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    'strCode = Mid(strAllCode, Len(GetParentCode(str上级ID, strTableName)) + 1)
    'GetMaxLocalCode = String(intCode - Len(strAllCode), "0") & strCode
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str上级ID, strTableName)) + 1)
    If GetMaxLocalCode = "" Then GetMaxLocalCode = "1"
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Function GetParentCode(ByVal str上级ID As String, ByVal strTableName As String) As String
    '功能描述：读取上级编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 上级编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        GetParentCode = ""
        Exit Function
    Else
        strSQL = "select 编码 from " & strTableName & " where ID=" & str上级ID
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetParentCode")
    
    If rsTemp.RecordCount = 0 Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("编码").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetLocalCodeLength(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：上级ID，表名
    '输出参数：成功返回 最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级ID is null" & strWhere
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级ID=" & str上级ID & strWhere
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetLocalCodeLength")
    
    If rsTemp.RecordCount = 0 Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Private Sub SetControlVisib()
    '设置控件的可见性
    Me.Caption = "开票点分类设置"
    Me.Height = 2250
    lblEdit(Edit_客户端).Visible = False: txtEdit(Edit_客户端).Visible = False
    lblEdit(Edit_位置).Visible = False: txtEdit(Edit_位置).Visible = False
    lblEdit(Edit_部门).Visible = False: txtEdit(Edit_部门).Visible = False
    lblStationNo.Visible = False: cmbStationNo.Visible = False
    cmd客户端.Enabled = False: cmd客户端.Visible = False
    cmd部门.Enabled = False: cmd部门.Visible = False
End Sub

Private Sub SetControlStation()
    '设置控件的位置
    Me.Caption = "开票点对码"
    Me.Height = 1900
    lblEdit(Edit_编码).Visible = False: txtEdit(Edit_编码).Visible = False: txtTemp.Visible = False
    lblEdit(Edit_上级).Visible = False: txtEdit(Edit_上级).Visible = False: cmd上级.Visible = False
    lblEdit(Edit_简码).Visible = False: txtEdit(Edit_简码).Visible = False
    
    lblEdit(Edit_客户端).Top = lblEdit(Edit_编码).Top: txtEdit(Edit_客户端).Top = txtTemp.Top: cmd客户端.Top = txtTemp.Top + 15
    lblEdit(Edit_收费员).Top = lblEdit(Edit_名称).Top: txtEdit(Edit_收费员).Top = txtEdit(Edit_名称).Top: cmd收费员.Top = txtEdit(Edit_名称).Top + 15

    lblEdit(Edit_名称).Top = lblEdit(Edit_上级).Top: txtEdit(Edit_名称).Top = txtEdit(Edit_上级).Top
    If mintMode = 2 Then Exit Sub
    cmdOK.Top = lblEdit(Edit_上级).Top: cmdCancel.Top = 600
    Me.Height = 1585
    If mintMode = 0 Then
        lblEdit(Edit_客户端).Top = 700: txtEdit(Edit_客户端).Top = 650: cmd客户端.Top = 665
        lblEdit(Edit_收费员).Visible = False: txtEdit(Edit_收费员).Visible = False: cmd收费员.Visible = False
    Else
        lblEdit(Edit_收费员).Top = 700: txtEdit(Edit_收费员).Top = 650: cmd收费员.Top = 665
        lblEdit(Edit_客户端).Visible = False: txtEdit(Edit_客户端).Visible = False: cmd客户端.Visible = False
    End If
End Sub

