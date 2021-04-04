VERSION 5.00
Begin VB.Form frmIllSortEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病分类编辑"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmIllSortEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   1110
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1464
      Width           =   3885
   End
   Begin VB.CommandButton cmd上级 
      Caption         =   "…"
      Height          =   240
      Left            =   4710
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1110
      MaxLength       =   6
      TabIndex        =   1
      Top             =   195
      Width           =   1395
   End
   Begin VB.CheckBox chk病人 
      Caption         =   "疾病疗效只能是其他(&S)"
      Height          =   195
      Left            =   2850
      TabIndex        =   2
      Top             =   248
      Width           =   2205
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1890
      Width           =   3885
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   14
      Top             =   2520
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -150
      TabIndex        =   15
      Top             =   2310
      Width           =   5445
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   13
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2610
      TabIndex        =   12
      Top             =   2520
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1110
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1041
      Width           =   3885
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1110
      MaxLength       =   150
      TabIndex        =   4
      Top             =   618
      Width           =   3885
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "编码范围(&R)"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   7
      Top             =   1530
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "序号(&T)"
      Height          =   180
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   255
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "上级(&D)"
      Height          =   180
      Index           =   3
      Left            =   420
      TabIndex        =   9
      Top             =   1950
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "简码(&J)"
      Height          =   180
      Index           =   2
      Left            =   420
      TabIndex        =   5
      Top             =   1101
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "名称(&N)"
      Height          =   180
      Index           =   1
      Left            =   420
      TabIndex        =   3
      Top             =   678
      Width           =   630
   End
End
Attribute VB_Name = "frmIllSortEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrID As String             '当前编辑的项目ID
Dim mstr上级项目ID As String     '当前编辑的上级项目ID
Dim mstr编码类别 As String

Dim mblnChange As Boolean  '已修改

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save项目() = False Then Exit Sub
    
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstrID = ""
    txtEdit(0).Text = ""
    txtEdit(1).Text = ""
    txtEdit(2).Text = ""
    txtEdit(4).Text = ""
    chk病人.Value = 0
    txtEdit(0).SetFocus
    mblnChange = False
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能:分析输入编码类别的内容是否有效
'参数:
'返回值:有效返回True,否则为False
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 4
        If i <> 3 Then
            strTemp = Trim(txtEdit(i).Text)
            If zlCommFun.StrIsValid(Trim(txtEdit(i).Text), txtEdit(i).MaxLength) = False Then
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        End If
    Next
    
    If Not IsNumeric(txtEdit(0).Text) Then
        MsgBox "请输入正整数。", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    If Val(txtEdit(0).Text) <= 0 Or Val(txtEdit(0).Text) > 999999 Then
        MsgBox "序号不能小于或等于零，且要小于1000000。", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    If Val(txtEdit(0).Text) <> Int(txtEdit(0).Text) Then
        MsgBox "请输入正整数。", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    If Len(Trim(txtEdit(1).Text)) = 0 Then
        MsgBox "名称不能为空。", vbExclamation, gstrSysName
        txtEdit(1).Text = ""
        txtEdit(1).SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save项目() As Boolean
'功能:保存编辑的内容到编码类别表中
'参数:
'返回值:成功返回True,否则为False
    Dim lng分类id As Long
    Dim nodTemp As Node
    On Error GoTo ErrHandle
    
    If mstrID = "" Then       '新增一条记录
        lng分类id = zlDatabase.GetNextId("疾病编码分类")
        
        gstrSQL = "ZL_疾病编码分类_INSERT(" & lng分类id & ",'" & mstr上级项目ID & "'," & txtEdit(0).Text & _
                ",'" & txtEdit(1).Text & "','" & UCase(txtEdit(2).Text) & "','" & txtEdit(4).Text & "','" & mstr编码类别 & "'," & IIF(chk病人.Value = 1, 0, 1) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Else    '修改
        lng分类id = mstrID
        gstrSQL = "ZL_疾病编码分类_UPDATE(" & lng分类id & ",'" & mstr上级项目ID & "'," & txtEdit(0).Text & _
                ",'" & txtEdit(1).Text & "','" & UCase(txtEdit(2).Text) & "','" & txtEdit(4).Text & "'," & IIF(chk病人.Value = 1, 0, 1) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    '更新管理窗口
    With frmIllManage.tvwMain_S
        If mstrID = "" Then
            '新增分类
            If mstr上级项目ID = "" Then
                Set nodTemp = .Nodes.Add(, , "K" & lng分类id, "【" & txtEdit(0).Text & "】" & Trim(txtEdit(1).Text), "Root", "Root")
            Else
                Set nodTemp = .Nodes.Add("K" & mstr上级项目ID, tvwChild, "K" & lng分类id, "【" & txtEdit(0).Text & "】" & Trim(txtEdit(1).Text), "Root", "Root")
            End If
        Else
            '修改分类
            Set nodTemp = .Nodes("K" & lng分类id)
            nodTemp.Text = "【" & txtEdit(0).Text & "】" & Trim(txtEdit(1).Text)
            
            If mstr上级项目ID = "" Then
                If Not nodTemp.Parent Is Nothing Then
                    '改变其分类
                    Call frmIllManage.FillTree
                End If
            Else
                If Not nodTemp.Parent Is .Nodes("K" & mstr上级项目ID) Then
                    '改变其分类
                    Set nodTemp.Parent = .Nodes("K" & mstr上级项目ID)
                End If
            End If
        End If
        .Nodes("K" & lng分类id).EnsureVisible
    End With
        
    Save项目 = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function 疾病编辑(ByVal str上级项目 As String, ByVal str上级项目ID As String, _
    ByVal str编码类别 As String, Optional ByVal strID As String = "") As Boolean
'功能:用来与调用的编码类别管理窗口进行通讯的程序
'参数:str上级项目     上级编码类别的名字
'     str上级项目ID   上级编码类别的ID
'     str编码类别     整个编码的类别
'     strID           本编码类别的的ID
'返回值:编辑成功返回True,否则为False
    
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mstr编码类别 = str编码类别
    
    mstrID = strID
    
    On Error GoTo ErrHandle
    If strID <> "" Then
        rsTemp.CursorLocation = adUseClient
        
        gstrSQL = "select A.ID,A.上级ID,A.名称,A.简码,A.编码范围,A.序号,A.是否病人,B.序号 as 上级序号,B.名称 as 上级名称 " & _
                " from 疾病编码分类 A,疾病编码分类 B " & _
                " where B.ID(+)=A.上级ID and A.ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
                
        txtEdit(0).Text = rsTemp("序号")
        txtEdit(1).Text = Trim(rsTemp("名称"))
        txtEdit(2).Text = IIF(IsNull(rsTemp("简码")), "", rsTemp("简码"))
        txtEdit(4).Text = IIF(IsNull(rsTemp("编码范围")), "", rsTemp("编码范围"))
        chk病人.Value = IIF(rsTemp("是否病人") = 1, 0, 1)
        mstr上级项目ID = IIF(IsNull(rsTemp("上级ID")), "", rsTemp("上级ID"))
        
        If IsNull(rsTemp("上级名称")) Then
            txtEdit(3).Text = "无"
        Else
            txtEdit(3).Text = "【" & rsTemp("上级序号") & "】" & Trim(rsTemp("上级名称"))
        End If
        
    Else
        mstr上级项目ID = str上级项目ID
        txtEdit(3).Text = str上级项目
    End If
    
    mblnChange = False
    frmIllSortEdit.Show vbModal
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmd上级_Click()
    Dim blnRe As Boolean
    Dim str名称 As String
    Dim str上级ID As String
    Dim str编码范围 As String
    
    str上级ID = mstr上级项目ID
    str名称 = txtEdit(3).Text
    blnRe = frmClassSel.ShowTree(str上级ID, str名称, str编码范围, mstr编码类别, mstrID)
    '成功返回
    If blnRe Then
        '新的本级的宽度
        mstr上级项目ID = str上级ID
        txtEdit(3).Text = str名称
        mblnChange = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk病人_Click()
    mblnChange = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        txtEdit(2).Text = zlStr.GetCodeByVB(txtEdit(1).Text)
    ElseIf Index = 2 Then
        txtEdit(2).Text = UCase(txtEdit(2).Text)
    End If
    mblnChange = True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        '要作报表名称，所以不能有怪字符
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf Index = 0 Then
        '序号只允许输入数字
        If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    ElseIf Index = 4 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
        '只能取这些字母
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.,-" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 1 Then
        zlCommFun.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Or Index = 0 Or Index = 4 Then
        zlCommFun.OpenIme False
    End If
End Sub
