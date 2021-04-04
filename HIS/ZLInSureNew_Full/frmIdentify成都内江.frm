VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmIdentify成都内江 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin ZL9BillEdit.BillEdit msf其他病种 
      Height          =   1275
      Left            =   915
      TabIndex        =   48
      Top             =   4845
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   2249
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.ComboBox cbo出院类别 
      Height          =   300
      ItemData        =   "frmIdentify成都内江.frx":0000
      Left            =   915
      List            =   "frmIdentify成都内江.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4035
      Width           =   2295
   End
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   285
      Left            =   6720
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4470
      Width           =   255
   End
   Begin VB.ComboBox cbo类别 
      Height          =   300
      ItemData        =   "frmIdentify成都内江.frx":0004
      Left            =   915
      List            =   "frmIdentify成都内江.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   885
      Width           =   2295
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5925
      TabIndex        =   9
      Top             =   6390
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4635
      TabIndex        =   8
      Top             =   6390
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   12
      Top             =   300
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -465
      TabIndex        =   10
      Top             =   6210
      Width           =   8340
   End
   Begin VB.TextBox TxtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   915
      MaxLength       =   20
      TabIndex        =   1
      Top             =   510
      Width           =   2295
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4605
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   510
      Width           =   2385
   End
   Begin VB.CommandButton cmd修改密码 
      Caption         =   "修改密码"
      Height          =   350
      Left            =   225
      TabIndex        =   11
      Top             =   6390
      Width           =   1100
   End
   Begin VB.TextBox txt病种 
      Height          =   315
      Left            =   930
      TabIndex        =   7
      Top             =   4440
      Width           =   6075
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "其他病种"
      Height          =   285
      Left            =   150
      TabIndex        =   47
      Top             =   4845
      Width           =   750
   End
   Begin VB.Label 出院类别 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "出院类别"
      Height          =   180
      Left            =   135
      TabIndex        =   46
      Top             =   4095
      Width           =   720
   End
   Begin VB.Label lbl病种 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "病种(&F)"
      Height          =   180
      Left            =   240
      TabIndex        =   37
      Top             =   4515
      Width           =   630
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   915
      TabIndex        =   44
      Top             =   3645
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "制卡单位"
      Height          =   180
      Index           =   17
      Left            =   180
      TabIndex        =   43
      Top             =   3690
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   4605
      TabIndex        =   42
      Top             =   3645
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "在职情况"
      Height          =   180
      Index           =   16
      Left            =   3825
      TabIndex        =   41
      Top             =   3690
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   4605
      TabIndex        =   40
      Top             =   3240
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   15
      Left            =   3825
      TabIndex        =   39
      Top             =   3285
      Width           =   720
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "请在密码栏输入病人的IC卡密码后回车,将读取病人的相关信息。"
      Height          =   180
      Left            =   720
      TabIndex        =   38
      Top             =   60
      Width           =   5130
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保卡号"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   555
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "个人编号"
      Height          =   180
      Index           =   1
      Left            =   3825
      TabIndex        =   36
      Top             =   945
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   2
      Left            =   540
      TabIndex        =   35
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   4185
      TabIndex        =   34
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   33
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出身日期"
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   32
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "工况类别"
      Height          =   180
      Index           =   6
      Left            =   3825
      TabIndex        =   31
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡有效期"
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   30
      Top             =   2895
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   8
      Left            =   540
      TabIndex        =   29
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "统筹编号"
      Height          =   180
      Index           =   9
      Left            =   3825
      TabIndex        =   28
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "补卡次数"
      Height          =   180
      Index           =   10
      Left            =   3825
      TabIndex        =   27
      Top             =   2895
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "制卡日期"
      Height          =   180
      Index           =   11
      Left            =   3825
      TabIndex        =   26
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位号码"
      Height          =   180
      Index           =   12
      Left            =   180
      TabIndex        =   25
      Top             =   3285
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4605
      TabIndex        =   24
      Top             =   900
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   915
      TabIndex        =   23
      Top             =   1305
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4605
      TabIndex        =   22
      Top             =   1305
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   915
      TabIndex        =   21
      Top             =   1695
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4605
      TabIndex        =   20
      Top             =   1695
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   915
      TabIndex        =   19
      Top             =   2085
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4605
      TabIndex        =   18
      Top             =   2085
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   915
      TabIndex        =   17
      Top             =   2475
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4605
      TabIndex        =   16
      Top             =   2475
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   915
      TabIndex        =   15
      Top             =   2850
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   4605
      TabIndex        =   14
      Top             =   2850
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   915
      TabIndex        =   13
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Index           =   13
      Left            =   4185
      TabIndex        =   2
      Top             =   555
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "交易类别"
      Height          =   180
      Index           =   14
      Left            =   180
      TabIndex        =   4
      Top             =   945
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify成都内江"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
Private mlng病人ID As Long
Private mstrReturn As String
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private mstr并发症 As String '20051026 陈东

Private Sub cbo出院类别_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd修改密码_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    Dim StrInput As String, strOutput As String
    
    If InitInfor_成都内江.读卡器_内江 = 0 Then
        '明华修改密码
    
        strNewPassWord = frm修改密码.ChangePassword(strOldPassWord, strOldPassWord)
        
        If strOldPassWord = strNewPassWord Then Exit Sub
        If strNewPassWord = "" Then Exit Sub
        '    a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
        '    b)  OldPassword：输入参数，为原密码，要求长度为6，字符串中只能包含0到9的数字；
        '    c)  NewPassword：输入参数，为新密码，要求长度为6，字符串中只能包含0到9的数字。
        StrInput = InitInfor_成都内江.串号号_内江
        StrInput = StrInput & vbTab & strOldPassWord
        StrInput = StrInput & vbTab & strNewPassWord
        
        If 业务请求_成都内江(更改密码_内江, StrInput, strOutput) = False Then Exit Sub
        TxtEdit(1).Text = strNewPassWord
    Else
        '其他只能读卡
    End If
    If ReadCardInFo() = False Then Exit Sub
    Call LoadCtrlData
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    '如果不是明华的读卡，则不需输入密码
    If InitInfor_成都内江.读卡器_内江 = 0 Then Exit Sub
    TxtEdit(1).Enabled = False
    TxtEdit(1).BackColor = TxtEdit(0).BackColor
    
    '读卡
    If ReadCardInFo() = False Then Exit Sub
    Call LoadCtrlData
    Me.cmd修改密码.Caption = "读卡(&R)"
    
    
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    'Beging 20051026 陈东
    Dim i As Long
    Dim vat并发症 As Variant
    
    With msf其他病种
        '设置行数及列数及列标题名称
        .Rows = 4
        .Cols = 1
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 4800
        .TextMatrix(0, 0) = "病种编码与名称"
        
        '设置各列的列值，以确定哪些列可操作、可编辑或非编辑列
        .ColData(0) = 1   '文本框输入且有命令按钮
        '未设置的列的列值均为0 (默认), 这些列将可以选择但不能修改
    End With
    
    If mstr并发症 <> "" Then
        If InStr(mstr并发症, "|") > 0 Then
            vat并发症 = Split(mstr并发症, "|")
            msf其他病种.Rows = UBound(vat并发症) + 1
            For i = 0 To UBound(vat并发症) - 1
                msf其他病种.TextMatrix(i + 1, 0) = "[" & Split(vat并发症(i), ";")(0) & "]"      '& Split(vat并发症(i), ";")(1)
            Next
        End If
    End If
    'End 20051026 陈东
End Sub


Private Sub msf其他病种_CommandClick()
    Dim str病种 As String
    Select Case msf其他病种.ColData(msf其他病种.Col)
        Case 0
            str病种 = msf其他病种.TextMatrix(msf其他病种.Row, msf其他病种.Col)
            str病种 = BZXZ_成都内江(str病种)
            If str病种 = "" Then Exit Sub
            msf其他病种.TextMatrix(msf其他病种.Row, msf其他病种.Col) = str病种
    End Select
End Sub

Private Sub msf其他病种_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim str病种 As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    str病种 = msf其他病种.Text
    
    If str病种 = "" And msf其他病种.Rows = msf其他病种.Row + 1 Then
        SendKeys "{Tab}"
    End If
    
    If str病种 = "" And msf其他病种.Rows = msf其他病种.Row + 2 Then
        If msf其他病种.TextMatrix(msf其他病种.Row + 1, msf其他病种.Col) = "" Then
            SendKeys "{Tab}"
        End If
    End If
    
    'Cancel = True
    str病种 = BZXZ_成都内江(str病种, 1)
    If str病种 <> "" Then
        msf其他病种.Text = str病种
        msf其他病种.TextMatrix(msf其他病种.Row, msf其他病种.Col) = str病种
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        TxtEdit(Index).Tag = ""
        g病人身份_成都内江.个人编号 = ""
        g病人身份_成都内江.卡号 = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    mblnChange = True
    
    If Index = 1 Then
        '密码输入完毕
        '需获取病人信息
         SetOKCtrl False
         If ReadCardInFo = False Then Exit Sub
        '初始值
        Call LoadCtrlData
        SetOKCtrl True
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
Private Function ReadCardInFo() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取读卡信息
    '--入参数:
    '--出参数:
    '--返  回:成功,返回true,否则返回False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String
     '读取病人信息
        '   a)  Port：输入参数，为通讯端口号，0、1、2、3分别代表串口1、2、3、4;并口为其I/O地址（如0x378）；建议将读卡器连接到串口1；
        '   b)  UserPassword：输入参数，为用户密码，要求长度为6，字符串中只能包含0到9的数字；
           
    ReadCardInFo = False
    StrInput = InitInfor_成都内江.串号号_内江
    If InitInfor_成都内江.读卡器_内江 = 0 Then
        '明华需输入密码
        If Trim(TxtEdit(1)) = "" Then
            ShowMsgbox "请输入IC卡密码!"
            If TxtEdit(1).Enabled Then TxtEdit(1).SetFocus
            Exit Function
        End If
        StrInput = StrInput & vbTab & TxtEdit(1).Text
    End If
    
    Err = 0
    On Error GoTo ErrHand:
    
    If 获取参保人员信息_成都内江(StrInput) = False Then Exit Function
    ReadCardInFo = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmd确定.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    IsValid = False
    If Trim(TxtEdit(0).Text) = "" Then
        MsgBox "还没有进行身份验证！", vbInformation, gstrSysName
        If TxtEdit(1).Enabled Then TxtEdit(1).SetFocus
        Exit Function
    End If
    
    If Trim(g病人身份_成都内江.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        If TxtEdit(1).Enabled Then TxtEdit(1).SetFocus
        Exit Function
    End If
    
    If cbo类别.Text = "" Then
        ShowMsgbox "交易类别未选择"
        Exit Function
    End If
    If cbo出院类别.Text = "" And mbytType = 4 Then
        ShowMsgbox "出院类别未选择"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查录前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=" & TYPE_成都内江 & " and 医保号='" & g病人身份_成都内江.统筹编号 & g病人身份_成都内江.个人编号 & "'"
            Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '设置
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_成都内江 & " and  医保号='" & g病人身份_成都内江.统筹编号 & g病人身份_成都内江.个人编号 & "'"
        
        zldatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
        End If
        rsTemp.Close
        mstrReturn = mlng病人ID
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim lng疾病ID As Long
    
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str类别 As String
    Dim int当前状态 As Integer
    Dim lng入院疾病ID As Long
    

    lng疾病ID = IIf(Val(Me.txt病种.Tag) = 0, 0, Val(Me.txt病种.Tag))
    
    If lng疾病ID <> 0 Then
        gstrSQL = "Select * From 保险病种 where id=" & lng疾病ID
        zldatabase.OpenRecordset rsTemp, gstrSQL, "获取病种"
        g病人身份_成都内江.病种编码 = Nvl(rsTemp!编码)
        g病人身份_成都内江.病种名称 = Nvl(rsTemp!名称)
    Else
        g病人身份_成都内江.病种编码 = ""
        g病人身份_成都内江.病种名称 = ""
        If mbytType = 1 Or mbytType = 4 Then
            ShowMsgbox "请必需选择病种!"
            Exit Sub
        End If
    End If
        
    If IsValid = False Then Exit Sub
    'Beging 20051026 其他病种
    If mbytType = 1 Or mbytType = 4 Then
        Dim str其他病种 As String, str其他病种编码 As String, i As Long
        
        For i = 1 To msf其他病种.Rows - 1
            str其他病种编码 = msf其他病种.TextMatrix(i, 0)
            If str其他病种编码 <> "" Then
                If InStr(str其他病种编码, "]") > 0 And InStr(str其他病种编码, "[") > 0 And InStr(str其他病种编码, "]") - InStr(str其他病种编码, "[") > 1 Then
                    str其他病种编码 = Mid(str其他病种编码, InStr(str其他病种编码, "[") + 1, InStr(str其他病种编码, "]") - InStr(str其他病种编码, "[") - 1)
                    str其他病种 = str其他病种 & str其他病种编码 & "|"
                End If
            End If
        Next
    End If
    'End 20051026 其他病种
    g病人身份_成都内江.交易类别 = Split(cbo类别.Text, "-")(0)
    If mbytType = 4 Then
    g病人身份_成都内江.出院类别 = Split(cbo出院类别.Text, "-")(0)
    End If
    int当前状态 = 0
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_成都内江 & " and  医保号='" & g病人身份_成都内江.统筹编号 & g病人身份_成都内江.个人编号 & "'"
        
        zldatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
            '>>Beging 陈东 20050601
            lng入院疾病ID = Nvl(rsTemp!病种ID, 0)
            '>> End
        End If
        rsTemp.Close
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证(工况类别);7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号(统筹地区编码|制卡日期|卡有效日期);16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_成都内江
        
        strIdentify = .卡号                               '0卡号
        strIdentify = strIdentify & ";" & .统筹编号 & .个人编号              '1医保号
        strIdentify = strIdentify & ";"                     '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & Decode(.性别, "1", "男", "2", "女", .性别)              '4性别
        strIdentify = strIdentify & ";" & .出生日期                '5出生日期
        strIdentify = strIdentify & ";" & .身份证号           '6身份证
        strIdentify = strIdentify & ";" & IIf(.单位号码 = "", "", "(" & .单位号码 & ")")            '7.单位名称(编码)
        strAddition = ";0"                                          '8.中心代码
        strAddition = strAddition & ";"                             '9.顺序号
        strAddition = strAddition & ";" & .交易类别                 '10人员身份
        strAddition = strAddition & ";" & .帐户余额                 '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态               '12当前状态
            'beging 陈东 20050601 出院时,是出院病种,不能将入院病种冲掉了
        If mbytType = 4 Then
            strAddition = strAddition & ";" & lng入院疾病ID                 '13病种ID
        Else
            strAddition = strAddition & ";" & lng疾病ID                 '13病种ID
        End If
            'End
        strAddition = strAddition & ";1"                            '14在职(1,2,3)
        strAddition = strAddition & ";" & .统筹编号 & "|" & .制卡日期 & "|" & .卡有效期 & "|" & .制卡单位 & "|" & .在职情况    '15退休证号
        strAddition = strAddition & ";" & .补卡次数                     '16年龄段
        strAddition = strAddition & ";" & .工况类别                            '17灰度级
        strAddition = strAddition & ";" & .帐户余额                             '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_成都内江)
    
    g病人身份_成都内江.lng病人ID = mlng病人ID
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
        'Beging 陈东 20050601 保存出院病种ID
        If mbytType = 4 Then
            gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_成都内江 & ",'出院病种ID','" & lng疾病ID & "')"
            Call zldatabase.ExecuteProcedure(gstrSQL, "更新出院病种")
            'beging 20051026 陈东
            gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_成都内江 & ",'出院其他病种','''" & str其他病种 & "''')"
            Call zldatabase.ExecuteProcedure(gstrSQL, "更新出院其他病种")
        End If
            '
        If mbytType = 1 Then
            gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_成都内江 & ",'其他病种','''" & str其他病种 & "''')"
            Call zldatabase.ExecuteProcedure(gstrSQL, "更新其他病种")
            gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_成都内江 & ",'附加诊断','''" & "" & "''')"
            Call zldatabase.ExecuteProcedure(gstrSQL, "更新附加诊断")
            'End 20051026 陈东
        End If
        'end
        
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    Dim rsTmp As New ADODB.Recordset
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    
    DebugTool "进入身份验证,并开始加入基本信息"
    
    If LoadBaseData = False Then
        DebugTool "加入失败(身份验证)"
        Exit Function
    End If
    DebugTool "加入成功(身份验证)"
    
    'Beging 20051026 陈东
    gstrSQL = "Select * from 保险帐户 Where 病人ID=" & lng病人ID & " And 险类=" & TYPE_成都内江
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "取并发症")
    If rsTmp.EOF = False Then
        If mbytType = 4 Then
            mstr并发症 = Nvl(rsTmp!出院其他病种)
        Else
            mstr并发症 = Nvl(rsTmp!其他病种)
        End If
    End If
    'End 20051026 陈东
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function
Private Function LoadBaseData() As Boolean
    '加载基础数据
    Dim rsTemp As New ADODB.Recordset
    LoadBaseData = False
    On Error GoTo ErrHand:
      
    If mbytType = 0 Or mbytType = 3 Or mbytType = 2 Then
        cbo类别.AddItem "0-普通门诊"
    Else
        cbo类别.AddItem "1-普通住院"
    End If
    cbo类别.ListIndex = cbo类别.NewIndex
    If mbytType = 4 Then
        cbo出院类别.AddItem "0-治愈出院"
        cbo出院类别.AddItem "1-好转出院"
        cbo出院类别.AddItem "2-未愈出院"
        cbo出院类别.AddItem "3-死亡"
        cbo出院类别.AddItem "4-自动出院"
        cbo出院类别.AddItem "5-转本统筹地区内的医院"
        cbo出院类别.AddItem "6-转本统筹地区外的医院"
        cbo出院类别.ListIndex = 0
    Else
        cbo出院类别.Enabled = False
    End If
    If mbytType = 0 Or mbytType = 3 Or mbytType = 2 Then
       msf其他病种.TabIndex = 47
    End If
    LoadBaseData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    With g病人身份_成都内江
        TxtEdit(0) = .卡号
        lblEdit(1) = .个人编号
        lblEdit(2) = .姓名
        lblEdit(3) = Decode(.性别, "1", "男", "2", "女", .性别)
        lblEdit(4) = .身份证号
        lblEdit(5) = .工况类别
        lblEdit(6) = .出生日期
        lblEdit(7) = .统筹编号
        lblEdit(8) = .年龄
        lblEdit(9) = .制卡日期
        lblEdit(10) = .卡有效期
        lblEdit(11) = .补卡次数
        lblEdit(12) = .单位号码
        lblEdit(13) = Format(.帐户余额, "####0.00;-#####0.00; ;")
        lblEdit(14) = .制卡单位
        lblEdit(15) = .在职情况
   End With
End Sub

Private Sub cmd病种_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & TYPE_成都内江
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txt病种.Text)
    If rsTemp.State = 0 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt病种.Text = rsTemp("名称")
        txt病种.Tag = rsTemp("ID")
        zlControl.TxtSelAll txt病种
    End If
    txt病种.SetFocus
End Sub

Private Sub txt病种_Change()
    txt病种.Tag = ""
    txt病种.ForeColor = &HC0&
End Sub

Private Sub txt病种_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt病种.Text = "" Or txt病种.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt病种.Text
    
    
    gstrSQL = "Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特殊病','普通病') 类别 " & _
             "   FROM 保险病种 A WHERE A.险类=" & TYPE_成都内江 & " And (" & _
                zlCommFun.GetLike("A", "编码", strText) & " or " & zlCommFun.GetLike("A", "名称", strText) & " or " & zlCommFun.GetLike("A", "简码", strText) & ")"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_成都内江, rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Text = rsTemp("名称")
        txt病种.Tag = rsTemp("ID")
        SendKeys "{TAB}"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub txt病种_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt病种.Text = ""
        txt病种.Tag = ""
    End If
End Sub

Function BZXZ_成都内江(ByVal StrInput As String, Optional strLoad As String = 0) As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmpSQL As String
    
    On Error Resume Next
   
    
    If StrInput = "" And strLoad = 1 Then Exit Function
    
    If StrInput = "" Then
        strTmpSQL = "Select ID,编码,名称 from 保险病种"
    Else
        strTmpSQL = "Select ID,编码,名称 from 保险病种" & _
                 " Where 编码 Like '%" & StrInput & "%' OR " & _
                 "名称 like '%" & StrInput & "%' Or " & _
                 "lower(简码) like lower('%" & StrInput & "%')"
    End If
    
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "病种", True, , , , False, gcnOracle)
    If rsTmp Is Nothing Then Exit Function
    BZXZ_成都内江 = "[" & rsTmp!编码 & "]" & rsTmp!名称
End Function

