VERSION 5.00
Begin VB.Form frmIdentify壁山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmIdentify壁山.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1170
      Width           =   3765
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   15
      TabIndex        =   4
      Top             =   1605
      Width           =   6150
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   4530
      TabIndex        =   3
      Top             =   1890
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   2850
      TabIndex        =   2
      Top             =   1890
      Width           =   1305
   End
   Begin VB.TextBox txtPwd 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   645
      Width           =   3765
   End
   Begin VB.Label Label1 
      Caption         =   "病种"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1485
      TabIndex        =   8
      Top             =   1185
      Width           =   660
   End
   Begin VB.Label lblPatiInfo 
      AutoSize        =   -1  'True
      Caption         =   "病人信息"
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   390
      TabIndex        =   7
      Top             =   945
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   450
      Picture         =   "frmIdentify壁山.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1485
      TabIndex        =   6
      Top             =   705
      Width           =   510
   End
   Begin VB.Label lblNote 
      Caption         =   "请在插入IC卡之后，输入个人密码。"
      Height          =   255
      Left            =   1095
      TabIndex        =   5
      Top             =   180
      Width           =   3645
   End
End
Attribute VB_Name = "frmIdentify壁山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'处理读卡器的普通函数
Private Declare Function IC_InitComm Lib "DCIC32.DLL" (ByVal Port%) As Long
Private Declare Function IC_ExitComm% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_Down% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_Pushout% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_InitType% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal TypeNo%)
Private Declare Function IC_Status% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_Erase% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%)
Private Declare Function IC_Read% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%, ByVal Databuffer$)
Private Declare Function IC_Read_Hex% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%, ByVal Databuffer$)
Private Declare Function IC_Read_Float% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, fdata As Single)
Private Declare Function IC_Read_Int% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, fdata As Long)
Private Declare Function IC_Write% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal Length%, ByVal Databuffer$)
Private Declare Function IC_Write_Hex% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal Length%, ByVal Databuffer$)
Private Declare Function IC_Write_Float% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal fdata As Single)
Private Declare Function IC_Write_Int% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal fdata As Long)
'专门处理4428卡的函数
Private Declare Function IC_ReadWithProtection% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%, ByVal ProtBuffer$)
Private Declare Function IC_WriteWithProtection% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal offset%, ByVal l%)
Private Declare Function IC_ReadCount_SLE4428% Lib "DCIC32.DLL" (ByVal icdev As Long)
Private Declare Function IC_CheckPass_SLE4428% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal Password$)
Private Declare Function IC_ChangePass_SLE4428% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal Password$)
Private Declare Function IC_CheckPass_4428hex% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal Password$)
Private Declare Function IC_ChangePass_4428hex% Lib "DCIC32.DLL" (ByVal icdev As Long, ByVal Password$)
 


Public mstr验证码 As String
Public mstrPatiInfo As String
Public mcur余额 As Currency
Public mstr特殊病
Private mcur卡内余额 As Currency        '适用于黔江医保

Private mstr医保号 As String
Private mstr卡号 As String
Private mstrRead As String * 25       '由于这个变量与演示程序不一致，可能需要进行修正
Private blnLoad As Boolean       '是否能够成功进行操作
Private mintst As Long
Private mlngIcdev As Long '当前正在进行通讯的串口设备
Private mstr余额 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'这个逻辑可能存在问题，需要及时进行调整
    Dim strPassWord As String * 20
    Dim intst As Integer, strSQL As String
    Dim strTime As String, rs壁山 As New ADODB.Recordset
    Dim strTmp As String, lngErrLine As Long
    
    On Error GoTo errHandle
    intst = IC_Status(mlngIcdev): lngErrLine = 1 '获取读卡器的状态
    If intst < 0 Then
        MsgBox "读卡器初始化失败,请检查串口", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    If intst = 0 Then
        lblPatiInfo.Caption = "正在读卡，请稍候...."
    End If
    If intst = 1 Then
        MsgBox "检测到读卡器之中缺卡，请检查", vbInformation, gstrSysName
        Exit Sub
    End If
    '对卡进行相关的操作
    intst = IC_InitType(mlngIcdev, 4): lngErrLine = 2
    If intst <> 0 Then
        MsgBox "IC卡的初始化失败,请检查", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    DoEvents
    
    intst = IC_ReadCount_SLE4428(ByVal mlngIcdev&): lngErrLine = 3
    If intst < 0 Then
        MsgBox "在读卡的过程之中发生错误", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    DoEvents
    '对卡的密码进行校验
    '提交时要改成B518
    strPassWord = mstr验证码: lngErrLine = 4 ' "B518" '这个密码属于在进行发卡的时候确定，所以需要修改
    intst = IC_CheckPass_4428hex(ByVal mlngIcdev&, ByVal strPassWord$): lngErrLine = 5
    If intst < 0 Then
        MsgBox "卡验证码发生错误，请到中心换卡", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    DoEvents
    '读出当前使用的IC卡的卡号
    mstrRead = String(25, " "): lngErrLine = 6
    intst = IC_Read(ByVal mlngIcdev, 80, 25, ByVal mstrRead$): lngErrLine = 7
    If intst <> 0 Then
        MsgBox "读卡器读卡失败，请重试[IC卡号]", vbInformation, gstrSysName
        mstrPatiInfo = ""
        Exit Sub
    End If
    DoEvents
    'Modified By 朱玉宝 下午 06:06:46
    '读取IC卡内余额
    If Val(Get保险参数_壁山("适用地区")) = 1 Then   '黔江地区使用的话，需从卡内读取帐户余额
        mstr余额 = String(6, "0"): lngErrLine = 8
        intst = IC_Read(ByVal mlngIcdev, 105, 6, ByVal mstr余额$): lngErrLine = 9
        If intst <> 0 Then
            MsgBox "读卡器读卡失败，请重试[IC卡余额]", vbInformation, gstrSysName
            mstrPatiInfo = ""
            Exit Sub
        End If
        If IsNumeric(mstr余额) Then
            mcur卡内余额 = Val(mstr余额) / 100: lngErrLine = 10       '按分保存，需要转换
        Else
            mcur卡内余额 = 0
        End If
    End If
    DoEvents
    mintst = IC_Down(ByVal mlngIcdev): lngErrLine = 11 '将读卡器下电
    If mintst < 0 Then
        lblPatiInfo.Caption = "注意：IC卡下电失败"
    End If
    
    
    '在数据库之中获取持卡病人的验证信息
    strTime = CStr(Format(zlDatabase.Currentdate, "yyyymmddhhmmss")) & "00": lngErrLine = 12
'    mstrRead = "1234510226200304250856132"
    strSQL = "insert into Check_doex_interface(Bill_no,App_code," & _
            "Ic_id) values('" & strTime & "','" & Mid(gstr医院编码, 1, 4) & _
            "','" & txtPwd.Text & mstrRead & "')": lngErrLine = 13
    gcn壁山.Execute strSQL: lngErrLine = 14
    '对身份验证进行请求
    
    strSQL = "insert into Check_bill_request(Bill_no,App_code," & _
            "Request_status) values('" & strTime & "','" & _
            Mid(gstr医院编码, 1, 4) & "','4')": lngErrLine = 15
    gcn壁山.Execute strSQL: lngErrLine = 16
    
    On Error Resume Next
    If Checkrequest(strTime) = False Then Exit Sub
    DoEvents
    If Requestinfo(strTime) <> "" Then
        DoEvents
        Me.Hide
    End If
    
    On Error GoTo errHandle
    '删除相关的请求
    
    strSQL = "delete from Check_bill_request where Bill_no = '" & _
            strTime & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": lngErrLine = 17
    gcn壁山.Execute strSQL: lngErrLine = 18
    strSQL = "delete from Check_doex_interface where Bill_no = '" & _
             strTime & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": lngErrLine = 19
    gcn壁山.Execute strSQL: lngErrLine = 20
    
    'Modified by ZYB 20040921
    '---------------------------------------------------------------------
    '发请求获取个人帐户余额
    If Val(Get保险参数_壁山("适用地区")) <> 1 Then
        strSQL = "insert into Check_doex_interface(Bill_no,App_code," & _
                "Ic_id) values('" & strTime & "','" & Mid(gstr医院编码, 1, 4) & _
                "','" & txtPwd.Text & mstrRead & "')": lngErrLine = 21
        gcn壁山.Execute strSQL: lngErrLine = 22
        strSQL = "insert into Check_bill_request(Bill_no,App_code," & _
                "Request_status) values('" & strTime & "','" & _
                Mid(gstr医院编码, 1, 4) & "','2')": lngErrLine = 23
        gcn壁山.Execute strSQL: lngErrLine = 24
        
        On Error Resume Next
        If Checkrequest(strTime) = False Then Exit Sub
        DoEvents
        strSQL = "select Ps_Bala " & _
                " from Check_Doex_Interface where Bill_no = '" & strTime & "'" & _
                " and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": lngErrLine = 1
        If rs壁山.State = adStateOpen Then rs壁山.Close: lngErrLine = 25
        rs壁山.Open strSQL, gcn壁山
        If rs壁山.RecordCount <> 0 Then mcur余额 = Nvl(rs壁山!Ps_Bala, 0)
        
        On Error GoTo errHandle
        '删除相关的请求
        strSQL = "delete from Check_bill_request where Bill_no = '" & _
                strTime & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": lngErrLine = 26
        gcn壁山.Execute strSQL: lngErrLine = 27
        strSQL = "delete from Check_doex_interface where Bill_no = '" & _
                 strTime & "' and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": lngErrLine = 28
        gcn壁山.Execute strSQL: lngErrLine = 29
    End If
    '---------------------------------------------------------------------
    
    
    If Val(Get保险参数_壁山("适用地区")) = 2 Or Val(Get保险参数_壁山("适用地区")) = 1 Then
        mstr特殊病 = Combo1.Text: lngErrLine = 21
        If Combo1.ListIndex = 0 Then
            gbln特殊门诊 = False: lngErrLine = 22
        Else
            gbln特殊门诊 = True: lngErrLine = 23
        End If
    Else
        mstr特殊病 = "": lngErrLine = 24
        gbln特殊门诊 = False: lngErrLine = 25
    End If
    
    '检查是否在院
    Dim rsTmp As New ADODB.Recordset
    gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆壁山, mstr医保号)
    If rsTmp.RecordCount > 0 Then
        If rsTmp("状态") > 0 Then
            MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    Exit Sub
errHandle:
    MsgBox "在[身份验证]窗体，[cmdOK_Click]事件中，第" & lngErrLine & "行发生错误", vbExclamation, gstrSysName
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Activate()
    If blnLoad = False Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rs壁山 As New ADODB.Recordset
    On Error GoTo errHandle
    
    'Modified by ZYB 20040921
    '---------------------------------------------------------------------
    '读取特殊病种
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
        strSQL = "Select esp_ID||'--'||esp_name 病种 From check_esp_interface Order by esp_id"
        If rs壁山.State = 1 Then rs壁山.Close
        rs壁山.CursorLocation = adUseClient
        rs壁山.Open strSQL, gcn壁山
    End If
    
    With Combo1
        .AddItem "普通病"
        If Val(Get保险参数_壁山("适用地区")) = 2 Then
            Do While Not rs壁山.EOF
                .AddItem rs壁山!病种
                rs壁山.MoveNext
            Loop
        Else
            .AddItem "特殊病"
        End If
'        .AddItem "01--癌症病人晚期的放疗、化疗、镇痛治疗"
'        .AddItem "02--肾功能衰竭病人透析治疗"
'        .AddItem "03--器官移植后的抗排异治疗"
'        .AddItem "04--急诊观察病人（3日内）的抢救治疗"
'        .AddItem "05--80岁以上老人的治疗型家庭病床（180天内）"
'        .AddItem "06--糖尿病、红斑狼疮"
'        .AddItem "07--慢性高血压、冠心病、风心病、脑卒中后遗症"
'        .AddItem "08--老年性慢性支气管哮喘、肺气肿、肺心病"
'        .AddItem "09--超声乳化白内障摘除术"
'        .AddItem "10--慢性肝硬化"
'        .AddItem "11--慢性再生障碍性贫血"
'        .AddItem "12--精神病"
'        .AddItem "13--结核病"
    End With
    '---------------------------------------------------------------------
    
    Combo1.ListIndex = 0
    If Val(Get保险参数_壁山("适用地区")) = 2 Then
        Combo1.Visible = True
    Else
        Combo1.Visible = False
    End If
    '对串口进行初始化
    mstrPatiInfo = ""
'    mintst = IC_ExitComm(mlngIcdev)  '关闭串口
    mlngIcdev = IC_InitComm(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "当前使用的串口", "0"))    '初始化串口 COM1
    If mlngIcdev <= 0 Then
        blnLoad = False
        MsgBox "串口初始化失败,请检查串口", vbInformation, gstrSysName
        Exit Sub
    End If
    mintst = IC_Status(mlngIcdev) '获取读卡器状态
    If mintst < 0 Then
        blnLoad = False
        MsgBox "串口初始化成功，但是读卡器初始化失败", vbInformation, gstrSysName
    Else
        blnLoad = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mintst = IC_ExitComm(mlngIcdev)  'Close COM
    blnLoad = False
End Sub

Private Function Requestinfo(Billno As String) As String
'功能：向数据库进行查询基础数据，从而得到需要的信息
    Dim strSQL As String, rs壁山 As New ADODB.Recordset, lngErrLine As Long
    
    On Error GoTo errHandle
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    strSQL = "select Ps_Code,Ps_Name,Ps_Sex,Ps_Bdate,Ps_Sfzh,Ep_id,Ps_Bala " & _
            " from Check_Doex_Interface where Bill_no = '" & Billno & "'" & _
            " and App_code = '" & Mid(gstr医院编码, 1, 4) & "'": lngErrLine = 1
    If rs壁山.State = adStateOpen Then rs壁山.Close: lngErrLine = 2
    mstrPatiInfo = "": lngErrLine = 3
    rs壁山.Open strSQL, gcn壁山, adOpenStatic, adLockReadOnly: lngErrLine = 4
    '构造当前需要进行使用的数据
    If Not rs壁山.BOF Then
        mstr医保号 = Nvl(rs壁山("Ps_Code"), "")
        mstrPatiInfo = mstrRead & ";" & Nvl(rs壁山("Ps_Code"), "") & ";" & _
                        txtPwd.Text & ";" & Nvl(rs壁山("Ps_Name"), "") & _
                        ";" & Nvl(rs壁山("Ps_Sex"), "") & ";" & _
                        CStr(Nvl(rs壁山("Ps_Bdate"), "")) & ";" & _
                        Nvl(rs壁山("Ps_Sfzh"), "") & ";" & Nvl(rs壁山("Ep_id"), ""): lngErrLine = 5
        mcur余额 = IIf(IsNull(rs壁山("Ps_Bala")), 0, rs壁山("Ps_Bala")): lngErrLine = 6
        
        'Modified By 朱玉宝 下午 06:07:01
        If Val(Get保险参数_壁山("适用地区")) = 1 Then   '黔江地区使用的话，需从卡内读取帐户余额
            mcur余额 = mcur卡内余额: lngErrLine = 7
        End If
    End If
    Requestinfo = mstrPatiInfo
    Exit Function
errHandle:
    MsgBox "在[身份验证]窗体，[RequestInfo]事件中第" & lngErrLine & "行发生错误", vbExclamation, "错误"
    If ErrCenter() = 1 Then Resume
End Function

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Combo1.Visible = True Then
            Combo1.SetFocus
        Else
            cmdOK_Click
        End If
    End If
End Sub
