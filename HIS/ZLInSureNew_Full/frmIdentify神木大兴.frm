VERSION 5.00
Begin VB.Form frmIdentify神木大兴 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd验卡 
      Caption         =   "重新获取(&R)"
      Height          =   350
      Left            =   105
      TabIndex        =   11
      Top             =   3105
      Width           =   1305
   End
   Begin VB.Frame fra 
      Height          =   2745
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   210
      Width           =   5715
      Begin VB.Frame Frame1 
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   345
         TabIndex        =   4
         Top             =   1620
         Width           =   5115
         Begin VB.OptionButton Opt性别 
            Caption         =   "男"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   840
            TabIndex        =   5
            Top             =   270
            Width           =   885
         End
         Begin VB.OptionButton Opt性别 
            Caption         =   "女"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1890
            TabIndex        =   6
            Top             =   270
            Width           =   885
         End
         Begin VB.OptionButton Opt性别 
            Caption         =   "未知"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   2790
            TabIndex        =   7
            Top             =   270
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   1545
         MaxLength       =   20
         TabIndex        =   3
         Top             =   967
         Width           =   3945
      End
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   1545
         MaxLength       =   16
         TabIndex        =   1
         Top             =   420
         Width           =   3945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "病人姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   1005
         Width           =   1260
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "医保卡号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   180
         TabIndex        =   0
         Top             =   495
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4605
      TabIndex        =   9
      Top             =   3105
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3450
      TabIndex        =   8
      Top             =   3105
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify神木大兴"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐,99-修改指定病人的卡号,88-按住院信息进行查询

Private mlng病人ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '第一次起动系统时调用
Private mblnChange As Boolean
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
   
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long


Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000
Private Sub cmd验卡_Click()
    Call Read病人信息
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    cmd确定.Enabled = False
    If mbytType = 99 Then
        If 获取参保人员信息 = False Then Unload Me: Exit Sub
    Else
        Call Read病人信息
    End If
    Call SetCtlEn
End Sub

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
    Dim StrInput As String, strOutput As String
    Dim lng状态 As Long
    
    IsValid = False
    If Trim(g病人身份_神木大兴.IC卡号) = "" Then
        MsgBox "未输入医保卡号！", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        Exit Function
    End If
    
    If zlCommFun.ActualLen(g病人身份_神木大兴.IC卡号) > 16 Then
        MsgBox "医保卡号超长,最多能输入16个字符或8个汉字！", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        Exit Function
    End If
    
    If Trim(g病人身份_神木大兴.姓名) = "" Then
        MsgBox "未输入病人姓名！", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    
    If zlCommFun.ActualLen(g病人身份_神木大兴.姓名) > 20 Then
        MsgBox "姓名超长,最多能输入20个字符或10个汉字！！", vbInformation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If mbytType = 99 Then
        IsValid = True
        Exit Function
    End If
      
    If mbytType <> 2 And mbytType <> 88 Then
        If mbytType = 4 Then
            '不检查当前状态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_陕西大兴, g病人身份_神木大兴.IC卡号)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
         '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_陕西大兴 & " and  医保号='" & g病人身份_神木大兴.IC卡号 & "'"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
        Else
            mlng病人ID = 0
        End If
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
    Dim strTmp As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim int当前状态 As Integer
    Dim i As Byte
    
    With g病人身份_神木大兴
        .IC卡号 = Trim(txtEdit(0).Text)
        .姓名 = Trim(txtEdit(1).Text)
        strTmp = ""
        For i = 0 To 2
            If Opt性别(i).Value = True Then
                strTmp = Decode(i, 0, "男", 1, "女", "未知")
                Exit For
            End If
        Next
        .性别 = strTmp
    End With
    
    
    If IsValid = False Then Exit Sub
    
    If mbytType = 99 Then       '更新卡号
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_陕西大兴 & ",'卡号','''" & g病人身份_神木大兴.IC卡号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "卡号")
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_陕西大兴 & ",'医保号','''" & g病人身份_神木大兴.IC卡号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "医保号")
        
        If gcnOracle_神木大兴 Is Nothing Then
           Open中间库_神木大兴
        ElseIf gcnOracle_神木大兴.State <> 1 Then
           Open中间库_神木大兴
        End If
        gstrSQL = "ZL_住院_UPDATE("
        gstrSQL = gstrSQL & "'" & txtEdit(0).Tag & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_神木大兴.IC卡号 & "')"
        
        ExecuteProcedure_神木大兴 "改变卡号"
        mstrReturn = g病人身份_神木大兴.IC卡号
        Unload Me
        Exit Sub
    End If
    int当前状态 = 0
    
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_陕西大兴 & " and  医保号='" & g病人身份_神木大兴.IC卡号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_神木大兴
        strIdentify = .IC卡号                        '0卡号
        strIdentify = strIdentify & ";" & .IC卡号    '1医保号
        strIdentify = strIdentify & ";"              '2密码
        strIdentify = strIdentify & ";" & .姓名      '3姓名
        strIdentify = strIdentify & ";" & .性别      '4性别
        strIdentify = strIdentify & ";"              '5出生日期
        strIdentify = strIdentify & ";"              '6身份证
        strIdentify = strIdentify & ";"              '7.单位名称(编码)
        strAddition = ";0"                           '8.中心代码
        strAddition = strAddition & ";"              '9.顺序号
        strAddition = strAddition & ";"              '10人员身份
        strAddition = strAddition & ";"              '11帐户余额

        strAddition = strAddition & ";"              '12当前状态
        strAddition = strAddition & ";"              '13病种ID
        strAddition = strAddition & ";1"             '14在职(1,2,3)
        strAddition = strAddition & ";"              '15退休证号
        strAddition = strAddition & ";"              '16年龄段
        strAddition = strAddition & ";"              '17灰度级
        strAddition = strAddition & ";"              '18帐户增加累计
        strAddition = strAddition & ";0"             '19帐户支出累计
        strAddition = strAddition & ";0"             '20上年工资总额
        strAddition = strAddition & ";"              '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_陕西大兴)
    
    If mlng病人ID = 0 Then
        ShowMsgbox "病人信息有误!"
        Exit Sub
    End If
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    g病人身份_神木大兴.病人ID = mlng病人ID
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Function 获取参保人员信息() As Boolean
    '获取参保人员信息
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    
    获取参保人员信息 = False
    Err = 0:    On Error GoTo errHand:
    If mbytType = 99 Then
        gstrSQL = "Select a.*,b.卡号,b.医保号 From 病人信息 a,保险帐户 b where a.病人id=b.病人id and a.病人id =" & mlng病人ID
    Else
        gstrSQL = "Select * From 病人信息 where 病人id in (Select 病人id From 保险帐户 where  医保号='" & txtEdit(0).Text & "')"
    End If
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    If rsTemp.EOF Then
        'txtEdit(1).Text = ""
        If mbytType = 99 Then
            ShowMsgbox "不存指定的医保病人"
        End If
        Exit Function
    End If
    If mbytType = 99 Then
        txtEdit(0).Text = Nvl(rsTemp!医保号)
        txtEdit(0).Tag = Nvl(rsTemp!医保号)
        
    End If
    txtEdit(1).Text = Nvl(rsTemp!姓名)
    Opt性别(Decode(Nvl(rsTemp!性别), "男", 0, "女", 1, 2)).Value = True
    
    获取参保人员信息 = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Sub ClearData()
    Dim i As Long
    '清除相关信息
    With g病人身份_神木大兴
        .IC卡号 = ""
        .姓名 = ""
        .性别 = ""
    End With
End Sub

Private Sub Opt性别_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If txtEdit(0).Text = "" Then
        SetOKCtrl False
    Else
        SetOKCtrl True
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 0 And mbytType <> 99 Then
            Call 获取参保人员信息
        End If
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtEdit(Index), KeyAscii, m文本式)
End Sub
Private Sub SetCtlEn()
    If mbytType = 99 Then
        txtEdit(1).BackColor = &H8000000F
        txtEdit(1).Enabled = False
        Frame1(1).Enabled = False
        Me.Caption = "修改医保卡号"
    End If
End Sub

Private Sub Read病人信息()
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '--功  能:通过配置文件读取病人信息
    '--入参数:
    '--出参数:
    '--返  回:字串
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim objFile As New FileSystemObject
    Dim objText As TextStream
    
    Dim strText As String
    Dim strFile As String
    Dim STRNAME As String
    Dim str医保证号 As String
    Dim int标志 As Integer
    Dim str性别 As String
    Dim STR姓名 As String
    

        
    strFile = Replace(UCase(Replace(InitInfor_神木大兴.病人目录, "\\", "\")), UCase("ReadYbInfo.INI"), UCase("ReadYbInfo.exe"))
    
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "该目录“" & strFile & "”文件不存在，请在参数中重新设置!"
        Exit Sub
    End If
    Call ExecuteReadCardExe(strFile)
    
    strFile = Replace(InitInfor_神木大兴.病人目录, "\\", "\")
    If Not Dir(strFile) <> "" Then
        ShowMsgbox "该目录“" & strFile & "”文件不存在，请在参数中重新设置!"
        Exit Sub
    End If
    
    Dim strArr
    
    Err = 0
    On Error GoTo errHand:
    If Dir(strFile) <> "" Then
            Set objText = objFile.OpenTextFile(strFile)
            Do While Not objText.AtEndOfStream
                strText = Trim(objText.ReadLine)
                If UCase("[ReadYlbxIcInfo]") <> UCase(strText) Then
                    If InStr(1, strText, "=") <> 0 Then
                        strArr = Split(strText, "=")
                        Select Case UCase(strArr(0))
                        Case UCase("pcode")
                            str医保证号 = Trim(strArr(1))
                        Case UCase("ycbz")
                            int标志 = Val(strArr(1))
                        Case UCase("xb")
                            str性别 = Trim(strArr(1))
                        Case UCase("xm")
                            STR姓名 = Trim(strArr(1))
                        End Select
                    End If
                End If
            Loop
            objText.Close
    End If
    If int标志 <> 0 Then
        ShowMsgbox "该医保病人无效,请检查!"
        Exit Sub
    End If
    txtEdit(0) = str医保证号
    txtEdit(1) = STR姓名
    Select Case str性别
    Case "男", "1"
        Opt性别(0).Value = True
    Case "女", "2"
        Opt性别(1).Value = True
    Case Else
        Opt性别(2).Value = True
    End Select
    
    Exit Sub
errHand:
    DebugTool Err.Description
    Exit Sub
End Sub

Private Function ExecuteReadCardExe(ByVal strFile As String) As Boolean
    '执行Exe文件
    Dim lngTask As Long, lngRet As Long, lngpHandle As Long
    lngTask = Shell(strFile, vbHide)
    lngpHandle = OpenProcess(SYNCHRONIZE, False, lngTask)
    lngRet = WaitForSingleObject(lngpHandle, INFINITE)
    lngRet = CloseHandle(lngpHandle)
    ExecuteReadCardExe = True
End Function

