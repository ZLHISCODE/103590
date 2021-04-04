VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentify莲合 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify莲合.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   120
      TabIndex        =   14
      Top             =   2985
      Width           =   6030
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   3765
      TabIndex        =   13
      Top             =   3135
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   2310
      TabIndex        =   12
      Top             =   3135
      Width           =   1305
   End
   Begin VB.TextBox txt身份证号 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2100
      MaxLength       =   18
      TabIndex        =   9
      Top             =   1890
      Width           =   2715
   End
   Begin VB.ComboBox cbo性别 
      Height          =   360
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   870
      Width           =   840
   End
   Begin VB.TextBox txt姓名 
      Height          =   360
      Left            =   2100
      TabIndex        =   3
      Top             =   870
      Width           =   1335
   End
   Begin VB.TextBox txtAccount 
      Height          =   360
      Left            =   2100
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtBanlance 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2100
      MaxLength       =   18
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   2715
   End
   Begin MSComCtl2.DTPicker dtpBirthday 
      Height          =   360
      Left            =   2100
      TabIndex        =   7
      Top             =   1380
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   635
      _Version        =   393216
      CustomFormat    =   "yyyy-mm-dd"
      Format          =   87031808
      CurrentDate     =   37243
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   555
      Picture         =   "frmIdentify莲合.frx":030A
      Top             =   390
      Width           =   480
   End
   Begin VB.Label lbl身份证号 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号"
      Height          =   240
      Left            =   1065
      TabIndex        =   8
      Top             =   1950
      Width           =   960
   End
   Begin VB.Label lbl出生日期 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期"
      Height          =   240
      Left            =   1065
      TabIndex        =   6
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label lbl性别 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      Height          =   240
      Left            =   3465
      TabIndex        =   4
      Top             =   930
      Width           =   600
   End
   Begin VB.Label lbl姓名 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   240
      Left            =   1515
      TabIndex        =   2
      Top             =   930
      Width           =   510
   End
   Begin VB.Label lblCard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "帐号"
      Height          =   240
      Left            =   1500
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblbanlance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "个人余额"
      Height          =   240
      Left            =   1020
      TabIndex        =   10
      Top             =   2460
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmIdentify莲合"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'编译常量不能定义成公共的，必须在使用到的地方单独定义，在编译时统一修改
#Const gverControl = 99  ' 0-不支持动态医保(9.19以前),1-支持动态医保无附加参数(9.22以前) , _
'    2-解决了虚拟结算与正式结算结果不一致;结算作废与原始结算结果不一致;门诊收费死锁的问题;3-公共部件增加GetNextNO();
'    99-所有交易增加附加参数(最新版)

Public mlng病人ID As Long
Public strPatiInfo As String

Private strCardMask As String
Private blnShowCard As Boolean
Private bytCardNOLen As Byte

Private rsTmp As New ADODB.Recordset
Private strSQL As String
Private mintHIS收费 As Integer

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    strPatiInfo = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim(txt姓名.Text) = "" Then
        MsgBox "未正确地输入姓名,请重输！", vbInformation, gstrSysName
        txt姓名.SetFocus
        
        Exit Sub
    End If
    If mintHIS收费 = 1 Then
        If Trim(txtAccount.Text) = "" Then
            MsgBox "必须输入医保卡号,请重输！", vbInformation, gstrSysName
            txtAccount.SetFocus
    
            Exit Sub
        End If
        
        If Trim(txtBanlance.Text) <> "" Then
            If Not IsNumeric(txtBanlance.Text) Then
                MsgBox "个人余额必须为数字型!", vbOKOnly, gstrSysName
                txtBanlance.SelStart = 0
                txtBanlance.SelLength = Len(txtBanlance.Text)
                txtBanlance.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "个人余额不能为空!", vbOKOnly + vbExclamation, gstrSysName
            txtBanlance.SelStart = 0
            txtBanlance.SelLength = Len(txtBanlance.Text)
            txtBanlance.SetFocus
            Exit Sub
        End If
    
    End If
    
    If Trim(txt身份证号.Text) <> "" Then
        If Not IsNumeric(txt身份证号.Text) Then
            MsgBox "身份证号必须为数字如1,2,3等", vbOKOnly, gstrSysName
            txt身份证号.SelStart = 0
            txt身份证号.SelLength = Len(txt身份证号.Text)
            txt身份证号.SetFocus
            Exit Sub
        End If
    End If
    
    Call SaveInfo
    Me.Hide
End Sub

Private Sub SaveInfo()
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
        '9中心;10.顺序号;11人员身份;12帐户余额;13当前状态;14病种ID;15在职(0,1);16退休证号;17年龄段;18灰度级
    
    Dim strKH As String
    Dim strSelfNo As String
    Dim strSelfPwd As String
    Dim STRNAME As String
    Dim strSex As String
    Dim strBirth As String
    Dim strSFZ As String
    Dim strDWMC As String
    Dim strdwbm As String
    Dim rsTemp As New ADODB.Recordset
    
    If mintHIS收费 = 1 Then
        strKH = Trim(txtAccount.Text)
        strSelfNo = Trim(txtAccount.Text)
        gcurBanlance = Trim(txtBanlance.Text)
    Else
        strKH = Format(Now, "yyyymmddHHMMSS")
        strSelfNo = Format(Now, "yyyymmddHHMMSS")
        gcurBanlance = 0
    End If
    mlng病人ID = Val(txt姓名.Tag)
    
    If mlng病人ID <> 0 Then
        '如果病人ID不为零，则提取该病人现有的医保号与卡号，避免再次产生医保病人档案与关联数据
        gstrSQL = "Select 卡号,医保号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取医保信息", TYPE_成都莲合, mlng病人ID)
        If rsTemp.RecordCount <> 0 Then
            strKH = rsTemp!卡号
            strSelfNo = rsTemp!医保号
        End If
    End If
    
    strSelfPwd = ""
    STRNAME = Trim(txt姓名.Text)
    strSex = Mid(cbo性别.List(cbo性别.ListIndex), InStr(1, cbo性别.List(cbo性别.ListIndex), "-") + 1)
    strBirth = Format(dtpBirthday.Value, "yyyy-mm-dd")
    strSFZ = Trim(txt身份证号.Text)
    strDWMC = ""
    strdwbm = ""
    
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    strPatiInfo = strKH & ";" & strSelfNo & ";" & strSelfPwd & ";" & _
                    STRNAME & ";" & strSex & ";" & _
                    strBirth & ";" & strSFZ & ";" & _
                    strDWMC & "(" & strdwbm & ")"
End Sub

Private Sub dtpBirthday_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New Recordset
    Dim i As Long
    
    mintHIS收费 = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("HIS收费"), 0)
    If mintHIS收费 = 1 Then
        lblCard.Visible = True
        txtAccount.Visible = True
        lblbanlance.Visible = True
        txtBanlance.Visible = True
    End If
    
    strSQL = "Select 编码,名称,简码,缺省标志 From 性别 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cbo性别.Clear
    If rsTmp.RecordCount <> 0 Then
        For i = 1 To rsTmp.RecordCount
            cbo性别.AddItem rsTmp!名称
            If rsTmp!缺省标志 Then
                cbo性别.ListIndex = i - 1
                cbo性别.ItemData(i - 1) = -1 '作标志
            End If
            rsTmp.MoveNext
        Next
        If cbo性别.ListIndex = -1 Then cbo性别.ListIndex = 0
    End If
    dtpBirthday.Value = Now()
    dtpBirthday.MaxDate = Now()
    gcurBanlance = 0
    Me.txt姓名.Tag = mlng病人ID
    
    '取系统参数
    bytCardNOLen = 7
    Dim strPar As String
    
    #If gverControl >= 4 Then
        blnShowCard = -Not Abs(Val(zlDatabase.GetPara(12, glngSys, , 0)))
        strCardMask = UCase(zlDatabase.GetPara(27, glngSys))
        strPar = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    #Else
        blnShowCard = -Not Abs(Val(GetPara(12, glngSys, , , 0)))
        strCardMask = UCase(GetPara(27, glngSys))
        strPar = GetPara(20, glngSys, , , "7|7|7|7|7")
    #End If
    If InStr(1, strPar, "|") <> 0 Then
        bytCardNOLen = Val(Split(strPar, "|")(4))
    Else
        bytCardNOLen = Mid(strPar, 5, 1)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strPatiInfo = ""
    mlng病人ID = 0
End Sub

Private Sub SeekCob(ByVal ConObj As ComboBox, ByVal strSeek As String)
    Dim intSeek As Integer
    
    ConObj.ListIndex = 0
    If strSeek = "" Then Exit Sub
    
    For intSeek = 0 To ConObj.ListCount
        If ConObj.List(intSeek) = strSeek Then
            ConObj.ListIndex = intSeek
            Exit For
        End If
    Next
End Sub

Private Function GetPatiRec(ByVal strAccount As String) As Recordset
    gstrSQL = "select a.卡号,a.医保号,a.密码,b.姓名,b.性别,b.出生日期,b.身份证号,b.工作单位 " _
        & " from 保险帐户 a,病人信息 b " _
        & " where a.病人id=b.病人id " _
        & " and a.卡号=[1] and a.险类=[2]"
        
        
    'New:0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码)
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strAccount, TYPE_成都莲合)
    Set GetPatiRec = rsTmp
End Function

Private Sub txtAccount_GotFocus()
    With txtAccount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
    Dim rsPati As New Recordset
    
    If KeyAscii = 13 And Trim(txtAccount.Text) <> "" Then
        Set rsPati = GetPatiRec(txtAccount.Text)
        If Not rsPati.EOF Then
            txt姓名.Text = IIf(IsNull(rsPati!姓名), "", rsPati!姓名)
            Call SeekCob(cbo性别, rsPati!性别)
            dtpBirthday.Value = Format(IIf(IsNull(rsPati!出生日期), zlDatabase.Currentdate, rsPati!出生日期), "yyyy-mm-dd")
            txt身份证号.Text = IIf(IsNull(rsPati!身份证号), "", rsPati!身份证号)
       
            txtBanlance.SetFocus
        Else
            txt姓名.SetFocus
            txt姓名.SelStart = 0
            txt姓名.SelLength = Len(txt姓名.Text)
        End If
    End If
End Sub

Private Sub txtBanlance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    End If
End Sub

Private Sub txt身份证号_GotFocus()
    With txt身份证号
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt身份证号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    End If
End Sub

Private Sub txt姓名_GotFocus()
    With txt姓名
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    Dim StrInput As String
    Dim blnCard As Boolean, blnRead As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    blnCard = InputIsCard(txt姓名, KeyAscii)
    If blnCard And Len(txt姓名.Text) = bytCardNOLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txt姓名.Text) <> "" And Mid(txt姓名.Text, 1, 1) <> "*" Then
        '通过就诊卡获取病人身份
        blnRead = True
        StrInput = txt姓名.Text & Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        If Mid(txt姓名.Text, 1, 1) = "*" Then '认为是门诊号或住院号
            blnRead = True
            StrInput = Val(Mid(txt姓名.Text, 2))
        End If
    End If
    If blnRead = False Then Exit Sub
    '如果即不是刷卡，也不是门诊号或住院号，因下面的SQL将其当做门诊号处理，所以将其转换为数字型
    If Not blnCard And Mid(txt姓名.Text, 1, 1) <> "*" Then StrInput = Val(StrInput)
    
    '根据门诊号提取病人基本信息
    gstrSQL = " Select 病人ID,姓名,性别,出生日期,身份证号 From 病人信息 " & _
              " Where " & IIf(blnCard, "就诊卡号=", IIf(Me.Tag = "0", "门诊号=", "住院号=")) & StrInput
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人基本信息")
    If rsTemp.RecordCount <> 0 Then
        Me.txt姓名.Tag = rsTemp!病人ID
        Me.txt姓名.Text = rsTemp!姓名
        Me.dtpBirthday.Value = Format(rsTemp!出生日期, "yyyy-MM-dd")
        Me.txt身份证号.Text = Nvl(rsTemp!身份证号)
        
        Select Case rsTemp!性别
        Case "男"
            Me.cbo性别.ListIndex = 0
        Case "女"
            Me.cbo性别.ListIndex = 1
        Case Else
            Me.cbo性别.ListIndex = 2
        End Select
    End If
    zlCommFun.PressKey (vbKeyTab)
End Sub

Public Function InputIsCard(txtInput As Object, KeyAscii As Integer) As Boolean
'功能：判断指定文本框中当前输入是否在刷卡,根据处理密文显示
    Dim strText As String, blnCard As Boolean
    Dim arrMask As Variant, i As Long

    '当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = UCase(strText & Chr(KeyAscii))
    End If
        
    '判断是否在刷卡
    blnCard = False
    If IsNumeric(strText) And IsNumeric(Left(strText, 1)) Then
        blnCard = True
    ElseIf strCardMask <> "" Then
        arrMask = Split(strCardMask, "|")
        For i = 0 To UBound(arrMask)
            If strText Like arrMask(i) & "*" Then
                If IsNumeric(Mid(strText, Len(arrMask(i)) + 1)) And IsNumeric(Mid(strText, Len(arrMask(i)) + 1, 1)) Then
                    blnCard = True
                End If
            End If
        Next
    End If
    
    '刷卡时卡号是否密文显示
    If blnCard Then
        txtInput.PasswordChar = IIf(blnShowCard, "", "*")
    Else
        txtInput.PasswordChar = ""
    End If
    
    InputIsCard = blnCard
End Function
