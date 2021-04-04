VERSION 5.00
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmIdentify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "病人身份验证"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "frmIdentify.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtMoney 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2325
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1350
      Width           =   3015
   End
   Begin VB.TextBox txtCard 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2325
      TabIndex        =   1
      Top             =   1912
      Width           =   3015
   End
   Begin VB.TextBox txtPass 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2340
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2475
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      Top             =   3450
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3450
      Width           =   1100
   End
   Begin VB.Frame fraDown 
      Height          =   30
      Left            =   -30
      TabIndex        =   9
      Top             =   3225
      Width           =   7290
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   0
      ScaleHeight     =   1050
      ScaleWidth      =   6540
      TabIndex        =   10
      Top             =   0
      Width           =   6540
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   1140
         X2              =   8715
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   1140
         X2              =   8715
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Label lblFamilyRest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家属余额:9999999.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   4260
         TabIndex        =   17
         Tag             =   "家属余额:"
         Top             =   750
         Width           =   2280
      End
      Begin VB.Label lblRest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人余额:9999999.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1140
         TabIndex        =   16
         Tag             =   "病人余额:"
         Top             =   750
         Width           =   2280
      End
      Begin VB.Label lblPatiType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人类型:普通患者"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1140
         TabIndex        =   15
         Tag             =   "病人类型:"
         Top             =   420
         Width           =   2040
      End
      Begin VB.Label lblFeeType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别:普通"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4740
         TabIndex        =   14
         Tag             =   "费别:"
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄:30岁"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4740
         TabIndex        =   13
         Tag             =   "年龄:"
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别:未知"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3330
         TabIndex        =   12
         Tag             =   "性别:"
         Top             =   90
         Width           =   1080
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:琪玛多吉"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   1140
         TabIndex        =   11
         Tag             =   "姓名:"
         Top             =   90
         Width           =   1560
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   240
         Picture         =   "frmIdentify.frx":058A
         Top             =   135
         Width           =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -105
         X2              =   7470
         Y1              =   1035
         Y2              =   1035
      End
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   420
      Left            =   1665
      TabIndex        =   7
      Top             =   1905
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   741
      Appearance      =   2
      IDKindStr       =   "就|就诊卡|0|0|0|0|0|;IC|IC卡号|1|0|0|0|0|"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      FontName        =   "宋体"
      IDKind          =   -1
      ShowPropertySet =   -1  'True
      NotContainFastKey=   ""
      BackColor       =   -2147483633
      SaveRegType     =   4
      ProductName     =   "一卡通消费支付"
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "刷卡金额"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1155
      TabIndex        =   5
      Top             =   1425
      Width           =   1140
   End
   Begin VB.Label lblCardNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1110
      TabIndex        =   6
      Top             =   1980
      Width           =   570
   End
   Begin VB.Label lblPass 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1425
      TabIndex        =   8
      Top             =   2580
      Width           =   870
   End
End
Attribute VB_Name = "frmIdentify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mintCount As Integer
Private mstr病人IDs As String
Private mlngSys As Long
Private mblnPreCard As Boolean
Private mobjCard As Card '当前处理的卡
'--------------------------------------------------
'卡相关:
Private mobjKeyboard As Object
Private mblnPassInputCardNo As Boolean  '是否密文输入卡号
Private mobjSquareCard As Object
Private mlng医疗卡长度 As Long
Private mlngModul As Long
Private mstrPassWord As String
Private mlngDefaultCardTypeID As Long '缺省的刷卡类别ID
Private mblnBrushCard As Boolean
Private Const VK_RETURN = &HD
Private mblnCheckPassWord As Boolean
Private mblnReadIDCard As Boolean  '读取的是身份证
Private mblnReadICCard As Boolean
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard  '问题:47945
Attribute mobjICCard.VB_VarHelpID = -1
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mstrRegSection As String
Private mlngPreBrushCardTypeID As Long '上次刷卡类别
'--------------------------------------------------
Public Function ShowMe(frmParent As Object, ByVal lngSys As Long, ByVal lng病人ID As Long, _
    ByVal cur金额 As Currency, Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0, _
    Optional lngDefaultCardTypeID As Long = 0, _
    Optional blnCheckPassWord As Boolean = True, _
    Optional blnFamilyMoney As Boolean, _
    Optional strFamilyPatiIDs As String = "", _
    Optional bln刷卡验证 As Boolean = True, _
    Optional bln无密码不验卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证窗体入口
    '入参:frmParent-调用的主窗体
    '       lngSys-系统号
    '       lng病人ID-指定的病人ID
    '       lngModul-模块号
    '       bytOperationType-业务类型(0-不区分;1-门诊;2-住院)
    '       mlngDefaultCardTypeID-缺省的刷卡类别ID
    '       blnCheckPassWord-验证密码(true-验证密码,false-只刷卡,不输入密码)
    '       blnFamilyMoney-是否读取家属预交余额
    '       strFamilyPatiIDs-病人家属的病人ID
    '       bln刷卡验证-是否进行刷卡验证，主要用于不刷卡验证时读取家属IDs
    '       bln无密码不验卡-病人的所有医疗卡都没有设置密码时是否验卡，当为True时，只要有一张卡设置了密码都要进行验卡,112418
    '出参:
    '返回:验证成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-10 16:35:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strSQL As String, intMouse As Integer
    mblnCheckPassWord = blnCheckPassWord
    mlngSys = lngSys: mlngModul = lngModul: mlngDefaultCardTypeID = lngDefaultCardTypeID
    mblnOK = False: mintCount = 3: mstr病人IDs = lng病人ID
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    
    '读取就诊卡信息
    On Error GoTo ErrH
    '病人信息及预交余额
    strSQL = "Select 病人id, Nvl(Sum(预交余额), 0) - Nvl(Sum(费用余额), 0) As 余额" & vbNewLine & _
            " From 病人余额" & vbNewLine & _
            " Where 病人id = [1] And 性质 = 1 And Decode([2],0,0,类型)=[2]" & vbNewLine & _
            " Group By 病人id"
    '病人不存在预交余额记录时，取病人ID用于读取病人信息
    strSQL = strSQL & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select To_Number([1]) As 病人ID, 0 As 余额 From Dual"
    If blnFamilyMoney Then
        '病人家属信息及预交余额
        strSQL = strSQL & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select b.病人id, Nvl(Sum(b.预交余额), 0) - Nvl(Sum(b.费用余额), 0) As 余额" & vbNewLine & _
                " From 病人家属 A, 病人余额 B" & vbNewLine & _
                " Where a.家属id = b.病人id And a.病人id = [1] And b.性质 = 1 And Decode([2],0,0,b.类型)=[2] " & vbNewLine & _
                "       And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & vbNewLine & _
                " Group By b.病人id"
        '病人家属不存在预交余额记录时，取病人家属ID用于读取病人信息
        strSQL = strSQL & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select 家属id, 0 As 余额 From 病人家属 Where 病人id = [1] " & vbNewLine & _
                "       And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)"
    End If
    strSQL = "Select a.病人id, a.姓名, a.性别, a.年龄, a.病人类型, a.费别, a.就诊卡号, a.卡验证码, Nvl(b.余额, 0) As 余额" & vbNewLine & _
            " From 病人信息 A, (" & strSQL & ") B" & vbNewLine & _
            " Where a.病人id = b.病人id And a.停用时间 Is Null"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取病人和家属的信息及预交余额", lng病人ID, bytOperationType)

    '1-家属余额及家属信息
    Dim cur家属余额 As Currency
    strFamilyPatiIDs = "": cur家属余额 = 0
    rsTmp.Filter = "病人id<>" & lng病人ID
    Do While Not rsTmp.EOF
        If InStr(strFamilyPatiIDs & ",", "," & gobjComLib.zlCommFun.NVL(rsTmp!病人ID) & ",") = 0 Then
            strFamilyPatiIDs = strFamilyPatiIDs & "," & gobjComLib.zlCommFun.NVL(rsTmp!病人ID)
            cur家属余额 = cur家属余额 + Val(gobjComLib.zlCommFun.NVL(rsTmp!余额))
        End If
        rsTmp.MoveNext
    Loop
    If strFamilyPatiIDs <> "" Then strFamilyPatiIDs = Mid(strFamilyPatiIDs, 2)
    
    '不用刷卡验证直接返回
    If Not bln刷卡验证 Then ShowMe = True: Exit Function
    
    If strFamilyPatiIDs <> "" Then mstr病人IDs = mstr病人IDs & "," & strFamilyPatiIDs
    '2-病人本人信息
    rsTmp.Filter = "病人id=" & lng病人ID
    If rsTmp.EOF Then
        MsgBox "病人信息不存在,请检查!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '检查病人及家属是否有卡，只要其中任何一人有卡都需要刷卡，79868
'    If gobjComLib.zlCommFun.NVL(rsTmp!就诊卡号) = "" Then
        '问题:43449，如果病人没有发卡的,则允许不输入密码及刷卡操作,直接进行扣款
        strSQL = _
        "Select Count(1) As 存在卡, Sum(Decode(密码, Null, 0, 1)) As 存在密码" & vbNewLine & _
        "From 病人医疗卡信息" & vbNewLine & _
        "Where 状态 = 0 And 病人id In (Select /*+cardinality(a,10)*/ Column_Value From Table(f_Num2list([1])) A)"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "检查病人或家属是否发卡", mstr病人IDs)
        If rsTemp.EOF Then
            '无记录,直接返回true,不用验卡
            ShowMe = True: Exit Function
        Else
            If Val(gobjComLib.zlCommFun.NVL(rsTemp!存在卡)) = 0 Then
                '未发卡,直接返回true,不用验卡
                ShowMe = True: Exit Function
            End If
            If Val(gobjComLib.zlCommFun.NVL(rsTemp!存在密码)) = 0 And bln无密码不验卡 Then
                '所有卡都无密码,直接返回true,不用验卡
                ShowMe = True: Exit Function
            End If
        End If
'    End If
    
    If Not rsTmp.EOF Then
        lblName.Caption = lblName.Tag & gobjComLib.zlCommFun.NVL(rsTmp!姓名)
        lblSex.Caption = lblSex.Tag & gobjComLib.zlCommFun.NVL(rsTmp!性别)
        lblAge.Caption = lblAge.Tag & gobjComLib.zlCommFun.NVL(rsTmp!年龄)
        lblPatiType.Caption = lblPatiType.Tag & gobjComLib.zlCommFun.NVL(rsTmp!病人类型)
        lblFeeType.Caption = lblFeeType.Tag & gobjComLib.zlCommFun.NVL(rsTmp!费别)
        
        lblRest.Caption = lblRest.Tag & Format(Val(gobjComLib.zlCommFun.NVL(rsTmp!余额)), "0.00")
    End If
    lblFamilyRest.Caption = lblFamilyRest.Tag & Format(cur家属余额, "0.00")
    
    txtMoney.Text = Format(cur金额, "0.00")
'        txtCard.Tag = .NVL(rsTmp!就诊卡号)
'        txtPass.Tag = .NVL(rsTmp!卡验证码)
    On Error GoTo 0
    Me.Show 1, frmParent
    ShowMe = mblnOK
    
    Screen.MousePointer = intMouse
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡的有效性
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-03-19 17:04:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPassWord As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim blnSucces As Boolean '输入成功
    Dim str名称 As String
    On Error GoTo errHandle
    
    If mobjCard Is Nothing Then Exit Function
    If mobjCard.名称 Like "*卡号" Then
        str名称 = mobjCard.名称
    ElseIf mobjCard.名称 Like "*身份证" Then
        str名称 = "身份证号"
    ElseIf mobjCard.名称 Like "*卡" Then
        str名称 = mobjCard.名称 & "卡号"
    Else
        str名称 = mobjCard.名称 & "卡卡号"
    End If

    If UCase(Trim(txtCard.Text)) = "" Then Exit Function
    If Not InStr("," & mstr病人IDs & ",", "," & Val(lblPass.Tag) & ",") > 0 Or Val(lblPass.Tag) = 0 Then
        MsgBox "当前" & str名称 & "与病人的" & str名称 & "不相符！", vbExclamation, gstrSysName
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Function
    End If
    
    If Not mblnCheckPassWord Then IsValied = True: Exit Function
    strPassWord = gobjComLib.zlCommFun.zlStringEncode(txtPass.Text)
    If strPassWord <> mstrPassWord Then
        If mintCount = 1 Then
            MsgBox "三次密码输入错误,不能再输入！", vbExclamation, gstrSysName
        Else
            MsgBox "密码输入错误！", vbExclamation, gstrSysName
        End If
        txtPass.Text = "": mintCount = mintCount - 1
        If mintCount = 0 Then
            Unload Me '密码错误，可输入2次
        ElseIf txtPass.Enabled Then
            txtPass.SetFocus
        End If
        Exit Function
    End If
    IsValied = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Sub cmdOK_Click()
    If IsValied = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call IDKind.ActiveFastKey
End Sub
Private Sub Form_Load()
    mstrRegSection = "私有模块\" & gstrDBUser & "\界面设置\" & Me.Name & Me.Name
    mlngPreBrushCardTypeID = GetSetting("ZLSOFT", mstrRegSection, "缺省卡类别ID", 0)

    Call CreateObjectKeyboard
    Call zlCardSquareObject
    Call SetCtrlVisible
    Call NewCardObject
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not IDKind.GetCurCard Is Nothing Then
         SaveSetting "ZLSOFT", mstrRegSection, "缺省卡类别ID", IDKind.GetCurCard.接口序号
    End If
    
    Set mobjKeyboard = Nothing
    Set mobjCard = Nothing
    Call zlCardSquareObject(True)
    Call CloseIDCard
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard Is Nothing Then Exit Sub
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        
        If mobjICCard Is Nothing Then Exit Sub
        txtCard.MaxLength = 0
        txtCard.Text = mobjICCard.Read_Card()
        If txtCard.Text = "" Then
            If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
            Exit Sub
        End If
        
            '问题号:42948
        If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            txtCard.Text = "": If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
        End If
        mblnReadICCard = True
        If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
        If txtCard.Text <> "" Then
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Exit Sub
        End If
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus: Exit Sub
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If mobjSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtCard.Text = strOutCardNO
    
    '问题号:42948
    If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
    End If
    
    If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
    End If
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
     If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
     End If
     If mblnCheckPassWord Then txtPass.SetFocus: Exit Sub
     Call cmdOK_Click
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    txtCard.PasswordChar = IIf(objCard.卡号密文规则 <> "", "*", "")
    '85565,李南春,2015/7/10:读卡性质
    mblnBrushCard = objCard.是否刷卡 Or objCard.是否扫描
    If txtCard.Text <> "" Then txtCard.Text = ""
    txtCard.Locked = Not mblnBrushCard
    If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)

    txtCard.Text = objPatiInfor.卡号
    If GetPatient(objCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
    End If
    
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
     If txtCard.Text = "" Then
        If txtCard.Enabled And txtCard.Visible Then txtCard.SetFocus
        Exit Sub
     End If
     If mblnCheckPassWord Then txtPass.SetFocus: Exit Sub
     Call cmdOK_Click
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    'IC卡读取
    
    If strCardNO = "" Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtCard.MaxLength = Len(strCardNO)
    txtCard.Text = strCardNO: mblnReadICCard = True
    If GetPatient(objCard, strCardNO) = False Then
         mblnReadICCard = False: Exit Sub
    End If
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    '显示卡信息
    If strID = "" Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("身份证号", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtCard.Text = strID: mblnReadICCard = True
    txtCard.MaxLength = Len(strID)
    If GetPatient(objCard, strID) = False Then
         mblnReadICCard = False: Exit Sub
    End If
    If Not mblnCheckPassWord Then cmdOK_Click: Exit Sub
    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
End Sub
Private Sub txtCard_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
    IDKind.SetAutoReadCard (False)
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub
Private Sub txtCard_Change()
    lblPass.Tag = "": txtCard.Tag = ""
    txtPass.Enabled = txtCard.Text <> ""
    If Not txtPass.Enabled Then txtPass.Text = ""
    mblnReadIDCard = False
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtCard.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtCard.Text = "")
    IDKind.SetAutoReadCard (txtCard.Text = "")
End Sub

Private Sub txtCard_GotFocus()
    Call gobjComLib.zlControl.TxtSelAll(txtCard)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtCard.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtCard.Text = "")
End Sub

Private Sub txtCard_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    mblnPreCard = False

    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(txtCard.Text) = IDKind.GetCurCard.卡号长度 - 1 And txtCard.SelLength <> Len(txtCard.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txtCard.Text = txtCard.Text & Chr(KeyAscii)
            txtCard.SelStart = Len(txtCard.Text)
        End If
        KeyAscii = 0
        If GetPatient(IDKind.GetCurCard, Trim(txtCard.Text)) = False Then
            If txtCard.Enabled Then txtCard.SetFocus
            gobjComLib.zlControl.TxtSelAll txtCard
            Exit Sub
        End If
        mblnPreCard = blnCard
        If mblnCheckPassWord Then
            If txtPass.Enabled Then txtPass.SetFocus
        Else
            Call cmdOK_Click: Exit Sub
        End If
    Else
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If

        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 And IDKind.GetCurCard.是否持卡消费 = True Then
            sngNow = timer
            If txtCard.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txtCard.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txtCard.Text = Chr(KeyAscii)
                txtCard.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub
Private Sub txtCard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtCard.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtCard.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtCard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtCard.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
Private Sub txtPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    glngTXTProc = GetWindowLong(txtPass.hWnd, GWL_WNDPROC)
    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, AddressOf WndMessage)
End Sub

Private Sub txtPass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button <> 2 Then Exit Sub
    Call SetWindowLong(txtPass.hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtPass_GotFocus()
    If txtCard.Text <> "" And mstrPassWord = "" Then Call cmdOK_Click: Exit Sub
    Call gobjComLib.zlControl.TxtSelAll(txtPass)
    OpenPassKeyboard txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If mblnPreCard Then
            '60580
            mblnPreCard = False
             If (GetAsyncKeyState(VK_RETURN) And &H1) <> 0 Then
                txtPass.Text = ""
                Exit Sub
             End If
        End If
        mblnPreCard = False
        Call cmdOK_Click
    ElseIf KeyAscii = 22 Then
        KeyAscii = 0 '不允许粘贴
    Else
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        End If
    End If
    '60580
    mblnPreCard = False
End Sub

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function OpenPassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function
 
 
Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String, i As Integer, intIdKind As Integer
    '只有:执行或退费时,才可能管结算卡的
    If blnClosed Then
       If Not mobjSquareCard Is Nothing Then
            Call mobjSquareCard.CloseWindows
            Set mobjSquareCard = Nothing
        End If
        Exit Sub
    End If
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitCompoent (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Call mobjSquareCard.zlInitComponents(Me, mlngModul, mlngSys, gstrDBUser, gcnOracle, False, strExpend)
    mobjSquareCard.mblnYLMgr = True
    Err = 0: On Error GoTo 0
    Call IDKind.zlInit(Me, mlngSys, mlngModul, gcnOracle, gstrDBUser, mobjSquareCard, "", txtCard)
    
    Err = 0: On Error Resume Next
     If mlngPreBrushCardTypeID <> 0 Then
        intIdKind = IDKind.GetKindIndex(mlngPreBrushCardTypeID)
        If intIdKind <> 0 Then
            IDKind.IDKind = intIdKind
        End If
     End If
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    Optional blnIDCard As Boolean = False, Optional blnICCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取病人信息
    '入参:objCard-按指定的卡类别进行读卡
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-26 00:20:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo ErrH
    
    mstrPassWord = ""
    Set mobjCard = Nothing
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Function
    '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
    If mobjSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID, Nothing, Me, False, True) = False Then
        '进行模糊查找:-1:医疗卡类别(但是如果当前的卡号长度不够的话,会存在问题)
        If mobjSquareCard.zlGetPatiID(-1, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID, Nothing, Me, False, True) = False Then
            GoTo NotFoundPati:
        End If
    End If
    If lng病人ID <= 0 Then GoTo NotFoundPati:
    If Not InStr("," & mstr病人IDs & ",", "," & lng病人ID & ",") > 0 Then
        If objCard.名称 Like "*卡号" Then
            MsgBox "当前" & objCard.名称 & "与病人所持有的" & objCard.名称 & "不相符,请检查！", vbExclamation, gstrSysName
        ElseIf objCard.名称 Like "*身份证" Then
            MsgBox "当前身份证号与病人所持有的身份证号不相符,请检查！", vbExclamation, gstrSysName
        ElseIf objCard.名称 Like "*卡" Then
            MsgBox "当前" & objCard.名称 & "卡号与病人所持有的" & objCard.名称 & "卡号不相符,请检查！", vbExclamation, gstrSysName
        Else
            MsgBox "当前" & objCard.名称 & "卡卡号与病人所持有的" & objCard.名称 & "卡卡号不相符,请检查！", vbExclamation, gstrSysName
        End If
        txtCard.Text = ""
        Exit Function '卡号不匹配，不准重试
    End If
    txtCard.Tag = strInput
    lblPass.Tag = lng病人ID
    mstrPassWord = strPassWord
    Set mobjCard = objCard
    GetPatient = True
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Exit Function
NotFoundPati:
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
    Else
        MsgBox "未找到当前卡的持有病人,请检查!", vbOKOnly + vbInformation, gstrSysName
        txtCard.Text = ""
    End If
    txtCard.Tag = "": lblPass.Tag = ""
End Function

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的visible属性
    '编制:刘兴洪
    '日期:2012-03-13 11:28:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    lblFamilyRest.Visible = InStr(mstr病人IDs, ",") > 0 '没有家属则隐藏家属余额的显示，79868
    lblPass.Visible = mblnCheckPassWord
    txtPass.Visible = mblnCheckPassWord
    If mblnCheckPassWord Then Exit Sub
    With txtCard
        .Top = picTop.Top + picTop.Height + (fraDown.Top - (picTop.Top + picTop.Height) - .Height) \ 2
        IDKind.Top = .Top
        lblCardNO.Top = .Top + (.Height - lblCardNO.Height) \ 2
    End With
    If Err <> 0 Then Err.Clear
End Sub

Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭自助读卡功能
    '编制:刘兴洪
    '日期:2012-03-09 16:26:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled (False)
        Set mobjICCard = Nothing
    End If
End Sub
Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化新的卡对象
    '编制:刘兴洪
    '日期:2012-03-09 16:28:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
    End If
    If mobjICCard Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Err = 0: On Error GoTo 0
    End If
End Sub



