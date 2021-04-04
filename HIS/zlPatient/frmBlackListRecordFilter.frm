VERSION 5.00
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#3.0#0"; "zlIDKind.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBlackListRecordFilter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "过滤条件"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6540
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraSplit 
      Height          =   90
      Left            =   -45
      TabIndex        =   23
      Top             =   3210
      Width           =   7455
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Index           =   1
      Left            =   1860
      TabIndex        =   4
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Index           =   2
      Left            =   1860
      TabIndex        =   7
      Top             =   1245
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin VB.ComboBox cbo加入原因 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   960
      TabIndex        =   14
      Text            =   "cbo加入原因"
      Top             =   2085
      Width           =   5400
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "按撤消时间查询"
      Height          =   180
      Index           =   2
      Left            =   195
      TabIndex        =   6
      Top             =   1335
      Width           =   1695
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   90
      TabIndex        =   19
      Top             =   3435
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5265
      TabIndex        =   18
      Top             =   3405
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4050
      TabIndex        =   17
      Top             =   3405
      Width           =   1100
   End
   Begin VB.TextBox txt撤消人 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   4245
      MaxLength       =   18
      TabIndex        =   12
      Top             =   1710
      Width           =   2100
   End
   Begin VB.TextBox txt登记人 
      Height          =   300
      IMEMode         =   1  'ON
      Left            =   960
      MaxLength       =   64
      TabIndex        =   10
      Top             =   1710
      Width           =   1830
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Index           =   0
      Left            =   4245
      TabIndex        =   2
      Top             =   480
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "按发生时间查询"
      Height          =   180
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   540
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Index           =   1
      Left            =   4245
      TabIndex        =   5
      Top             =   870
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "按加入时间查询"
      Height          =   180
      Index           =   1
      Left            =   195
      TabIndex        =   3
      Top             =   930
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Index           =   2
      Left            =   4245
      TabIndex        =   8
      Top             =   1275
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   61341699
      CurrentDate     =   36588
   End
   Begin zlIDKind.PatiIdentify patiFind 
      Height          =   345
      Left            =   960
      TabIndex        =   16
      Top             =   2505
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindStr       =   $"frmBlackListRecordFilter.frx":0000
      BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IDKindAppearance=   2
      InputAppearance =   2
      ShowSortName    =   -1  'True
      DefaultCardType =   "就诊卡"
      IDKindWidth     =   555
      BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowAutoCommCard=   -1  'True
      AllowAutoICCard =   -1  'True
      AllowAutoIDCard =   -1  'True
      NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人"
      Height          =   180
      Left            =   540
      TabIndex        =   15
      Top             =   2595
      Width           =   360
   End
   Begin VB.Label lblRangDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Index           =   2
      Left            =   3990
      TabIndex        =   22
      Top             =   1335
      Width           =   180
   End
   Begin VB.Label lblRangDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Index           =   1
      Left            =   3990
      TabIndex        =   21
      Top             =   930
      Width           =   180
   End
   Begin VB.Label lbl加入原因 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "加入原因"
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lblRangDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "至"
      Height          =   180
      Index           =   0
      Left            =   3990
      TabIndex        =   20
      Top             =   540
      Width           =   180
   End
   Begin VB.Label lbl撤消人 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "撤消人"
      Height          =   180
      Left            =   3630
      TabIndex        =   11
      Top             =   1770
      Width           =   540
   End
   Begin VB.Label lbl登记人 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "登记人"
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   1770
      Width           =   540
   End
End
Attribute VB_Name = "frmBlackListRecordFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcllFilter As Collection
Private mlngModule As Long
Private mlng病人ID As Long
Public Function zlShowEdit(ByVal frmMain As Object, ByVal lngModule As Long, ByRef cllFilter As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:编辑常用原因
    '入参:frmMain-调用的主窗体
    '    cllFilter-过滤条件
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-08 17:01:16
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    Set mcllFilter = cllFilter: mlngModule = lngModule
    If mcllFilter Is Nothing Then Set mcllFilter = New Collection
    mblnOK = False
    Me.Show 1, frmMain
    Set cllFilter = mcllFilter
    zlShowEdit = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub LoadDefalutFilterValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载缺省条件值
    '编制:刘兴洪
    '日期:2018-02-28 14:07:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtSysdate As Date
    Dim i As Integer
    
    On Error GoTo errHandle
    
    dtSysdate = zlDatabase.Currentdate
    For i = 0 To dtpBegin.UBound
        dtpBegin(i).MaxDate = Format(dtSysdate, "yyyy-MM-dd 23:59:59")
        dtpBegin(i).Value = Format(dtSysdate - 7, "yyyy-MM-dd 00:00:00")
        dtpEnd(i).Value = dtpBegin(i).MaxDate
        dtpEnd(i).MaxDate = dtpBegin(i).MaxDate
    Next
    
    chkDate(0).Value = 1: chkDate(1).Value = 0: chkDate(2).Value = 0
    txt登记人.Text = ""
    txt撤消人.Text = ""
    cbo加入原因.Text = ""
    patiFind.Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Function GetConsFilter(ByRef cllFilter_Out As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取查询条件
    '入参:
    '出参:cllFilter-返回相关的条件集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-11-14 14:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHandle
    
    Set cllFilter_Out = New Collection
    If patiFind.Text <> "" And Val(patiFind.Tag) = 0 Then
        MsgBox "未找到指定的病人，请在病人处按回车查找病人!", vbInformation + vbOKOnly, gstrSysName
        patiFind.SetFocus
        Exit Function
    End If
    
    
    '暂时不开放这个，因为不良记录是一个小表，数据量不大，暂可以不使用索引
'    If Val(patiFind.Tag) = 0 And chkDate(0).Value <> 1 And chkDate(1).Value <> 1 And chkDate(2).Value <> 1 Then
'        MsgBox "病人、发生时间、撤消时间、加入时间必须要至少要设置一个条件！", vbInformation + vbOKOnly, gstrSysName
'        Exit Function
'    End If
    
    If Val(patiFind.Tag) <> 0 Then
        cllFilter_Out.Add Array("病人ID", Val(patiFind.Tag)), "病人ID"
    End If
    If Trim(txt登记人.Text) <> "" Then
        cllFilter_Out.Add Array("登记人", Trim(txt登记人.Text)), "登记人"
    End If
    If Trim(txt撤消人.Text) <> "" Then
        cllFilter_Out.Add Array("撤消人", Trim(txt撤消人.Text)), "撤消人"
    End If
    
    If Trim(cbo加入原因.Text) <> "" Then
        cllFilter_Out.Add Array("加入原因", Trim(cbo加入原因.Text)), "加入原因"
    End If
    If chkDate(0).Value = 1 Then
        cllFilter_Out.Add Array("发生时间", dtpBegin(0).Value, dtpEnd(0).Value), "发生时间"
    End If
    If chkDate(1).Value = 1 Then
        cllFilter_Out.Add Array("加入时间", dtpBegin(1).Value, dtpEnd(1).Value), "加入时间"
    End If
    If chkDate(2).Value = 1 Then
        cllFilter_Out.Add Array("撤消时间", dtpBegin(1).Value, dtpEnd(1).Value), "撤消时间"
    End If
    GetConsFilter = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo加入原因_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll cbo加入原因
    
End Sub

Private Sub cbo加入原因_LostFocus()
     zlCommFun.OpenIme False
End Sub

Private Sub chkDate_Click(Index As Integer)
    dtpBegin(Index).Enabled = chkDate(Index).Value = 1
    dtpEnd(Index).Enabled = chkDate(Index).Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim cllFilter As Collection
    If GetConsFilter(cllFilter) = False Then Exit Sub
    Set mcllFilter = cllFilter
    mblnOK = True
    Unload Me
End Sub

Private Sub dtpEnd_Change(Index As Integer)
    dtpBegin(Index).MaxDate = dtpEnd(Index).Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is patiFind Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub patiFind_Change()
    mlng病人ID = 0
    patiFind.Tag = ""
End Sub

Private Sub patiFind_FindPatiArfter(ByVal objCard As zlOneCardComLib.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlOneCardComLib.clsPatientInfo, objCardData As zlOneCardComLib.clsPatientInfo, strErrMsg As String, blnCancel As Boolean)
    Dim cllFilter As Collection, lngPatiID As Long
    
    If objHisPati Is Nothing Then
        lngPatiID = 0
    Else
        lngPatiID = objHisPati.病人ID
    End If
    mlng病人ID = lngPatiID
    patiFind.Tag = lngPatiID
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
    
End Sub

Private Sub LoadDataFromcllFilter()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据界面置，加载缺省条件
    '编制:刘兴洪
    '日期:2018-02-28 14:07:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtSysdate As Date, rsTemp As ADODB.Recordset
    Dim i As Integer, lng病人id As Long, strSQL As String
    Dim varData As Variant
    
    On Error GoTo errHandle
    For i = 1 To mcllFilter.Count
        varData = mcllFilter(i)
        Select Case varData(0)
        Case "病人ID"
            lng病人id = Val(varData(1))
            strSQL = "Select 姓名 From 病人信息 where 病人ID=[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人id)
            If Not rsTemp.EOF Then
                patiFind.Text = Nvl(rsTemp!姓名)
                patiFind.Tag = lng病人id
            End If
        Case "加入时间"
            dtpBegin(1).Value = Format(CDate(varData(1)), "yyyy-mm-dd HH:MM:SS")
            dtpEnd(1).Value = Format(CDate(varData(2)), "yyyy-mm-dd HH:MM:SS")
            chkDate(1).Value = 1
        Case "撤消时间"
            dtpBegin(2).Value = Format(CDate(varData(1)), "yyyy-mm-dd HH:MM:SS")
            dtpEnd(2).Value = Format(CDate(varData(2)), "yyyy-mm-dd HH:MM:SS")
            chkDate(2).Value = 1
        Case "发生时间"
            dtpBegin(0).Value = Format(CDate(varData(1)), "yyyy-mm-dd HH:MM:SS")
            dtpEnd(0).Value = Format(CDate(varData(2)), "yyyy-mm-dd HH:MM:SS")
            chkDate(0).Value = 1
        Case "加入原因"
            cbo加入原因.Text = varData(1)
        Case "登记人"
            txt登记人.Text = varData(1)
        Case "撤消人"
            txt撤消人.Text = varData(1)
        End Select
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面
    '编制:刘兴洪
    '日期:2018-11-14 14:28:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim objCards As zlOneCardComLib.Cards, i As Integer, strFind As String, strKindstr As String
    Dim objCard As zlOneCardComLib.Card
    
    strSQL = "Select 编码,名称,简码 From 常用不良行为原因  Order by  编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With cbo加入原因
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            rsTemp.MoveNext
        Loop
        .ListIndex = -1
    End With
    
    strKindstr = "姓|姓名或就诊卡|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;住|住院号|0;手|手机号|0"
    Call patiFind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKindstr, gstrProductName)
    Set objCards = patiFind.objIDKind.Cards
    If Not objCards Is Nothing Then
        strFind = Val(zlDatabase.GetPara("上次查找类别", glngSys, mlngModule, ""))  '查找缺省项
        If strFind <> "" Then
            For i = 1 To objCards.Count
                Set objCard = objCards(i)
                If objCard.名称 = strFind Then
                    If patiFind.GetKindIndex(objCard.接口序号) >= 0 Then
                        patiFind.IDKindIDX = i + 1
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    Call LoadDefalutFilterValue
End Sub

Private Sub cmdDef_Click()
    Call LoadDefalutFilterValue
End Sub

Private Sub Form_Load()
    Call InitFace   '初始化界面
    Call LoadDataFromcllFilter
End Sub

Private Sub patiFind_FindPatiBefore(ByVal objCard As zlOneCardComLib.Card, blnCard As Boolean, strShowText As String, objCardData As zlOneCardComLib.clsPatientInfo, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim strPati As String, vRect As RECT, rsTmp As ADODB.Recordset
    Dim strName As String, lng病人id As Long
    If objCard.名称 Like "*姓*名*" And Not blnCard And InStr("-*+/.", Left(Trim(patiFind.Text), 1)) = 0 Then
        If zlstr.ActualLen(patiFind.Text) < 4 Then
            MsgBox "根据病人姓名查找,请输入至少两个汉字或四个字符!", vbInformation + vbOKOnly, gstrSysName
            blnCancel = True: Exit Sub
        End If
        strName = Trim(patiFind.Text)
        strPati = _
            " Select A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
            "           A.住院号,B.名称 as 科室,A.当前床号 as 床号," & _
            "           A.出生日期,A.身份证号,A.家庭地址,decode(Nvl(A.在院,0),0,'','√') As 在院标志 " & _
            " From 病人信息 A,部门表 B " & _
            " Where  Rownum <101 And A.停用时间 is NULL And A.当前科室ID=B.ID(+) And A.姓名 Like [1] " & _
            "   Order by A.姓名"
        vRect = zlControl.GetControlRect(patiFind.hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, _
            vRect.Left, vRect.Top, patiFind.Height, blnCancel, False, True, _
            strName & "%", "bytSize=1")
        If Not rsTmp Is Nothing Then
            If Nvl(rsTmp!ID) = 0 Then '当作新病人
                blnCancel = True: Exit Sub
            Else '以病人ID读取
                lng病人id = Nvl(rsTmp!ID)
            End If
        Else '取消选择
            blnCancel = True: Exit Sub
        End If
        If zlGetPati(lng病人id, objCardData) = False Then blnCancel = True: Exit Sub
        blnFindPatied = True
    End If
End Sub

Private Sub txt撤消人_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt撤消人
End Sub

Private Sub txt撤消人_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt登记人_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt登记人
End Sub

Private Sub txt登记人_LostFocus()
    zlCommFun.OpenIme False
End Sub

Public Function zlGetPati(ByVal lng病人id As String, ByRef objPati As clsPatientInfo) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID,重新获取数据
    '返回:合法,返回true,否则返回False
    '编制:焦博
    '日期:2019-09-04 10:22:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    Set objPati = New clsPatientInfo
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select a.病人id, a. 门诊号, a.住院号, a.卡验证码, a.费别, a.医疗付款方式,p.编码 as 医疗付款方式编码, a. 姓名, a.性别, a. 年龄, a.出生日期, a.出生地点, a.身份证号, a.其他证件,  " & _
    "        a.职业, a.民族, a.国籍, a.区域, a.学历, a.婚姻状况, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.监护人, a.联系人姓名, a.联系人关系, a.联系人地址, a.联系人电话, " & _
    "        a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号,a.户口地址, a.户口地址邮编, a.籍贯, '' as 卡号, '' as 密码" & _
    "   From 病人信息 A,医疗付款方式 P" & _
    "   Where  a.医疗付款方式=P.名称(+) And 病人ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取病人信息", lng病人id)
    If rsTemp.EOF Then Exit Function
    objPati.病人ID = rsTemp!病人ID
    objPati.门诊号 = Nvl(rsTemp!门诊号)
    objPati.姓名 = Nvl(rsTemp!姓名)
    objPati.性别 = Nvl(rsTemp!性别)
    objPati.年龄 = Nvl(rsTemp!年龄)
    objPati.出生日期 = Format(rsTemp!出生日期, "yyyy-mm-dd")
    objPati.出生地址 = Nvl(rsTemp!出生地点)
    objPati.身份证号 = Nvl(rsTemp!身份证号)
    objPati.其他证件 = Nvl(rsTemp!其他证件)
    objPati.职业 = Nvl(rsTemp!职业)
    objPati.民族 = Nvl(rsTemp!民族)
    objPati.国籍 = Nvl(rsTemp!国籍)
    objPati.学历 = Nvl(rsTemp!学历)
    objPati.婚姻状况 = Nvl(rsTemp!婚姻状况)
    objPati.区域 = Nvl(rsTemp!区域)
    objPati.家庭地址 = Nvl(rsTemp!家庭地址)
    objPati.家庭电话 = Nvl(rsTemp!家庭电话)
    objPati.家庭邮编 = Nvl(rsTemp!家庭地址邮编)
    objPati.监护人 = Nvl(rsTemp!监护人)
    objPati.联系人 = Nvl(rsTemp!联系人姓名)
    objPati.联系人关系 = Nvl(rsTemp!联系人关系)
    objPati.联系人地址 = Nvl(rsTemp!联系人地址)
    objPati.联系人电话 = Nvl(rsTemp!联系人电话)
    objPati.工作单位 = Nvl(rsTemp!工作单位)
    objPati.工作单位电话 = Nvl(rsTemp!单位电话)
    objPati.工作单位邮编 = Nvl(rsTemp!单位邮编)
    objPati.工作单位开户行 = Nvl(rsTemp!单位开户行)
    objPati.工作单位开户行帐户 = Nvl(rsTemp!单位帐号)
    objPati.户口地址 = Nvl(rsTemp!户口地址)
    objPati.户口地址邮编 = Nvl(rsTemp!户口地址邮编)
    objPati.籍贯 = Nvl(rsTemp!籍贯)
    objPati.密码 = Nvl(rsTemp!密码)
    objPati.医疗付款方式编码 = Nvl(rsTemp!医疗付款方式编码)
    objPati.医疗付款方式 = Nvl(rsTemp!医疗付款方式)
    zlGetPati = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

