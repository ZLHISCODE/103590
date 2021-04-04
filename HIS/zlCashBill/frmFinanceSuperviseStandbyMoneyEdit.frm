VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFinanceSuperviseStandbyMoneyEdit 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "备用金领用单"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtBackTime 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3795
      Width           =   2145
   End
   Begin VB.TextBox txtBackPerson 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3795
      Width           =   1785
   End
   Begin VB.ComboBox cboNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5295
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1035
      Width           =   2040
   End
   Begin VB.ComboBox cboPerson 
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Text            =   "cboPerson"
      Top             =   1890
      Width           =   2040
   End
   Begin VB.TextBox txtMemo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2820
      Width           =   6120
   End
   Begin VB.TextBox txtTime 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3315
      Width           =   2145
   End
   Begin VB.TextBox txtInputPerson 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3315
      Width           =   1785
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6315
      TabIndex        =   13
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5055
      TabIndex        =   12
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印设置(&S)"
      Height          =   350
      Left            =   90
      TabIndex        =   14
      Top             =   4800
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   330
      Left            =   1320
      TabIndex        =   3
      Top             =   2340
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   101056515
      CurrentDate     =   41520
   End
   Begin VB.TextBox txtMoney 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2355
      Width           =   2160
   End
   Begin VB.Label lblBackTime 
      AutoSize        =   -1  'True
      Caption         =   "登记时间"
      Height          =   210
      Left            =   4395
      TabIndex        =   21
      Top             =   3855
      Width           =   840
   End
   Begin VB.Label lblBackPerson 
      AutoSize        =   -1  'True
      Caption         =   "回收人"
      Height          =   210
      Left            =   690
      TabIndex        =   20
      Top             =   3855
      Width           =   630
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "领用金额(&M)"
      Height          =   210
      Left            =   4125
      TabIndex        =   4
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   15
      X2              =   10455
      Y1              =   1485
      Y2              =   1485
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5070
      TabIndex        =   17
      Top             =   1080
      Width           =   210
   End
   Begin VB.Label lblPerson 
      AutoSize        =   -1  'True
      Caption         =   "领用人(&P)"
      Height          =   210
      Left            =   375
      TabIndex        =   0
      Top             =   1965
      Width           =   945
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "摘  要(&Z)"
      Height          =   210
      Left            =   375
      TabIndex        =   6
      Top             =   2910
      Width           =   945
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "领用时间(&T)"
      Height          =   210
      Left            =   165
      TabIndex        =   2
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "登记时间"
      Height          =   210
      Left            =   4395
      TabIndex        =   10
      Top             =   3375
      Width           =   840
   End
   Begin VB.Label lblInputPerson 
      AutoSize        =   -1  'True
      Caption         =   "登记人"
      Height          =   210
      Left            =   690
      TabIndex        =   8
      Top             =   3360
      Width           =   630
   End
   Begin VB.Line linMain 
      BorderColor     =   &H8000000C&
      X1              =   -285
      X2              =   10155
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Label lblTittle 
      Alignment       =   2  'Center
      Caption         =   "备用金领用单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   135
      TabIndex        =   16
      Top             =   210
      Width           =   7170
   End
End
Attribute VB_Name = "frmFinanceSuperviseStandbyMoneyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mlngID As Long
Private mstr领用人 As String
Public Enum gEditCard
    EM_ED_增加 = 0
    EM_ED_上岗 = 1
    EM_ED_查看 = 2
End Enum
Private mEditType As gEditCard
Private mrsChargePerson As ADODB.Recordset
Private mblnFirst As Boolean, mblnOK As Boolean
Private mblnChange  As Boolean
Private mblnUnload As Boolean

Public Function EditCard(ByVal frmMain As Object, _
    ByVal EditType As gEditCard, _
    ByVal lngModuel As Long, ByVal strPrivs As String, _
    ByVal str领用人 As String, Optional ByVal lngID As Long = 0) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的主窗体
    '       EditType-编辑类型
    '       lngID-暂存ID(查看时传入)
    '       str领用人-领用人
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-12 15:37:32
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mlngModule = lngModuel: mstrPrivs = strPrivs
    mlngID = lngID: mblnOK = False: mstr领用人 = str领用人
    If frmMain Is Nothing Then
         Me.Show vbModal
    Else
         Me.Show vbModal, frmMain
    End If
    EditCard = mblnOK
End Function

Private Sub ClearCtrlData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除控件数据
    '编制:刘兴洪
    '日期:2013-10-12 15:44:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtMemo.Text = "": txtMoney.Text = ""
    txtInputPerson.Text = ""
    txtTime.Text = ""
    txtBackPerson.Text = ""
    txtBackTime.Text = ""
End Sub
Private Sub SetCtrlEnabled()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enabled属性
    '编制:刘兴洪
    '日期:2013-10-12 16:02:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    Dim lngBackColor As Long
    
    blnEnabled = mEditType = EM_ED_增加 Or mEditType = EM_ED_上岗
    lngBackColor = IIf(blnEnabled, &H80000005, &H8000000F)
    txtMemo.Enabled = blnEnabled: txtMemo.BackColor = lngBackColor
    txtMoney.Enabled = blnEnabled: txtMoney.BackColor = lngBackColor
    cboNO.Enabled = blnEnabled: cboNO.BackColor = lngBackColor
    dtpDate.Enabled = blnEnabled
    dtpDate.Value = zlDatabase.Currentdate
    cboPerson.Enabled = blnEnabled: cboPerson.BackColor = lngBackColor
    txtInputPerson.Enabled = False
    txtTime.Enabled = False
    txtBackPerson.Enabled = False
    txtBackTime.Enabled = False
    cmdOK.Visible = mEditType <> EM_ED_查看
End Sub

Private Function LoadCardData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载卡片数据
    '返回:加载成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-12 15:43:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblMoney As Double
    On Error GoTo errHandle
    If mEditType = EM_ED_增加 Or mEditType = EM_ED_上岗 Then
        Call ClearCtrlData
        txtInputPerson.Text = UserInfo.姓名
        txtTime.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        dblMoney = Val(zlDatabase.GetPara("缺省领用备用金额", glngSys, mlngModule, 1000, Array(txtMoney, lblMoney), InStr(1, mstrPrivs, ";参数设置;") > 0))
        txtMoney.Text = Format(dblMoney, "0.00")
        If txtMoney.Enabled Then txtMoney.BackColor = &H80000005
        If mEditType = EM_ED_上岗 Then
            lblTittle.Caption = "备用金领用单(上岗)"
        Else
            lblTittle.Caption = "备用金领用单"
        End If
        LoadCardData = LoadPerson: Exit Function
    End If

    strSQL = "" & _
    "   Select ID,收缴ID,记录性质,NO,结算方式,金额,收款员 as 领用人,领用时间, " & _
    "           收回人,收回时间,备注,登记人,登记时间  " & _
    "   From 人员暂存记录 " & _
    "   Where ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If rsTemp.EOF Then
        MsgBox "未找到指定的备用金领用记录,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    cmdCancel.Caption = "退出(&X)"
    With cboPerson
        .Clear
        .AddItem Nvl(rsTemp!领用人)
        .ListIndex = .NewIndex
    End With
    dtpDate.Value = Format(rsTemp!领用时间, "yyyy-mm-dd")
    txtTime.Text = Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
    txtBackTime.Text = Format(rsTemp!收回时间, "yyyy-mm-dd HH:MM:SS")
    txtInputPerson.Text = Nvl(rsTemp!登记人)
    txtBackPerson.Text = Nvl(rsTemp!收回人)
    txtMoney.Text = Format(Val(Nvl(rsTemp!金额)), "###0.00;-###0.00;0.00;-0.00")
    txtMoney.Tag = Nvl(rsTemp!结算方式)
    txtMemo.Text = Nvl(rsTemp!备注)
    cboNO.AddItem Nvl(rsTemp!NO)
    cboNO.ListIndex = cboNO.NewIndex
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Private Sub cboPerson_Click()
    mblnChange = True
End Sub
  

Private Sub cboPerson_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim rsTemp As ADODB.Recordset
    If KeyAscii <> 13 Then Exit Sub
    
    If cboPerson.Locked Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    strText = UCase(cboPerson.Text)
    If cboPerson.ListIndex <> -1 Then
        '弹出列表时,又在文本框输入了内容
        If strText <> UCase(cboPerson.List(cboPerson.ListIndex)) Then Call zlcontrol.CboSetIndex(cboPerson.hWnd, -1)
    End If
    If strText = "" Then cboPerson.ListIndex = -1: Exit Sub
    If cboPerson.ListIndex >= 0 Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    
    intIdx = -1
    '先复制记录集
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsChargePerson)
    strCompents = Replace(gstrLike, "%", "*") & strText & "*"
    If IsNumeric(strText) Then
        intInputType = 0 '0-输入的是全数字
    ElseIf zlCommFun.IsCharAlpha(strText) Then
        intInputType = 1 '1-输入的是全字母
    Else
        intInputType = 2 '2-其他
    End If
    mrsChargePerson.Filter = 0: iCount = 0
    With mrsChargePerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not mrsChargePerson.EOF
            Select Case intInputType
            Case 0  '输入的是全数字
                '如果输入的数字,需要检查:
                '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                If Nvl(!编号) = strText Then strResult = Nvl(!姓名): iCount = 0: Exit Do
                '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                If Val(Nvl(!编号)) = Val(strText) Then
                    If iCount = 0 Then strResult = Nvl(!姓名)
                    iCount = iCount + 1
                End If
                '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                 If Val(mrsChargePerson!编号) Like strText & "*" Then
                    If CheckPersonExists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
                 End If
                 
            Case 1  '输入的是全字母
                '规则:
                ' 1.输入的简码相等,则直接定位
                ' 2.根据参数来匹配相同数据
                
                '1.输入的简码相等,则直接定位
                If Trim(Nvl(!简码)) = strText Then
                    If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同简码
                    iCount = iCount + 1
                End If
                '2.根据参数来匹配相同数据
                If Trim(Nvl(!简码)) Like strCompents Then
                    If CheckPersonExists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
                End If
            Case Else  ' 2-其他
                '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                '1.编码\简码相等,直接定位
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                
                '1.编码\简码相等,直接定位
                If Trim(!编号) = strText Or Trim(!简码) = strText Or Trim(!姓名) = strText Then
                    If iCount = 0 Then strResult = Nvl(!姓名)   '可能存在多个相同的多个
                    iCount = iCount + 1
                End If
                '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                If Trim(!编号) Like strText & "*" Or Trim(Nvl(!简码)) Like strCompents Or Trim(Nvl(!姓名)) Like strCompents Then
                    If CheckPersonExists(Nvl(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
                End If
            End Select
            mrsChargePerson.MoveNext
        Loop
    End With
    
    If iCount > 1 Then strResult = ""
    If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!姓名)
    '直接定位
    If strResult <> "" Then
        rsTemp.Close: Set rsTemp = Nothing
        If CheckPersonExists(strResult, True) Then zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
     If rsTemp.RecordCount = 0 Then
        '未找到
        rsTemp.Close: Set rsTemp = Nothing
        KeyAscii = 0: zlcontrol.TxtSelAll cboPerson: Exit Sub
     End If
     
    '先按某种方式进行排序
    Select Case intInputType
    Case 0 '输入全数字
        rsTemp.Sort = "编号"
    Case 1 '输入全拼音
        rsTemp.Sort = "简码"
    Case Else
        '根据选择来定
        rsTemp.Sort = "编号"
    End Select
    '弹出选择器
    Dim rsReturn As ADODB.Recordset
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, cboPerson, rsTemp, True, "", "", rsReturn) Then
        If cboPerson.Enabled Then cboPerson.SetFocus
        If Not rsReturn Is Nothing Then
            If rsReturn.RecordCount <> 0 Then
                '进行定位
                If CheckPersonExists(Nvl(rsReturn!姓名), True) Then
                    'zlCommFun.PressKey vbKeyTab
                End If
            End If
        End If
    End If
    rsTemp.Close: Set rsTemp = Nothing
End Sub

Private Sub cboPerson_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub cboPerson_Validate(Cancel As Boolean)
    If cboPerson.Text <> "" Then
        If cbo.FindIndex(cboPerson, zlStr.NeedName(cboPerson.Text), True) = -1 Then cboPerson.ListIndex = -1: cboPerson.Text = ""
    End If
    If cboPerson.Text = "" Then Call cboPerson_KeyPress(vbKeyReturn)
    '有数据，必须输入
    If cboPerson.ListIndex = -1 And cboPerson.ListCount <> 0 Then Cancel = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim strNO As String
    If isValied = False Then Exit Sub
    If SaveData(strNO) = False Then Exit Sub
    Call BillPrint(strNO)
    MsgBox "备用金发放成功!", vbOKOnly + vbInformation, gstrSysName
    Call LoadCardData
    If cboPerson.Enabled And cboPerson.Visible Then cboPerson.SetFocus
    mblnChange = False: mblnOK = True
End Sub

Private Sub BillPrint(ByVal strNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收款收据打印
    '编制:刘兴洪
    '日期:2013-09-11 11:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    blnPrint = False
    If Not zlStr.IsHavePrivs(mstrPrivs, "备用金领用单") Then Exit Sub
    Select Case Val(zlDatabase.GetPara("备用金领用单打印方式", glngSys, mlngModule))     '使用医生站的相关参数
    Case 0    '不打印
        Exit Sub
    Case 1    '自助动打印
        blnPrint = True
    Case 2    '选择打印
        If MsgBox("你是否要打印缴款收据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            blnPrint = True
        End If
    End Select
    If blnPrint = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500_1", Me, "NO=" & strNO, 2)
End Sub

Private Function SaveData(ByRef strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '出参:strNo-数据保存成功，返回单据号
    '返回:数据保存成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-12 17:07:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strSQL As String
    On Error GoTo errHandle
    
    lngID = zlDatabase.GetNextId("人员暂存记录")
    strNO = zlDatabase.GetNextNo(141)
    '    Zl_人员暂存记录_Insert
    strSQL = "Zl_人员暂存记录_Insert("
    '  Id_In       In 人员暂存记录.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  No_In       In 人员暂存记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  金额_In     In 人员暂存记录.金额%Type,
    strSQL = strSQL & "" & Val(txtMoney.Text) & ","
    '  领用人_In   In 人员暂存记录.收款员%Type,
    strSQL = strSQL & "'" & zlStr.NeedName(cboPerson.Text) & "',"
    '  领用时间_In In 人员暂存记录.领用时间%Type,
    strSQL = strSQL & "to_date('" & Format(dtpDate.Value, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  备注_In     In 人员暂存记录.备注%Type,
    strSQL = strSQL & IIf(Trim(txtMemo.Text) = "", "NULL", "'" & Trim(txtMemo.Text) & "'") & ","
    '  登记人_In   In 人员暂存记录.登记人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  登记时间_In In 人员暂存记录.登记时间%Type
    strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  记录性质_In In 人员暂存记录.记录性质%Type
    If mEditType = EM_ED_上岗 Then
        strSQL = strSQL & "" & 1 & ")"
    ElseIf mEditType = EM_ED_增加 Then
        strSQL = strSQL & "" & 11 & ")"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function isValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-12 16:56:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If zlCommFun.ActualLen(txtMemo.Text) > 50 Then
         MsgBox "摘要超长,最多只能输入25个字符或50个汉字", vbInformation, gstrSysName
         If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
         Exit Function
    End If
    If InStr(1, txtMemo.Text, "'") > 0 Then
        MsgBox "摘要中不能包含单引号!", vbInformation, gstrSysName
        If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
        Exit Function
    End If
    If cboPerson.ListIndex < 0 Then
        MsgBox "未选择领用人!", vbInformation, gstrSysName
        If cboPerson.Visible And cboPerson.Enabled Then cboPerson.SetFocus
        Exit Function
    End If
    
'    If Val(txtMoney.Text) = 0 Then
'        MsgBox "必须输入的金额!", vbInformation, gstrSysName
'        If txtMoney.Visible And txtMoney.Enabled Then txtMoney.SetFocus
'        Exit Function
'    End If
    
    If Val(txtMoney.Text) > 99999999 Or Val(txtMoney.Text) < 0 Then
        MsgBox "输入的金额必须在0-99999999范围之内!", vbInformation, gstrSysName
        If txtMoney.Visible And txtMoney.Enabled Then txtMoney.SetFocus
        Exit Function
    End If
    If Format(dtpDate.Value, "yyyy-MM-dd HH:mm:ss") > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then
        MsgBox "输入的领用日期不能大于" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "!", vbInformation, gstrSysName
        If dtpDate.Visible And dtpDate.Enabled Then dtpDate.SetFocus
        Exit Function
    End If
    
    If mEditType = EM_ED_上岗 Then
        strSQL = "Select 1 From 人员缴款余额 where 收款员=[1] and 性质=1 and 余额<>0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(zlStr.NeedName(cboPerson.Text)))
        If Not rsTemp.EOF Then
            MsgBox "领用人:" & Trim(cboPerson.Text) & " 已经产生收款记录,无法领用上岗备用金", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSQL = "Select 1 From 人员暂存记录 where 收款员=[1] and 收回时间 is null And MOD(记录性质,10)=1  and Rownum<2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(zlStr.NeedName(cboPerson.Text)))
    If Not rsTemp.EOF Then
        If MsgBox("领用人:" & Trim(cboPerson.Text) & " 已经领用过备用金,是否继续领用?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then
            If cboPerson.Visible And cboPerson.Enabled Then cboPerson.SetFocus
            Exit Function
        End If
    End If
    
'    strSQL = "Select Count(1) As 数量 From 人员暂存记录 Where 收款员=[1] And 收回时间 Is Null And 记录性质=11"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(NeedName(cboPerson.Text)))
'    If Val(Nvl(rsTemp!数量)) < 1 Then
'        strSQL = "Select 1 From 人员缴款余额 where 收款员=[1] and 性质=1 and 余额<>0"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(NeedName(cboPerson.Text)))
'        If Not rsTemp.EOF Then
'            MsgBox "领用人:" & Trim(cboPerson.Text) & " 已经产生收款记录,无法领用备用金", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdPrintSet_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500_1", Me
End Sub
Private Sub dtpDate_Change()
    mblnChange = True
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnload Then Unload Me: Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is cboPerson Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Function LoadPerson() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收费员信息
    '入参:blnFilter-是否进行过滤
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-23 11:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsReturn As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim strSQL As String, strIcon As String
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct A.ID,A.编号,A.姓名,A.简码,M.名称 as 所属部门,a.性别" & _
    "   From 人员表 A,人员性质说明 B, 部门人员 C,部门表 M" & _
    "   Where A.id = B.人员ID And B.人员性质 In ('门诊挂号员','门诊收费员','预交收款员','住院结帐员','入院登记员','发卡登记人')  " & _
    "               And A.ID=C.人员ID and C.部门ID=M.ID(+) And C.缺省(+)=1 " & _
    "               And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "               And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
    "   Order By 编号"
    Set mrsChargePerson = zlDatabase.OpenSQLRecord(strSQL, "获取收费员信息")
    If mrsChargePerson.RecordCount = 0 Then
        MsgBox "不存在一个人员性质为: " & vbCrLf & _
                      "     门诊挂号员,门诊收费员,预交收款员,住院结帐员,入院登记员,发卡登记人 " & vbCrLf & _
                      "的收费人员,请在[人员管理]中进行设置!", vbInformation + vbOKOnly, gstrSysName
        cboPerson.Clear
        Exit Function
    End If
    With cboPerson
        .Clear
        Do While Not mrsChargePerson.EOF
            .AddItem Nvl(mrsChargePerson!编号) & "-" & Nvl(mrsChargePerson!姓名)
            .ItemData(.NewIndex) = Val(Nvl(mrsChargePerson!ID))
            If .ListIndex < 0 Then .ListIndex = .NewIndex
            If Nvl(mrsChargePerson!姓名) = mstr领用人 Then .ListIndex = .NewIndex
            mrsChargePerson.MoveNext
        Loop
        'If .ListCount <> 0 Then .ListIndex = 0
    End With
    LoadPerson = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Load()
    mblnFirst = True
    Call SetCtrlEnabled
    mblnUnload = Not LoadCardData
    If mblnUnload Then Exit Sub
    mblnChange = False
End Sub

Private Sub txtMemo_Change()
    mblnChange = True
End Sub
 
Private Sub txtMemo_GotFocus()
    zlCommFun.OpenIme True
    zlcontrol.TxtSelAll txtMemo
End Sub

Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    zlcontrol.TxtCheckKeyPress txtMemo, KeyAscii, m文本式
End Sub

Private Sub txtMemo_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtMoney_Change()
    mblnChange = True
End Sub
Private Function CheckPersonExists(ByVal str姓名 As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在你收费员下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '出参:
    '返回:存在返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cboPerson.ListCount - 1
        If zlStr.NeedName(cboPerson.List(i)) = str姓名 Then
            If blnLocateItem Then cboPerson.ListIndex = i
            CheckPersonExists = True
            Exit Function
        End If
    Next
End Function

Private Sub txtMoney_GotFocus()
    zlCommFun.OpenIme False
    zlcontrol.TxtSelAll txtMoney
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    zlcontrol.TxtCheckKeyPress txtMoney, KeyAscii, m金额式
End Sub
