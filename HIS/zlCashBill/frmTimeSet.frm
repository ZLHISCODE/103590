VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTimeSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置过滤条件"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmTimeSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraFilter 
      Caption         =   "过滤条件"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3765
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1830
         TabIndex        =   4
         Top             =   870
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   170852355
         CurrentDate     =   36279
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1830
         TabIndex        =   2
         Top             =   390
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   170852355
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1830
         TabIndex        =   6
         Text            =   "cboOperator"
         Top             =   1320
         Width           =   1785
      End
      Begin VB.Label lblOperator 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "领用人(&P)"
         Height          =   180
         Left            =   960
         TabIndex        =   5
         Top             =   1395
         Width           =   810
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   150
         Picture         =   "frmTimeSet.frx":000C
         Top             =   420
         Width           =   480
      End
      Begin VB.Label lblTimeStart 
         AutoSize        =   -1  'True
         Caption         =   "开始时间(&B)"
         Height          =   180
         Left            =   780
         TabIndex        =   1
         Top             =   450
         Width           =   990
      End
      Begin VB.Label lblTimeStop 
         AutoSize        =   -1  'True
         Caption         =   "结束时间(&E)"
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   930
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   7
      Top             =   240
      Width           =   1100
   End
End
Attribute VB_Name = "frmTimeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mbytInFun As Byte '0-收费缴款过滤，1-票据使用过滤
Private mdatBegin As Date, mdatEnd As Date
Private mstrOperator As String, mstrPrivs As String
Private mrsPerson As ADODB.Recordset
Private mlngModule  As Long
Private mlngPreID As Long
Private mblnDateMoved As Boolean '是否在转出日期之前
 
Private Sub cboOperator_Click()
    If cboOperator.ListIndex >= 0 Then mlngPreID = cboOperator.ItemData(cboOperator.ListIndex)
End Sub

Private Sub cboOperator_KeyPress(KeyAscii As Integer)
   Dim lngIdx As Long, lng医生ID As Long
     '刘兴洪 问题:27378 日期:2010-01-27 16:20:02
    Dim strAllCaption As String
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cboOperator.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mbytInFun <> 0 And InStr(mstrPrivs, ";所有操作员;") > 0 Then
        strAllCaption = "所有人员"
    Else
    End If

    If mrsPerson Is Nothing Then Exit Sub
    If zlPersonSelect(Me, mlngModule, cboOperator, mrsPerson, _
        cboOperator.Text, True, strAllCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub
Private Sub cboOperator_Validate(Cancel As Boolean)
    If cboOperator.ListIndex < 0 Then zlControl.CboLocate cboOperator, mlngPreID, True
    If cboOperator.ListIndex < 0 And cboOperator.Text <> "" Then cboOperator.Text = ""
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpEnd.SetFocus
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "开始时间不应大于结束时间。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    mblnDateMoved = zlDatabase.DateMoved(Format(dtpBegin.Value, "yyyy-MM-dd hh:mm:ss"), , , Me.Caption)
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    
    If cboOperator.Text <> "所有人员" Then
        mstrOperator = zlCommFun.GetNeedName(cboOperator.Text)
    Else
        mstrOperator = ""
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function ShowMe(ByVal frmOwner As Form, ByVal bytInFun As Byte, _
    ByVal bytInvoiceKind As gBillType, ByVal lngModule As Long, ByVal strPrivs As String, _
    datBegin As Date, datEnd As Date, strOperator As String, blnDateMoved As Boolean, _
    Optional strPersonelKind As String, Optional blnOnlyHave As Boolean) As Boolean
'参数：
'    bytInFun:0-收费缴款过滤，1-票据使用过滤
'    bytInvoiceKind:当bytInFun=1时，1-收费收据,2-预交收据,3-结帐收据,4-挂号收据,5-就诊卡,6-消费卡,7-会员卡
'    strPersonelKind:人员性质，为空时表示所有性质
'    blnOnlyHave:只包含有余额的人员

    mbytInFun = bytInFun
    mstrPrivs = strPrivs: mlngModule = lngModule
                        
    dtpBegin.Value = TruncateDate(datBegin)
    dtpEnd.Value = TruncateDate(datEnd)
    dtpBegin.MaxDate = TruncateDate(zlDatabase.Currentdate)
    dtpEnd.MaxDate = dtpBegin.MaxDate
                    
    If mbytInFun = 0 Then
        lblOperator.Caption = "缴款人(&P)"
        Call Fill缴款人(strOperator, strPersonelKind, blnOnlyHave)
    Else
        lblOperator.Caption = "领用人(&P)"
        Call FillOperator(bytInvoiceKind)
    End If
    
    frmTimeSet.Show vbModal, frmOwner
    ShowMe = mblnOK
    If mblnOK = True Then
        datBegin = mdatBegin
        datEnd = mdatEnd
        strOperator = mstrOperator
        blnDateMoved = mblnDateMoved
    End If
End Function

Private Sub Fill缴款人(strOperator As String, strPersonelKind As String, blnOnlyHave As Boolean)
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    cboOperator.Clear
    
    If strPersonelKind = "" Then
        strSQL = " And C.人员性质 in " & _
                "       ('门诊挂号员','门诊收费员','预交收款员','住院结帐员','入院登记员','发卡登记人')"
    Else
        strSQL = " And C.人员性质=[1]"
    End If
                
    If blnOnlyHave Then
        '在指点定期间内有暂存金的操作员
        strSQL = _
            "Select Distinct B.ID,B.编号, B.姓名,B.简码" & vbNewLine & _
            "From 人员缴款余额 A,人员表 B,人员性质说明 C" & vbNewLine & _
            "Where A.收款员=B.姓名 And B.id=C.人员ID And a.余额<>0" & vbNewLine & _
            "      And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                   strSQL & vbNewLine & _
            "Order by 姓名"
    Else
        '所有期间内操作员
        strSQL = _
            "Select Distinct A.ID,A.编号, A.姓名,A.简码" & vbNewLine & _
            "From 人员表 A,人员性质说明 C " & vbNewLine & _
            "Where A.ID=C.人员ID And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            "      And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & vbNewLine & _
                   strSQL & vbNewLine & _
            "Order by 姓名"
    End If
    Set mrsPerson = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPersonelKind)
    If mrsPerson.EOF Then Exit Sub
    
    For i = 1 To mrsPerson.RecordCount
        cboOperator.AddItem mrsPerson!编号 & "-" & mrsPerson!姓名
        cboOperator.ItemData(cboOperator.NewIndex) = Val(Nvl(mrsPerson!ID))
        If strOperator = mrsPerson!姓名 Then cboOperator.ListIndex = cboOperator.NewIndex
        mrsPerson.MoveNext
    Next
    If cboOperator.ListIndex = -1 And cboOperator.ListCount > 0 Then cboOperator.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillOperator(ByVal bytInvoiceKind As gBillType)
    Dim strSQL As String
    Dim strValue As String, i As Long, strID As Long
    
    If InStr(mstrPrivs, "所有操作员") = 0 Then
        cboOperator.Clear
        cboOperator.AddItem UserInfo.编号 & "-" & UserInfo.姓名
        cboOperator.ItemData(cboOperator.NewIndex) = UserInfo.ID
        cboOperator.ListIndex = 0
    Else
        If bytInvoiceKind > 0 And bytInvoiceKind <= 7 Then
            '如果是入院登记员，则需要同时设置对应的发卡或预交人员属性这里才显示，病人信息管理同样也有这两项功能了
            strValue = Choose(bytInvoiceKind, "门诊收费员", "预交收款员", "住院结帐员", "门诊挂号员", _
                "发卡登记人", "发卡登记人", "发卡登记人")
        End If
        strSQL = _
            "Select Distinct A.ID, A.编号, A.姓名,A.简码" & vbNewLine & _
            "From 人员表 A, 人员性质说明 B" & vbNewLine & _
            "Where A.ID = B.人员id And B.人员性质 = [1] " & vbNewLine & _
            "      And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
            "      And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)"

        On Error GoTo errH
        Set mrsPerson = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue)
       
        cboOperator.Clear
        cboOperator.AddItem "所有人员"
        For i = 1 To mrsPerson.RecordCount
            cboOperator.AddItem mrsPerson!编号 & "-" & mrsPerson!姓名
            cboOperator.ItemData(cboOperator.NewIndex) = Val(Nvl(mrsPerson!ID))
            mrsPerson.MoveNext
        Next
        cboOperator.ListIndex = 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
