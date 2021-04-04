VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.UserControl usrOperGeneral 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   LockControls    =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   8040
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   2
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   660
      Width           =   2145
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   10
      Left            =   915
      MaxLength       =   20
      TabIndex        =   7
      Top             =   990
      Width           =   2145
   End
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "麻醉"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   15
      TabIndex        =   10
      Top             =   1755
      Width           =   690
   End
   Begin VB.CheckBox chk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "输氧"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   21
      Top             =   3120
      Width           =   675
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2370
      Width           =   2145
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   0
      Left            =   915
      MaxLength       =   4
      TabIndex        =   9
      Top             =   1320
      Width           =   1875
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      ItemData        =   "usrOperGeneral.ctx":0000
      Left            =   915
      List            =   "usrOperGeneral.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2715
      Width           =   2145
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Index           =   3
      Left            =   945
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1845
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Index           =   1
      Left            =   945
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2055
      Width           =   1845
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Index           =   2
      Left            =   945
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3105
      Width           =   1845
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Index           =   4
      Left            =   945
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3450
      Width           =   1845
   End
   Begin ZL9BillEdit.BillEdit bill 
      Height          =   1380
      Index           =   0
      Left            =   3960
      TabIndex        =   31
      Top             =   0
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2434
      Appearance      =   0
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
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
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   0
      Left            =   915
      TabIndex        =   1
      Top             =   0
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   1
      Left            =   915
      TabIndex        =   3
      Top             =   330
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   2
      Left            =   915
      TabIndex        =   13
      Top             =   1680
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   3
      Left            =   915
      TabIndex        =   16
      Top             =   2025
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   4
      Left            =   915
      TabIndex        =   24
      Top             =   3060
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   5
      Left            =   915
      TabIndex        =   27
      Top             =   3405
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67567619
      CurrentDate     =   37908
   End
   Begin ZL9BillEdit.BillEdit bill 
      Height          =   1380
      Index           =   1
      Left            =   3960
      TabIndex        =   33
      Top             =   1365
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2434
      Appearance      =   0
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
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
   Begin ZL9BillEdit.BillEdit bill2 
      Height          =   1380
      Index           =   0
      Left            =   3960
      TabIndex        =   35
      Top             =   2730
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   2434
      Appearance      =   0
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
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
   Begin ZL9BillEdit.BillEdit bill2 
      Height          =   1800
      Index           =   1
      Left            =   3960
      TabIndex        =   37
      Top             =   4095
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   3175
      Appearance      =   0
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
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
   Begin ZL9BillEdit.BillEdit bill 
      Height          =   2145
      Index           =   2
      Left            =   900
      TabIndex        =   29
      Top             =   3750
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   3784
      Appearance      =   0
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "手术规模"
      Height          =   180
      Index           =   17
      Left            =   120
      TabIndex        =   4
      Top             =   690
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "手 术 间"
      Height          =   180
      Index           =   16
      Left            =   120
      TabIndex        =   6
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "手术时间"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "到"
      Height          =   180
      Index           =   5
      Left            =   690
      TabIndex        =   2
      Top             =   375
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "拟行手术"
      Height          =   180
      Index           =   10
      Left            =   3180
      TabIndex        =   30
      Top             =   30
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已行手术"
      Height          =   180
      Index           =   9
      Left            =   3180
      TabIndex        =   32
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "到"
      Height          =   180
      Index           =   6
      Left            =   675
      TabIndex        =   14
      Top             =   2070
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "从"
      Height          =   180
      Index           =   7
      Left            =   690
      TabIndex        =   11
      Top             =   1755
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "麻醉方式"
      Height          =   180
      Index           =   8
      Left            =   105
      TabIndex        =   17
      Top             =   2415
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "麻醉质量"
      Height          =   180
      Index           =   13
      Left            =   105
      TabIndex        =   19
      Top             =   2775
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "到"
      Height          =   180
      Index           =   12
      Left            =   675
      TabIndex        =   25
      Top             =   3465
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "从"
      Height          =   180
      Index           =   11
      Left            =   675
      TabIndex        =   22
      Top             =   3120
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输液总量                      ML"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1380
      Width           =   2880
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "手术人员"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   28
      Top             =   3750
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "术后诊断"
      Height          =   180
      Index           =   2
      Left            =   3180
      TabIndex        =   36
      Top             =   4080
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "术前诊断"
      Height          =   180
      Index           =   3
      Left            =   3180
      TabIndex        =   34
      Top             =   2730
      Width           =   720
   End
End
Attribute VB_Name = "usrOperGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private i As Long
Private strSQL As String
Private rsTmp As New ADODB.Recordset

Private mlng病历id As Long                      '外界传入
Private mlng医嘱id As Long                      '外界传入
Private mstr性别 As String

Private mlng手术记录id As Long

Private mlng手术间id As Long                    '暂存变量
Private mlngOrderID As Long                     '暂存变量

Private Const STR_COMPART = "|';"
Private Const LAWLChar = "';`|,"""

Private mblnMode As Boolean '为真是表示是用户进行的编辑，这时才赋值

Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String

Private mblnLoaded As Boolean

'-------------------------------------------------------------------------------------------------------------------
'公共方法、属性
Public Property Get DispMode() As Boolean
    '是否为显示模式
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    mDispMode = New_DispMode
    ShowOperGeneral mlng病历id, Not mDispMode
    PropertyChanged "DispMode"
    
    If mDispMode Then
        
        dtp(0).Enabled = False
        dtp(1).Enabled = False
        dtp(2).Enabled = False
        dtp(3).Enabled = False
        dtp(4).Enabled = False
        dtp(5).Enabled = False
        
        cbo(0).Locked = True
        cbo(1).Locked = True
        cbo(2).Locked = True
        
        txt(0).Locked = True
        txt(10).Locked = True
        
        bill(0).Active = False
        bill(1).Active = False
        bill(2).Active = False
        
        bill2(0).Active = False
        bill2(1).Active = False
        
        chk(0).Enabled = False
        chk(1).Enabled = False
        
    End If
    
End Property

Public Property Get ID病人病历() As Long
    '返回病人病历ID
    
    ID病人病历 = mlng病历id
End Property

Public Property Let ID病人病历(ByVal New_ID病人病历 As Long)
    '设置病人病历ID,并检查该病历是不是存在
    
    mlng病历id = New_ID病人病历
    ShowOperGeneral mlng病历id, Not mDispMode
    
End Property

Public Property Let Get医嘱id(ByVal New_医嘱ID As Long)
        
    mlng医嘱id = New_医嘱ID
        
End Property

Public Property Get Get医嘱id() As Long
        
    Get医嘱id = mlng医嘱id
        
End Property

Private Sub SetErr(lngErrNum As Long, strErr As String)
    '设置错误描述及错误号
    '如果lngErrNum=-1 表示 控件自己定义的错误
    mReturnErrnumber = lngErrNum
    mReturnErrDescription = strErr
End Sub

Public Property Get ReturnErrNumber() As Long
    '返回最后一次的错误号
    ReturnErrNumber = mReturnErrnumber
End Property

Public Property Get ReturnErrDescription() As String
    '返回最后一次错误描述字符串
    ReturnErrDescription = mReturnErrDescription
End Property

'-------------------------------------------------------------------------------------------------------------------
Private Function CheckStrValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByRef strError As String) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
        
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        strError = "所输入内容含有非法字符。"
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            strError = "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。"
            Exit Function
        End If
    End If
    
    CheckStrValid = True
End Function

Private Function PopSelect(ByVal objBill As BillEdit, Optional ByVal str性别 As String = "0") As Boolean
    '----------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim sglX As Single
    Dim sglY As Single
    Dim strNote As String
    Dim strLvw As String
           
    On Error GoTo errHand
    
    CalcPosition sglX, sglY, objBill
    
    Select Case objBill.Col
    Case 1
        gstrSql = "SELECT ID," & _
                        "上级ID," & _
                        "0 AS 末级," & _
                        "编码," & _
                        "名称 " & _
                "FROM 疾病诊断分类 " & _
                "START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "B.分类id AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称 " & _
                "FROM 疾病诊断目录 A,疾病诊断属类 B " & _
                "WHERE A.ID=B.诊断ID "
                    
        strNote = "请选择一个疾病诊断项目"
        strLvw = "编码,1200,0,1;名称,2400,0,2"
    Case 0
        gstrSql = "SELECT ID," & _
                        "上级ID," & _
                        "0 AS 末级," & _
                        "NULL AS 编码," & _
                        "名称," & _
                        "NULL AS 简码 " & _
                "FROM 疾病编码分类 " & _
                "WHERE 类别='D' " & _
                "START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.分类id AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称, " & _
                        "A.简码 " & _
                "FROM 疾病编码目录 A " & _
                "WHERE 类别='D' " & _
                    "AND DECODE(性别限制,'男',1,'女',2,0) IN (" & str性别 & ") "
                    
        strNote = "请选择一个疾病编码项目"
        strLvw = "编码,1200,0,1;名称,2400,0,2;简码,810,0,0"
    End Select
    
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要"
    
    If rs.BOF Then Exit Function
    
    If frmSelectTree.ShowSelect(Screen, _
                                rs, _
                                sglX, sglY, 5400, 2400, _
                                objBill.MsfObj.CellHeight, _
                                "诊断提示_2", _
                                strLvw, _
                                strNote) Then
        
        PopSelect = True
        
        objBill.Text = zlCommFun.Nvl(rs("编码").Value)
        
        Select Case objBill.Col
        Case 0
            objBill.TextMatrix(objBill.Row, 0) = objBill.Text
            objBill.TextMatrix(objBill.Row, 4) = zlCommFun.Nvl(rs("ID").Value)
        Case 1
            objBill.TextMatrix(objBill.Row, 1) = objBill.Text
            objBill.TextMatrix(objBill.Row, 3) = zlCommFun.Nvl(rs("ID").Value)
        End Select
        objBill.TextMatrix(objBill.Row, 2) = zlCommFun.Nvl(rs("名称").Value)
        
        objBill.RowData(objBill.Row) = "1"
        
        '搜索对应的疾病诊断目录或疾病编码
        MatchDiagnoses Val(objBill.TextMatrix(objBill.Row, 4)), Val(objBill.TextMatrix(objBill.Row, 3)), objBill
                
    End If
   
    Exit Function
   
errHand:
   If ErrCenter = 1 Then Resume
End Function

Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As BillEdit)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.MsfObj.hwnd, objPoint)
    
    x = objPoint.x * 15 + objBill.MsfObj.CellLeft - 45
    y = objPoint.y * 15 + objBill.MsfObj.CellTop + objBill.MsfObj.CellHeight - 30
End Sub

Private Sub MatchDiagnoses(ByVal lngCodeKey As Long, ByVal lngListKey As Long, objMsf As BillEdit, Optional ByVal strCaption As String)
    Dim rs As New ADODB.Recordset
    '----------------------------------------------------------------------
    '1.知道疾病编码，求对应的疾病诊断目录
    '2.知道疾病诊断目录，求对应的疾病编码
    '----------------------------------------------------------------------
    gstrSql = "SELECT A.疾病ID,A.诊断ID,B.名称 AS 疾病编码,C.名称 AS 疾病诊断 " & _
                "FROM 疾病诊断对照 A,疾病编码目录 B,疾病诊断目录 C " & _
                "WHERE A.疾病ID=B.ID AND A.诊断ID=C.ID AND (A.疾病ID=" & lngCodeKey & " OR A.诊断ID=" & lngListKey & ")"
                
    zlDatabase.OpenRecordset rs, gstrSql, strCaption
    If rs.BOF = False Then
        If rs.RecordCount > 0 Then
            objMsf.TextMatrix(objMsf.Row, 0) = zlCommFun.Nvl(rs("疾病诊断"))
            objMsf.TextMatrix(objMsf.Row, 1) = zlCommFun.Nvl(rs("疾病编码"))
            objMsf.TextMatrix(objMsf.Row, 2) = zlCommFun.Nvl(rs("诊断ID"))
            objMsf.TextMatrix(objMsf.Row, 3) = zlCommFun.Nvl(rs("疾病ID"))
        End If
    End If
End Sub

Private Sub ReDimArray(ByRef LngCount As Long, ByRef strArray() As String)
    
    '功能：重新定义数组
    LngCount = LngCount + 1
    ReDim Preserve strArray(1 To LngCount)
        
End Sub

Private Function ShowDownList2(ByVal frmMain As Object, _
                            ByVal bytMode As Byte, _
                            objMsf As BillEdit, _
                            ByVal x As Single, _
                            ByVal y As Single, _
                            Optional ByVal blnWhere As Boolean = False) As Boolean
    '----------------------------------------------------------------------
    '功能:显示输入提示对话框
    '参数:
    '返回:
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strLvw As String
    Dim strInput As String
    Dim sglWidth As Single
    Dim sglHeight As Single
    Dim strPath As String
        
    On Error GoTo errHand
    
    If InStr(objMsf.Text, "'") > 0 Then
        MsgBox "内容中有非法字符！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strInput = "'%" & objMsf.Text & "%'"
                
    Select Case bytMode
    Case 1
        gstrSql = "SELECT   编码," & _
                           "名称," & _
                           "简码," & _
                           "附码," & _
                           "ID " & _
                    "FROM 疾病编码目录 " & _
                    "WHERE 类别='D' " & _
                        "AND DECODE(性别限制,'男',1,'女',2,0) IN (" & mstr性别 & ") " & _
                        "AND (编码 LIKE " & strInput & " OR 名称 LIKE " & strInput & " OR 简码 LIKE " & strInput & ")"
        sglWidth = 5100
        sglHeight = 2400
        strLvw = "编码,1200,0,0;名称,2400,0,0;简码,900,0,0;附码,900,0,0"
        strPath = "手术诊断_编码"
    Case 2
        gstrSql = "SELECT A.编码," & _
                           "A.名称," & _
                           "A.ID " & _
                    "FROM 疾病诊断目录 A " & _
                    "Where A.类别 = 1 " & _
                          "AND (编码 LIKE " & strInput & " OR 名称 LIKE " & strInput & " " & _
                          "OR A.id IN (SELECT B.诊断id " & _
                                        "FROM 疾病诊断别名 B " & _
                                        "WHERE 1=1 " & _
                                            "AND (名称 LIKE " & strInput & " OR 简码 LIKE " & strInput & ")))"
        sglWidth = 4500
        sglHeight = 2400
        strLvw = "编码,1200,0,0;名称,3000,0,0"
        strPath = "手术诊断_目录"
    End Select
            
    Call zlDatabase.OpenRecordset(rs, gstrSql, "手术概要")
    If rs.BOF Then
        If blnWhere Then objMsf.Text = ""
        Exit Function
    End If
    If blnWhere And rs.RecordCount = 1 Then
        ShowDownList2 = True
        GoTo FillPoint
        Exit Function
    End If
        
    If frmSelectList.ShowSelect(Screen, rs, strLvw, x, y, sglWidth, sglHeight, "手术概要\" & strPath, "请从下面选择一个项目") Then
        ShowDownList2 = True
        GoTo FillPoint
    Else
        If blnWhere Then objMsf.Text = ""
    End If
    
    Exit Function
    
FillPoint:

    objMsf.Text = zlCommFun.Nvl(rs("编码"))
    Select Case bytMode
    Case 1
        objMsf.TextMatrix(objMsf.Row, 0) = objMsf.Text
        objMsf.TextMatrix(objMsf.Row, 4) = zlCommFun.Nvl(rs("ID"))
    Case 2
        objMsf.TextMatrix(objMsf.Row, 1) = objMsf.Text
        objMsf.TextMatrix(objMsf.Row, 3) = zlCommFun.Nvl(rs("ID"))
    End Select
    objMsf.TextMatrix(objMsf.Row, 2) = zlCommFun.Nvl(rs("名称"))
    objMsf.RowData(objMsf.Row) = "1"
    
    '搜索对应的疾病诊断目录或疾病编码
    Call MatchDiagnoses(Val(objMsf.TextMatrix(objMsf.Row, 4)), Val(objMsf.TextMatrix(objMsf.Row, 3)), objMsf, "手术概要卡")
            
    Exit Function
errHand:
    objMsf.Text = ""
End Function

Private Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'功能: 装载数据入指定的组合下拉框或网格中的下拉框中
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Private Function PopOperateSelect(ByVal objBill As BillEdit, ByVal bytMode As Byte) As Boolean
    '----------------------------------------------------------------------
    '功能:
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim sglX As Single
    Dim sglY As Single
    Dim strNote As String
    Dim strLvw As String
    Dim str性别 As String
           
    On Error GoTo errHand
    
    
    '查询出性别
    str性别 = "0,1,2"
    If mlng病历id > 0 Then
        gstrSql = "SELECT B.性别 FROM 病人病历记录 A,病人信息 B WHERE A.病人ID=B.病人ID  and A.ID=" & mlng病历id
    Else
        gstrSql = "SELECT B.性别 FROM 病人医嘱记录 A,病人信息 B WHERE A.病人ID=B.病人ID  and A.ID=" & mlng医嘱id
    End If
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要专用纸"
    If rs.BOF = False Then
        If zlCommFun.Nvl(rs("性别").Value, "") Like "*男*" Then str性别 = "0,1"
        If zlCommFun.Nvl(rs("性别").Value, "") Like "*女*" Then str性别 = "0,2"
    End If
    
    CalcPosition sglX, sglY, objBill
    
    Select Case bytMode
    Case 1          '诊疗项目
        gstrSql = "SELECT ID," & _
                        "上级ID," & _
                        "0 AS 末级," & _
                        "编码," & _
                        "名称," & _
                        "NULL AS 单位 " & _
                "FROM 诊疗分类目录 " & _
                "WHERE 类型=5 " & _
                "START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.分类id AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称, " & _
                        "A.计算单位 AS 单位 " & _
                "FROM 诊疗项目目录 A "
        gstrSql = gstrSql & _
                "WHERE (撤档时间 = TO_DATE('30000101', 'YYYYMMDD') OR 撤档时间 IS NULL) " & _
                    "AND 服务对象 IN (2, 3) " & _
                    "AND 类别 = 'F' " & _
                    "AND NVL(适用性别,0) IN (" & str性别 & ")"
                    
        strNote = "请选择一个手术诊疗项目"
        strLvw = "编码,1200,0,1;名称,2400,0,2;单位,900,0,0"
    Case 2
        gstrSql = "SELECT ID," & _
                        "上级ID," & _
                        "0 AS 末级," & _
                        "NULL AS 编码," & _
                        "名称," & _
                        "NULL AS 简码 " & _
                "FROM 疾病编码分类 " & _
                "WHERE 类别='D' " & _
                "START WITH 上级ID is NULL CONNECT BY PRIOR ID = 上级ID " & _
                "UNION ALL " & _
                "SELECT A.ID, " & _
                        "A.分类id AS 上级ID, " & _
                        "1 AS 末级, " & _
                        "A.编码, " & _
                        "A.名称, " & _
                        "A.简码 " & _
                "FROM 疾病编码目录 A " & _
                "WHERE 类别='S' " & _
                    "AND DECODE(性别限制,'男',1,'女',2,0) IN (" & str性别 & ") "
                    
        strNote = "请选择一个疾病编码项目"
        strLvw = "编码,1200,0,1;名称,2400,0,2;简码,810,0,0"
    End Select
    
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要"
    
    If rs.BOF Then Exit Function
    
    If frmSelectTree.ShowSelect(Screen, _
                                rs, _
                                sglX, _
                                sglY, _
                                9000, _
                                3000, _
                                objBill.MsfObj.CellHeight, _
                                "手术提示_2", _
                                strLvw, _
                                strNote) Then
        
        If CheckHave(objBill, objBill.Row, zlCommFun.Nvl(rs("ID").Value)) Then
            MsgBox "在列表中已经存在此项目[" & zlCommFun.Nvl(rs("名称").Value) & "]！", vbInformation, gstrSysName
            Exit Function
        End If
        
        PopOperateSelect = True
        
        objBill.Text = zlCommFun.Nvl(rs("名称").Value)
        objBill.TextMatrix(objBill.Row, 1) = objBill.Text
        objBill.RowData(objBill.Row) = zlCommFun.Nvl(rs("ID").Value)

    End If
   
    Exit Function
   
errHand:
   If ErrCenter = 1 Then Resume
End Function

Private Function ShowDownListOperate(ByVal frmMain As Object, ByVal bytMode As Byte, objMsf As BillEdit, Optional ByVal blnWhere As Boolean = False, Optional ByVal str性别 As String = "0") As Boolean
    '----------------------------------------------------------------------
    '功能:显示输入提示对话框
    '参数:
    '返回:
    '----------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strLvw As String
    Dim strInput As String
    Dim sglWidth As Single
    Dim sglHeight As Single
    Dim strPath As String
    Dim sglY As Single
    Dim sglX As Single
        
    On Error GoTo errHand
    
    If InStr(objMsf.Text, "'") > 0 Then
        MsgBox "内容中有非法字符！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strInput = "'%" & objMsf.Text & "%'"
                
    Select Case bytMode
    Case 2
        gstrSql = "SELECT   编码," & _
                           "名称," & _
                           "简码," & _
                           "附码," & _
                           "ID " & _
                    "FROM 疾病编码目录 " & _
                    "WHERE 类别='S' " & _
                        "AND DECODE(性别限制,'男',1,'女',2,0) IN (" & str性别 & ") " & _
                        "AND (编码 LIKE " & strInput & " OR 名称 LIKE " & strInput & " OR 简码 LIKE " & strInput & ")"
        sglWidth = 5100
        sglHeight = 2400
        strLvw = "编码,1200,0,0;名称,2400,0,0;简码,900,0,0;附码,900,0,0"
        strPath = "手术_疾病"
    Case 1
        gstrSql = "SELECT A.编码," & _
                           "A.名称," & _
                           "A.ID " & _
                    "FROM 诊疗项目目录 A " & _
                    "Where A.类别 = 'F' " & _
                          "AND A.服务对象 IN (2, 3) " & _
                          "AND NVL(适用性别,0) IN (" & str性别 & ") " & _
                          "AND (编码 LIKE " & strInput & " OR 名称 LIKE " & strInput & " " & _
                          "OR A.id IN (SELECT B.诊疗项目ID " & _
                                        "FROM 诊疗项目别名 B " & _
                                        "WHERE (名称 LIKE " & strInput & " OR 简码 LIKE " & strInput & ")))"
        sglWidth = 9000
        sglHeight = 3000
        strLvw = "编码,1200,0,0;名称,3000,0,0"
        strPath = "手术_诊疗"
    End Select
            
    Call zlDatabase.OpenRecordset(rs, gstrSql, "手术概要")
    If rs.BOF Then
        If blnWhere Then objMsf.Text = ""
        Exit Function
    End If
    If blnWhere And rs.RecordCount = 1 Then
        ShowDownListOperate = True
        GoTo FillPoint
        Exit Function
    End If
    
    CalcPosition sglX, sglY, objMsf
    
    If frmSelectList.ShowSelect(Screen, rs, strLvw, sglX, sglY, sglWidth, sglHeight, "手术概要\" & strPath, "请从下面选择一个项目") Then
        ShowDownListOperate = True
        GoTo FillPoint
    Else
        If blnWhere Then objMsf.Text = ""
    End If
    
    Exit Function
    
FillPoint:
    If CheckHave(objMsf, objMsf.Row, zlCommFun.Nvl(rs("ID").Value)) Then
        MsgBox "在列表中已经存在此项目[" & zlCommFun.Nvl(rs("名称").Value) & "]！", vbInformation, gstrSysName
        Exit Function
    End If
    
    objMsf.Text = zlCommFun.Nvl(rs("名称"))
    objMsf.TextMatrix(objMsf.Row, 1) = objMsf.Text
    objMsf.RowData(objMsf.Row) = zlCommFun.Nvl(rs("ID"))
    
    Exit Function
errHand:
    objMsf.Text = ""
End Function

Public Function ShowDownListPerson(ByVal frmMain As Object, _
                                objMsf As BillEdit, _
                                ByVal lngDeptKey As Long, _
                                ByVal x As Single, _
                                ByVal y As Single, _
                                Optional ByVal blnWhere As Boolean = False, _
                                Optional ByVal blnFlag As Boolean = False, _
                                Optional ByVal lngKey As Long = 0) As Boolean
    '----------------------------------------------------------------------
    '功能:显示输入提示对话框
    '参数:
    '返回:
    '----------------------------------------------------------------------
    Dim strInput As String
    Dim strSelected As String
    Dim lngLoop As Long
    Dim strClass As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    If InStr(objMsf.Text, "'") > 0 Then
        MsgBox "内容中有非法字符！", vbInformation, gstrSysName
        Exit Function
    End If
    
    strInput = "'%" & objMsf.Text & "%'"
                        
    strSelected = "0"
    For lngLoop = 1 To objMsf.Rows - 1
        If lngLoop <> objMsf.Row Then
            strSelected = strSelected & "," & objMsf.RowData(lngLoop)
        End If
    Next
    
    strClass = "医生"
    If InStr("洗手护士;巡回护士", Trim(objMsf.TextMatrix(objMsf.Row, 0))) > 0 Then strClass = "护士"
        
    gstrSql = "SELECT   A.编号," & _
                       "A.姓名," & _
                       "A.简码," & _
                       "C.名称 AS 科室," & _
                       "DECODE(C.ID," & lngDeptKey & ",1,2) AS 序号," & _
                       "A.ID " & _
                "FROM 人员表 A,人员性质说明 B,部门表 C,部门人员 D " & _
                "WHERE A.ID=B.人员id AND C.ID=D.部门id AND D.人员id=A.ID AND D.缺省=1 " & _
                    "AND A.ID NOT IN (" & strSelected & ") " & _
                    "AND B.人员性质='" & strClass & "' " & _
                    "AND (A.编号 LIKE " & strInput & " OR A.姓名 LIKE " & strInput & " OR A.简码 LIKE " & strInput & ") " & _
                "ORDER BY 序号"
    
    Call zlDatabase.OpenRecordset(rs, gstrSql, "手术概要")
    If rs.BOF Then
        If blnWhere Then objMsf.Text = ""
        Exit Function
    End If
    If blnWhere And rs.RecordCount = 1 Then
        ShowDownListPerson = True
        GoTo FillPoint
    End If
    
    If frmSelectList.ShowSelect(Screen, rs, "编号,1200,0,0;姓名,2400,0,0;简码,900,0,0;科室,900,0,0", x, y, 8100, 3000, "手术概要" & "手术_人员", "请从下表中选择一个人员") Then
        ShowDownListPerson = True
        GoTo FillPoint
    Else
        If blnWhere Then objMsf.Text = ""
    End If
    
    Exit Function
    
FillPoint:
    objMsf.Text = zlCommFun.Nvl(rs("姓名"))
    objMsf.TextMatrix(objMsf.Row, 1) = objMsf.Text
    objMsf.TextMatrix(objMsf.Row, 2) = zlCommFun.Nvl(rs("编号"))
    objMsf.RowData(objMsf.Row) = zlCommFun.Nvl(rs("ID"))
    
    Exit Function
errHand:
    objMsf.Text = ""
End Function
'------------------------------------------------------------------------------------------------------------

Private Sub ShowOperGeneral(lng病历ID As Long, Optional ByVal blnEditMode As Boolean = False)
    '------------------------------------------------------------------------------------------------------------
    '功能：外部调用显示手术概要的过程
    '------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
    
    mlng病历id = lng病历ID
    mDispMode = Not blnEditMode
    
    mstr性别 = "0,1,2"
    
    '按逻辑应先初始控件
    InitData
    
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Sub

    '检查那份病历是不是存在
    strSQL = _
        "SELECT a.ID" & vbCrLf & _
        "  FROM 病人病历内容 A" & vbCrLf & _
        " WHERE a.元素类型 = 4 and " & vbCrLf & _
        "      a.元素编码 IN" & vbCrLf & _
        "      (SELECT 编码" & vbCrLf & _
        "         FROM 病历元素目录" & vbCrLf & _
        "        WHERE 类型 = 4 AND 名称 = '手术概要记录卡')" & vbCrLf & _
        " AND A.id=" & mlng病历id
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "手术概要记录卡")
    
    If rsTmp.RecordCount = 0 And mlng医嘱id = 0 Then
        SetErr -1, "该病历不存在无调用手术概要记录卡！"
'        Exit Sub
    End If
    
    Call ReadData
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function LocalCheck是否非法(txt As Control, ByVal strLawlChar As String) As Boolean
'功能:检查是不是包含strLawlChar里的字符串,如果有就返回为真否则就返回否
On Error GoTo ErrHandle
    Dim strSour As String
    
    If TypeOf txt Is TextBox Or TypeOf txt Is ComboBox Then
        If TypeOf txt Is ComboBox Then
            If txt.Style <> 0 Then
                '不管ComboBox为选择的情况，只管输入的情况
                LocalCheck是否非法 = True
                Exit Function
            End If
        End If
        strSour = txt.Text
        If Len(strSour) > 0 Then
            For i = 1 To Len(strLawlChar)
                If InStr(strSour, Mid(strLawlChar, i, 1)) > 0 Then
                    txt.SelStart = InStr(strSour, Mid(strLawlChar, i, 1))
                    txt.SelLength = 1
                    MsgBox "文本里包含有非法字符！", vbInformation, gstrSysName
                    LocalCheck是否非法 = True
                    Exit Function
                End If
            Next
            If VarType(txt.Tag) = vbLong Or VarType(txt.Tag) = vbInteger Then
                If zlCommFun.ActualLen(strSour) > txt.Tag And txt.Tag > 0 Then
                    MsgBox "您所输入的文本超长！", vbInformation, gstrSysName
                    LocalCheck是否非法 = True
                End If
            ElseIf VarType(txt.Tag) = vbString And IsNumeric(txt.Tag) Then
                If zlCommFun.ActualLen(strSour) > CLng(txt.Tag) And CLng(txt.Tag) > 0 Then
                    MsgBox "您所输入的文本超长！", vbInformation, gstrSysName
                    LocalCheck是否非法 = True
                End If
            End If
        End If
    End If
    Exit Function
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub SetgcnOracle()
    '-------------------------------------------------------------------------------------------------
    '接口
    '-------------------------------------------------------------------------------------------------
    
    Call InitCommon(gcnOracle)
End Sub

Private Sub InitData()
    '初始化窗体
    
    Dim strTmp As String
    
    On Error GoTo ErrHandle
        
    If Not gcnOracle Is Nothing Then
        If Not gcnOracle.State <> adStateOpen Then
            If Ambient.UserMode = True Then
                strSQL = "select * FROM 诊疗项目目录 where 类别 ='G'"
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, "手术概要记录卡")
                If rsTmp.RecordCount > 0 Then
                    rsTmp.MoveFirst
                    For i = 0 To rsTmp.RecordCount - 1
                        cbo(0).AddItem rsTmp("编码") & "-" & rsTmp("名称") & Space(200) & zlCommFun.Nvl(rsTmp("操作类型"))
                        rsTmp.MoveNext
                    Next
                    cbo(0).ListIndex = 0
                End If
            End If
        End If
    
    End If
    With cbo(1)
        .Clear
        .AddItem "1-优"
        .AddItem "2-佳"
        .AddItem "3-劣"
        .AddItem "4-危(急)"
        .ListIndex = 0
    End With
    
    strTmp = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59"
    
    dtp(0).Value = strTmp
    dtp(1).Value = strTmp
    dtp(2).Value = strTmp
    dtp(3).Value = strTmp
    dtp(4).Value = strTmp
    dtp(5).Value = strTmp
    
    mblnLoaded = True
    
    Exit Sub
    
ErrHandle:

    If Ambient.UserMode = False Or InDesign = False Then
        SetErr Err.Number, Err.Description
        Exit Sub
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------
    '功能：读出数据库里的数据
    '------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
        
    On Error GoTo ErrHandle
    
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Function
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Function
    
    gstrSql = "SELECT 名称,0 FROM 诊疗手术规模"
    Call zlDatabase.OpenRecordset(rs, gstrSql, "手术概要")
    If rs.RecordCount = 0 Then
        MsgBox "系统数据不完整，没有诊疗手术规模基础字典！", vbInformation, gstrSysName
        Exit Function
    End If
    If rs.BOF = False Then Call AddComboData(cbo(2), rs)
    
    
    mlng手术记录id = 0
        
    If mlng医嘱id = 0 Then
        strTmp = "SELECT A.医嘱id FROM 病人病历记录 A,病人病历内容 B WHERE B.病历记录id=A.ID AND B.ID=" & mlng病历id
        zlDatabase.OpenRecordset rs, strTmp, "手术概要卡"
        If rs.BOF = False Then mlng医嘱id = zlCommFun.Nvl(rs("医嘱id").Value, 0)
    End If
        
    If mlng医嘱id > 0 Then
        '表明是从手术麻醉系统调用
        
        '通过医嘱id查找病人手术id
        strTmp = "SELECT ID FROM 病人手术记录 WHERE 医嘱id=" & mlng医嘱id
        zlDatabase.OpenRecordset rs, strTmp, "手术概要卡"
        If rs.BOF = False Then mlng手术记录id = zlCommFun.Nvl(rs("ID").Value, 0)
                    
    Else
        '表明是从病历工作站调用，就要确定是新增还是修改
        
        strTmp = "SELECT ID FROM 病人手术记录 WHERE 病历id=" & mlng病历id
        zlDatabase.OpenRecordset rs, strTmp, "手术概要卡"
        If rs.BOF = False Then mlng手术记录id = zlCommFun.Nvl(rs("ID").Value, 0)
        
    End If
    
    strTmp = "SELECT 医嘱id,手术间,手术室id FROM 病人手术记录 WHERE ID=" & mlng手术记录id
    zlDatabase.OpenRecordset rs, strTmp, "手术概要卡"
    If rs.BOF = False Then
        mlng手术间id = zlCommFun.Nvl(rs("手术室id").Value, 0)
        mlngOrderID = zlCommFun.Nvl(rs("医嘱id").Value, 0)
    End If
    
    '1.读取手术基本资料
    gstrSql = "SELECT A.* FROM 病人手术记录 A WHERE A.ID=" & mlng手术记录id
    
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要卡"
    If rs.BOF = False Then
        dtp(0).Value = Format(zlCommFun.Nvl(rs("手术开始时间")), "YYYY-MM-DD HH:MM")
        dtp(1).Value = Format(zlCommFun.Nvl(rs("手术结束时间")), "YYYY-MM-DD HH:MM")
        If IsNull(rs("麻醉开始时间")) = False Then
            chk(0).Value = 1
            dtp(2).Value = Format(zlCommFun.Nvl(rs("麻醉开始时间")), "YYYY-MM-DD HH:MM")
            dtp(3).Value = Format(zlCommFun.Nvl(rs("麻醉结束时间")), "YYYY-MM-DD HH:MM")
        Else
            dtp(2).Value = Format(zlCommFun.Nvl(rs("手术开始时间")), "YYYY-MM-DD HH:MM")
            dtp(3).Value = Format(zlCommFun.Nvl(rs("手术结束时间")), "YYYY-MM-DD HH:MM")
        End If
        If IsNull(rs("输氧开始时间")) = False Then
            chk(1).Value = 1
            dtp(4).Value = Format(zlCommFun.Nvl(rs("输氧开始时间")), "YYYY-MM-DD HH:MM")
            dtp(5).Value = Format(zlCommFun.Nvl(rs("输氧结束时间")), "YYYY-MM-DD HH:MM")
        Else
            dtp(4).Value = Format(zlCommFun.Nvl(rs("手术开始时间")), "YYYY-MM-DD HH:MM")
            dtp(5).Value = Format(zlCommFun.Nvl(rs("手术结束时间")), "YYYY-MM-DD HH:MM")
        End If
                        
                        
        zlControl.CboLocate cbo(0), zlCommFun.Nvl(rs("麻醉方式"))
        zlControl.CboLocate cbo(1), zlCommFun.Nvl(rs("麻醉质量"))
        zlControl.CboLocate cbo(2), zlCommFun.Nvl(rs("手术规模"))
        txt(10).Text = zlCommFun.Nvl(rs("手术间"))
        txt(0).Text = zlCommFun.Nvl(rs("输液总量"))
    End If
    
    bill(0).Rows = 2
    bill(1).Rows = 2
    bill(2).Rows = 2
    bill2(0).Rows = 2
    bill2(1).Rows = 2
    ClearSpecRowCol bill(0), 1, Array()
    ClearSpecRowCol bill(1), 1, Array()
    ClearSpecRowCol bill(2), 1, Array()
    ClearSpecRowCol bill2(0), 1, Array()
    ClearSpecRowCol bill2(1), 1, Array()
    
    '2.读取拟行手术记录
    gstrSql = "SELECT DECODE(A.诊疗项目ID,null,'2-疾病','1-诊疗') AS 手术来源," & _
                    "A.手术名称," & _
                    "A.缺省," & _
                    "DECODE(A.诊疗项目id,NULL,A.手术操作ID,A.诊疗项目id) AS ID " & _
                "FROM 病人手术情况 A,病人手术记录 B " & _
                "WHERE A.记录id=B.ID " & _
                        "AND A.性质=1 " & _
                        "AND B.ID=" & mlng手术记录id
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要卡"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill(0).RowData(1) > 0 Then bill(0).Rows = bill(0).Rows + 1
            
            bill(0).RowData(bill(0).Rows - 1) = zlCommFun.Nvl(rs("ID").Value, 0)
            bill(0).TextMatrix(bill(0).Rows - 1, 0) = zlCommFun.Nvl(rs("手术来源").Value)
            bill(0).TextMatrix(bill(0).Rows - 1, 1) = zlCommFun.Nvl(rs("手术名称").Value)
            bill(0).TextMatrix(bill(0).Rows - 1, 2) = IIf(zlCommFun.Nvl(rs("缺省").Value) = 1, "√", "")
            
            rs.MoveNext
        Loop
    End If
            
    
    '3.读取已行手术记录
    gstrSql = "SELECT DECODE(A.诊疗项目ID,null,'2-疾病','1-诊疗') AS 手术来源," & _
                    "A.手术名称," & _
                    "A.缺省," & _
                    "DECODE(A.诊疗项目id,NULL,A.手术操作ID,A.诊疗项目id) AS ID " & _
                "FROM 病人手术情况 A,病人手术记录 B " & _
                "WHERE A.记录id=B.ID " & _
                        "AND A.性质=2 " & _
                        "AND B.ID=" & mlng手术记录id
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要卡"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill(1).RowData(1) > 0 Then bill(1).Rows = bill(1).Rows + 1
            
            bill(1).RowData(bill(1).Rows - 1) = zlCommFun.Nvl(rs("ID").Value, 0)
            bill(1).TextMatrix(bill(1).Rows - 1, 0) = zlCommFun.Nvl(rs("手术来源").Value)
            bill(1).TextMatrix(bill(1).Rows - 1, 1) = zlCommFun.Nvl(rs("手术名称").Value)
            bill(1).TextMatrix(bill(1).Rows - 1, 2) = IIf(zlCommFun.Nvl(rs("缺省").Value) = 1, "√", "")
            
            rs.MoveNext
        Loop
    End If
    If bill(1).RowData(1) = 0 Then CopyMsfGrid bill(0), bill(1)
        
    '3.读取术前诊断记录
    If mlng医嘱id > 0 Then
        gstrSql = "select 诊断ID," & _
                          "疾病ID," & _
                          "(select 编码 FROM 疾病诊断目录 where id = 诊断ID) AS 诊断编码," & _
                          "(select 编码 FROM 疾病编码目录 where id = 疾病ID) AS 疾病编码," & _
                          "诊断描述 " & _
                     "From 病人诊断记录 " & _
                    "where 医嘱id = " & mlng医嘱id & " and 诊断类型 = 8"
    Else
        gstrSql = "select 诊断ID," & _
                          "疾病ID," & _
                          "(select 编码 FROM 疾病诊断目录 where id = 诊断ID) AS 诊断编码," & _
                          "(select 编码 FROM 疾病编码目录 where id = 疾病ID) AS 疾病编码," & _
                          "诊断描述 " & _
                     "From 病人诊断记录 " & _
                    "where 病历id = " & mlng病历id & " and 诊断类型 = 8"
    End If
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要卡"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill2(0).RowData(1) > 0 Then bill2(0).Rows = bill2(0).Rows + 1
            
            bill2(0).RowData(bill2(0).Rows - 1) = "1"
            bill2(0).TextMatrix(bill2(0).Rows - 1, 1) = zlCommFun.Nvl(rs("诊断编码").Value)
            bill2(0).TextMatrix(bill2(0).Rows - 1, 0) = zlCommFun.Nvl(rs("疾病编码").Value)
            bill2(0).TextMatrix(bill2(0).Rows - 1, 2) = zlCommFun.Nvl(rs("诊断描述").Value)
            bill2(0).TextMatrix(bill2(0).Rows - 1, 3) = zlCommFun.Nvl(rs("诊断ID").Value)
            bill2(0).TextMatrix(bill2(0).Rows - 1, 4) = zlCommFun.Nvl(rs("疾病ID").Value)
            
            rs.MoveNext
        Loop
    End If
    
     '3.读取术后诊断记录
    If mlng医嘱id > 0 Then
        gstrSql = "select 诊断ID," & _
                          "疾病ID," & _
                          "(select 编码 FROM 疾病诊断目录 where id = 诊断ID) AS 诊断编码," & _
                          "(select 编码 FROM 疾病编码目录 where id = 疾病ID) AS 疾病编码," & _
                          "诊断描述 " & _
                     "From 病人诊断记录 " & _
                    "where 医嘱id = " & mlng医嘱id & " and 诊断类型 = 9"
    Else
        gstrSql = "select 诊断ID," & _
                          "疾病ID," & _
                          "(select 编码 FROM 疾病诊断目录 where id = 诊断ID) AS 诊断编码," & _
                          "(select 编码 FROM 疾病编码目录 where id = 疾病ID) AS 疾病编码," & _
                          "诊断描述 " & _
                     "From 病人诊断记录 " & _
                    "where 病历id = " & mlng病历id & " and 诊断类型 = 9"
    End If
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要卡"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill2(1).RowData(1) > 0 Then bill2(1).Rows = bill2(1).Rows + 1
            
            bill2(1).RowData(bill2(1).Rows - 1) = "1"
            bill2(1).TextMatrix(bill2(1).Rows - 1, 1) = zlCommFun.Nvl(rs("诊断编码").Value)
            bill2(1).TextMatrix(bill2(1).Rows - 1, 0) = zlCommFun.Nvl(rs("疾病编码").Value)
            bill2(1).TextMatrix(bill2(1).Rows - 1, 2) = zlCommFun.Nvl(rs("诊断描述").Value)
            bill2(1).TextMatrix(bill2(1).Rows - 1, 3) = zlCommFun.Nvl(rs("诊断ID").Value)
            bill2(1).TextMatrix(bill2(1).Rows - 1, 4) = zlCommFun.Nvl(rs("疾病ID").Value)
            
            rs.MoveNext
        Loop
    End If
    If bill2(1).RowData(1) = 0 Then CopyMsfGrid bill2(0), bill2(1)
    
    '3.读取手术人员记录
    gstrSql = "SELECT A.人员id," & _
                    "DECODE(A.岗位,'主刀医生','1-主刀医生','助手医生',2,'麻醉医生',3,'洗手护士',4,5) AS 序号," & _
                    "D.名称 AS 岗位," & _
                    "A.姓名 " & _
                "FROM 病人手术人员 A,手术岗位 D " & _
                "WHERE  D.名称=A.岗位 " & _
                        "AND A.记录ID=" & mlng手术记录id & " " & _
                "ORDER BY 序号"
                
    zlDatabase.OpenRecordset rs, gstrSql, "手术概要卡"
    If rs.BOF = False Then
        Do While Not rs.EOF
            If bill(2).RowData(1) > 0 Then bill(2).Rows = bill(2).Rows + 1
            
            bill(2).RowData(bill(2).Rows - 1) = zlCommFun.Nvl(rs("人员id").Value, 0)
            bill(2).TextMatrix(bill(2).Rows - 1, 0) = zlCommFun.Nvl(rs("岗位").Value)
            bill(2).TextMatrix(bill(2).Rows - 1, 1) = zlCommFun.Nvl(rs("姓名").Value)
            
            rs.MoveNext
        Loop
    End If
    
    Exit Function
    
ErrHandle:
    
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function

Public Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
'功能: 清除指定网格的指定行指定列的数据
'参数: obj=要操作的网格控件
'      intRow=要清除的行号
'      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Private Sub CopyMsfGrid(ByVal objFrom As Object, ByRef objTo As Object)
    Dim lngRow As Long
    Dim lngCol As Long
    
    objTo.Rows = objFrom.Rows
    objTo.Cols = objFrom.Cols
    
    For lngRow = 1 To objFrom.Rows - 1
        objTo.RowData(lngRow) = objFrom.RowData(lngRow)
        For lngCol = 0 To objFrom.Cols - 1
            objTo.TextMatrix(lngRow, lngCol) = objFrom.TextMatrix(lngRow, lngCol)
        Next
    Next
End Sub

Private Function CheckDataValid(ByRef strError As String) As Boolean
    '----------------------------------------------------------------------
    '功能：对新增、修改的数据进行合法性校验
    '返回：校验合法返回True，否则返回False
    '----------------------------------------------------------------------
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    CheckDataValid = False
        
    strError = ""
    
    If mDispMode Then
        strError = "当前为显示模式不能保存数据！"
        SetErr -1, "当前为显示模式不能保存数据！"
        Exit Function
    End If
    
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Function
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Function
    
    If CheckStrValid(txt(0).Text, txt(0).MaxLength, strError) = False Then
        zlControl.TxtSelAll txt(0)
        txt(0).SetFocus
        Exit Function
    End If
    
    If CheckStrValid(txt(10).Text, txt(10).MaxLength, strError) = False Then
        zlControl.TxtSelAll txt(10)
        txt(10).SetFocus
        Exit Function
    End If
    
    If dtp(0).Value > dtp(1).Value Then
        strError = "手术开始时间不能大于手术结束时间！"
        dtp(0).SetFocus
        GoTo errHand
    End If
    
    If Abs(DateDiff("h", CDate(Format(dtp(0).Value, "YYYY-MM-DD HH:MM")), CDate(Format(dtp(1).Value, "YYYY-MM-DD HH:MM")))) > 12 Then
        strError = "手术开始时间和手术结束时间之间不能大于12小时！"
        dtp(0).SetFocus
        GoTo errHand
    End If
    
    
    If dtp(2).Value > dtp(3).Value And chk(0).Value = 1 Then
        strError = "麻醉开始时间不能大于麻醉结束时间！"
        dtp(2).SetFocus
        GoTo errHand
    End If
    
    If chk(0).Value = 1 And cbo(0).ListIndex = -1 Then
        strError = "必须指明麻醉方式！"
        cbo(0).SetFocus
        GoTo errHand
    End If
    
    If chk(0).Value = 1 And cbo(1).ListIndex = -1 Then
        strError = "必须指明麻醉质量！"
        cbo(1).SetFocus
        GoTo errHand
    End If
    
    If dtp(4).Value > dtp(5).Value And chk(1).Value = 1 Then
        strError = "输氧开始时间不能大于输氧结束时间！"
        dtp(4).SetFocus
        GoTo errHand
    End If
    
    If CheckAllNumber(txt(0).Text) = False Then
        strError = "输液总量必须为全数字！"
        
        zlControl.TxtSelAll txt(0)
        txt(0).SetFocus
        GoTo errHand
    End If
    
    Dim LngCount As Long
    
    LngCount = 0
    For lngLoop = 1 To bill(2).Rows - 1
        If bill(2).RowData(lngLoop) > 0 And InStr(bill(2).TextMatrix(lngLoop, 0), "主刀医生") > 0 Then
            LngCount = LngCount + 1
            If LngCount > 1 Then
                strError = "主刀医生只能一个！"
                bill(2).SetFocus
                GoTo errHand
            End If
            If LngCount > 2 Then Exit For
        End If
    Next
    If LngCount < 1 Then
        strError = " 必须指定手术的主刀医生！"
        bill(2).SetFocus
        GoTo errHand
    End If
        
    '检查手术名称是否有非法字符、超长、手术个数
    For lngIndex = 0 To 1
        For lngLoop = 1 To bill(lngIndex).Rows - 1
            If bill(lngIndex).RowData(lngLoop) > 0 Then
                Exit For
            End If
        Next
            
        If lngLoop = bill(lngIndex).Rows Then
            If lngIndex = 0 Then
                strError = "至少有一个拟行手术！"
            Else
                strError = "至少有一个已行手术！"
            End If
                        
            bill(lngIndex).SetFocus
            GoTo errHand
        End If
        
        For lngLoop = 1 To bill(lngIndex).Rows - 1
            If bill(lngIndex).RowData(lngLoop) > 0 Then
                If CheckStrValid(bill(lngIndex).TextMatrix(lngLoop, 1), 50, strError) = False Then
                    bill(lngIndex).Col = 1
                    bill(lngIndex).Row = lngLoop
                    bill(lngIndex).SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
    
    '检查诊断描述是否有非法字符、超长
    For lngIndex = 0 To 1
        For lngLoop = 1 To bill2(lngIndex).Rows - 1
            If bill2(lngIndex).RowData(lngLoop) > 0 Then
                If CheckStrValid(bill2(lngIndex).TextMatrix(lngLoop, 2), 100, strError) = False Then
                    bill2(lngIndex).Col = 2
                    bill2(lngIndex).Row = lngLoop
                    bill2(lngIndex).SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
               
    '检查人员编码、人员姓名是否有非法字符、超长
    For lngLoop = 1 To bill(2).Rows - 1
        If bill(2).RowData(lngLoop) > 0 Then
            If CheckStrValid(bill(2).TextMatrix(lngLoop, 1), 20, strError) = False Then
                bill(2).Col = 1
                bill(2).Row = lngLoop
                bill(2).SetFocus
                Exit Function
            End If
            
            If CheckStrValid(bill(2).TextMatrix(lngLoop, 2), 10, strError) = False Then
                bill(2).Col = 1
                bill(2).Row = lngLoop
                bill(2).SetFocus
                Exit Function
            End If
            
        End If
    Next
    
    CheckDataValid = True
    
    Exit Function
    
errHand:
    
End Function

Public Function SaveData(lng病人ID As Long, lng主页ID As Long, lng病历ID As Long, ReturnStrSQL As String, strError As String) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能：保存数据，这里只是形成SQL语句，需要在调用此专用单的窗体来执行
    '参数：
    '---------------------------------------------------------------------------------------------------------
    Dim str麻醉方式 As String
    Dim str麻醉类型 As String
    Dim strTmp As String
    Dim LngCount As Long
    Dim rs As New ADODB.Recordset
    
    Dim strSQL() As String
    Dim lngLoop As Long
    
    On Error GoTo ErrHandle
    
    '检查输入数据的有效性
    If CheckDataValid(strError) = False Then Exit Function
    
    ReDimArray LngCount, strSQL
    strSQL(LngCount) = "ZL_病人手术情况_DELETE(" & mlng手术记录id & ",1)"
    
    ReDimArray LngCount, strSQL
    strSQL(LngCount) = "ZL_病人手术情况_DELETE(" & mlng手术记录id & ",2)"
    
    If mlng医嘱id > 0 Then
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_病人诊断记录_DELETE2(" & mlng医嘱id & ",8)"
    
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_病人诊断记录_DELETE2(" & mlng医嘱id & ",9)"
        
        mlngOrderID = mlng医嘱id
    Else
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_病人诊断记录_DELETE(" & lng病人ID & "," & lng主页ID & ",1," & lng病历ID & ",'8,9')"
    End If
    
    ReDimArray LngCount, strSQL
    strSQL(LngCount) = "ZL_病人手术人员_CANCELPERSON(" & mlng手术记录id & ")"
    
    
    'ReDimArray lngCount, strSQL
    'strSQL(lngCount) = "ZL_病人手术记录_DELETE(" & mlng手术记录id & ")"
    
    '始终新增,因为在调用此接品前可能已删除
    
    If cbo(0).Text <> "" Then
        str麻醉方式 = Trim(Mid(cbo(0).Text, 1, InStr(cbo(0).Text, Space(200)) - 1))
        str麻醉类型 = Trim(Mid(cbo(0).Text, InStr(cbo(0).Text, Space(200)) + 200))
    End If
    
    If mlng手术记录id = 0 Then
        mlng手术记录id = zlDatabase.GetNextId("病人手术记录")
        
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_病人手术记录_INSERT(" & mlng手术记录id & "," & _
                                    lng病人ID & "," & _
                                    IIf(lng主页ID = 0, "NULL", lng主页ID) & "," & _
                                    IIf(mlngOrderID > 0, mlngOrderID, "Null") & "," & _
                                    lng病历ID & "," & _
                                    "TO_DATE('" & Format(dtp(0).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    "TO_DATE('" & Format(dtp(0).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    "TO_DATE('" & Format(dtp(1).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    IIf(chk(0).Value = 1, "TO_DATE('" & Format(dtp(2).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "TO_DATE('" & Format(dtp(3).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & zlCommFun.GetNeedName(str麻醉方式) & "'", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & str麻醉类型 & "'", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & zlCommFun.GetNeedName(cbo(1).Text) & "'", "NULL") & ",'" & _
                                    txt(0).Text & "'," & _
                                    IIf(chk(1).Value = 1, "TO_DATE('" & Format(dtp(4).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(1).Value = 1, "TO_DATE('" & Format(dtp(5).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & ",'" & _
                                    txt(10).Text & "'," & _
                                    mlng手术间id & ",'" & _
                                    cbo(2).Text & "')"
    Else
        ReDimArray LngCount, strSQL
        strSQL(LngCount) = "ZL_病人手术记录_UPDATE(" & mlng手术记录id & "," & _
                                    "TO_DATE('" & Format(dtp(0).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    "TO_DATE('" & Format(dtp(0).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    "TO_DATE('" & Format(dtp(1).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                    IIf(chk(0).Value = 1, "TO_DATE('" & Format(dtp(2).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "TO_DATE('" & Format(dtp(3).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & zlCommFun.GetNeedName(str麻醉方式) & "'", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & str麻醉类型 & "'", "NULL") & "," & _
                                    IIf(chk(0).Value = 1, "'" & zlCommFun.GetNeedName(cbo(1).Text) & "'", "NULL") & ",'" & _
                                    txt(0).Text & "'," & _
                                    IIf(chk(1).Value = 1, "TO_DATE('" & Format(dtp(4).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                    IIf(chk(1).Value = 1, "TO_DATE('" & Format(dtp(5).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & ",'" & _
                                    txt(10).Text & "'," & _
                                    mlng手术间id & ",'" & _
                                    cbo(2).Text & "'," & lng病历ID & ")"
    End If
    
    '填写拟行手术
    For lngLoop = 1 To bill(0).Rows - 1
        If bill(0).RowData(lngLoop) > 0 Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_病人手术情况_INSERT(" & mlng手术记录id & ",1," & IIf(bill(0).TextMatrix(lngLoop, 2) = "√", 1, 0) & ",'" & bill(0).TextMatrix(lngLoop, 1) & "'," & IIf(Val(Mid(bill(0).TextMatrix(lngLoop, 0), 1, 1)) = 1, "NULL," & bill(0).RowData(lngLoop), bill(0).RowData(lngLoop) & ",NULL") & ")"
        End If
    Next
    '填写已行手术
    For lngLoop = 1 To bill(1).Rows - 1
        If bill(1).RowData(lngLoop) > 0 Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_病人手术情况_INSERT(" & mlng手术记录id & ",2," & IIf(bill(1).TextMatrix(lngLoop, 2) = "√", 1, 0) & ",'" & bill(1).TextMatrix(lngLoop, 1) & "'," & IIf(Val(Mid(bill(1).TextMatrix(lngLoop, 0), 1, 1)) = 1, "NULL," & bill(1).RowData(lngLoop), bill(1).RowData(lngLoop) & ",NULL") & ")"
        End If
    Next
    
    '填写术前诊断
    For lngLoop = 1 To bill2(0).Rows - 1
        If bill2(0).RowData(lngLoop) > 0 And (Val(bill2(0).TextMatrix(lngLoop, 3)) > 0 Or Val(bill2(0).TextMatrix(lngLoop, 4)) > 0) Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_病人诊断记录_INSERT(" & lng病人ID & "," & IIf(lng主页ID = 0, "NULL", lng主页ID) & ",1," & lng病历ID & ",8," & Val(bill2(0).TextMatrix(lngLoop, 4)) & "," & Val(bill2(0).TextMatrix(lngLoop, 3)) & ",NULL,'" & bill2(0).TextMatrix(lngLoop, 2) & "',NULL,NULL,NULL,SYSDATE," & IIf(mlngOrderID = 0, "NULL", mlngOrderID) & ")"
        End If
    Next
    
    '填写术后诊断
    For lngLoop = 1 To bill2(1).Rows - 1
        If bill2(1).RowData(lngLoop) > 0 And (Val(bill2(1).TextMatrix(lngLoop, 3)) > 0 Or Val(bill2(1).TextMatrix(lngLoop, 4)) > 0) Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_病人诊断记录_INSERT(" & lng病人ID & "," & IIf(lng主页ID = 0, "NULL", lng主页ID) & ",1," & lng病历ID & ",9," & Val(bill2(1).TextMatrix(lngLoop, 4)) & "," & Val(bill2(1).TextMatrix(lngLoop, 3)) & ",NULL,'" & bill2(1).TextMatrix(lngLoop, 2) & "',NULL,NULL,NULL,SYSDATE," & IIf(mlngOrderID = 0, "NULL", mlngOrderID) & ")"
        End If
    Next
    
    '填写手术人员
    For lngLoop = 1 To bill(2).Rows - 1
        If bill(2).RowData(lngLoop) > 0 Then
            ReDimArray LngCount, strSQL
            strSQL(LngCount) = "ZL_病人手术记录_PERSON(" & mlng手术记录id & ",'" & bill(2).TextMatrix(lngLoop, 0) & "'," & bill(2).RowData(lngLoop) & ",'" & bill(2).TextMatrix(lngLoop, 2) & "','" & bill(2).TextMatrix(lngLoop, 1) & "')"
        End If
    Next
    
    strTmp = ""
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then
            If strTmp = "" Then
                strTmp = strSQL(lngLoop)
            Else
                strTmp = strTmp & Chr(9) & strSQL(lngLoop)
            End If
        End If
    Next
    
    '返回SQL语句
    ReturnStrSQL = strTmp
        
    SaveData = True
    
    Exit Function
    
ErrHandle:
    
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State <> adStateOpen Then Exit Function
    strError = Err.Description
    Call SaveErrLog
    
End Function

Private Sub SetDefault(ByVal objBill As BillEdit, ByVal intCol As Integer)
    '
    '功能：
    '
    Dim lngLoop As Long
    
    For lngLoop = 1 To objBill.Rows - 1
        If objBill.RowData(lngLoop) > 0 Then
            If objBill.TextMatrix(lngLoop, intCol) = "√" Then
                Exit For
            End If
        End If
    Next
    
    If lngLoop = objBill.Rows And objBill.RowData(1) > 0 Then
        objBill.TextMatrix(1, intCol) = "√"
    End If
    
End Sub

Private Sub bill_AfterDeleteRow(Index As Integer)
    If Index <> 2 Then SetDefault bill(Index), 2
End Sub

Private Sub bill_CellCheck(Index As Integer, Row As Long, Col As Long)
    Dim lngLoop As Long
    
    If bill(Index).TextMatrix(Row, Col) = "" Then
        SetDefault bill(Index), 2
    Else
        For lngLoop = 1 To bill(Index).Rows - 1
            If lngLoop <> Row Then bill(Index).TextMatrix(lngLoop, Col) = ""
        Next
    End If
End Sub

Private Sub bill_CommandClick(Index As Integer)
    Dim sglX As Single
    Dim sglY As Single
    
    If bill(Index).TextMatrix(bill(Index).Row, 0) <> "" Then
        Select Case Index
        Case 0, 1
            If PopOperateSelect(bill(Index), Val(Mid(bill(Index).TextMatrix(bill(Index).Row, 0), 1, 1))) Then
                SetDefault bill(Index), 2
            End If
        Case 2
            CalcPosition sglX, sglY, bill(Index)
            
            Call ShowDownListPerson(Screen, bill(Index), 0, sglX, sglY, False)
            
        End Select
    End If
    
End Sub

Private Sub bill_EditKeyPress(Index As Integer, KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub bill_EnterCell(Index As Integer, Row As Long, Col As Long)
    If Index = 0 Or Index = 1 Then SetDefault bill(Index), 2
End Sub

Private Sub bill_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim sglX As Single
    Dim sglY As Single
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If bill(Index).Col = 1 And Trim(bill(Index).TextMatrix(bill(Index).Row, 1)) = "" And bill(Index).TxtVisible = False Then
        zlCommFun.PressKey vbKeyTab
    End If
    
    Select Case bill(Index).Col
    Case 0
        If bill(Index).TextMatrix(bill(Index).Row, 0) <> bill(Index).List(bill(Index).ListIndex) Then
            
            Call ClearSpecRowCol(bill(Index), bill(Index).Row, Array())
            
            bill(Index).TextMatrix(bill(Index).Row, 0) = bill(Index).List(bill(Index).ListIndex)
                        
        End If
    End Select
    
    If bill(Index).TxtVisible = False Then Exit Sub
            
    If Trim(bill(Index).Text) <> "" Then
        Select Case Index
        Case 0, 1
            Cancel = Not ShowDownListOperate(Screen, Val(Mid(bill(Index).TextMatrix(bill(Index).Row, 0), 1, 1)), bill(Index), True, mstr性别)
            If Cancel = False Then Call SetDefault(bill(Index), 2)
        Case 2
            Call CalcPosition(sglX, sglY, bill(Index))
            Cancel = Not ShowDownListPerson(Screen, bill(Index), 0, sglX, sglY, True)
            If Cancel = False Then
                
            End If
        End Select
    Else
        If bill(Index).Col = 1 And bill(Index).RowData(bill(Index).Row) = 0 Then zlCommFun.PressKey vbKeyTab
    End If
    
End Sub

Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub bill_LostFocus(Index As Integer)
    bill(Index).CmdVisible = False
'    bill(Index).CboVisible = False
End Sub


Private Sub bill2_CommandClick(Index As Integer)
    If PopSelect(bill2(Index), mstr性别) Then
        
    End If
End Sub

Private Sub bill2_EditKeyPress(Index As Integer, KeyAscii As Integer)
        
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub bill2_EnterCell(Index As Integer, Row As Long, Col As Long)
    If bill2(Index).TextMatrix(Row, Col) = "" And Col <> 2 Then bill2(Index).TextMatrix(Row, Col) = " "
End Sub

Private Sub bill2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim sglX As Single
    Dim sglY As Single
        
    If KeyCode <> 13 Then Exit Sub
    
    If bill2(Index).TxtVisible = False Then Exit Sub
            
    Call CalcPosition(sglX, sglY, bill2(Index))

    If Trim(bill2(Index).Text) <> "" And bill2(Index).Col = 0 Or bill2(Index).Col = 1 Then
        Cancel = Not ShowDownList2(Screen, bill2(Index).Col + 1, bill2(Index), sglX, sglY, True)
        If Cancel = False Then
            
        End If
    End If
    
End Sub

Private Sub bill2_LostFocus(Index As Integer)
    bill2(Index).CmdVisible = False
    bill2(Index).CboVisible = False
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_Click(Index As Integer)
    If Index = 0 Then
        dtp(2).Enabled = chk(Index).Value
        dtp(3).Enabled = chk(Index).Value
        
        cbo(0).Enabled = dtp(2).Enabled
        cbo(1).Enabled = dtp(2).Enabled
        
        If cbo(1).Enabled = False Then
            cbo(1).ListIndex = -1
        ElseIf cbo(1).ListIndex = -1 Then
            cbo(1).ListIndex = 0
        End If
        
        txt(3).Visible = Not dtp(2).Enabled
        txt(1).Visible = Not dtp(3).Enabled
        
    Else
        dtp(4).Enabled = chk(Index).Value
        dtp(5).Enabled = chk(Index).Value
        txt(2).Visible = Not dtp(4).Enabled
        txt(4).Visible = Not dtp(5).Enabled
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub UserControl_Initialize()

    '初始化控件属性
    Dim lngLoop As Long
    
    On Error GoTo ErrHandle
    
    For lngLoop = 0 To 1
        With bill(lngLoop)
            .Cols = 3
            .TextMatrix(0, 0) = "编码方式"
            .TextMatrix(0, 1) = "手术名称"
            .TextMatrix(0, 2) = "缺省"
            .ColWidth(0) = 855
            .ColWidth(1) = 2220
            .ColWidth(2) = 600
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 4
            .ColData(0) = 3
            .ColData(1) = 1
            .ColData(2) = -1
            .AddItem "1-诊疗"
            .AddItem "2-疾病"
            .MsfObj.GridLinesFixed = flexGridFlat
            .MsfObj.GridColor = &H8000000C
            .MsfObj.GridColorFixed = &H8000000C
            .MsfObj.BackColorFixed = &HFFFFFF
            .Active = True
        End With
    Next
    
    For lngLoop = 0 To 1
        With bill2(lngLoop)
            .Cols = 5
            .TextMatrix(0, 0) = "疾病编码"
            .TextMatrix(0, 1) = "诊断编码"
            .TextMatrix(0, 2) = "诊断描述"
            .TextMatrix(0, 3) = "诊断ID"
            .TextMatrix(0, 4) = "疾病ID"
            .ColWidth(0) = 900
            .ColWidth(1) = 900
            .ColWidth(2) = 1890
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColAlignment(0) = 1
            .ColAlignment(1) = 1
            .ColAlignment(2) = 1
            
            .ColData(0) = 1
            .ColData(1) = 1
            .ColData(2) = 4
            .ColData(3) = 5
            .ColData(4) = 5
            .PrimaryCol = 2
            .MsfObj.GridLinesFixed = flexGridFlat
            .MsfObj.GridColor = &H8000000C
            .MsfObj.GridColorFixed = &H8000000C
            .MsfObj.BackColorFixed = &HFFFFFF
            .Active = True
        End With
    Next
    
    With bill(2)
        .Cols = 3
        
        .TextMatrix(0, 0) = "岗位"
        .TextMatrix(0, 1) = "姓名"
        .TextMatrix(0, 2) = "编号"
        
        .ColWidth(0) = 900
        .ColWidth(1) = 900
        .ColWidth(2) = 0
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColData(0) = 3
        .ColData(1) = 1
        .ColData(2) = 5
        .MsfObj.GridLinesFixed = flexGridFlat
        .MsfObj.GridColor = &H8000000C
        .MsfObj.GridColorFixed = &H8000000C
        .MsfObj.BackColorFixed = &HFFFFFF
        bill(2).Active = True
    End With
    
    Exit Sub
    
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function InDesign() As Boolean
    
    '功能：判断当前运行程序是否在VB的工程环境中
    
    On Error Resume Next
    
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
    
End Function

Private Sub UserControl_InitProperties()
    '初始病人病历为0
    mlng病历id = 0
    mDispMode = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDispMode = PropBag.ReadProperty("DispMode", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", BorderStyleSettings.flexBorderNone)
End Sub

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Resize()
    UserControl.Width = 8130
    UserControl.Height = 5985
End Sub

Private Sub UserControl_Terminate()
    If rsTmp.State = adStateOpen Then rsTmp.Close
    Set rsTmp = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DispMode", mDispMode, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, BorderStyleSettings.flexBorderNone)
End Sub

Private Sub UserControl_Show()
    Dim objCtl As Control
    Dim rs As New ADODB.Recordset
    
    
    '填写手术岗位列表
    gstrSql = "SELECT 名称,0 FROM 手术岗位"
    Call zlDatabase.OpenRecordset(rs, gstrSql, "手术概要")
    If rs.RecordCount = 0 Then
        MsgBox "系统数据不完整，没有手术岗位数据！", vbInformation, gstrSysName
        Exit Sub
    End If
    If rs.BOF = False Then Call AddComboData(bill(2), rs, False)
    
    '只在运行时显示
    If Ambient.UserMode = True And InDesign = False Then
        If mDispMode Then
            For Each objCtl In Controls
                If UCase(TypeName(objCtl)) <> UCase("ImageList") Then
                    objCtl.Enabled = False
                End If
            Next
        End If
    End If
    
    If mblnLoaded = False Then
        InitData
        Call ReadData
    End If
    
    mblnLoaded = True
End Sub

Public Property Get Text() As String
    '为每一个控件加上文本转储属性
    Dim lngLoop As Long
    Dim strTmp As String
    
    '通过用户输入的内容得到转储文本
    strTmp = "术前诊断：" & vbCrLf
    For lngLoop = 1 To bill2(0).Rows - 1
        If bill2(0).RowData(lngLoop) > 0 Then
            strTmp = strTmp & "          " & bill2(0).TextMatrix(lngLoop, 2) & vbCrLf
        End If
    Next
    
    strTmp = strTmp & "术后诊断：" & vbCrLf
    For lngLoop = 1 To bill2(1).Rows - 1
        If bill2(1).RowData(lngLoop) > 0 Then
            strTmp = strTmp & "          " & bill2(1).TextMatrix(lngLoop, 2) & vbCrLf
        End If
    Next
    
    strTmp = strTmp & "拟行手术：" & vbCrLf
    For lngLoop = 1 To bill(0).Rows - 1
        If bill(0).RowData(lngLoop) > 0 Then
            strTmp = strTmp & "          " & bill(0).TextMatrix(lngLoop, 1) & vbCrLf
        End If
    Next
    
    strTmp = strTmp & "已行手术：" & vbCrLf
    For lngLoop = 1 To bill(1).Rows - 1
        If bill(1).RowData(lngLoop) > 0 Then
            strTmp = strTmp & "          " & bill(1).TextMatrix(lngLoop, 1) & vbCrLf
        End If
    Next

    strTmp = strTmp & "手术开始：" & Format(dtp(0).Value, "YYYY年MM月DD日 HH时MM分") & "   "
    strTmp = strTmp & "手术终止：" & Format(dtp(1).Value, "YYYY年MM月DD日 HH时MM分") & vbCrLf

    If chk(0).Value <> 0 Then
        strTmp = strTmp & "麻醉开始：" & Format(dtp(2).Value, "YYYY年MM月DD日 HH时MM分") & "   "
        strTmp = strTmp & "麻醉终止：" & Format(dtp(3).Value, "YYYY年MM月DD日 HH时MM分") & vbCrLf
        strTmp = strTmp & "麻醉方式：" & cbo(0).Text & "   "
        strTmp = strTmp & "麻醉质量：" & cbo(1).Text & vbCrLf
    End If
    
    If chk(1).Value <> 0 Then
        strTmp = strTmp & "氧气开始：" & Format(dtp(4).Value, "YYYY年MM月DD日 HH时MM分") & "   "
        strTmp = strTmp & "氧气终止：" & Format(dtp(5).Value, "YYYY年MM月DD日 HH时MM分") & vbCrLf
    End If
    
    strTmp = strTmp & "手术人员：" & vbCrLf
    For lngLoop = 1 To bill(2).Rows - 1
        If bill(2).RowData(lngLoop) > 0 Then
            strTmp = strTmp & bill(2).TextMatrix(lngLoop, 0) & "          " & bill(2).TextMatrix(lngLoop, 1) & vbCrLf
        End If
    Next

    Text = strTmp
    
End Property

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    
    UserControl.Parent.CallBack_GotFocus
    
End Sub

Private Function CheckAllNumber(ByVal strKey As String) As Boolean
    
    Dim lngLoop As Long
    
    For lngLoop = 1 To Len(strKey)
        If Mid(strKey, lngLoop, 1) < "0" Or Mid(strKey, lngLoop, 1) > "9" Then
            Exit Function
        End If
    Next
    
    CheckAllNumber = True
End Function

Private Function CheckHave(ByVal objBill As Object, ByVal intRow As Integer, ByVal lngKey As Long) As Boolean
    Dim lngLoop As Long
    
    For lngLoop = 1 To objBill.Rows - 1
        If objBill.RowData(lngLoop) = lngKey And lngLoop <> intRow Then
            CheckHave = True
            Exit Function
        End If
    Next
    
    CheckHave = False
End Function

