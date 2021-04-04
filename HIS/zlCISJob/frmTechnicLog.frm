VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTechnicLog 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "执行情况"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmTechnicLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboResult 
      Height          =   300
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1995
      Width           =   1095
   End
   Begin VB.TextBox txt发送数次 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H80000011&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3795
      TabIndex        =   1
      Top             =   225
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4110
      TabIndex        =   7
      Top             =   2670
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2985
      TabIndex        =   6
      Top             =   2670
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -180
      TabIndex        =   14
      Top             =   2415
      Width           =   5970
   End
   Begin VB.TextBox txt执行摘要 
      Height          =   945
      Left            =   1005
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   975
      Width           =   4185
   End
   Begin VB.ComboBox cbo执行人 
      Height          =   300
      Left            =   1005
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1995
      Width           =   2070
   End
   Begin MSComCtl2.DTPicker dtp执行时间 
      Height          =   300
      Left            =   1005
      TabIndex        =   2
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   246153219
      CurrentDate     =   38082
   End
   Begin VB.TextBox txt本次数次 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3795
      TabIndex        =   3
      Top             =   600
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker dtp要求时间 
      Height          =   300
      Left            =   1005
      TabIndex        =   0
      Top             =   225
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   246153219
      CurrentDate     =   38082
   End
   Begin VB.Label lbl开嘱时间 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行结果"
      Height          =   180
      Left            =   3240
      TabIndex        =   17
      Top             =   2055
      Width           =   720
   End
   Begin VB.Label lbl单位 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位"
      Height          =   180
      Index           =   1
      Left            =   4845
      TabIndex        =   16
      Top             =   660
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "发送数次"
      Height          =   180
      Left            =   3030
      TabIndex        =   15
      Top             =   285
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行摘要"
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行人"
      Height          =   180
      Left            =   420
      TabIndex        =   12
      Top             =   2055
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行时间"
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   660
      Width           =   720
   End
   Begin VB.Label lbl单位 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单位"
      Height          =   180
      Index           =   0
      Left            =   4845
      TabIndex        =   10
      Top             =   285
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "本次数次"
      Height          =   180
      Left            =   3030
      TabIndex        =   9
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "要求时间"
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   285
      Width           =   720
   End
End
Attribute VB_Name = "frmTechnicLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Enum_Inside_Program
Private mlng科室ID As Long
Private mlng医嘱ID As Long
Private mlng发送号 As Long
Private mlng执行科室ID As Long
Private mstr执行时间 As String
Private mbln单独执行 As Boolean
Private mblnOK As Boolean
Private mdate执行终止时间  As Date
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr执行分类 As String
Private mstr操作类型 As String
Private mstr诊疗类别 As String
Private mstrPrivs    As String
Private mstrName     As String
Private mlng最大已销 As Long '对于单独执行医嘱，是指当前医嘱的此次发送的单据中被退费的最大次数，否则，是指此组医嘱的此次发送的单据中被退费的最大次数
Private mlng完全执行 As Long '执行结果为完全执行的总次数
Private mlng本次执行结果Old  As Long '更新时取原始本次执行结果，默认为 1 表示执行，0/2/3 均表示未执行
Private mlng本次次数Old     As Long '更新时取原始本次执行次数
Private mstrNO As String '发送记录对应的单据号NO
Private mint血袋数 As Integer '一共输几袋血
Private mbln血库流程 As Boolean
Private mbln叮嘱发送执行 As Boolean
Private mobjESign As Object '电子签名接口部件


Public Function ShowMe(ByVal frmParent As Object, ByVal lngModul As Enum_Inside_Program, ByVal lng科室ID As Long, ByVal lng医嘱ID As Long, _
    ByVal lng发送号 As Long, ByVal bln单独执行 As Boolean, Optional ByVal str执行时间 As String, Optional ByVal lng执行科室ID As Long, Optional ByVal strName As String, Optional ByVal strPrivs As String) As Boolean
'功能：登记或调整执行情况
'参数：lng科室ID=当前医技科室ID
'      str执行时间=调整时用(yyyy-MM-dd HH:mm:ss)
'返回：是否取消
    mlngModul = lngModul
    mlng科室ID = lng科室ID
    mlng医嘱ID = lng医嘱ID
    mlng发送号 = lng发送号
    mlng执行科室ID = lng执行科室ID
    mbln单独执行 = bln单独执行
    mstr执行时间 = str执行时间
    mstrPrivs = strPrivs
    mstrName = strName
    
    On Error Resume Next
    Me.Show 1, frmParent
    
    ShowMe = mblnOK
End Function

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrTime As Variant, strTime As String
    Dim vDate As Date, strPause As String
    Dim lngAllCount As Long, lngCurCount As Long, strCurDate As String, lng组ID As Long
    Dim blnFind As Long, dblTmp As Double
    Dim rs血库 As ADODB.Recordset
    
    mblnOK = False
        
    On Error GoTo errH
    
    '读取叮嘱发送执行参数
    mbln叮嘱发送执行 = Val(zlDatabase.GetPara("叮嘱需要发送执行", glngSys)) = 1
    
    '读取开嘱时间
    If mlng医嘱ID <> 0 Then
        strSQL = "select B.开嘱时间 from 病人医嘱记录 B where B.id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
        If Not rsTmp.EOF Then
            lbl开嘱时间.Caption = "开嘱时间：" & Format(Nvl(rsTmp!开嘱时间), "yyyy-MM-dd HH:mm")
        End If
    End If
    
    '读取执行人(本科人员)
    strSQL = "Select A.ID,A.编号,A.姓名,A.简码 From 人员表 A,部门人员 B" & _
        " Where A.ID=B.人员ID And B.部门ID=[1]" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng科室ID)
    For i = 1 To rsTmp.RecordCount
        cbo执行人.AddItem rsTmp!编号 & "-" & rsTmp!姓名
        cbo执行人.ItemData(cbo执行人.NewIndex) = rsTmp!ID
        If mstr执行时间 = "" Then
            If rsTmp!ID = UserInfo.ID Then
                cbo执行人.ListIndex = cbo执行人.NewIndex
                blnFind = True
            End If
        Else
            If rsTmp!姓名 = mstrName Then
                cbo执行人.ListIndex = cbo执行人.NewIndex
                blnFind = True
            End If
        End If
        rsTmp.MoveNext
    Next
    
    If InStr(mstrPrivs, "执行他科项目") > 0 And blnFind = False Then
        cbo执行人.AddItem UserInfo.编号 & "-" & UserInfo.姓名
        cbo执行人.ItemData(cbo执行人.NewIndex) = UserInfo.ID
        cbo执行人.ListIndex = cbo执行人.NewIndex
    End If
    
    If mlngModul = p医技工作站 Then
        If Val(zlDatabase.GetPara(51, glngSys)) = 1 Then
            Me.cbo执行人.Enabled = False
        End If
    End If

    '执行结果下拉菜单初始化
    cboResult.Clear
    cboResult.AddItem "未执行"
    cboResult.AddItem "完成"
    cboResult.AddItem "拒绝"
    cboResult.AddItem "外出"
    cboResult.ListIndex = 1

    mlng最大已销 = 0
    mlng完全执行 = 0
    mlng本次执行结果Old = 0
    mlng本次次数Old = 0
    '读取执行情况
    If mstr执行时间 = "" Then
        '上次已执行到的一些数据
        strSQL = "Select " & _
            " Max(执行时间) as LastDate," & _
            " Max(要求时间) as curDate," & _
            " Count(要求时间) as curCount," & _
            " Sum(本次数次) as curNum" & _
            " From 病人医嘱执行" & _
            " Where 医嘱ID=[1] And 发送号=[2] and 本次数次>0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号)
        If Not rsTmp.EOF Then
            dtp执行时间.Tag = Format(Nvl(rsTmp!LastDate), "yyyy-MM-dd HH:mm:ss") '上次实际执行时间
            txt发送数次.Tag = Nvl(rsTmp!curNum, 0) '上次为止实际已执行的数次总量
            strCurDate = Format(Nvl(rsTmp!curDate), "yyyy-MM-dd HH:mm:ss") '上次执行的要求时间
            lngCurCount = Nvl(rsTmp!curCount, 0) '上次为止实际已执行的次数
        End If
        
        '计算本次执行应该的要求时间
        strSQL = "Select A.发送数次,Nvl(B.相关id, B.ID) 组ID,C.计算单位,A.首次时间,A.末次时间,Decode(B.病人来源, 2, Decode(A.记录性质, 1, 1, Decode(A.门诊记帐, 1, 1, 2)), 1) 费用性质," & _
            " B.开始执行时间,Decode(B.医嘱期效,0,B.执行终止时间,null) as 执行终止时间,B.上次执行时间,B.执行时间方案," & _
            " B.执行频次,B.频率次数,B.频率间隔,B.间隔单位,B.病人ID,b.主页ID,c.类别,c.操作类型,c.执行分类,C.计算方式,B.医嘱期效,Nvl(b.总给予量, 1) as 总给予量,NVL(B.单次用量,1) AS 单次用量,A.NO " & _
            " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C" & _
            " Where A.医嘱ID=B.ID And B.诊疗项目ID=C.ID(+)" & _
            " And A.医嘱ID=[1] And A.发送号=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号)
        mstrNO = rsTmp!NO & ""
        '查询单据中最大的已经退费或销帐的医嘱执行次数
        mlng最大已销 = Get最大已销(mbln单独执行, mlng医嘱ID, rsTmp!组ID, rsTmp!类别 & "", Val(rsTmp!费用性质 & ""))
        
        lbl单位(0).Caption = Nvl(rsTmp!计算单位)
        lbl单位(1).Caption = Nvl(rsTmp!计算单位)
        txt发送数次.Text = Nvl(rsTmp!发送数次)
        dtp执行时间.Value = zlDatabase.Currentdate
        mdate执行终止时间 = CDate(Nvl(rsTmp!执行终止时间, 0))
        mlng病人ID = Val(rsTmp!病人ID & "")
        mlng主页ID = Val(rsTmp!主页ID & "")
        mstr诊疗类别 = rsTmp!类别 & ""
        mstr操作类型 = rsTmp!操作类型 & ""
        mstr执行分类 = rsTmp!执行分类 & ""
        
        '新增执行记录时，输血医嘱单独处理
        If gbln血库系统 And mstr诊疗类别 = "E" And mstr操作类型 = "8" Then
            mbln血库流程 = True
            strSQL = "select zl_Get_输血执行次数([1]) as 数量 from dual"
            Set rs血库 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!组ID & ""))
            If Not rs血库.EOF Then mint血袋数 = Val(rs血库!数量 & "")
            lbl单位(0).Caption = "袋"
            lbl单位(1).Caption = "袋"
            txt发送数次.Text = mint血袋数
            Label1.Tag = txt发送数次.Tag
            txt发送数次.Tag = FormatEx(Val(txt发送数次.Tag) * mint血袋数, 0) '上次执行总量。已经有5位小数，用四舍五入就可以满足
            If mlng最大已销 = 1 Then '对于输血医嘱 这个 mlng最大已销 的值只能是 0 或 1 因为只能销一次。
                MsgBox "该医嘱相关单据已经退费或销帐 " & IIf(mbln单独执行, "不能再执行。", "不能再一并执行，请单独执行。"), vbInformation, gstrSysName
                Unload Me: Exit Sub
            ElseIf mlng最大已销 = 0 Then
                If Val(txt发送数次.Tag) >= Val(txt发送数次.Text) Then
                    MsgBox "该医嘱本次发送允许执行 " & txt发送数次.Text & "袋，当前已经执行了 " & Val(txt发送数次.Tag) & " 袋，不能再执行。", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
            dtp要求时间.Value = rsTmp!开始执行时间 '输血医嘱都为一次性执行的临嘱
            txt本次数次.Text = 1 '每次执行默认为一袋
            Exit Sub
        Else
            mbln血库流程 = False
            mint血袋数 = 0
        End If
        
        
        '当前实际已经执行了要求的次数,不准再执行
        If Val(txt发送数次.Tag) + mlng最大已销 >= Val(txt发送数次.Text) And (Not (mbln叮嘱发送执行 And mstr执行分类 = "")) Then
            MsgBox "该医嘱本次发送允许执行 " & txt发送数次.Text & IIf(mlng最大已销 <> 0, " 次，" & "相关单据已经退费或销帐" & mlng最大已销, "") & "次，当前已经执行了 " & Val(txt发送数次.Tag) & " 次，" & IIf(mbln单独执行, "不能再执行。", "不能再一并执行，请单独执行。"), vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        If rsTmp!执行频次 & "" = "一次性" Or rsTmp!执行频次 & "" = "需要时" Or (mbln叮嘱发送执行 And mstr执行分类 = "") Then
            '为一次性执行的临嘱
            dtp要求时间.Value = rsTmp!开始执行时间
        ElseIf strCurDate = "" And lngCurCount = 0 Then
            '第一次执行时,就为首次时间
            dtp要求时间.Value = rsTmp!首次时间
        Else
            '根据执行频率分解时间
            strPause = GetAdvicePause(mlng医嘱ID)
            If IsNull(rsTmp!执行时间方案) And (Nvl(rsTmp!频率次数, 0) = 0 Or Nvl(rsTmp!频率间隔, 0) = 0 Or IsNull(rsTmp!间隔单位)) Then
                '持续性长嘱
                lngAllCount = 0: strTime = ""
                vDate = Format(rsTmp!首次时间, "yyyy-MM-dd")
                Do While vDate <= Format(rsTmp!末次时间, "yyyy-MM-dd")
                    If Not DateIsPause(vDate, strPause) Then
                        lngAllCount = lngAllCount + 1
                        If Format(vDate, "yyyy-MM-dd") > Format(strCurDate, "yyyy-MM-dd") And strTime = "" Then
                            strTime = Format(vDate, "yyyy-MM-dd")
                        End If
                    End If
                    vDate = vDate + 1
                Loop
                
                '当前实际已经执行了要求的次数,不准再执行
                If lngCurCount + mlng最大已销 >= lngAllCount And (Not (mbln叮嘱发送执行 And mstr执行分类 = "")) Then
                    MsgBox "该医嘱本次发送允许执行 " & lngAllCount & "次，当前已经执行了 " & lngCurCount & " 次，不能再执行。", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
                
                dtp要求时间.Value = CDate(strTime)
            Else
                vDate = Calc本周期开始时间(rsTmp!开始执行时间, rsTmp!首次时间, rsTmp!频率间隔, rsTmp!间隔单位)
                strTime = Calc段内分解时间(vDate, rsTmp!末次时间, strPause, rsTmp!执行时间方案 & "", rsTmp!频率次数, rsTmp!频率间隔, rsTmp!间隔单位, rsTmp!开始执行时间)
                arrTime = Split(strTime, ",")
                lngAllCount = 0
                For i = 0 To UBound(arrTime)
                    If CDate(arrTime(i)) >= rsTmp!首次时间 Then
                        lngAllCount = lngAllCount + 1
                    End If
                Next
                '当前实际已经执行了要求的次数,不准再执行
                If lngCurCount + mlng最大已销 >= lngAllCount Then
                    MsgBox "该医嘱本次发送允许执行 " & lngAllCount & "次，当前已经执行了 " & lngCurCount & " 次，不能再执行。", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
                
                dtp要求时间.Value = rsTmp!开始执行时间
                For i = 0 To UBound(arrTime)
                    If arrTime(i) > strCurDate Then
                        dtp要求时间.Value = CDate(arrTime(i))
                        Exit For '以第一个时间为要求时间
                    End If
                Next
                If i > UBound(arrTime) Then
                    dtp要求时间.Value = CDate(arrTime(0))
                End If
            End If
        End If
        If Val(rsTmp!计算方式 & "") = 2 Or Val(rsTmp!计算方式 & "") = 1 Then
            If rsTmp!医嘱期效 = 0 Then
                '1、长嘱可选频率、持续性、必要时和不定时以单量作为数次。
                txt本次数次.Text = Val(rsTmp!单次用量 & "")
            ElseIf InStr("一次性,需要时", rsTmp!执行频次 & "") And rsTmp!执行频次 & "" <> "" Then
                '2、临嘱一次性和需要时频率的医嘱取总量作为数次。
                If mstr诊疗类别 = "E" And mstr操作类型 = "8" Or mstr诊疗类别 = "K" Then
                    txt本次数次.Text = Val(rsTmp!总给予量 & "") - Val(txt发送数次.Tag)
                Else
                    txt本次数次.Text = 1
                End If
            Else
                txt本次数次.Text = Get本次数次(mlng医嘱ID, mlng发送号, dtp要求时间.Value, Val(rsTmp!总给予量 & ""), Val(rsTmp!单次用量 & ""))
            End If
        Else
            dblTmp = Val(txt发送数次.Text) - Val(txt发送数次.Tag) - mlng最大已销
            txt本次数次.Text = IIf(dblTmp > 1, 1, dblTmp)
        End If
        If mbln叮嘱发送执行 And mstr执行分类 = "" Then
            txt本次数次.Text = 1
        End If
        If Mid(txt本次数次.Text, 1, 1) = "." Then txt本次数次.Text = "0" & txt本次数次.Text
        If gbln血库系统 And mstr诊疗类别 = "K" Then txt本次数次.Text = Val(txt本次数次.Text) - Val(txt发送数次.Tag)
    Else
        '上次已执行到的一些数据(不算本次)
        strSQL = "Select " & _
            " Max(执行时间) as LastDate," & _
            " Sum(本次数次) as curNum" & _
            " From 病人医嘱执行" & _
            " Where 执行时间<[3] And 医嘱ID=[1] And 发送号=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号, CDate(mstr执行时间))
        If Not rsTmp.EOF Then
            txt发送数次.Tag = Nvl(rsTmp!curNum, 0) '上次为止实际已执行的数次总量
            dtp执行时间.Tag = Format(Nvl(rsTmp!LastDate), "yyyy-MM-dd HH:mm:ss") '上次实际执行时间
        End If
    
        strSQL = "Select A.要求时间,Nvl(C.相关id, C.ID) 组ID,A.执行时间,A.本次数次,A.执行摘要,nvl(A.执行结果,1) as 执行结果,A.执行人,B.发送数次,Decode(C.病人来源, 2, Decode(B.记录性质, 1, 1, Decode(B.门诊记帐, 1, 1, 2)), 1) 费用性质,D.计算单位,Decode(c.医嘱期效,0,c.执行终止时间,null) as 执行终止时间 ,d.类别,d.操作类型,d.执行分类,c.病人ID,c.主页ID,B.NO" & _
            " From 病人医嘱执行 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 D" & _
            " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And B.医嘱ID=C.ID And C.诊疗项目ID=D.ID(+)" & _
            " And A.医嘱ID=[1] And A.发送号=[2] And A.执行时间=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号, CDate(mstr执行时间))
        
        '查询单据的中医嘱的最大执行次数
        mlng完全执行 = Get完全执行(mbln单独执行, mlng医嘱ID, rsTmp!组ID, mlng发送号)
        mstrNO = rsTmp!NO & ""
        '查询单据中最大的已经退费或销帐的医嘱执行次数
        mlng最大已销 = Get最大已销(mbln单独执行, mlng医嘱ID, rsTmp!组ID, rsTmp!类别 & "", Val(rsTmp!费用性质 & ""))
        
        dtp要求时间.Value = rsTmp!要求时间
        txt发送数次.Text = Nvl(rsTmp!发送数次)
        lbl单位(0).Caption = Nvl(rsTmp!计算单位)
        mdate执行终止时间 = CDate(Nvl(rsTmp!执行终止时间, 0))
        mlng病人ID = Val(rsTmp!病人ID & "")
        mlng主页ID = Val(rsTmp!主页ID & "")
        mstr诊疗类别 = rsTmp!类别 & ""
        mstr操作类型 = rsTmp!操作类型 & ""
        mstr执行分类 = rsTmp!执行分类 & ""
        mlng本次次数Old = FormatEx(Nvl(rsTmp!本次数次), 5)
        mlng本次执行结果Old = Val(rsTmp!执行结果 & "")
        
        dtp执行时间.Value = rsTmp!执行时间
        txt本次数次.Text = FormatEx(Nvl(rsTmp!本次数次), 5)
        
        lbl单位(1).Caption = Nvl(rsTmp!计算单位)
        
        txt执行摘要.Text = Nvl(rsTmp!执行摘要)
        '修改时获取执行结果
        cboResult.ListIndex = Val(rsTmp!执行结果 & "")
        
        mlng完全执行 = mlng完全执行 - IIf(Val(rsTmp!执行结果 & "") = 1, Val(txt本次数次.Text), 0)
        
        Cbo.SeekIndex cbo执行人, rsTmp!执行人
        '修改执行记录时，输血医嘱单独处理
        If gbln血库系统 And mstr诊疗类别 = "E" And mstr操作类型 = "8" Then
            mbln血库流程 = True
            strSQL = "select zl_Get_输血执行次数([1]) as 数量 from dual"
            Set rs血库 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp!组ID & ""))
            If Not rs血库.EOF Then mint血袋数 = Val(rs血库!数量 & "")
            '只需要处理以下几个量， mlng本次次数Old ，mlng本次执行结果Old 等不用单独处理。这些变量用于做特殊检查，输血医嘱不用做这些检查
            lbl单位(0).Caption = "袋"
            lbl单位(1).Caption = "袋"
            txt发送数次.Text = mint血袋数
            txt发送数次.Tag = FormatEx(Val(txt发送数次.Tag) * mint血袋数, 0)
            txt本次数次.Text = FormatEx(Val("" & rsTmp!本次数次) * mint血袋数, 0)
            Exit Sub
        Else
            mbln血库流程 = False
            mint血袋数 = 0
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Get本次数次(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByVal dat要求时间 As Date, ByVal dbl总量 As Double, ByVal dbl单量 As Double) As Double
'功能：根据医嘱信息执行时间，查出临时计时计量医嘱本次数次
    Dim strSQL As String, rsTmp As Recordset
    Dim lng当前次数 As Long, i As Long
    Dim dbl总量Tmp As Double, dbl数量 As Double
    
    strSQL = "Select 要求时间 From 医嘱执行时间 Where 医嘱id = [1] And 发送号 = [2] Order By 要求时间"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Get本次数次", lng医嘱ID, lng发送号)
    dbl总量Tmp = dbl总量
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp.RecordCount = 1 Then
                dbl数量 = dbl总量
            Else
                If i = rsTmp.RecordCount Then
                    dbl数量 = dbl总量Tmp
                Else
                    If dbl总量Tmp >= dbl单量 Then
                        dbl数量 = dbl单量
                    Else
                        dbl数量 = dbl总量Tmp
                    End If
                    dbl总量Tmp = dbl总量Tmp - dbl数量
                End If
            End If
            If CDate(Format(rsTmp!要求时间 & "", "YYYY-MM-DD HH:mm:ss")) = dat要求时间 Then
                Get本次数次 = dbl数量
                Exit For
            End If
            rsTmp.MoveNext
        Next
    Else
        Get本次数次 = 1
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim blnTrans As Boolean
    Dim lng次数 As Long
    Dim strTmp As String
    Dim dbl本次数次 As Double
    Dim dbl已发次数 As Double
    
    If zlCommFun.ActualLen(txt执行摘要.Text) > txt执行摘要.MaxLength Then
        MsgBox "执行摘要内容过多，最多允许 " & txt执行摘要.MaxLength \ 2 & " 个汉字或 " & txt执行摘要.MaxLength & " 个字符。", vbInformation, gstrSysName
        txt执行摘要.SetFocus: Exit Sub
    End If
    
    If cboResult.ListIndex = -1 Then
        MsgBox "请确定执结果。", vbInformation, gstrSysName
        cboResult.SetFocus
        Exit Sub
    End If
    
    dbl本次数次 = Val(txt本次数次.Text)
  
    If dbl本次数次 < 0 Or dbl本次数次 = 0 And cboResult.ListIndex <= 1 Then
        MsgBox "请确认本次执行的" & IIf(mbln血库流程, "输血袋数。", "数次。"), vbInformation, gstrSysName
        txt本次数次.SetFocus: Exit Sub
    End If
     
    If cbo执行人.Text = "" Then
        MsgBox "请确定执行人。", vbInformation, gstrSysName
        If cbo执行人.Enabled Then cbo执行人.SetFocus
        Exit Sub
    End If
    If dtp要求时间.Value > mdate执行终止时间 And mdate执行终止时间 <> CDate(0) Then
        MsgBox "要求时间超过了医嘱终止时间，请确认医嘱是否提前停止。", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    '检查本次执行时间是否大于上次执行时间
    If IsDate(dtp执行时间.Tag) Then
        If dtp执行时间.Value <= CDate(Format(dtp执行时间.Tag, "yyyy-MM-dd HH:mm:ss")) Then
            MsgBox "本次执行时间应晚于上次执行时间 " & Format(dtp执行时间.Tag, "yyyy-MM-dd HH:mm") & "。", vbInformation, gstrSysName
            dtp执行时间.SetFocus: Exit Sub
        End If
    End If 
    
    '检查每次执行数次是否超过总的发送数次
    If Val(txt发送数次.Text) <> 0 And Not mbln血库流程 Then '未填写发送数次的不限制(如持续性长嘱)
        If dbl本次数次 + Val(txt发送数次.Tag) > Val(txt发送数次.Text) And Not (mbln叮嘱发送执行 And mstr执行分类 = "") Then
            MsgBox "包括本次执行数次的所有执行数次超过了医嘱发送数次 " & FormatEx(txt发送数次.Text, 5) & " " & lbl单位(0).Caption & "。", vbInformation, gstrSysName
            txt本次数次.SetFocus: Exit Sub
        End If
        '存在退费的单据需要判断账面数量与实际数量，实际数量是指不包含本次执行的执行完成次数+退费数次。账面数量是指，不包含本次执行的所有执行登记次数+退费数次
        If mlng最大已销 <> 0 And (mlng本次执行结果Old <> 1 And cboResult.ListIndex = 1 Or mlng本次次数Old <> dbl本次数次) Then
            If (mlng完全执行 + mlng最大已销 >= Val(txt发送数次.Text)) Or (Val(txt发送数次.Tag) + mlng最大已销 > Val(txt发送数次.Text)) Then
                MsgBox "该医嘱已经部分退费或销帐,没有剩余执行次数，不允许修改本次的执行次数与执行结果。 ", vbInformation, gstrSysName
                txt本次数次.Text = mlng本次次数Old
                cboResult.ListIndex = mlng本次执行结果Old
                cmdCancel.SetFocus: Exit Sub
            ElseIf (mlng完全执行 + mlng最大已销 + IIf(cboResult.ListIndex = 1, dbl本次数次, 0) > Val(txt发送数次.Text)) Or (Val(txt发送数次.Tag) + mlng最大已销 + dbl本次数次 > Val(txt发送数次.Text)) Then
                lng次数 = IIf((Val(txt发送数次.Text) - mlng完全执行 - mlng最大已销) > (Val(txt发送数次.Text) - Val(txt发送数次.Tag) - mlng最大已销), (Val(txt发送数次.Text) - Val(txt发送数次.Tag) - mlng最大已销), (Val(txt发送数次.Text) - mlng完全执行 - mlng最大已销))
                If lng次数 > 0 Then
                    MsgBox "该医嘱本次发送允许执行 " & txt发送数次.Text & "次,相关单据已经退费或销帐" & mlng最大已销 & "次，" & _
                            IIf(Val(txt发送数次.Tag) = 0, "", "当前已经执行了 " & Val(txt发送数次.Tag) & " 次，") & _
                            "还允许执行" & lng次数 & "次。", vbInformation, gstrSysName
                    txt本次数次.Text = lng次数
                    cboResult.ListIndex = mlng本次执行结果Old
                    cmdCancel.SetFocus: Exit Sub
                Else
                    MsgBox "该医嘱已经部分退费或销帐,没有剩余执行次数，不允许修改本次的执行次数与执行结果。 ", vbInformation, gstrSysName
                    txt本次数次.Text = mlng本次次数Old
                    cboResult.ListIndex = mlng本次执行结果Old
                    cmdCancel.SetFocus: Exit Sub
                End If
            End If
        End If
    ElseIf mbln血库流程 Then
        strTmp = FormatEx(dbl本次数次, 5)
        If InStr(strTmp, ".") > 0 Then
            MsgBox "输血袋数不应包含小数。", vbInformation, gstrSysName
            txt本次数次.SetFocus: Exit Sub
        End If
        
        If dbl本次数次 + Val(txt发送数次.Tag) > Val(txt发送数次.Text) Then
            MsgBox "本次执行" & dbl本次数次 & "袋，已经执行" & Val(txt发送数次.Tag) & "袋，超过了所有需要执行的输血袋数 " & FormatEx(txt发送数次.Text, 5) & " " & lbl单位(0).Caption & "。", vbInformation, gstrSysName
            txt本次数次.SetFocus: Exit Sub
        End If
        
        dbl本次数次 = FormatEx(dbl本次数次 / mint血袋数, 5)
        dbl已发次数 = Val(Label1.Tag)
        If dbl本次数次 + dbl已发次数 > 1 Then
            dbl本次数次 = 1 - dbl已发次数
        End If
    End If
    
    If mstr诊疗类别 = "E" And mstr操作类型 = "1" And gintCA > 0 And Mid(gstrESign, 2, 1) = "1" Then
        If Not Check电子签名 Then Exit Sub
    End If
    '保存数据
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    
    If mstr执行时间 = "" Then
        If mlng执行科室ID <> 0 Then
            strSQL = "Zl_病人医嘱发送_科室变更(" & mlng医嘱ID & "," & mlng发送号 & "," & mlng执行科室ID & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        
        strSQL = "ZL_病人医嘱执行_Insert(" & mlng医嘱ID & "," & mlng发送号 & "," & _
            "To_Date('" & Format(dtp要求时间.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            dbl本次数次 & ",'" & txt执行摘要.Text & "','" & zlCommFun.GetNeedName(cbo执行人.Text) & "'," & _
            "To_Date('" & Format(dtp执行时间.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            IIf(mbln单独执行, 1, 0) & "," & "0," & cboResult.ListIndex & ",'','" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlng执行科室ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        strSQL = "ZL_病人医嘱执行_Update(To_Date('" & mstr执行时间 & "','YYYY-MM-DD HH24:MI:SS')," & mlng医嘱ID & "," & mlng发送号 & "," & _
            "To_Date('" & Format(dtp要求时间.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
            dbl本次数次 & ",'" & txt执行摘要.Text & "','" & zlCommFun.GetNeedName(cbo执行人.Text) & "'," & _
            "To_Date('" & Format(dtp执行时间.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')" & "," & cboResult.ListIndex & ",NULL," & IIf(mbln单独执行, 1, 0) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlng执行科室ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    mbln血库流程 = False
    mint血袋数 = 0
    Set mobjESign = Nothing
End Sub

Private Sub txt本次数次_GotFocus()
    Call zlControl.TxtSelAll(txt本次数次)
End Sub

Private Sub txt本次数次_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt本次数次_Validate(Cancel As Boolean)
    If Not IsNumeric(txt本次数次.Text) Then
        txt本次数次.Text = ""
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub txt发送数次_GotFocus()
    Call zlControl.TxtSelAll(txt发送数次)
End Sub

Private Sub txt执行摘要_GotFocus()
    Call zlControl.TxtSelAll(txt执行摘要)
End Sub

Private Function Get最大已销(ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng组ID As Long, ByVal str诊疗类别 As String, ByVal int费用性质 As Integer) As Long
'功能：获取某条医嘱，或某组医嘱的最大已销帐的医嘱执行次数
'       bln单独执行 是否单独执行，检验检查类存在单据的医嘱的单独执行某一部位，某一部分检查
'       lng医嘱ID 该条医嘱ID
'       lng组ID 没有父医嘱，或者父医嘱时为医嘱ID,子医嘱为相关ID
'       str诊疗类别 该医嘱的诊疗类别
'       int费用性质 1-门诊费用，2-住院费用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strTable As String
    Dim rs计价 As ADODB.Recordset
    Dim lngRes As Long
    Dim dblTmp As Double
    
    On Error GoTo errH
    strTable = IIf(int费用性质 = 1, "门诊费用记录", "住院费用记录")
    If bln单独执行 Then
        lng组ID = lng医嘱ID
        strSQL = "Select -1 * Sum(Nvl(a.付数, 1) * a.数次 / b.数量) As 最大已销数" & vbNewLine & _
                "From " & strTable & " A, 病人医嘱计价 B" & vbNewLine & _
                "Where a.医嘱序号 = [1] And A.NO=[3] And b.医嘱id = a.医嘱序号 And b.收费细目id = a.收费细目id And Nvl(B.费用性质,0)=0 And a.记录状态 = 2 And a.记录性质 in(1,2,11) And a.价格父号 Is Null And" & vbNewLine & _
                "      a.收费类别 Not In ('5', '6', '7') And Not Exists" & vbNewLine & _
                " (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1)"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng组ID, str诊疗类别, mstrNO)
        If rsTmp.RecordCount <> 0 Then
            lngRes = Val(rsTmp!最大已销数 & "")
        End If
    Else
        strSQL = "Select a.医嘱id,a.收费细目id,count(1) as 次数 From 医嘱执行计价 a Where a.医嘱id = [1] And a.发送号 = [2] and a.数量>0 group by a.医嘱id,a.收费细目id"
        Set rs计价 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号)
        
        '取出销帐消得最多的那一个，即收费次数最少的那一个
        strSQL = "select a.医嘱id,a.收费细目id,a.收费次数 from (Select a.医嘱序号 as 医嘱id,a.收费细目id,Sum(Nvl(a.付数, 1) * a.数次 / b.数量) As 收费次数" & vbNewLine & _
                "       From " & strTable & " A, 病人医嘱计价 B" & vbNewLine & _
                "       Where a.医嘱序号 In (Select ID From 病人医嘱记录 Where (ID = [1] Or 相关id = [1]) And A.NO=[3] And 诊疗类别 = [2]) And b.医嘱id = a.医嘱序号 And" & vbNewLine & _
                "             b.收费细目id = a.收费细目id And Nvl(B.费用性质,0)=0  And a.记录性质 in(1,2) And a.价格父号 Is Null And a.收费类别 Not In ('5', '6', '7') And" & vbNewLine & _
                "             Not Exists" & vbNewLine & _
                "        (Select 1 From 材料特性 Where 材料id = a.收费细目id And 跟踪在用 = 1) " & vbNewLine & _
                "       Group By  a.医嘱序号,a.收费细目id) a order by a.收费次数"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng组ID, str诊疗类别, mstrNO)
        
        '剩余最少那一个即为第一条数据
        If Not rsTmp.EOF Then
            rs计价.Filter = "医嘱id=" & rsTmp!医嘱ID & " and 收费细目id=" & rsTmp!收费细目id
            If Not rs计价.EOF Then
                dblTmp = Val(rs计价!次数 & "") - Val(rsTmp!收费次数 & "")
                If dblTmp > 0 Then
                    lngRes = IntEx(dblTmp)
                End If
            End If
        End If
    End If
    Get最大已销 = lngRes
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get完全执行(ByVal bln单独执行 As Boolean, ByVal lng医嘱ID As Long, ByVal lng组ID As Long, ByVal lng发送号 As Long) As Long
'功能：获取医嘱的完全执行次数
'       bln单独执行 是否单独执行，检验检查类存在单据的医嘱的单独执行某一部位，某一部分检查
'       lng医嘱ID 该条医嘱ID
'       lng组ID 没有父医嘱，或者父医嘱时为医嘱ID,子医嘱为相关ID
'       lng发送号 本次医嘱发送的发送号

    Dim rsTmp As New ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    If bln单独执行 Then
        lng组ID = lng医嘱ID
        strSQL = "Select Sum(Nvl(b.本次数次,a.发送数次)) 完全执行次数" & vbNewLine & _
            "From 病人医嘱发送 A, 病人医嘱执行 B" & vbNewLine & _
            "Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And a.医嘱id = [1] And a.发送号 = [2] And Nvl(b.执行结果, 1)=1"
    Else
        strSQL = "Select Max(C.完全执行次数) 完全执行次数" & vbNewLine & _
            "From (Select  Sum(Nvl(b.本次数次,a.发送数次)) 完全执行次数" & vbNewLine & _
            "       From 病人医嘱发送 A, 病人医嘱执行 B" & vbNewLine & _
            "       Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And" & vbNewLine & _
            "             a.医嘱id In (Select ID From 病人医嘱记录 Where ID = [1] Or 相关id = [1]) And a.发送号 = [2] And" & vbNewLine & _
            "             Nvl(b.执行结果, 1)=1" & vbNewLine & _
            "       Group By a.医嘱id) C"

    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng组ID, lng发送号)
    If rsTmp.RecordCount <> 0 Then
        Get完全执行 = Val(rsTmp!完全执行次数 & "")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check电子签名() As Boolean
    '判断是否启用数字签名
    Check电子签名 = True
    If gintCA > 0 And CheckSign(2, mlng科室ID, , , , False, mobjESign) Then
        If mobjESign Is Nothing Then
            On Error Resume Next
            Set mobjESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not mobjESign Is Nothing Then
                Call mobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If mobjESign Is Nothing Then
            MsgBox "电子签名部件未能正确安装，签名操作不能继续。", vbInformation, gstrSysName
            Check电子签名 = False
            Exit Function
        Else
            If Not mobjESign.CheckCertificate(UserInfo.用户名) Then
                Check电子签名 = False
                Exit Function
            End If
        End If
    End If
End Function
