VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChildQuestionFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤条件"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frmChildQuestionFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo 
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   195
      Width           =   2430
   End
   Begin VB.TextBox txt反馈人 
      Height          =   300
      Left            =   1230
      TabIndex        =   7
      Top             =   2115
      Width           =   2070
   End
   Begin VB.CommandButton cmd反馈人 
      Height          =   300
      Left            =   3345
      Picture         =   "frmChildQuestionFilter.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2100
      Width           =   300
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   255
      TabIndex        =   9
      Top             =   2685
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.ComboBox cbo抽查次数 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1620
      Width           =   2430
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2475
      TabIndex        =   11
      Top             =   2685
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1365
      TabIndex        =   10
      Top             =   2685
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   0
      Left            =   1230
      TabIndex        =   1
      Top             =   615
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   288227331
      CurrentDate     =   38083
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   300
      Index           =   1
      Left            =   1230
      TabIndex        =   3
      Top             =   1050
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   288161795
      CurrentDate     =   38083
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "反馈人(&3)"
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   6
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "审查次数(&2)"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "审查时间(&1)"
      Height          =   180
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "～"
      Height          =   180
      Index           =   9
      Left            =   810
      TabIndex        =   2
      Top             =   1080
      Width           =   180
   End
End
Attribute VB_Name = "frmChildQuestionFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################

Private mblnDataChanged As Boolean
Private mblnOK As Boolean
Private mstr抽查开始时间 As String
Private mstr抽查结束时间 As String
Private mstr日期选择 As String
Private mstr反馈人    As String
Private mlngCurNum As Long

Private mblnDataExecute As Boolean


'######################################################################################################################
Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function ShowPara(ByVal frmMain As Object, ByRef str抽查开始时间 As String, ByRef str抽查结束时间 As String, ByRef str日期选择 As String, ByRef lngCurNum As Long, ByRef str反馈人 As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnOK = False
    mblnDataExecute = True
    With cbo(0)
        .Clear
        .AddItem "所  有"
        .AddItem "自定义"
        .AddItem "今  天"
        .AddItem "昨  天"
        .AddItem "本  周"
        .AddItem "本  月"
        .AddItem "本  季"
        .AddItem "本半年"
        .AddItem "本  年"
        .AddItem "前三天"
        .AddItem "前一周"
        .AddItem "前半月"
        .AddItem "前一月"
        .AddItem "前二月"
        .AddItem "前三月"
        .AddItem "前半年"
        .AddItem "前一年"
        .AddItem "前二年"
        .Text = "前一月"
    End With
    mblnDataExecute = False
    
    If str日期选择 <> "" Then
        cbo(0).Text = str日期选择
    End If
    
    If cbo(0).Text = "自定义" Then
        If str抽查开始时间 = "" Then
            dtp(0).Value = Format(Now, "YYYY-MM-DD 00:00:00")
        Else
            dtp(0).Value = CDate(str抽查开始时间)
        End If
        
        If str抽查结束时间 = "" Then
            dtp(1).Value = Format(Now, "YYYY-MM-DD 23:59:59")
        Else
            dtp(1).Value = CDate(str抽查结束时间)
        End If
   
        Call Init抽查次数(str抽查开始时间, str抽查结束时间)
    
        If lngCurNum = 0 Then
            cbo抽查次数.ListIndex = 0
        Else
    '        cbo抽查次数.Text = lngCurNum
            cbo抽查次数.ListIndex = 0
        End If
    End If
    
    If str反馈人 = "" Then
        txt反馈人.Text = ""
    Else
        txt反馈人.Text = str反馈人
    End If
    
    Call SetCob(lngCurNum)
    
    
    Me.Show 1, frmMain
    
    If mblnOK Then
        str抽查开始时间 = mstr抽查开始时间
        str抽查结束时间 = mstr抽查结束时间
        str日期选择 = mstr日期选择
        lngCurNum = mlngCurNum
        str反馈人 = mstr反馈人
        ShowPara = mblnOK
    End If
    
End Function

Private Sub cbo_Click(Index As Integer)
    
    If mblnDataExecute Then Exit Sub
    
    Select Case Index
    Case 0
        Select Case cbo(Index).Text
        Case "所  有"
            dtp(0).Enabled = False
            dtp(1).Enabled = False
            dtp(0).Value = Format("2000-01-01 00:00:00", dtp(0).CustomFormat)
            dtp(1).Value = Format("3000-01-01 23:59:59", dtp(1).CustomFormat)
        Case "自定义"
            dtp(0).Enabled = True
            dtp(1).Enabled = True
        Case Else
            If dtp(0).Enabled = False Then
                dtp(0).Enabled = True
                dtp(1).Enabled = True
            End If
            dtp(0).Value = Format(GetBasePeriod(cbo(0).Text, 1), dtp(0).CustomFormat)
            dtp(1).Value = Format(GetBasePeriod(cbo(0).Text, 2), dtp(1).CustomFormat)
        End Select
        
         Call Init抽查次数(dtp(0).Value, dtp(1).Value)
         DataChanged = True
    End Select
    
    Dim strTempNum As String
    strTempNum = CLng(cbo抽查次数.ItemData(cbo抽查次数.ListIndex))
    Call Init抽查次数(dtp(0).Value, dtp(1).Value)
    If strTempNum = "" Then Exit Sub
'    cbo抽查次数.Text = strTempNum
    
    
    
End Sub

Private Sub cbo抽查次数_Change()
    DataChanged = True
End Sub

Private Sub cbo抽查次数_Click()
'    Dim strTempNum As String
'    strTempNum = CLng(cbo抽查次数.ItemData(cbo抽查次数.ListIndex))
'    Call Init抽查次数(dtp(0).Value, dtp(1).Value)
'    If strTempNum = "" Then Exit Sub
''    cbo抽查次数.Text = strTempNum
End Sub

Private Sub cbo抽查次数_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii < Asc(0) Or KeyAscii > Asc(9) Then KeyAscii = 0
End Sub

Private Sub cbo抽查次数_Validate(Cancel As Boolean)
     DataChanged = True
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    If DataChanged Then
        On Error Resume Next
        mstr抽查开始时间 = CStr(dtp(0).Value)
        mstr抽查结束时间 = CStr(dtp(1).Value)
        mstr日期选择 = cbo(0).Text
        mstr反馈人 = CStr(txt反馈人.Text)
        mlngCurNum = CLng(cbo抽查次数.ItemData(cbo抽查次数.ListIndex))
        
        mblnOK = True
        DataChanged = False
    End If
    Unload Me
End Sub

Private Sub cmdRef_Click()
    Call Init抽查次数(dtp(0).Value, dtp(1).Value)
End Sub



Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub dtp_Change(Index As Integer)
    Call zlControl.CboLocate(cbo(0), "自定义")
    DataChanged = True
End Sub

'''Private Sub Init抽查次数(ByVal str抽查开始时间 As String, ByVal str抽查结束时间 As String)
'''    On Error GoTo errH
'''        Dim rs As ADODB.Recordset
'''        cbo抽查次数.Clear
'''        gstrSQL = "select distinct(反馈次数) as 次数 from 病案反馈记录 A where A.反馈时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400 and A.反馈次数 is not null order by A.反馈次数"
'''        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(str抽查开始时间, "yyyy-mm-dd"), Format(str抽查结束时间, "yyyy-mm-dd"))
'''        If rs.RecordCount > 0 Then
'''            If rs.BOF = False Then
'''                Call AddComboData(cbo抽查次数, rs, "次数", "次数", , False)
'''            End If
'''        Else
'''            cbo抽查次数.AddItem "1"
'''        End If
'''    Exit Sub
'''errH:
'''    Err.Clear
'''    Exit Sub
'''End Sub

'获取抽查次数汇总信息
Private Sub Init抽查次数(ByVal str抽查开始时间 As String, ByVal str抽查结束时间 As String)
    On Error GoTo errH
        Dim rs As ADODB.Recordset
        Dim lngCount As Long '记录次数
        cbo抽查次数.Clear
        lngCount = 0
        gstrSQL = "select distinct(反馈次数),Sum(A.分值) as 总扣分数,Min(A.反馈时间) as 最早反馈时间 from 病案反馈记录 A where A.反馈时间 Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400 group by A.反馈次数 order by A.反馈次数 ASC"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(str抽查开始时间, "yyyy-mm-dd"), Format(str抽查结束时间, "yyyy-mm-dd"))
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            cbo抽查次数.AddItem "所有"
            Do Until rs.EOF
                If NVL(rs!反馈次数, 0) = 0 Then
                    cbo抽查次数.AddItem "第" & NVL(rs!反馈次数, 0) & "次-" & Format(NVL(rs!最早反馈时间, Now()), "YYYY-MM-DD") & "(" & NVL(rs!总扣分数, 0) & ")"
                    cbo抽查次数.ItemData(cbo抽查次数.NewIndex) = NVL(rs!反馈次数, 0)
                End If
                rs.MoveNext
            Loop
            
            rs.MoveFirst
            Do Until rs.EOF
                    If lngCount >= 10 Then Exit Do
'                        Call AddComboData(cbo抽查次数, rs, "最早反馈时间", "次数", , False)
                        If NVL(rs!反馈次数, 0) <> 0 Then
                            cbo抽查次数.AddItem "第" & NVL(rs!反馈次数, 0) & "次-" & Format(NVL(rs!最早反馈时间, Now()), "YYYY-MM-DD") & "(" & NVL(rs!总扣分数, 0) & ")"
                            cbo抽查次数.ItemData(cbo抽查次数.NewIndex) = NVL(rs!反馈次数, 0)
                        End If
                    lngCount = lngCount + 1
                    rs.MoveNext
            Loop
            cbo抽查次数.ListIndex = 0
        Else
            cbo抽查次数.AddItem "所有"
            cbo抽查次数.ItemData(cbo抽查次数.NewIndex) = 0
            cbo抽查次数.ListIndex = 0
        End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

Private Sub dtp_Validate(Index As Integer, Cancel As Boolean)
    
    Dim strTempNum As String
    If cbo(0).Text = "自定义" Then
        strTempNum = CLng(cbo抽查次数.ItemData(cbo抽查次数.ListIndex))
        Call Init抽查次数(dtp(0).Value, dtp(1).Value)
        If strTempNum = "" Then Exit Sub
    End If
'    cbo抽查次数.Text = strTempNum
End Sub

Private Sub txt反馈人_Change()
    DataChanged = True
End Sub

Private Sub cmd反馈人_Click()
    On Error GoTo errH
    SelectDoctor
    Exit Sub
errH:
    Err.Clear
    Exit Sub
End Sub

Private Sub txt反馈人_KeyPress(KeyAscii As Integer)
    If Trim(txt反馈人.Text) = "" Then Exit Sub
    If InStr(1, "—'|[](){}*%", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SelectDoctor txt反馈人.Text
    End If
End Sub

'选择医生
Private Sub SelectDoctor(Optional strShortName As String = "")
    Dim rsTmp           As ADODB.Recordset
    Dim rsResult        As ADODB.Recordset
    Dim bytRet          As Byte
On Error GoTo errH
    gstrSQL = ""
    If strShortName <> "" Then
        gstrSQL = gstrSQL & vbCrLf & "select distinct (A.反馈人) as 名称,B.ID as id,B.编号 From 病案反馈记录 A,人员表 B"
        gstrSQL = gstrSQL & vbCrLf & "Where A.反馈时间 Between To_Date([1], 'yyyy-mm-dd') And"
        gstrSQL = gstrSQL & vbCrLf & "To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400"
        gstrSQL = gstrSQL & vbCrLf & "And (B.简码 like '%'||[3]||'%')"
        gstrSQL = gstrSQL & vbCrLf & "And A.反馈人 = B.姓名"
        gstrSQL = gstrSQL & vbCrLf & "order by A.反馈人"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(dtp(0).Value, "yyyy-mm-dd"), Format(dtp(1).Value, "yyyy-mm-dd"), UCase(strShortName))
        bytRet = ShowPubSelect(Me, txt反馈人, 2, "编号,1200,0,;名称,1200,0,", Me.Name & "\反馈人选择", "请从下表中选择一个或多个反馈人", rsTmp, rsResult, 5000, 4500, True)
    Else
        Dim strTemp As String
        gstrSQL = gstrSQL & vbCrLf & "select distinct (A.反馈人) as 名称,B.ID as id,B.编号 From 病案反馈记录 A,人员表 B"
        gstrSQL = gstrSQL & vbCrLf & "Where A.反馈时间 Between To_Date([1], 'yyyy-mm-dd') And"
        gstrSQL = gstrSQL & vbCrLf & "To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400"
        gstrSQL = gstrSQL & vbCrLf & "And A.反馈人 = B.姓名"
        gstrSQL = gstrSQL & vbCrLf & "order by A.反馈人"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Format(dtp(0).Value, "yyyy-mm-dd"), Format(dtp(1).Value, "yyyy-mm-dd"))
        bytRet = ShowPubSelect(Me, txt反馈人, 2, "编号,1200,0,;名称,1200,0,", Me.Name & "\反馈人选择", "请从下表中选择一个或多个反馈人", rsTmp, rsResult, 5000, 4500, True)
        
    End If
     
    If rsResult Is Nothing Then
'        txt反馈人.Text = ""
    ElseIf rsResult.EOF Or rsResult.BOF Then
        txt反馈人.Text = ""
    Else
        rsResult.MoveFirst
        Do Until rsResult.EOF
            If Len(txt反馈人.Text) = 0 Then
                txt反馈人.Text = rsResult("名称").Value
            Else
                If InStrRev(txt反馈人.Text, rsResult("名称").Value, -1) = 0 Then
                    txt反馈人.Text = txt反馈人.Text & "," & rsResult("名称").Value
                End If
            End If
            rsResult.MoveNext
        Loop
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err.Clear
    Exit Sub
End Sub

Private Function GetBasePeriod(ByVal strMode As String, Optional ByVal bytFlag As Byte = 1) As String
    '******************************************************************************************************************
    '功能:获取特殊时间
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim intDay As Integer
    Dim varValue As Variant
    
    If Left(strMode, 3) = "自定义" Then
        '自定义:3,4
        varValue = Split(Mid(strMode, 5), ",")
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", Val(varValue(0)), zlDatabase.Currentdate), "yyyy-MM-dd") & " 00:00:00"
        Else
            If UBound(varValue) < 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "yyyy-MM-dd") & " 23:59:59"
            Else
                GetBasePeriod = Format(DateAdd("d", Val(varValue(1)), zlDatabase.Currentdate), "yyyy-MM-dd") & " 23:59:59"
            End If
        End If
            
        Exit Function
    End If
    
    Select Case strMode
    Case "当  时"      '当时
        GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    Case "今  天"       '当天
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  周"       '本周,bytFlag=1,本周开始时间,=2,本周结束时间
        intDay = Weekday(CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD")))
        
        If intDay = 1 Then
            intDay = 7
        Else
            intDay = intDay - 1
        End If
        
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 0 - intDay + 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 7 - intDay, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  月"       '本月
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM") & "-01 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM") & "-01"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "本  季"      '本季度
        Select Case Format(zlDatabase.Currentdate, "MM")
        Case "01", "02", "03"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-03-31 23:59:59"
            End If
        Case "04", "05", "06"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-04-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Case "07", "08", "09"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-09-30 23:59:59"
            End If
        Case "10", "11", "12"
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-10-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End Select
    Case "本半年"      '本半年
        If Val(Format(zlDatabase.Currentdate, "MM")) < 7 Then
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-06-30 23:59:59"
            End If
        Else
            If bytFlag = 1 Then
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-07-01 00:00:00"
            Else
                GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
            End If
        End If
    Case "本  年"   '全年
        If bytFlag = 1 Then
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-01-01 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY") & "-12-31 23:59:59"
        End If
    Case "昨  天"       '昨天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", -1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "明  天"       '明天
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(DateAdd("d", 1, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
        End If
    Case "前三天"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -3, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一周"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -7, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前半月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -15, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前一月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -30, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前二月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -60, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    Case "前三月"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -90, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    
    Case "前半年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -180, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前一年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
        
    Case "前二年"
        If bytFlag = 1 Then
            GetBasePeriod = Format(DateAdd("d", -365 * 2, CDate(Format(zlDatabase.Currentdate, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 00:00:00"
        Else
            GetBasePeriod = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
        End If
    End Select
    
End Function

Private Sub SetCob(ByVal lngCurNum As Long)
    Dim i As Integer
    For i = 0 To cbo抽查次数.ListCount - 1
        If cbo抽查次数.ItemData(i) = lngCurNum Then
            cbo抽查次数.ListIndex = i
            Exit Sub
        End If
    Next
End Sub
