VERSION 5.00
Begin VB.Form frmSelectReceiver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "收件人选择"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmSelectReceiver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra 
      Height          =   75
      Index           =   2
      Left            =   0
      TabIndex        =   22
      Top             =   1433
      Width           =   8145
   End
   Begin VB.CommandButton cmdFind 
      Height          =   315
      Left            =   5700
      Picture         =   "frmSelectReceiver.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1980
      Width           =   390
   End
   Begin VB.TextBox txt简码 
      Height          =   300
      Left            =   720
      TabIndex        =   20
      Top             =   1995
      Width           =   4965
   End
   Begin VB.OptionButton optPick 
      Caption         =   "指定科室(&S)"
      Height          =   195
      Index           =   7
      Left            =   4275
      TabIndex        =   19
      Top             =   1645
      Width           =   1305
   End
   Begin VB.OptionButton optPick 
      Caption         =   "人员性质(&X)"
      Height          =   195
      Index           =   5
      Left            =   270
      TabIndex        =   6
      Top             =   1101
      Width           =   1590
   End
   Begin VB.OptionButton optPick 
      Caption         =   "在线人员(&N)"
      Height          =   195
      Index           =   4
      Left            =   2205
      TabIndex        =   18
      Top             =   1645
      Width           =   1365
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   330
      TabIndex        =   17
      Top             =   6435
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3675
      TabIndex        =   16
      Top             =   6435
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4990
      TabIndex        =   15
      Top             =   6435
      Width           =   1100
   End
   Begin VB.OptionButton optPick 
      Caption         =   "指定人员(&I)"
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   7
      Top             =   1645
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.Frame fra 
      Height          =   3945
      Index           =   1
      Left            =   165
      TabIndex        =   8
      Top             =   2400
      Width           =   6045
      Begin VB.CommandButton cmdFunc 
         Caption         =   ">>"
         Height          =   350
         Index           =   3
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         ToolTipText     =   "全部新增"
         Top             =   540
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "&>"
         Height          =   350
         Index           =   2
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         ToolTipText     =   "部分新增"
         Top             =   915
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "&<"
         Height          =   350
         Index           =   1
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         ToolTipText     =   "部分移除"
         Top             =   2160
         Width           =   540
      End
      Begin VB.CommandButton cmdFunc 
         Caption         =   "<<"
         Height          =   350
         Index           =   0
         Left            =   2760
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         ToolTipText     =   "全部移除"
         Top             =   2580
         Width           =   540
      End
      Begin VB.ListBox lst 
         Height          =   3480
         Index           =   1
         ItemData        =   "frmSelectReceiver.frx":685E
         Left            =   3450
         List            =   "frmSelectReceiver.frx":6860
         TabIndex        =   12
         Top             =   270
         Width           =   2385
      End
      Begin VB.ListBox lst 
         Height          =   3480
         Index           =   0
         ItemData        =   "frmSelectReceiver.frx":6862
         Left            =   240
         List            =   "frmSelectReceiver.frx":6864
         TabIndex        =   9
         Top             =   300
         Width           =   2385
      End
   End
   Begin VB.OptionButton optPick 
      Caption         =   "本科室人员(&F)"
      Height          =   195
      Index           =   2
      Left            =   4275
      TabIndex        =   5
      Top             =   769
      Width           =   1485
   End
   Begin VB.OptionButton optPick 
      Caption         =   "本部门人员(&D)"
      Height          =   195
      Index           =   1
      Left            =   2205
      TabIndex        =   4
      Top             =   769
      Width           =   1485
   End
   Begin VB.OptionButton optPick 
      Caption         =   "所有人员(&A)"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   3
      Top             =   769
      Width           =   1365
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   4365
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -240
      TabIndex        =   0
      Top             =   557
      Width           =   8145
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "简码"
      Height          =   180
      Left            =   270
      TabIndex        =   23
      Top             =   2025
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "用户所属系统(&S)"
      Height          =   180
      Left            =   270
      TabIndex        =   1
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frmSelectReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean

Private mstr收件人 As String           '收件人的描述

Private mrs人员 As New ADODB.Recordset '保存人员清单
Private mrs系统 As New ADODB.Recordset '保存着系统

Private mrsUser As New ADODB.Recordset
Private mrsTemp As ADODB.Recordset  '保存lst(0)中的信息
Private mlngOptPick As Long  '用于存储当前选择的是哪一个optPick，主要用于区分optPick(3)和optPick(4)

Private Sub cmbSystem_Click()
    Dim strOwner As String
    On Error GoTo ErrH
    mrs系统.Filter = "编号=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrs系统("所有者")
    
    gstrSQL = "Select A.编码 As 部门编号, B.姓名, D.用户名" & vbNewLine & _
            "From " & strOwner & ".部门表 A, " & strOwner & ".部门人员 C, " & strOwner & ".上机人员表 D, " & strOwner & ".人员表 B" & vbNewLine & _
            "Where A.ID = C.部门id And B.ID = C.人员id And C.人员id = D.人员id And C.缺省 = 1 And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) " & vbNewLine & _
            "Order By B.姓名"
    Set mrs人员 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Call optPick_Click(0)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim strOwner As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    If Trim(txt简码.Text) = "" Then Exit Sub
    lst(0).Clear
    mrs系统.Filter = "编号=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrs系统("所有者")
    
    If optPick(3).Value = True Then
    
        gstrSQL = "select DISTINCT B.姓名,D.用户名 " & _
                  " from " & strOwner & ".部门表 A," & strOwner & ".人员表 B," & _
                  strOwner & ".部门人员 C," & strOwner & ".上机人员表 D " & _
                  "  where A.ID=C.部门ID and B.ID=C.人员ID and C.人员ID=D.人员ID And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) and C.缺省=1 " & _
                  " And Upper(B.简码) Like [1] order by B.姓名"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(Trim(txt简码.Text)) & "%")
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("姓名") & "(" & rsTemp("用户名") & ")"
            rsTemp.MoveNext
        Loop
    ElseIf optPick(4).Value = True Then
        gstrSQL = "select DISTINCT B.姓名,D.用户名 " & _
                      " from " & strOwner & ".部门表 A," & strOwner & ".人员表 B," & _
                      strOwner & ".部门人员 C," & strOwner & ".上机人员表 D,V$session S " & _
                      "  where A.ID=C.部门ID and B.ID=C.人员ID and C.人员ID=D.人员ID And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) and C.缺省=1 AND D.用户名=S.USERNAME " & _
                      " And Upper(B.简码) Like [1] order by B.姓名"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(Trim(txt简码.Text)) & "%")
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("姓名") & "(" & rsTemp("用户名") & ")"
            rsTemp.MoveNext
        Loop
    ElseIf optPick(7).Value = True Then
        gstrSQL = "Select Distinct A.编码,A.名称 From " & strOwner & ".部门表 A Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                  " And Upper(A.简码) Like [1] order by A.编码,A.名称"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "%" & UCase(Trim(txt简码.Text)) & "%")
        Do Until rsTemp.EOF
            lst(0).AddItem rsTemp("编码") & "-" & rsTemp("名称")
            rsTemp.MoveNext
        Loop
        
    End If
    If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    lst(0).SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Err = 0
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim intPos  As Long
    Dim strTemp As String
    Dim strOwner As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    mrs系统.Filter = "编号=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrs系统("所有者")

'    mstr用户 = ""
'    mstr姓名 = ""
    mstr收件人 = ""
    
    Set mrsUser = Rec.CopyNew(Nothing, True, , Array("用户名", adVarWChar, 30, Empty, "姓名", adVarWChar, 30, Empty, "收件人", adVarWChar, 30, Empty))

    
    If optPick(3).Value = True Or optPick(4).Value = True Then
        
        '根据列表框得到人员名单
        For i = 0 To lst(1).ListCount - 1
            If lst(1).List(i) <> "" Then
                '去掉两边的括号
                mrsUser.AddNew
                intPos = InStr(lst(1).List(i), "(")
                strTemp = Mid(lst(1).List(i), intPos + 1)
                strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
                mrsUser.Fields("用户名") = strTemp
                '括号前为用户姓名
                strTemp = Mid(lst(1).List(i), 1, intPos - 1)
                mstr收件人 = mstr收件人 & strTemp & ","
                mrsUser.Fields("姓名") = strTemp
                mrsUser.Fields("收件人") = strTemp
            End If
        Next
        If mstr收件人 <> "" Then
            mstr收件人 = Mid(mstr收件人, 1, Len(mstr收件人) - 1)
        End If
    ElseIf optPick(5).Value = True Then
        '人员性质:以分号分隔
        For i = 0 To lst(1).ListCount - 1
            mstr收件人 = mstr收件人 & lst(1).List(i) & ";"
        Next
        If mstr收件人 <> "" Then
           
            gstrSQL = "Select Distinct B.姓名, D.用户名" & vbNewLine & _
                    "From " & strOwner & ".人员性质说明 E, " & strOwner & ".上机人员表 D, " & strOwner & ".人员表 B" & vbNewLine & _
                    "Where B.ID = E.人员id And B.ID = D.人员id And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) And Instr('" & mstr收件人 & "', E.人员性质) > 0"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("用户名") = rsTemp.Fields("用户名")
                mrsUser.Fields("姓名") = rsTemp.Fields("姓名")
                rsTemp.MoveNext
            Loop
            mstr收件人 = "[" & Mid(mstr收件人, 1, Len(mstr收件人) - 1) & "]"
            
        End If
    ElseIf optPick(7).Value = True Then
        For i = 0 To lst(1).ListCount - 1
            mstr收件人 = mstr收件人 & lst(1).List(i) & ";"
        Next
        If mstr收件人 <> "" Then
            
            gstrSQL = "select DISTINCT B.姓名,D.用户名 " & _
                      " from " & strOwner & ".部门表 A," & strOwner & ".人员表 B," & _
                      strOwner & ".部门人员 C," & strOwner & ".上机人员表 D " & _
                      "  where A.ID=C.部门ID and B.ID=C.人员ID and C.人员ID=D.人员ID And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) " & _
                      "  And Instr('" & mstr收件人 & "', A.编码||'-'||A.名称 ) > 0" & _
                      " order by B.姓名"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            Do Until rsTemp.EOF
                mrsUser.AddNew
                mrsUser.Fields("用户名") = rsTemp.Fields("用户名")
                mrsUser.Fields("姓名") = rsTemp.Fields("姓名")
                rsTemp.MoveNext
            Loop
            mstr收件人 = "{" & Mid(mstr收件人, 1, Len(mstr收件人) - 1) & "}"
            
        End If
    Else
        If optPick(2).Value = True Then
        '从数据库中得到人员名单
            mstr收件人 = "本科室人员"
            mrs人员.Filter = "部门编号='" & gstrDeptCode & "'"
        ElseIf optPick(1).Value = True Then
            mstr收件人 = "本部门人员"
            If gstrDeptCode = "" Then
                mrs人员.Filter = "部门编号='无'"
            Else
                mrs人员.Filter = "部门编号 like '" & gstrDeptCode & "%'"
            End If
        Else
            mstr收件人 = "所有人员"
            mrs人员.Filter = 0
        End If
        Do Until mrs人员.EOF
            mrsUser.AddNew
            mrsUser.Fields("收件人") = mstr收件人
            mrsUser.Fields("用户名") = mrs人员("用户名")
            mrsUser.Fields("姓名") = mrs人员("姓名")
            
            mrs人员.MoveNext
        Loop
    End If
        
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdFunc_Click(Index As Integer)
    '调整指定人员的选择
    Dim strTemp As String
    Dim i, j, k As Long
    Dim lngRownum As Long '用于存储数据的rownum
    
    Select Case Index
        Case 0 '向左移全部
            If lst(1).ListCount = 0 Then Exit Sub
            lst(0).Clear
            lst(1).Clear
            mrsTemp.Filter = ""
            Do Until mrsTemp.EOF
                lst(0).AddItem mrsTemp("名称")
                mrsTemp!位置 = 0
                mrsTemp.MoveNext
            Loop
            If lst(0).ListCount > 0 And lst(1).ListIndex < 0 Then lst(0).ListIndex = 0
        Case 1 '向左移一个
            i = lst(1).ListIndex
            If i < 0 Then Exit Sub
            '获取要插入lst(0)的位置
            mrsTemp.Filter = "名称='" & lst(1).List(i) & "'"
            
            If mrsTemp.RecordCount = 1 Then
                lngRownum = mrsTemp!rownum
                mrsTemp.Filter = "位置 = 0"
                If mrsTemp.RecordCount = 0 Then
                    lst(0).AddItem lst(1).List(i)
                Else
                    For j = 1 To mrsTemp.RecordCount
                        If mrsTemp!rownum > lngRownum Then
                            '左边加一条数据
                            For k = 0 To lst(0).ListCount - 1
                                If lst(0).List(k) = mrsTemp!名称 Then
                                    Exit For
                                End If
                            Next
                            lst(0).AddItem lst(1).List(i), k
                            Exit For
                        End If
                        '当所有数据都遍历完了，还没有找到数据，那说明要添加的数据的rownum是最大的，应添加到lst(0)最后
                        If j = mrsTemp.RecordCount Then
                            '左边加一条数据
                            lst(0).AddItem lst(1).List(i)
                        End If
                        mrsTemp.MoveNext
                    Next
                End If
                '定位刚加入的那条记录，并将其位置设为0
                mrsTemp.Filter = "rownum = " & lngRownum
                mrsTemp!位置 = 0
            End If
            '右边减一条数据
            lst(1).RemoveItem i
            If i > lst(1).ListCount - 1 Then
                lst(1).ListIndex = lst(1).ListCount - 1
            Else
                lst(1).ListIndex = i
            End If
            lst(0).ListIndex = lst(0).NewIndex
        Case 2 '向右移一个
            i = lst(0).ListIndex
            If i < 0 Then Exit Sub
            mrsTemp.Filter = "名称 = '" & lst(0).List(i) & "'"
            mrsTemp!位置 = 1
            mrsTemp.Filter = "位置 = 0"
            strTemp = lst(0).List(i)
            lst(0).RemoveItem lst(0).ListIndex
            If i > lst(0).ListCount - 1 Then
                lst(0).ListIndex = lst(0).ListCount - 1
            Else
                lst(0).ListIndex = i
            End If
            lst(1).AddItem strTemp
            lst(1).ListIndex = lst(1).NewIndex
        Case 3 '向右移全部
            If lst(0).ListCount = 0 Then Exit Sub
            lst(0).Clear
            lst(1).Clear
            mrsTemp.Filter = ""
            Do Until mrsTemp.EOF
                lst(1).AddItem mrsTemp("名称")
                mrsTemp!位置 = 1
                mrsTemp.MoveNext
            Loop
            If lst(1).ListIndex < 0 And lst(1).ListCount > 0 Then lst(1).ListIndex = 0
    End Select
End Sub

Private Sub Form_Load()
    '当当前选择为人员性质时，不提供搜索功能
    If optPick(5).Value = True Then
        cmdFind.Enabled = False
        txt简码.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsTemp = Nothing
End Sub

Private Sub lst_DblClick(Index As Integer)
    If Index = 0 Then
        cmdFunc_Click 2
    Else
        cmdFunc_Click 1
    End If
End Sub

Private Sub lst_GotFocus(Index As Integer)
    cmdFunc(2).Default = True
End Sub

Private Sub lst_LostFocus(Index As Integer)
    cmdOK.Default = True
End Sub

Private Sub optPick_Click(Index As Integer)
    If mrs人员.State = 0 Then Exit Sub
    Dim strOwner As String
    Dim var收件人 As Variant, strTmp As String, i As Integer
    Dim lngOptPick As Long

    Dim blnList As Boolean
    On Error GoTo errHandle
    mrs系统.Filter = "编号=" & cmbSystem.ItemData(cmbSystem.ListIndex)
    strOwner = mrs系统("所有者")
    
    blnList = optPick(3).Value Or optPick(4).Value
    fra(1).Enabled = blnList
    lst(0).Enabled = blnList
    lst(1).Enabled = blnList
    cmdFunc(0).Enabled = blnList
    cmdFunc(1).Enabled = blnList
    cmdFunc(2).Enabled = blnList
    cmdFunc(3).Enabled = blnList
    
    cmdFind.Enabled = False
    txt简码.Enabled = False
    txt简码.Text = ""
    
    '不需要列表
    lst(0).Clear

    Set mrsTemp = New ADODB.Recordset
    
    If blnList = True Then
        If optPick(3).Value = True Then
            '从所有人员中选取，其中位置字段用于判断数据是应该放在lst(0)还是lst(1)中，0表示在lst(0)中，1表示在lst(1)中
            gstrSQL = "select rownum,姓名 || '(' || 用户名 || ')' 名称, 0 位置 from (select DISTINCT B.姓名,D.用户名 " & _
                      " from " & strOwner & ".部门表 A," & strOwner & ".人员表 B," & _
                      strOwner & ".部门人员 C," & strOwner & ".上机人员表 D " & _
                      "  where A.ID=C.部门ID and B.ID=C.人员ID and C.人员ID=D.人员ID And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') " & _
                      "Or B.撤档时间 Is Null) and C.缺省=1 order by B.姓名)"
            lngOptPick = 3
        Else
            '从在线人员中选取，其中位置字段用于判断数据是应该放在lst(0)还是lst(1)中，0表示在lst(0)中，1表示在lst(1)中
            gstrSQL = "select rownum,姓名 || '(' || 用户名 || ')' 名称, 0 位置 from (select DISTINCT B.姓名,D.用户名 " & _
                      " from " & strOwner & ".部门表 A," & strOwner & ".人员表 B," & _
                      strOwner & ".部门人员 C," & strOwner & ".上机人员表 D,V$session S " & _
                      "  where A.ID=C.部门ID and B.ID=C.人员ID and C.人员ID=D.人员ID And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') " & _
                      "Or B.撤档时间 Is Null) and C.缺省=1 AND D.用户名=S.USERNAME order by B.姓名)"
            lngOptPick = 4
        End If
        Call zlDatabase.OpenRecordset(mrsTemp, gstrSQL, Me.Caption, adOpenStatic, adLockBatchOptimistic)
        
        '若只是指定人员和在线人员之间点来点去的话，不清空lst(1)列表，只修改lst(0)列表
        If mlngOptPick = 3 Or mlngOptPick = 4 Then
            For i = 0 To lst(1).ListCount - 1
                mrsTemp.Filter = "名称 = '" & lst(1).List(i) & "'"
                If mrsTemp.RecordCount <> 0 Then
                    mrsTemp!位置 = 1
                End If
            Next
        Else
            lst(1).Clear
            If Not mrsUser Is Nothing Then
                If mrsUser.State = adStateOpen Then
                    If mrsUser.RecordCount > 0 Then mrsUser.MoveFirst
                    Do Until mrsUser.EOF
                        mrsTemp.Filter = "名称 = '" & mrsUser.Fields("姓名") & "(" & mrsUser.Fields("用户名") & ")" & "'"
                        If mrsTemp.RecordCount <> 0 Then
                            mrsTemp!位置 = 1
                        End If
                        lst(1).AddItem mrsUser.Fields("姓名") & "(" & mrsUser.Fields("用户名") & ")"
                        mrsUser.MoveNext
                    Loop
                End If
            End If
        End If
        mrsTemp.Filter = "位置 = 0"
        Do Until mrsTemp.EOF
            lst(0).AddItem mrsTemp("名称")
            mrsTemp.MoveNext
        Loop
        
        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
        cmdFind.Enabled = True
        txt简码.Enabled = True
        mlngOptPick = lngOptPick
    Else
        mlngOptPick = 0
    End If
    
    If optPick(5).Value = True Then
        lst(1).Clear
        fra(1).Enabled = True
        lst(0).Enabled = True
        lst(1).Enabled = True
        cmdFunc(0).Enabled = True
        cmdFunc(1).Enabled = True
        cmdFunc(2).Enabled = True
        cmdFunc(3).Enabled = True
        
        '从人员性质中选取，其中位置字段用于判断数据是应该放在lst(0)还是lst(1)中，0表示在lst(0)中，1表示在lst(1)中
        gstrSQL = "select rownum,编码,名称, 0 位置 from (Select 编码,名称 From " & strOwner & ".人员性质分类)"
        Call zlDatabase.OpenRecordset(mrsTemp, gstrSQL, Me.Caption, adOpenStatic, adLockBatchOptimistic)
        
        If InStr(mstr收件人, "]") > 0 And InStr(mstr收件人, "[") > 0 Then
            strTmp = Mid(mstr收件人, 2, Len(mstr收件人) - 2)
            If InStr(strTmp, ";") > 0 Then
                var收件人 = Split(strTmp, ";")
                For i = LBound(var收件人) To UBound(var收件人)
                    mrsTemp.Filter = "名称 = '" & var收件人(i) & "'"
                    mrsTemp!位置 = 1
                    lst(1).AddItem var收件人(i)
                Next
            Else
                mrsTemp.Filter = "名称 = '" & strTmp & "'"
                mrsTemp!位置 = 1
                lst(1).AddItem strTmp
            End If
        End If
        
        '填充人员性质分类到lst(0)中
        mrsTemp.Filter = "位置 = 0"
        Do Until mrsTemp.EOF
            lst(0).AddItem mrsTemp("名称")
            mrsTemp.MoveNext
        Loop
        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    End If
    
    If optPick(7).Value = True Then
        lst(1).Clear
        fra(1).Enabled = True
        lst(0).Enabled = True
        lst(1).Enabled = True
        cmdFunc(0).Enabled = True
        cmdFunc(1).Enabled = True
        cmdFunc(2).Enabled = True
        cmdFunc(3).Enabled = True
        cmdFind.Enabled = True
        txt简码.Enabled = True
        
        '从指定科室中选取，其中位置字段用于判断数据是应该放在lst(0)还是lst(1)中，0表示在lst(0)中，1表示在lst(1)中
        gstrSQL = "select rownum,编码 || '-' || 名称 名称, 0 位置 from (Select Distinct A.编码,A.名称 From " & strOwner & _
                ".部门表 A Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
                " order by A.编码,A.名称)"
        Call zlDatabase.OpenRecordset(mrsTemp, gstrSQL, Me.Caption, adOpenStatic, adLockBatchOptimistic)

        If InStr(mstr收件人, "}") > 0 And InStr(mstr收件人, "{") > 0 Then
            strTmp = Mid(mstr收件人, 2, Len(mstr收件人) - 2)
            If InStr(strTmp, ";") > 0 Then
                var收件人 = Split(strTmp, ";")
                For i = LBound(var收件人) To UBound(var收件人)
                    mrsTemp.Filter = "名称 = '" & var收件人(i) & "'"
                    mrsTemp!位置 = 1
                    lst(1).AddItem var收件人(i)
                Next
            Else
                mrsTemp.Filter = "名称 = '" & strTmp & "'"
                mrsTemp!位置 = 1
                lst(1).AddItem strTmp
            End If
        End If

        '填充科室列表到lst(0)中
        mrsTemp.Filter = "位置 = 0"
        Do Until mrsTemp.EOF
            lst(0).AddItem mrsTemp("名称")
            mrsTemp.MoveNext
        Loop
        If lst(0).ListCount > 0 Then lst(0).ListIndex = 0
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function Get收件人(str收件人 As String, rsUser As ADODB.Recordset) As Boolean
    
    Dim var收件人 As Variant, strTmp As String, i As Integer
    On Error GoTo errHandle
    mblnOK = False
    mstr收件人 = str收件人
    
    Set mrsUser = rsUser
    '-----------------------------------
    '根据传来的参数进行显示
    lst(1).Clear
    Select Case str收件人
        Case "所有人员"
            optPick(0).Value = True
        Case "本部门人员"
            optPick(1).Value = True
        Case "本科室人员"
            optPick(2).Value = True
        Case Else
            If mlngOptPick = 3 Or str收件人 = "" Then
                optPick(3).Value = True
            ElseIf mlngOptPick = 4 Then
                optPick(4).Value = True
            End If
            If InStr(str收件人, "[") > 0 And InStr(str收件人, "]") > 0 Then
                optPick(5).Value = True
            ElseIf InStr(str收件人, "{") > 0 And InStr(str收件人, "}") > 0 Then
                optPick(7).Value = True
            End If
            
            If optPick(5).Value = True Or optPick(7).Value = True Then
                '病人性质,指定科室
                lst(1).Clear
                strTmp = Mid(str收件人, 2, Len(str收件人) - 2)
                If InStr(strTmp, ";") > 0 Then
                    var收件人 = Split(strTmp, ";")
                    For i = 0 To UBound(var收件人)
                        lst(1).AddItem var收件人(i)
                    Next
                Else
                    lst(1).AddItem strTmp
                End If
            Else
                If Not rsUser Is Nothing Then
                    If rsUser.State = adStateOpen Then
                        If rsUser.RecordCount > 0 Then rsUser.MoveFirst
                        Do Until rsUser.EOF
                            lst(1).AddItem rsUser.Fields("姓名") & "(" & rsUser.Fields("用户名") & ")"
                            rsUser.MoveNext
                        Loop
                    End If
                End If
            End If
            If lst(1).ListCount > 0 Then lst(1).ListIndex = 0
    End Select
    
    '得到系统
    gstrSQL = "select A.编号,A.名称 ||'（'||A.编号||'）' as 名称,A.所有者 from zlsystems A, (select owner from all_tables where " & _
               " table_name in ('部门表','人员表','部门人员','上机人员表') " & _
               " group by owner " & _
               " having count(table_name)=4) B " & _
               " Where A.所有者 = B.owner"
    Call zlDatabase.OpenRecordset(mrs系统, gstrSQL, Me.Caption)
    
    If mrs系统.EOF Then
        MsgBox "你不具有选择收件人的权限，不能使用本功能。", vbInformation, gstrSysName
        Exit Function
    End If
    cmbSystem.Clear
    Do Until mrs系统.EOF
        cmbSystem.AddItem mrs系统("名称")
        cmbSystem.ItemData(cmbSystem.NewIndex) = mrs系统("编号")
        mrs系统.MoveNext
    Loop
    
    If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
    If cmbSystem.ListCount = 1 Then cmbSystem.Enabled = False
    
    '通过cmbSystem的选择已经得到人员清单
    
    frmSelectReceiver.Show vbModal
    Get收件人 = mblnOK
    If mblnOK = True Then
        str收件人 = mstr收件人
        Set rsUser = mrsUser
    End If
    If mrs人员.State = 1 Then mrs人员.Close
    Set mrs人员 = Nothing
    If mrs系统.State = 1 Then mrs系统.Close
    Set mrs系统 = Nothing
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub txt简码_GotFocus()
    cmdFind.Default = True
End Sub

Private Sub txt简码_LostFocus()
    cmdOK.Default = True
End Sub
