VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmQueueShow 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timerLCD 
      Interval        =   2000
      Left            =   7200
      Top             =   1920
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgQueuingData 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11775
      _cx             =   20770
      _cy             =   13150
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   65280
      BackColorFixed  =   16777215
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   0
      BackColorAlternate=   0
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483627
      TreeColor       =   -2147483633
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmQueueShow.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label labInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5835
      TabIndex        =   2
      Top             =   8880
      Width           =   105
   End
   Begin VB.Label labTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5850
      TabIndex        =   1
      Top             =   0
      Width           =   105
   End
End
Attribute VB_Name = "frmQueueShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngQueuingItemsCount As Long    '排队队列显示的数量
Private mlngQueuingDocItemsCount As Long    '排队队列每个医生下显示的数量
Private mlngLedLoopQueryTime As Long '轮询间隔时间长度
Private mlngQueuingingColor As Long     '排队中颜色
Private mlngCallingColor As Long        '呼叫中颜色
Private mlngEmergColor As Long      '急诊字体显示色
Private mstrGreeting As String
Private mlngTitleColor As Long
Private mlngVisitColor As Long
Private mlngCalledColor As Long

Private mstr队列名称() As String
Private mint有效天数 As Integer
Private mstr诊室条件 As String, mstr医生条件 As String, mstrExcludeData As String

Private mintViewDataType As Integer '数据显示类型
Private mblnComeBackFirst As Boolean    '回诊病人是否优先排队
Private mstrDelString As String '需要删除的字符，使用“,”号分割
Private mlngRoomsCount As Long      '每页显示的科室数量
Private mlngCurPageIndex As Long   '当前显示的页索引
Private mlngPageSwitchTime As Long  '页面切换时间长度
Private mblnStartSwitch As Boolean  '是否开始切换页面

Private mlngQueueId() As Long      '各个科室的当前显示队列ID
Private mstrQueueKey() As String
Private mlngBackColor As Long      '背景颜色设置

'判断数组是否为空
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Public Function zlShowMe(cnOracle As ADODB.Connection, str队列名称() As String, _
    Optional str诊室 As String = "", Optional str医生 As String = "", _
    Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0, _
    Optional blnComeBackFirst As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：显示排队情况
    '入参：str队列名称():传入的指定队列数组(从1开始)
    '         strCur队列名称-当前队列名称
    '         lngCur业务ID-业务ID
    '         str诊室-限制为指定的诊室,可以为多个诊室:如"一诊室,二诊室,..."
    '         str医生-限制为制定的医生,可以传多个医生,用逗号分隔,如"张三,李四,..."
    '         strExcludeData-排队的指定业务ID
    '         intViewDataType数据显示类型(由医生站的"接诊范围"来控制)，0显示当前科室下的所有数据，
    '                                      1显示诊室为当前科室，或者医生姓名等于当前医生，或者诊室为空和医生为空的数据
    '                                      2显示诊室为当前诊室，或医生姓名等于当前医生的数据
    '                                      3显示当前医生的数据
    '         blnComebackFirst回诊病人是否优先排队
    '编制：刘兴洪
    '日期：2010-06-11 20:54:55
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    mstr队列名称 = str队列名称
    mstr诊室条件 = str诊室
    mstr医生条件 = str医生
    mstrExcludeData = strExcludeData
    
    '将intViewDataType=1的情况处理为0
    If intViewDataType = 1 Then
        mintViewDataType = 0
    Else
        mintViewDataType = intViewDataType
    End If
    mblnComeBackFirst = blnComeBackFirst
    Call GetDepartKey(mstr队列名称, mstrQueueKey)
    Me.Show
End Function

Public Function zlSetPara(str队列名称() As String, _
    Optional str诊室 As String = "", Optional str医生 As String = "", _
    Optional strExcludeData As String, Optional blnComeBackFirst As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：显示排队情况
    '入参：str队列名称():传入的指定队列数组(从1开始)
     '         str诊室-限制为指定的诊室,可以为多个诊室:如"一诊室,二诊室,..."
    '         str医生-限制为制定的医生,可以传多个医生,用逗号分隔,如"张三,李四,..."
    '         strExcludeData-排队的指定业务ID
    '编制：刘兴洪
    '日期：2010-06-11 20:54:55
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    mstr队列名称 = str队列名称
    mstr诊室条件 = str诊室
    mstr医生条件 = str医生
    mstrExcludeData = strExcludeData
    mblnComeBackFirst = blnComeBackFirst
    Call GetDepartKey(mstr队列名称, mstrQueueKey)
    
End Function

Private Sub MultiRoomsDisplay()
'**************************************************************************
'显示多个科室的叫号信息
'**************************************************************************
    Dim i As Integer, j As Integer
    Dim intCurPageRoomIndex As Integer '保存当前屏幕页的科室索引
    Dim blnSwitchPage As Boolean
    Dim blnAllowRoll As Boolean    '判断当前所显示的数据是否需要滚动，如果数据条数小于当前能够显示的记录数量，则不滚动
    Dim strCurRoomKey As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsSource As ADODB.Recordset   '呼叫数据
    Dim strValues(0 To 10) As String, strValue As String, strUninTable As String
    Dim strFilter As String, strCurExcludeData() As String
    Dim strDoc As String, intPati As Integer
    
    err = 0: On Error GoTo errHandle
    If SafeArrayGetDim(mstr队列名称) > 0 Then
        j = 0
        
        strFilter = ""
        strValue = ""
        strUninTable = ""
        
        For i = 1 To UBound(mstr队列名称)
            If j > 10 Then
                strFilter = strFilter & " Or A.队列名称 ='" & mstr队列名称(i) & "'"
            Else
                If gobjCommFun.ActualLen(strValue) > 2000 Then
                     strValues(j) = Mid(strValue, 2)
                     strUninTable = strUninTable & " Union ALL  Select  Column_Value as 队列名称 From Table(Cast(f_Str2list([" & j + 3 & "]) As zlTools.t_Strlist))  " & vbCrLf
                     strValue = "": j = j + 1
                End If
                strValue = strValue & "," & mstr队列名称(i)
            End If
        Next i
        If strValue <> "" Then
            strValues(j) = Mid(strValue, 2)
            strUninTable = strUninTable & " Union ALL  Select  Column_Value as 队列名称 From Table(Cast(f_Str2list([" & j + 3 & "]) As zlTools.t_Strlist))  " & vbCrLf
        End If
    End If
        
    
    If strUninTable <> "" Then
        strUninTable = Mid(strUninTable, 11)
    Else
       strUninTable = " Select  Column_Value as 队列名称 From Table(Cast(f_Str2list([3]) As zlTools.t_Strlist)) "
    End If
    If strFilter <> "" Then strFilter = "( " & Mid(strFilter, 4) & ")"
    
    '从数据库中查询队列排队情况 将a.排队标记 || to_char(a.排队号码,'FM0000') As 号码 去掉排队标记;程序中强制去掉电诊科执行科室
    '0:排队中，1:呼叫中，2：已弃号(接诊)，3：暂停，4：完成就诊，6：回诊，7：已呼叫
    '曾明春(20150715):眉山市人民医院要求增加正在就诊的显示
    strSQL = "Select /*+ Rule*/  to_Number(a.ID) as ID, a.队列名称, b.名称 as 科室名称, to_char(a.排队号码,'FM0000') As 号码,to_number(a.排队号码) 排队号码," & _
             "a.患者姓名, a.医生姓名, a.诊室,To_Char(m.发生时间,'HH24:MI') As 候诊时间,m.急诊,m.预约,m.NO,m.性别,m.年龄,r.专业技术职务, " & _
             "decode(a.排队状态,0,'候诊中',1,'呼叫中',2,'就诊中',3,'暂停',4,'完成就诊',6,'回诊',7,'已呼叫') as 排队状态, to_Number(a.优先) as 优先, a.排队时间, a.呼叫时间, to_Number(回诊序号) as 回诊序号, to_Number(a.业务类型) as 业务类型, a.业务ID, " & _
              IIf(mblnComeBackFirst, "to_Number(Nvl(a.回诊序号, 9999999999)) as 回诊排序号", "0 as 回诊排序号") & _
             " From 排队叫号队列 a, 部门表 b ,病人挂号记录 m,人员表 r, (" & strUninTable & ") E " & _
                IIf(mintViewDataType = 1, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 2, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 3, ", Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D", "") & _
             " Where a.科室id = b.Id And a.业务ID=m.ID And a. 队列名称=E.队列名称  And (a.排队状态 in (0, 1, 7) Or (a.排队状态=2 and m.执行状态<>0)) and nvl(m.记录标志,0)=0 and 排队时间 <= trunc(sysdate + 1) - 1/24/60/60 And 业务类型=0" & _
                IIf(mintViewDataType = 1, " and  ((a.诊室=C.Column_Value and a.医生姓名 is null) or a.医生姓名=D.Column_Value or (a.诊室 is null and a.医生姓名 is null))", "") & _
                IIf(mintViewDataType = 2, " and ((a.诊室=C.Column_Value or a.医生姓名=D.Column_Value) ", "") & _
                IIf(mintViewDataType = 3, " and a.医生姓名=D.Column_Value", "") & " and r.姓名=a.医生姓名" & _
             " Order By 医生姓名,排队状态 desc,a.优先 desc, 回诊排序号, to_number(a.排队号码) "

    Set rsSource = gobjDatabase.OpenSQLRecord(strSQL, "显示排队情况", mstr诊室条件, mstr医生条件, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    Set rsTemp = gobjDatabase.CopyNewRec(rsSource)
    
    If rsTemp.EOF Then
        vfgQueuingData.Cell(flexcpText, 2, 0, vfgQueuingData.Rows - 1, vfgQueuingData.Cols - 1) = ""
        Exit Sub
    End If
     
    While Not rsTemp.EOF
        If InStr(1, mstrExcludeData, rsTemp!业务类型 & ":" & rsTemp!业务ID) > 0 Then
            rsTemp.Delete
        End If
        rsTemp.MoveNext
    Wend
    
    '注释如下语句,不允许在实时显示的时候，根据实际的文字长度调整列的宽度
'    Call vfgQueuingData.AutoSize(0, 2)

    '显示当前科室候诊的数据
    If SafeArrayGetDim(mstrQueueKey) <= 0 Then
        Exit Sub
    End If
    
    '如果mstrQueueKey的科室全部为空，则不执行
    For i = 1 To UBound(mstrQueueKey)
        If Trim(mstrQueueKey(i)) <> "" Then
            GoTo Start
        End If
    Next i
    
    Exit Sub
    
Start:
    
    '当页面将数据全部显示后，再计算继续显示的时间
    If mblnStartSwitch Then
        If mlngPageSwitchTime > 0 Then
            mlngPageSwitchTime = mlngPageSwitchTime - 1
            Exit Sub
        Else
            mblnStartSwitch = False
        End If
    End If
    
    blnSwitchPage = True

    '设置候诊队列表格合并方式,仅表头合并
    vfgQueuingData.MergeCellsFixed = flexMergeFree
    vfgQueuingData.MergeCells = flexMergeNever
    
    '遍历当需要显示的科室名称，并读取信息显示
    For i = (mlngCurPageIndex - 1) * mlngRoomsCount + 1 To mlngCurPageIndex * mlngRoomsCount
        
        '取得当前页面对应的科室列索引
        intCurPageRoomIndex = i - (mlngCurPageIndex - 1) * mlngRoomsCount
        
        '设置LCD显示的科室名称
        If i <= UBound(mstrQueueKey) Then
            strCurRoomKey = mstrQueueKey(i)
            vfgQueuingData.Cell(flexcpText, 0, (intCurPageRoomIndex - 1) * 5 + 1) = strCurRoomKey
           '显示科室时,将第一列合并显示
            Call vfgMergeRowCol(0, (intCurPageRoomIndex - 1) * 5 + 1)
        Else
            strCurRoomKey = ""
            vfgQueuingData.Cell(flexcpText, 0, (intCurPageRoomIndex - 1) * 5 + 1) = ""
            '显示科室时,将第一列合并显示
            Call vfgMergeRowCol(0, (intCurPageRoomIndex - 1) * 5 + 1)
        End If
        
        If Trim(strCurRoomKey) <> "" Then
                        
            '过滤出当前科室所需要显示的候诊数据
            'rsTemp.Filter = "排队状态='排队中' and 科室名称='" & strCurRoomKey & "' and ID>" & mlngQueueId(intCurPageRoomIndex)
            'rsTemp.Filter = "科室名称='" & strCurRoomKey & "' and ID>" & mlngQueueId(intCurPageRoomIndex)
            'rsTemp.Sort = "医生姓名,排队状态,优先 desc, 回诊排序号 asc, 排队时间 asc, 排队号码 asc"                          '此处控制显示顺序
            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
            
            '如果记录条数小于当前能够显示的记录条数，且允许滚动，则blnAllowRoll可直接设置为true
            blnAllowRoll = IIf(rsTemp.RecordCount <= mlngQueuingItemsCount, False, True)  'blnAllowRoll = True
        
            If Not rsTemp.EOF Then
                '当实际记录数小于每页能够显示的记录时，则不进行滚动
                If blnAllowRoll Then
                    mlngQueueId(intCurPageRoomIndex) = rsTemp!ID
                    
                    '如果没有结束，则不允许翻页
                    blnSwitchPage = False
                End If
                
                j = 0
                '显示指定条数的记录，如果记录数据不够显示，则显示空数据
                While j < mlngQueuingItemsCount
                    If Not rsTemp.EOF Then
                        If strDoc <> rsTemp!医生姓名 Then       '单独一行显示医生
                            Call SetRoomsData(1, j + 2, intCurPageRoomIndex, "", "", "", "", ""): j = j + 1
                            Call SetRoomsData(1, j + 2, intCurPageRoomIndex, " 医生：" & Nvl(rsTemp!医生姓名) & "  " & Nvl(rsTemp!专业技术职务), " 医生：" & Nvl(rsTemp!医生姓名) & "  " & Nvl(rsTemp!专业技术职务), " 医生：" & Nvl(rsTemp!医生姓名) & "  " & Nvl(rsTemp!专业技术职务), " 医生：" & Nvl(rsTemp!医生姓名) & "  " & Nvl(rsTemp!专业技术职务), " 医生：" & Nvl(rsTemp!医生姓名) & "  " & Nvl(rsTemp!专业技术职务))
                            Call vfgMergeRowCol(j + 2, (intCurPageRoomIndex - 1) * 5 + 1)
                            j = j + 1: strDoc = rsTemp!医生姓名: intPati = 1
                        End If
                        If intPati - 1 < mlngQueuingDocItemsCount Then
                            Call SetRoomsData(2, j + 2, intCurPageRoomIndex, Nvl(rsTemp!排队号码), Nvl(rsTemp!患者姓名), Nvl(rsTemp!性别), Nvl(rsTemp!年龄), Nvl(rsTemp!排队状态) & " " & IIf(Nvl(rsTemp!急诊) = "1", "急诊", IIf(Nvl(rsTemp!预约) = "1", "预约", "")))
                            intPati = intPati + 1
                            If Nvl(rsTemp!急诊) = "1" Then
                                Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngEmergColor)
                            Else
                                Select Case Nvl(rsTemp!排队状态)
                                    Case "呼叫中"
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngCallingColor)
                                    Case "候诊中"
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngQueuingingColor)
                                    Case "就诊中"
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngVisitColor)
                                    Case "已呼叫"
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngCalledColor)
                                    Case Else
                                        Call SetRoomsColor(j + 2, intCurPageRoomIndex, mlngQueuingingColor)
                                End Select
                            End If
                            
                            j = j + 1
                        End If
                        rsTemp.MoveNext
                    Else
                        Call SetRoomsData(2, j + 2, intCurPageRoomIndex, "", "", "", "", "")
                        j = j + 1
                    End If
                    DoEvents
                Wend
            Else
                For j = 0 To mlngQueuingItemsCount - 2
                    Call SetRoomsData(2, j + 3, intCurPageRoomIndex, "", "", "", "", "")
                    DoEvents
                Next j
            End If
        Else
            '如果没有对应的科室可显示，则使数据为空
            For j = 0 To mlngQueuingItemsCount - 1
                Call SetRoomsData(2, j + 3, intCurPageRoomIndex, "", "", "", "", "")
                DoEvents
            Next j
        End If
        DoEvents
    Next i
           
    '如果blnSwitchPage为真，则进入下一个页面的显示
    If blnSwitchPage Then
        mlngCurPageIndex = mlngCurPageIndex + 1
        
        Dim intPageCount As Integer
        '计算页面数据
        If UBound(mstrQueueKey) Mod mlngRoomsCount <> 0 Then
            intPageCount = Int(UBound(mstrQueueKey) / mlngRoomsCount) + 1
        Else
            intPageCount = UBound(mstrQueueKey) / mlngRoomsCount
        End If
        
        '判断是否已经显示了最后的页，如果是，则重新显示第一页
        If mlngCurPageIndex > intPageCount Then
            mlngCurPageIndex = 1
        End If
        
        For i = 1 To UBound(mlngQueueId)
            mlngQueueId(i) = -1
        Next i
        
        mblnStartSwitch = True
        mlngPageSwitchTime = 3  '重新设置页面继续显示的时间长度
    End If
    
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    
    Exit Sub
errHandle:
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strReg As String
    
    strReg = "公共模块\排队叫号\液晶电视"
    
    mlngRoomsCount = CLng(GetSetting("ZLSOFT", strReg, "页面显示列", "1"))
    
    mlngLedLoopQueryTime = Val(GetSetting("ZLSOFT", strReg, "LED轮询时间", "2"))
    timerLCD.Interval = mlngLedLoopQueryTime * 1000
    mlngQueuingItemsCount = Val(GetSetting("ZLSOFT", strReg, "排队记录显示数", "6"))
    mlngQueuingDocItemsCount = Val(GetSetting("ZLSOFT", strReg, "排队记录显示数", "6"))
    mlngQueuingingColor = GetSetting("ZLSOFT", strReg, "排队中颜色", vbGreen)
    mlngCallingColor = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
    mlngVisitColor = GetSetting("ZLSOFT", strReg, "就诊中颜色", vbGreen)
    mlngCalledColor = GetSetting("ZLSOFT", strReg, "已呼叫颜色", vbGreen)
     
    mlngEmergColor = GetSetting("ZLSOFT", strReg, "急诊颜色", vbRed)
    mstrGreeting = GetSetting("ZLSOFT", strReg, "祝福语", "祝你早日康复！")
    mlngTitleColor = GetSetting("ZLSOFT", strReg, "标题颜色", vbRed)
    
    mlngCurPageIndex = 1
    mlngPageSwitchTime = 3 '10秒
    mblnStartSwitch = False
    
    Me.BackColor = vbBlack
    
    '设置数组大小
    ReDim mlngQueueId(mlngRoomsCount)
    ReDim mstrQueueKey(mlngRoomsCount)
    
    For i = 1 To mlngRoomsCount
      mlngQueueId(i) = -1
      mstrQueueKey(i) = ""
    Next i
    
    Call GetDepartKey(mstr队列名称, mstrQueueKey)
     
    '设置显示字体
    Call SetFaceFont
    '设置显示位置
    Call SetFacePostion
    '设置背景颜色
    Call SetBackColor
    
    Call InitFace(mlngRoomsCount, mlngQueuingItemsCount)
End Sub

Private Sub Form_Resize()
    Call InitFace(mlngRoomsCount, mlngQueuingItemsCount)
End Sub

Public Sub SetFaceFont()
'************************************************************************************
'设置界面显示的字体样式
'************************************************************************************
    Dim strReg As String
    Dim curFontSize As Currency
    
    On Error Resume Next
    '从注册表中，读取显示参数
    strReg = "公共模块\排队叫号\液晶电视"

    vfgQueuingData.Font.Name = GetSetting("ZLSOFT", strReg, "字体", "宋体")
    vfgQueuingData.Font.Bold = GetSetting("ZLSOFT", strReg, "粗体", "False")
    vfgQueuingData.Font.Italic = GetSetting("ZLSOFT", strReg, "斜体", "False")
    
    curFontSize = GetSetting("ZLSOFT", strReg, "字号", "14")
    vfgQueuingData.Font.Size = curFontSize * 4.5 / 5
    
    
    labTitle.Font.Name = vfgQueuingData.Font.Name
    labTitle.Font.Bold = vfgQueuingData.Font.Bold
    labTitle.Font.Italic = vfgQueuingData.Font.Italic
    labTitle.Font.Size = curFontSize + 1
    labTitle.Caption = GetSetting("ZLSOFT", "注册信息", "单位名称", "") & "门诊病人候诊一览表"
    labTitle.ForeColor = mlngTitleColor
    
    labInfo.Font.Name = vfgQueuingData.Font.Name
    labInfo.Font.Bold = vfgQueuingData.Font.Bold
    labInfo.Font.Italic = vfgQueuingData.Font.Italic
    labInfo.Font.Size = curFontSize + 0.5
    labInfo.ForeColor = mlngTitleColor
End Sub

Public Sub SetFacePostion()
'************************************************************************************
'
'设置界面的显示位置
'
'
'************************************************************************************
    Dim strReg As String
    
    On Error Resume Next
        
    '从注册表中，读取显示参数
    strReg = "公共模块\排队叫号\液晶电视"
    
    '设置显示参数
    Me.Left = GetSetting("ZLSOFT", strReg, "排队屏幕左", "1024") * Screen.TwipsPerPixelX
    Me.Top = GetSetting("ZLSOFT", strReg, "排队屏幕顶", "0") * Screen.TwipsPerPixelY
    Me.Width = GetSetting("ZLSOFT", strReg, "宽度", "1024") * Screen.TwipsPerPixelX
    Me.Height = GetSetting("ZLSOFT", strReg, "高度", "768") * Screen.TwipsPerPixelY
    
End Sub

Private Sub InitFace(ByVal lngRoomsCount As Long, ByVal lngWaitItemsCount As Long)
'************************************************************************************
'初始化界面显示
'lngRoomsCount: 每屏幕能够显示的科室数量
'lngWaitItemsCount：候诊队列显示的行数量
'************************************************************************************
    Dim i As Integer
    
    On Error GoTo errHandle
    
    '设置排队显示-------------------------------------
    labTitle.Top = Round(Me.ScaleHeight * 0.01)
    labTitle.Left = 0
    'labTitle.Height = Round(Me.ScaleHeight * 0.042)
    labTitle.Width = Me.ScaleWidth
    
    'vfgQueuingData.Top = Round(Me.ScaleHeight * 0.052)
    vfgQueuingData.Top = labTitle.Height + 100
    vfgQueuingData.Left = 100
    vfgQueuingData.Width = Me.ScaleWidth - 300
    'vfgQueuingData.Height = Round(Me.ScaleHeight * 0.907)
    vfgQueuingData.Height = Me.ScaleHeight - labTitle.Height - labInfo.Height - 200

    'labInfo.Top = Round(Me.ScaleHeight * 0.961)
    labInfo.Top = labTitle.Height + vfgQueuingData.Height + 100
    labInfo.Left = 0
    labInfo.Width = Me.ScaleWidth
    'labInfo.Height = Round(Me.ScaleHeight * 0.039)
    
    vfgQueuingData.Cols = lngRoomsCount * 5
    vfgQueuingData.Rows = Int(vfgQueuingData.Height / vfgQueuingData.Cell(flexcpHeight, 0, 0))
    
    '自动设置可显示的行数
    mlngQueuingItemsCount = vfgQueuingData.Rows - 2
    vfgQueuingData.ForeColor = mlngQueuingingColor
    vfgQueuingData.Enabled = False
    
    '设置候诊显示样式--------------------------------------
    For i = 0 To lngRoomsCount - 1
        vfgQueuingData.ColWidth(i * 4 + 0 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 0.75)
        vfgQueuingData.ColWidth(i * 4 + 1 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 1.1)
        vfgQueuingData.ColWidth(i * 4 + 2 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 0.85)
        vfgQueuingData.ColWidth(i * 4 + 3 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 0.9)
        vfgQueuingData.ColWidth(i * 4 + 4 + i) = Int(vfgQueuingData.Width / vfgQueuingData.Cols * 1.4)
        
        '设置科室的显示栏目为居中对齐
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 0 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 1 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 2 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 3 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 0, i * 4 + 4 + i) = flexAlignCenterCenter

        '设置数据列名称显示栏目为居中对齐
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 0 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 1 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 2 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 3 + i) = flexAlignCenterCenter
        vfgQueuingData.Cell(flexcpAlignment, 1, i * 4 + 4 + i) = flexAlignCenterCenter
        
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 0 + i) = "  序号  "
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 1 + i) = "  姓名  "
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 2 + i) = "  性别  "
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 3 + i) = "  年龄  "
        vfgQueuingData.Cell(flexcpText, 1, i * 4 + 4 + i) = "  就诊状态  "
    Next i
    
    Call DrawBorder
        
    '显示信息--------------------------------------
    If Split(mstrGreeting & "|", "|")(1) <> "" Then
        labInfo.Caption = Split(mstrGreeting & "|", "|")(0) & " " & Format(Now, "yyyy-mm-dd hh:mm") & " 星期" & GetTodayNum & " " & Split(mstrGreeting & "|", "|")(1)
    Else
        labInfo.Caption = Split(mstrGreeting & "|", "|")(0) & " " & Format(Now, "yyyy-mm-dd hh:mm") & " 星期" & GetTodayNum
    End If
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Function GetTodayNum()
    On Error Resume Next
    
    Select Case Weekday(Date, vbMonday)
        Case 1: GetTodayNum = "一"
        Case 2: GetTodayNum = "二"
        Case 3: GetTodayNum = "三"
        Case 4: GetTodayNum = "四"
        Case 5: GetTodayNum = "五"
        Case 6: GetTodayNum = "六"
        Case 7: GetTodayNum = "日"
    End Select
End Function

Public Sub timerLCD_Timer()
        On Error GoTo errHandle
        Dim blnTimer As Boolean

        blnTimer = timerLCD.Enabled
        timerLCD.Enabled = False
        
        labInfo.Caption = Split(mstrGreeting & "|", "|")(0) & "  " & Format(Now, "yyyy-mm-dd  hh:mm") & "  星期" & GetTodayNum & "  " & Split(mstrGreeting & "|", "|")(1)
        
        Call MultiRoomsDisplay
        
        timerLCD.Enabled = blnTimer
    Exit Sub
errHandle:
    Call gobjComLib.SaveErrLog
    
    timerLCD.Enabled = blnTimer
End Sub

Private Function setTextWidth(strText As String, iLen As Integer, intWay As Integer) As String
'**************************************************************************
'
'设置文本长度达到制定长度，如果不足，则补充空格
'
'strText：需要设置的文本串
'
'iLen：文本长度
'
'intWay：对齐方向
'
'**************************************************************************
    
    On Error GoTo errHandle
    
    If Len(strText) >= iLen Then
        setTextWidth = Mid(strText, 1, iLen)
        Exit Function
    End If
    
    Select Case intWay
      Case 1
        setTextWidth = Space(iLen - Len(strText)) & strText
      Case 2
        setTextWidth = strText & Space(iLen - Len(strText))
      Case 3
        setTextWidth = Space((iLen - Len(strText)) - Int((iLen - Len(strText)) / 2)) & strText & Space(Int((iLen - Len(strText)) / 2))
    End Select
    
    Exit Function
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Private Sub GetDepartKey(str队列名称() As String, strDepartKey() As String)
'**************************************************************************
'
'取得排队叫号系统中涉及到的科室名称
'
'str队列名称()：叫号系统中的队列名称数组
'
'strDepartKey()：保存科室名称
'
'**************************************************************************

    Dim strSQL As String
    Dim i As Integer
    Dim rsDepart As ADODB.Recordset
    Dim strDepartId As String
    
    On Error GoTo errHandle
    
    If SafeArrayGetDim(str队列名称) <= 0 Then
        Exit Sub
    End If
    
    If UBound(str队列名称) <= 0 Then
        Exit Sub
    End If
    
    '取得需要检索的科室ID
    strDepartId = ""
    For i = 1 To UBound(str队列名称)
        If Trim(str队列名称(i)) <> "" Then
            If Trim(strDepartId) <> "" Then strDepartId = strDepartId & ","
            strDepartId = strDepartId & Mid(str队列名称(i) & ":", 1, InStr(1, str队列名称(i) & ":", ":") - 1)
        End If
    Next i
    
    If Trim(strDepartId) = "" Then
      Exit Sub
    End If
    
    strSQL = "select /*+ Rule*/ distinct 名称, id from 部门表 a, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) b where a.id =b.Column_Value order by Id"
    Set rsDepart = gobjDatabase.OpenSQLRecord(strSQL, "读取科室信息", strDepartId)

    
    If rsDepart.RecordCount <= 0 Then
        Exit Sub
    End If
    
    ReDim strDepartKey(rsDepart.RecordCount)
    
    For i = 1 To rsDepart.RecordCount
        strDepartKey(i) = rsDepart!名称
        rsDepart.MoveNext
    Next i
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub vfgMergeRowCol(ByVal lngrow As Long, ByVal lngcol As Long)
'2010-07-09 ZHQ 强制第一行指定列前后两列进行合并，务必在将此三个单元格内容填写为相同内容
'               直接以第(0,lngcol)数据为准进行合并
    
    Dim strTemp As String
    strTemp = vfgQueuingData.TextMatrix(lngrow, lngcol)
    If strTemp = "" Then strTemp = " "
    
    With vfgQueuingData
        .TextMatrix(lngrow, lngcol - 1) = strTemp
        .TextMatrix(lngrow, lngcol) = strTemp
        .TextMatrix(lngrow, lngcol + 1) = strTemp
        .TextMatrix(lngrow, lngcol + 2) = strTemp
        .TextMatrix(lngrow, lngcol + 3) = strTemp
        .MergeRow(lngrow) = True
        .MergeCol(lngcol - 1) = True
        .MergeCol(lngcol) = True
        .MergeCol(lngcol + 1) = True
        .MergeCol(lngcol + 2) = True
        .MergeCol(lngcol + 3) = True
        .MergeCells = flexMergeRestrictRows
    End With
End Sub

Private Sub DrawBorder()
'**************************************************************************
'绘制表头边框
'**************************************************************************

    Dim i As Integer
    
    On Error GoTo errHandle
    
    For i = 0 To mlngRoomsCount - 1
        With vfgQueuingData
            .Select 0, i * 5, 0, i * 5 + 4
            .CellBorder vbWhite, 1, 1, 1, 1, 0, 0
        End With
    Next i
    
    For i = 0 To vfgQueuingData.Cols - 1
        With vfgQueuingData
            .Select 1, i, 1, i
            .CellBorder vbWhite, 1, 1, 1, 1, 0, 0
        End With
    Next i
    
    For i = 0 To mlngRoomsCount - 1
        With vfgQueuingData
            .Select 2, i * 5, .Rows - 1, i * 5 + 4
            .CellBorder vbWhite, 1, 1, 1, 1, 0, 0
        End With
    Next i
    
    Exit Sub
errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub SetRoomsData(ByVal Align As Integer, ByVal intRowIndex As Integer, ByVal intRoomIndex As Integer, _
    ByVal strName As String, ByVal strSex As String, ByVal strAge, ByVal strState As String, ByVal strDocter As String)
'**************************************************************************
'设置显示数据
'intRowIndex：当前数据显示行
'intRoomIndex：当前显示科室索引，从1开始
'strTime： 时间，strName：姓名，strSex：
'**************************************************************************
    On Error GoTo errHandle
        
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5) = False
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5 + 1) = False
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5 + 2) = False
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5 + 3) = False
        vfgQueuingData.MergeCol((intRoomIndex - 1) * 5 + 4) = False
        
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5) = strName
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5 + 1) = strSex
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5 + 2) = strAge
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5 + 3) = strState
        vfgQueuingData.Cell(flexcpText, intRowIndex, (intRoomIndex - 1) * 5 + 4) = strDocter
        
        If Align = 1 Then                       '左对齐
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 0) = flexAlignLeftCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 1) = flexAlignLeftCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 2) = flexAlignLeftCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 3) = flexAlignLeftCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 4) = flexAlignLeftCenter
        ElseIf Align = 3 Then                   '右对齐
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 0) = flexAlignRightCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 1) = flexAlignRightCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 2) = flexAlignRightCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 3) = flexAlignRightCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 4) = flexAlignRightCenter
        Else                                    '居中对齐
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 0) = flexAlignCenterCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 1) = flexAlignCenterCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 2) = flexAlignCenterCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 3) = flexAlignCenterCenter
            vfgQueuingData.Cell(flexcpAlignment, intRowIndex, (intRoomIndex - 1) * 5 + 4) = flexAlignCenterCenter
        End If
        
    Exit Sub
errHandle:
    err.Clear
End Sub

Private Sub SetRoomsColor(ByVal intRowIndex As Integer, ByVal intRoomIndex As Integer, fcolor As Long)
'**************************************************************************
'设置特定行的显示色
'**************************************************************************
    On Error GoTo errHandle
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5) = fcolor
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5 + 1) = fcolor
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5 + 2) = fcolor
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5 + 3) = fcolor
        vfgQueuingData.Cell(flexcpForeColor, intRowIndex, (intRoomIndex - 1) * 5 + 4) = fcolor
        
    Exit Sub
errHandle:
    err.Clear
End Sub

Public Sub SetBackColor()
'************************************************************************************
'设置界面背景色
'************************************************************************************
    Dim strReg As String
    Dim mlngBackColor As Long
    
    On Error Resume Next
    '从注册表中，读取显示参数
    strReg = "公共模块\排队叫号\液晶电视"
    mlngBackColor = GetSetting("ZLSOFT", strReg, "背景颜色", vbBlack)
    
    With vfgQueuingData
        .BackColor = mlngBackColor
        .BackColorAlternate = mlngBackColor
        .BackColorBkg = mlngBackColor
        .SheetBorder = mlngBackColor
    End With
    labTitle.BackColor = mlngBackColor
    labInfo.BackColor = mlngBackColor
End Sub

