VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLCDShow 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12315
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFace 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   11535
      TabIndex        =   1
      Top             =   7200
      Width           =   11535
      Begin VB.Label labInf 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   975
         Left            =   1320
         TabIndex        =   2
         Top             =   120
         Width           =   11775
      End
   End
   Begin VB.Timer timerLCD 
      Interval        =   2000
      Left            =   7320
      Top             =   2040
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgCallingData 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _cx             =   20770
      _cy             =   12303
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
      FormatString    =   $"frmLCDShow.frx":0000
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
End
Attribute VB_Name = "frmLCDShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr队列名称() As String
Private mint有效天数 As Integer
Private mstr诊室条件 As String, mstr医生条件 As String, mstrExcludeData As String

Private mlngCallItemsCount As Long    '呼叫队列显示的数量
Private mintViewDataType As Integer '数据显示类型
Private mblnComeBackFirst As Boolean    '回诊病人是否优先排队
Private mlngLedLoopQueryTime As Long '轮询间隔时间长度
Private mlngCallingColor As Long
Private mlngCalledColor As Long
Private mstrDelString As String '需要删除的字符，使用“,”号分割




'显示举例

'----------------------------------------------------------------------------------------------------------------------------------
'|   号码    姓名      科室      诊室     状态
'|
'|   7001    张三      放射       CT      呼叫中
'|   7002    李四      放射       DR      呼叫中
'|
'|
'|
'|
'|
'|
'|
'|
'|  |-------------------------|-------------------------|-------------------------|-------------------------|
'|  |       放射科            |       检验科            |       皮肤科            |       其他科            |
'|  |-------------------------|-------------------------|-------------------------|-------------------------|
'|  |  (号码)姓名     诊室    |  (号码)姓名     诊室    |  (号码)姓名     诊室    |  (号码)姓名     诊室    |
'|  |-------------------------|-------------------------|-------------------------|-------------------------|
'|  |  (7001)张三      CT(回) |                         |                         |                         |
'|  |  (7002)李四      DR     |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  |                         |                         |                         |                         |
'|  ---------------------------------------------------------------------------------------------------------
'|
'|      祝您早日康复！2010-06-04 13:51 星期一
'|
'|
'|
'|
'|
'----------------------------------------------------------------------------------------------------------------------------------


Private Const colSplit = &H808080



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
    '         intViewDataType数据显示类型，0显示当前科室下的所有数据，
    '                                      1显示诊室为当前诊室且医生姓名为空，或者医生姓名等于当前医生，或者诊室为空和医生为空的数据
    '                                      2显示诊室为当前诊室和医生姓名为空或医生姓名等于当前医生的数据
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
    mintViewDataType = intViewDataType
    mblnComeBackFirst = blnComeBackFirst
    
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
End Function


Private Sub Form_Load()
    Dim i As Integer
    Dim strReg As String
    
    strReg = "公共模块\排队叫号\液晶电视"

    mlngLedLoopQueryTime = Val(GetSetting("ZLSOFT", strReg, "LED轮询时间", "2"))
    timerLCD.Interval = mlngLedLoopQueryTime * 1000
    
    mlngCallItemsCount = Val(GetSetting("ZLSOFT", strReg, "呼叫记录显示数", "6"))
    
    mlngCallingColor = GetSetting("ZLSOFT", strReg, "呼叫中颜色", vbGreen)
    mlngCalledColor = GetSetting("ZLSOFT", strReg, "已呼叫颜色", &H408000)
    mstrDelString = GetSetting("ZLSOFT", strReg, "删除字符", "")

    Me.BackColor = vbBlack
    
    '设置显示字体
    Call SetFaceFont
    '设置显示位置
    Call SetFacePostion
    
    Call InitFace(mlngCallItemsCount)
End Sub


Private Sub Form_Resize()
    Call InitFace(mlngCallItemsCount)
End Sub


Public Sub SetFaceFont()
'************************************************************************************
'
'设置界面显示的字体样式
'
'
'************************************************************************************
    Dim strReg As String
    Dim curFontSize As Currency
    
    On Error Resume Next
    '从注册表中，读取显示参数
    strReg = "公共模块\排队叫号\液晶电视"

    vfgCallingData.Font.Name = GetSetting("ZLSOFT", strReg, "字体", "宋体")
    vfgCallingData.Font.Bold = GetSetting("ZLSOFT", strReg, "粗体", "False")
    vfgCallingData.Font.Italic = GetSetting("ZLSOFT", strReg, "斜体", "False")
    
    curFontSize = GetSetting("ZLSOFT", strReg, "字号", "14")
    vfgCallingData.Font.Size = curFontSize * 4.5 / 5
    
    
    labInf.Font.Name = vfgCallingData.Font.Name
    labInf.Font.Bold = vfgCallingData.Font.Bold
    labInf.Font.Italic = vfgCallingData.Font.Italic
    
    If curFontSize > 48 Then
        labInf.Font.Size = curFontSize - 16
    Else
        labInf.Font.Size = curFontSize
    End If
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
    Me.Left = GetSetting("ZLSOFT", strReg, "左", "1024") * Screen.TwipsPerPixelX
    Me.Top = GetSetting("ZLSOFT", strReg, "顶", "0") * Screen.TwipsPerPixelY
    Me.Width = GetSetting("ZLSOFT", strReg, "宽度", "1024") * Screen.TwipsPerPixelX
    Me.Height = GetSetting("ZLSOFT", strReg, "高度", "768") * Screen.TwipsPerPixelY
End Sub



Private Sub InitFace(ByVal lngCallItemsCount As Long)
'************************************************************************************
'
'初始化界面显示
'
'lngRoomsCount: 每屏幕能够显示的科室数量
'
'lngCallItemsCount：叫号队列显示的行数量
'lngWaitItemsCount：候诊队列显示的行数量
'
'************************************************************************************
    Dim i As Integer
    
    On Error GoTo errHandle
    
    '设置呼叫显示-------------------------------------
    vfgCallingData.Top = 0
    vfgCallingData.Left = 140
    vfgCallingData.Width = Me.ScaleWidth - 140
    vfgCallingData.Height = Round(Me.ScaleHeight * 0.9)
    vfgCallingData.Rows = lngCallItemsCount + 3
    vfgCallingData.ForeColor = vbGreen
    vfgCallingData.Enabled = False
    
    vfgCallingData.Cols = 3
'
''    vfgCallingData.Cell(flexcpText, 0, 0) = "号  码"
'    vfgCallingData.Cell(flexcpAlignment, 0, 0) = flexAlignLeftCenter
''    vfgCallingData.ColWidth(0) = Round(Me.Width / 5) + 1000
'
''    vfgCallingData.Cell(flexcpText, 0, 1) = "姓  名"
'    vfgCallingData.Cell(flexcpAlignment, 0, 1) = flexAlignLeftCenter
''    vfgCallingData.ColWidth(1) = Round(Me.Width / 5) - 1000
'
''    vfgCallingData.Cell(flexcpText, 0, 2) = "就诊科室"
'    vfgCallingData.Cell(flexcpAlignment, 0, 2) = flexAlignLeftCenter
''    vfgCallingData.ColWidth(2) = Round(Me.Width / 5) * 3
'
'
'    Call vfgCallingData.AutoSize(0, vfgCallingData.Cols - 1)

    
    picFace.Top = vfgCallingData.Height + 40
    picFace.Left = 0
    picFace.Width = Me.ScaleWidth
    picFace.Height = Round(Me.ScaleHeight * 0.1)
        

        
    '显示信息--------------------------------------
    labInf.Caption = "祝您早日康复！  " & Format(Now, "yyyy-mm-dd  hh:mm") & "  星期" & GetTodayNum
    
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub




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
    Set rsDepart = zlDatabase.OpenSQLRecord(strSQL, "读取科室信息", strDepartId)

    
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
    If ErrCenter = 1 Then Resume
End Sub








'滚动方式显示
Private Sub MultiRoomsDisplay()
'**************************************************************************
'
'显示多个科室的叫号信息
'
'
'**************************************************************************
    Dim i As Integer, j As Integer
    Dim blnSwitchPage As Boolean
    Dim blnAllowRoll As Boolean    '判断当前所显示的数据是否需要滚动，如果数据条数小于当前能够显示的记录数量，则不滚动
    Dim strCurRoomKey As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsSource As ADODB.Recordset   '呼叫数据
    Dim strValues(0 To 10) As String, strValue As String, strUninTable As String
    Dim strFilter As String, strCurExcludeData() As String
    Dim aryDelStr() As String
    Dim strRoomName As String
    
    
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
                If zlCommFun.ActualLen(strValue) > 2000 Then
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
    
    
    '从数据库中查询队列排队情况,只查询呼叫中和已呼叫的数据
    
    strSQL = "" & _
    "   Select /*+ Rule*/  to_Number(a.ID) as ID, a.队列名称, b.名称 as 科室名称, a.排队号码 As 号码, a.排队号码, a.患者姓名, a.医生姓名, a.诊室, " & _
    "               decode (a.排队状态,0,'排队中',1,'呼叫中',7,'已呼叫') as 排队状态, to_Number(a.优先) as 优先, a.排队时间, a.呼叫时间, to_Number(回诊序号) as 回诊序号, to_Number(a.业务类型) as 业务类型, a.业务ID, " & _
                    IIf(mblnComeBackFirst, "to_Number(Nvl(a.回诊序号, 9999999999)) as 回诊排序号", "0 as 回诊排序号") & _
    " From 排队叫号队列 a, 部门表 b , (" & strUninTable & ") E " & _
                IIf(mintViewDataType = 1, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 2, ", Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) C, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D ", "") & _
                IIf(mintViewDataType = 3, ", Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) D", "") & _
    " Where a.科室id = b.Id  And a. 队列名称=E.队列名称  And (a.排队状态 = 1 or a.排队状态=7) and 排队时间 <= trunc(sysdate + 1) - 1/24/60/60 " & _
                IIf(mintViewDataType = 1, " and  ((a.诊室=C.Column_Value and a.医生姓名 is null) or a.医生姓名=D.Column_Value or (a.诊室 is null and a.医生姓名 is null))", "") & _
                IIf(mintViewDataType = 2, " and ((a.诊室=C.Column_Value and a.医生姓名 is Null) or a.医生姓名=D.Column_Value) ", "") & _
                IIf(mintViewDataType = 3, " and a.医生姓名=D.Column_Value", "") & _
    " Order By a.排队状态, a.呼叫时间 desc"
    
    Set rsSource = zlDatabase.OpenSQLRecord(strSQL, "显示排队情况", mstr诊室条件, mstr医生条件, strValues(0), strValues(1), strValues(2), strValues(3), strValues(4), strValues(5), strValues(6), strValues(7), strValues(8), strValues(9), strValues(10))
    
    On Error GoTo errCopyData
        Set rsTemp = zlDatabase.CopyNewRec(rsSource)
        GoTo readData
        
errCopyData:
        If Not rsTemp Is Nothing Then Set rsTemp = Nothing
        
        Call SaveErrLog
        
        Exit Sub
readData:
        
    If rsTemp.EOF Then
        vfgCallingData.Cell(flexcpText, 0, 0, vfgCallingData.Rows - 1, vfgCallingData.Cols - 1) = ""
        Exit Sub
    End If
        
        
    While Not rsTemp.EOF
        If InStr(1, mstrExcludeData, rsTemp!业务类型 & ":" & rsTemp!业务ID) > 0 Then
            rsTemp.Delete
        End If
        
        rsTemp.MoveNext
    Wend
    
        
    '在电视屏幕中显示前几条呼叫的数据

    rsTemp.Sort = "排队状态 asc, 呼叫时间 desc"
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    
    '读取需要删除的字符到数组中
    If Trim(mstrDelString) <> "" Then
        mstrDelString = mstrDelString & ","
        aryDelStr() = Split(mstrDelString, ",")
    End If
    
    vfgCallingData.Redraw = flexRDNone
    For i = 0 To mlngCallItemsCount - 1
        '删除呼叫时间为空的数据
        While Not rsTemp.EOF
            If Nvl(rsTemp!呼叫时间) = "" Then
                rsTemp.MoveNext
            Else
                GoTo AddCallingData
            End If
        Wend
        
AddCallingData:
        
        If Not rsTemp.EOF Then
            vfgCallingData.Cell(flexcpText, i, 0) = "请 " & rsTemp!号码 & "号"
            vfgCallingData.Cell(flexcpText, i, 1) = rsTemp!患者姓名
            
            If Trim(mstrDelString) <> "" Then
                strRoomName = rsTemp!科室名称 & rsTemp!诊室 & "就诊" & IIf(Nvl(rsTemp!回诊序号, 0) = 0, "", "(回)")
                
                
                For j = LBound(aryDelStr()) To UBound(aryDelStr())
                    strRoomName = Replace(strRoomName, aryDelStr(j), "")
                Next j
                
                vfgCallingData.Cell(flexcpText, i, 2) = "到" & strRoomName
            Else
                vfgCallingData.Cell(flexcpText, i, 2) = "到" & rsTemp!科室名称 & rsTemp!诊室 & "就诊" & IIf(Nvl(rsTemp!回诊序号, 0) = 0, "", "(回)")
            End If

            vfgCallingData.Cell(flexcpAlignment, i, 0, i, 2) = flexAlignLeftCenter
            
            If rsTemp!排队状态 = "呼叫中" Then
                vfgCallingData.Cell(flexcpForeColor, i, 0, i, 2) = mlngCallingColor
            Else
                vfgCallingData.Cell(flexcpForeColor, i, 0, i, 2) = mlngCalledColor
            End If
            
            Call vfgCallingData.AutoSize(0, vfgCallingData.Cols - 1)
            
            rsTemp.MoveNext
        Else
            vfgCallingData.Cell(flexcpText, i, 0) = ""
            vfgCallingData.Cell(flexcpText, i, 1) = ""
            vfgCallingData.Cell(flexcpText, i, 2) = ""

            vfgCallingData.Cell(flexcpAlignment, i, 0, i, 2) = flexAlignLeftCenter
        End If
    Next i
    
    vfgCallingData.Redraw = flexRDBuffered

    
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    
    Exit Sub
errHandle:
    If Not rsTemp Is Nothing Then Set rsTemp = Nothing
    '在这里不能进行错误提示，否则造成程序不能正常工作
    Call SaveErrLog
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


Private Sub picFace_Resize()
    labInf.Top = 40
    labInf.Left = 0
    labInf.Width = picFace.Width
    labInf.Height = picFace.Height - 40
End Sub

Public Sub timerLCD_Timer()
        On Error GoTo errHandle
        Dim blnTimer As Boolean
                
        
        blnTimer = timerLCD.Enabled
        timerLCD.Enabled = False
        
        labInf.Caption = "祝您早日康复！  " & Format(Now, "yyyy-mm-dd  hh:mm") & "  星期" & GetTodayNum
        
        Call MultiRoomsDisplay
        
        timerLCD.Enabled = blnTimer
    Exit Sub
errHandle:
    Call SaveErrLog
    
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
    If ErrCenter = 1 Then Resume
End Function

