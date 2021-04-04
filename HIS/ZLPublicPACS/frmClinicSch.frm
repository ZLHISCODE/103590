VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmClinicSch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "临床检查预约"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9195
   Icon            =   "frmClinicSch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame frmSchedInfo 
      Caption         =   "此检查已完成预约，预约信息如下："
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   9015
      Begin VB.Label lblSchedInfo 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   8655
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出"
      Height          =   375
      Left            =   7200
      TabIndex        =   22
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H80000013&
      Caption         =   "大后天"
      Height          =   375
      Index           =   4
      Left            =   6840
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDate 
      Caption         =   "下月 >>"
      Height          =   375
      Index           =   5
      Left            =   8030
      TabIndex        =   19
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H80000013&
      Caption         =   "后天"
      Height          =   375
      Index           =   3
      Left            =   5640
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   18
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H80000013&
      Caption         =   "明天"
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   17
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDate 
      BackColor       =   &H80000013&
      Caption         =   "今天"
      Height          =   375
      Index           =   1
      Left            =   1200
      MaskColor       =   &H8000000F&
      TabIndex        =   16
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdChangeDate 
      Caption         =   "<< 上月"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFC0&
      Caption         =   "预约"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   7200
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSchSegment 
      Height          =   1530
      Left            =   120
      TabIndex        =   13
      Top             =   4680
      Width           =   9015
      _cx             =   15901
      _cy             =   2699
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VSFlex8Ctl.VSFlexGrid vsfSchDate 
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   9015
      _cx             =   15901
      _cy             =   4895
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox cboSchDevice 
         Height          =   300
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblSchTime 
         Caption         =   "预约时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label lblSchDate 
         Caption         =   "预约日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   671
         Width           =   1335
      End
      Begin VB.Label lblSchDevice 
         Caption         =   "预约设备"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   8
         Top             =   263
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "预约时间段："
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "预约日期："
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   671
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "预约设备："
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   263
         Width           =   975
      End
      Begin VB.Label lblOrder 
         Caption         =   "医嘱内容"
         Height          =   675
         Left            =   1320
         TabIndex        =   4
         Top             =   671
         Width           =   2415
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "医嘱内容："
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   671
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "姓名"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   263
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "患者姓名："
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   263
         Width           =   975
      End
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      Caption         =   "2019年6月"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   1500
      Width           =   2055
   End
End
Attribute VB_Name = "frmClinicSch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAdviceID As Long            '打开窗体时的医嘱ID
Private mblnCanSchedule As Boolean      '找到预约时间和设备
Private mblnRefreshDevice As Boolean    '是否刷新预约设备
Private mblnOK As Boolean               '预约成功

'XML返回的内容
Private mstrOrderID As String
Private mstr诊疗项目ID As String
Private mstr诊疗项目名称 As String
Private mstr医嘱内容 As String
Private mstr预约设备名称 As String
Private mstr预约设备ID As String
Private mstr预约日期 As String
Private mstr预约开始时间 As String
Private mstr预约结束时间 As String
Private mrsTimes As ADODB.Recordset
    
'检查日历
Private Enum schDateColTitle
    col_SchDate_周一 = 0
    col_SchDate_周二 = 1
    col_SchDate_周三 = 2
    col_SchDate_周四 = 3
    col_SchDate_周五 = 4
    col_SchDate_周六 = 5
    col_SchDate_周日 = 6
End Enum

'检查时间段
Private Enum schTimeSegColTitle
    col_SchTimeSeg_序号 = 0
    col_SchTimeSeg_设备 = 1
    col_SchTimeSeg_开始时间 = 2
    col_SchTimeSeg_结束时间 = 3
End Enum

Public Function zlShowMe(objParent As Object, blnShowModal As Boolean, lngAdviceID As Long, blnModify As Boolean) As Boolean
'-----------------------------------------------------------
'功能:显示临床检查预约窗口
'入参:  objParent -- 父窗体
'       blnShowModal -- 是否模式窗体
'       lngAdviceID -- 医嘱ID
'       blnModify -- 修改预约
'返回:
'-----------------------------------------------------------
    
    On Error GoTo err
    
    mblnOK = False
    mlngAdviceID = lngAdviceID
            
    If refreshDate(Format(Now, "YYYY-MM-DD")) = False Then
        If mblnCanSchedule = False Then
            zlShowMe = True
        Else
            zlShowMe = False
        End If
        Unload Me
        Exit Function
    End If
    
    Call loadPatInfo             '先查询日期后加载患者基本信息
    
    If blnModify = True Then
        frmSchedInfo.Visible = True
        Call loadSchedInfo
        Me.Caption = "修改临床检查预约"
    Else
        frmSchedInfo.Visible = False
        vsfSchSegment.Height = vsfSchSegment.Height + frmSchedInfo.Height
    End If
    
    mblnRefreshDevice = True
    
    Call Show(IIf(blnShowModal, 1, 0), objParent)
    
    zlShowMe = mblnOK
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cboSchDevice_Click()
On Error GoTo err
    If mblnRefreshDevice = True Then
        mstr预约设备ID = cboSchDevice.ItemData(cboSchDevice.ListIndex)
        cmdChangeDate_Click (1) '刷新今天的日历
    End If
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChangeDate_Click(index As Integer)
    Dim dtDate As Date
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Integer
    Dim j As Integer
    Dim strSelDate As String
On Error GoTo err
    strSelDate = ""
    'index =0-上月；1-今天；2-明天；3-后天；4-大后天；5-下月
    Select Case index
    Case 0:     '上月
        dtDate = Format(mstr预约开始时间, "YYYY-MM-DD")
        If GetOtherMonth(False, dtDate) = False Then
            Exit Sub
        End If
        
    Case 1:  '今天
        dtDate = Now
    Case 2:     '明天
        dtDate = Now + 1
        strSelDate = Format(dtDate, "YYYY-MM-DD")
    Case 3:     '后天
        dtDate = Now + 2
        strSelDate = Format(dtDate, "YYYY-MM-DD")
    Case 4:     '大后天
        dtDate = Now + 3
        strSelDate = Format(dtDate, "YYYY-MM-DD")
    Case 5:     '下月
        dtDate = Format(mstr预约开始时间, "YYYY-MM-DD")
        If GetOtherMonth(True, dtDate) = False Then
            Exit Sub
        End If
    End Select
    
    Call refreshDate(dtDate, CLng(mstr预约设备ID), strSelDate)
    
    If index = 1 Or index = 2 Or index = 3 Or index = 4 Then
        If Format(dtDate, "YYYY-MM-DD") <> Format(mstr预约日期, "YYYY-MM-DD") Then
            MsgBox IIf(index = 1, "今天", IIf(index = 2, "明天", IIf(index = 3, "后天", "大后天"))) _
                & "预约已满，只能从" & mstr预约日期 & "开始预约。", vbOKOnly, "检查预约"
        End If
    End If
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strIn As String
    Dim strOut As String
    Dim objXml As Object  'zl9ComLib.clsXML
    Dim strResult As String
    
    On Error GoTo err
    
    strIn = "<IN><ADVICEID>" & mstrOrderID & "</ADVICEID><MACHINEID>" _
        & mstr预约设备ID & "</MACHINEID><SCHBEGINTIME>" & mstr预约开始时间 _
        & "</SCHBEGINTIME><SCHENDTIME>" & mstr预约结束时间 & "</SCHENDTIME></IN>"
    strOut = gobjComLib.zlDatabase.CallProcedure("zl_影像预约_ScheduleInsert", Me.Caption, strIn, Empty)
    
    '解析返回的XML串
    If strOut = "" Then
        MsgBox "保存预约信息出错，请重新选择时间段后，再次预约。", vbOKOnly, "检查预约"
        '刷新预约日历
        Call refreshDate(CDate(Format(mstr预约日期, "YYYY-MM-DD")))
    End If
    '  --成功：
    '  --<OUTPUT>
    '  --  <RESULT>true</RESULT>
    '  --</OUTPUT>
    '
    '  --失败：
    '  --<OUTPUT>
    '  --  <RESULT>false</RESULT>
    '  --  <ERROR>
    '  --    <MSG>详细错误提示</MSG>
    '  --  </ERROR>
    '  --</OUTPUT>
    Set objXml = CreateObject("zl9ComLib.clsXML")
    Call objXml.OpenXMLDocument(strOut)
    Call objXml.GetSingleNodeValue("RESULT", strResult)
    If strResult = "true" Then
        mblnOK = True
        '预约成功，打印预约单后退出
        Call PrintSchedule(Me, mlngAdviceID)
        Unload Me
    Else
        '预约失败，提示，刷新列表
        Call objXml.GetSingleNodeValue("MSG", strResult)
        If InStr(strResult, "[ZLSOFT]") > 0 Then
            strResult = Split(strResult, "[ZLSOFT]")(1)
        End If
        MsgBox "保存预约信息出错，请重新选择时间段后，再次预约。" & vbCrLf & vbCrLf _
            & "错误信息：" & strResult, vbOKOnly, "检查预约"
        '刷新预约日历
        Call refreshDate(CDate(Format(mstr预约日期, "YYYY-MM-DD")))
    End If
    
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub loadPatInfo()
'-----------------------------------------------------------
'功能:加载患者基本信息
'入参:
'返回:
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim lngIndex As Long
    
    On Error GoTo err
    
    lngIndex = -1
    
    strSQL = "select a.姓名,nvl(a.婴儿,0) as 婴儿,a.医嘱内容,a.病人ID,a.主页ID from 病人医嘱记录 a where id =[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询患者信息", mlngAdviceID)
    
    If rsTemp.EOF = False Then
        If rsTemp!婴儿 <> 0 Then
            strSQL = "Select Decode(a.婴儿姓名, Null, b.姓名 || '之子' || Trim(To_Char(a.序号, '9')), a.婴儿姓名) As 婴儿姓名, " _
                    & " From 病人新生儿记录 A, 病人信息 B Where a.病人id = [ 1 ] And a.主页id = [ 2 ] " _
                    & " And a.病人id = b.病人id And a.序号 = [ 3 ]"
            Set rsBaby = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询患者信息", rsTemp!病人ID, rsTemp!主页ID, rsTemp!婴儿)
            lblName.Caption = rsBaby!婴儿姓名
        Else
            lblName.Caption = rsTemp!姓名
        End If
        lblOrder.Caption = rsTemp!医嘱内容
    End If
    
    strSQL = "select b.id ,b.设备名称 from 影像预约项目 A ,影像预约设备 B,病人医嘱记录 C,影像预约方案 D WHERE c.id=[1] " _
            & " and c.诊疗项目id = a.诊疗项目id and a.预约设备id = b.id and b.是否启用=1 and B.ID=D.预约设备ID(+) and D.是否启用=1 "
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询预约设备", mlngAdviceID)
    
    If rsTemp.RecordCount > 1 Then
        cboSchDevice.Clear
        While rsTemp.EOF = False
            cboSchDevice.AddItem rsTemp!设备名称
            cboSchDevice.ItemData(cboSchDevice.NewIndex) = rsTemp!ID
            If rsTemp!ID = mstr预约设备ID Then
                lngIndex = cboSchDevice.ListCount - 1
            End If
            rsTemp.MoveNext
        Wend
        If lngIndex <> -1 Then
            cboSchDevice.ListIndex = lngIndex
        Else
            cboSchDevice.ListIndex = 0
        End If
        
        cboSchDevice.Visible = True
        lblSchDevice.Visible = False
    Else
        cboSchDevice.Visible = False
        lblSchDevice.Visible = True
    End If
    
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub loadSchedInfo()
'-----------------------------------------------------------
'功能:加载已经预约的信息
'入参:
'返回:
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err

    strSQL = "SELECT a.预约设备名称,a.诊室名称,a.预约开始时间,a.预约结束时间 FROM 影像预约记录 a where a.医嘱ID = [1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询患者预约信息", mlngAdviceID)
    
    If rsTemp.EOF = False Then
        lblSchedInfo = "预约设备：" & rsTemp!预约设备名称 & "     预约日期：" & Format(rsTemp!预约开始时间, "YYYY-MM-DD") _
            & "      预约时间：" & Format(rsTemp!预约开始时间, "hh:mm:ss") & " - " & Format(rsTemp!预约结束时间, "hh:mm:ss")
    Else
        lblSchedInfo = "无"
    End If

    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Sub loadSchDate(strSelDate As String)
'-----------------------------------------------------------
'功能:加载预约日历
'入参:  strSelDate -- 需要选中的日期，如果为空，则选中mstr预约开始时间
'返回:
'-----------------------------------------------------------
    Dim i As Integer
    Dim dt预约日期 As Date
    Dim lng1号周几 As Long
    Dim lng总天数 As Long
    Dim lng周几 As Long
    Dim lngRow As Long
    Dim dt当天 As Date
    Dim dt选中日期 As Date
    
    On Error GoTo err
    
    If mblnCanSchedule = False Then Exit Sub
    
    '加载预约日历表
    With vsfSchDate
        .Rows = 1
        .Cols = 1
        .Rows = 6
        .Cols = 7
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .ColWidthMin = 1280
        .AllowUserResizing = flexResizeNone
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .SelectionMode = flexSelectionFree
        .AllowSelection = False

        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        
        .TextMatrix(0, 0) = "周一"
        .TextMatrix(0, 1) = "周二"
        .TextMatrix(0, 2) = "周三"
        .TextMatrix(0, 3) = "周四"
        .TextMatrix(0, 4) = "周五"
        .TextMatrix(0, 5) = "周六"
        .TextMatrix(0, 6) = "周日"
        
        '获取当前月的第一天是星期几
        dt预约日期 = Format(CDate(mstr预约开始时间), "YYYY-MM-DD")
        dt当天 = dt预约日期 - Day(dt预约日期) + 1
        If strSelDate <> "" And Format(dt选中日期, "YYYY-MM-DD") > Format(strSelDate, "YYYY-MM-DD") Then
            dt选中日期 = Format(strSelDate, "YYYY-MM-DD")
        Else
            dt选中日期 = dt预约日期
        End If
        
        lng1号周几 = Weekday(dt当天, vbMonday)
        
        lng总天数 = Day(DateSerial(Year(dt预约日期), Month(dt预约日期) + 1, 0))
        
        lng周几 = lng1号周几
        lngRow = 1
        For i = 1 To lng总天数
            .TextMatrix(lngRow, lng周几 - 1) = i
            '如果这天>=预约日期，可以预约，单击后再则显示已经预约的天数和总容量
            If DateCanSch(dt当天, CLng(mstr预约设备ID)) = True Then
                .Cell(flexcpBackColor, lngRow, lng周几 - 1) = &HFFFFC0  ' &HC0FFC0
            Else
                .Cell(flexcpBackColor, lngRow, lng周几 - 1) = vbBlack
            End If
            
            If dt当天 = dt选中日期 And .Cell(flexcpBackColor, lngRow, lng周几 - 1) <> vbBlack Then
                .Select lngRow, lng周几 - 1
            End If
            
            dt当天 = dt当天 + 1
            lng周几 = lng周几 + 1
            If lng周几 > 7 Then
                lngRow = lngRow + 1
                lng周几 = 1
            End If
            If lngRow = 6 Then
                .Rows = 7
            End If
        Next i
        
        .Rows = lngRow + 1  '最后修正行数
        .RowHeightMin = IIf(lngRow = 5, 450, 382)
        '选中 dt预约日期
        .Refresh

    End With
    
    '更新年月显示
    lblDate.Caption = Format(dt预约日期, "YYYY年MM月")
    Call loadSchSegment
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub loadSchSegment()
'-----------------------------------------------------------
'功能:加载预约日历
'入参:  rsTimes -- 预约时间段数据集
'返回:
'-----------------------------------------------------------
    Dim lngRow As Long
    Dim dtBeginTime As Date
    Dim dtEndTime As Date
    
    On Error GoTo err

    If mrsTimes.EOF = True Then
        vsfSchSegment.Rows = 1
        Exit Sub
    End If
    
    lngRow = 1
    '加载预约时间段
    With vsfSchSegment
        .Rows = mrsTimes.RecordCount
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .ColWidthMin = 500
        .AllowUserResizing = flexResizeNone
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        .SelectionMode = flexSelectionByRow
        
        .ColWidth(col_SchTimeSeg_设备) = 1500
        .ColWidth(col_SchTimeSeg_开始时间) = 2000
        .ColWidth(col_SchTimeSeg_结束时间) = 2000
        .TextMatrix(0, col_SchTimeSeg_序号) = "序号"
        .TextMatrix(0, col_SchTimeSeg_设备) = "预约设备"
        .TextMatrix(0, col_SchTimeSeg_开始时间) = "开始时间段"
        .TextMatrix(0, col_SchTimeSeg_结束时间) = "结束时间段"
        While mrsTimes.EOF = False
            If mrsTimes!node_name = "SEGBEGINTIME" Then
                dtBeginTime = Format(mrsTimes!node_value, "HH:MM:SS")
                mrsTimes.MoveNext
                dtEndTime = Format(mrsTimes!node_value, "HH:MM:SS")
                If dtBeginTime >= Format(mstr预约开始时间, "HH:MM:SS") Then
                    .TextMatrix(lngRow, col_SchTimeSeg_序号) = lngRow
                    .TextMatrix(lngRow, col_SchTimeSeg_设备) = mstr预约设备名称
                    .TextMatrix(lngRow, col_SchTimeSeg_开始时间) = dtBeginTime
                    .TextMatrix(lngRow, col_SchTimeSeg_结束时间) = dtEndTime
                    If dtBeginTime = Format(mstr预约开始时间, "HH:MM:SS") Then
                        .Cell(flexcpChecked, lngRow, col_SchTimeSeg_设备) = 1
                    Else
                        .Cell(flexcpChecked, lngRow, col_SchTimeSeg_设备) = 2
                    End If
                    lngRow = lngRow + 1
                End If
            End If
            mrsTimes.MoveNext
        Wend
        .Rows = lngRow
    End With
    
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ChangeSchDate(dtDate As Date, Optional lngSchDeviceID As Long = 0) As Boolean
'-----------------------------------------------------------
'功能:改变预约日期
'入参:  dtDate -- 最早可预约的日期
'       lngSchDeviceID -- 预约设备ID，如果不指定设备，传0
'返回: True -- 成功； False -- 失败
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsTimes As ADODB.Recordset
    Dim strInXML As String
    Dim strOutXML As String
    Dim objXml As Object  'zl9ComLib.clsXML
    Dim strError As String
    
    Dim strOrderIDOld As String
    Dim str诊疗项目IDOld As String
    Dim str诊疗项目名称Old As String
    Dim str医嘱内容Old As String
    Dim str预约设备名称Old As String
    Dim str预约设备IDOld As String
    Dim str预约开始时间Old As String
    Dim str预约结束时间Old As String
    Dim str预约日期Old As String
    
    On Error GoTo err
    
    
    Set rsTimes = Nothing
        
    If GetSchDate(dtDate, strOutXML, lngSchDeviceID) = False Then
        mblnCanSchedule = False
        Exit Function
    End If
    
    '解析XML串
'  --  <OUTPUT>
'  --    <ERROR>
'  --      <MSG>错误信息</MSG>
'  --    </ERROR>
'  --    <SCHINFO>
'  --      <ADVICEID>医嘱ID</ADVICEID>
'  --      <CHECKID>诊疗项目ID</CHECKID>
'  --      <CHECKNAME>诊疗项目名称</CHECKNAME>
'  --      <ADVICEDOC>医嘱内容</ADVICEDOC>
'  --      <MACHINENAME>预约设备名称</MACHINENAME>
'  --      <MACHINEID>预约设备ID</MACHINEID>
'  --      <SCHBEGINTIME>开始时间段</SCHBEGINTIME>
'  --      <SCHENDTIME>结束时间段</SCHENDTIME>
'  --    </SCHINFO>
'  --    <SCHTIMES>
'  --      <SEGBEGINTIME>开始时间段1</SEGBEGINTIME>
'  --      <SEGENDTIME>结束时间段1</SEGENDTIME>
'  --    </SCHTIMES>
'  --    <SCHTIMES>
'  --      <SEGBEGINTIME>开始时间段2</SEGBEGINTIME>
'  --      <SEGENDTIME>结束时间段2</SEGENDTIME>
'  --    </SCHTIMES>
'  --  </OUTPUT>
    Set objXml = CreateObject("zl9ComLib.clsXML")
    Call objXml.OpenXMLDocument(strOutXML)
    
    Call objXml.GetSingleNodeValue("ADVICEID", mstrOrderID)
    Call objXml.GetSingleNodeValue("CHECKID", mstr诊疗项目ID)
    Call objXml.GetSingleNodeValue("CHECKNAME", mstr诊疗项目名称)
    Call objXml.GetSingleNodeValue("ADVICEDOC", mstr医嘱内容)
    Call objXml.GetSingleNodeValue("MACHINENAME", mstr预约设备名称)
    Call objXml.GetSingleNodeValue("MACHINEID", mstr预约设备ID)
    Call objXml.GetSingleNodeValue("SCHBEGINTIME", mstr预约开始时间)
    Call objXml.GetSingleNodeValue("SCHENDTIME", mstr预约结束时间)
    
    Call objXml.GetMultiNodeRecord("OUTPUT/SCHTIMES", mrsTimes)
        
    mstr预约日期 = Format(mstr预约开始时间, "YYYY-MM-DD")
    
    '查询预约时间出错，提取错误信息
    If mstrOrderID = "" Then
        Call objXml.GetSingleNodeValue("MSG", strError)
        If strError <> "无可用的预约设备。" Then
            MsgBox "提取预约方案出现错误：" & strError, vbOKOnly, "检查预约"
        Else
            MsgBox "无可用的预约设备。", vbOKOnly, "检查预约"
        End If
        mblnCanSchedule = False
        ChangeSchDate = False
        mstrOrderID = strOrderIDOld
        mstr诊疗项目ID = str诊疗项目IDOld
        mstr诊疗项目名称 = str诊疗项目名称Old
        mstr医嘱内容 = str医嘱内容Old
        mstr预约设备名称 = str预约设备名称Old
        mstr预约设备ID = str预约设备IDOld
        mstr预约开始时间 = str预约开始时间Old
        mstr预约结束时间 = str预约结束时间Old
        mstr预约日期 = str预约日期Old
    Else
        mblnCanSchedule = True
        ChangeSchDate = True
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub loadSchInfo()
'-----------------------------------------------------------
'功能:加载预约信息
'入参:
'返回:
'-----------------------------------------------------------

    On Error GoTo err
    mblnRefreshDevice = False
    '填写预约信息
    If mblnCanSchedule = True Then
        lblSchDevice.Caption = mstr预约设备名称
        lblSchDate.Caption = Format(mstr预约开始时间, "YYYY-MM-DD")
        lblSchTime.Caption = Format(mstr预约开始时间, "hh:mm:ss") & " -- " & Format(mstr预约结束时间, "hh:mm:ss")
    Else
        lblSchDevice.Caption = ""
        lblSchDate.Caption = ""
        lblSchTime.Caption = ""
    End If
    mblnRefreshDevice = True
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfSchDate_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    Dim strDate As String
    On Error GoTo err
    '如果当天不能预约，则禁止更改
    strDate = vsfSchDate.TextMatrix(NewRowSel, NewColSel)
    If strDate = "" Then
        Cancel = True
    ElseIf vsfSchDate.Cell(flexcpBackColor, NewRowSel, NewColSel) = vbBlack Then
            Cancel = True
    End If
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfSchDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strDate As String
On Error GoTo err
    With vsfSchDate
        If .RowSel >= 1 Then
            strDate = .TextMatrix(.RowSel, .ColSel)
            If strDate <> "" Then
                Call ChangeSchDate(CDate(Format(mstr预约开始时间, "YYYY-MM") & "-" & Format(strDate, "00")), CLng(mstr预约设备ID))
                Call loadSchSegment
                Call loadSchInfo
            End If
        End If
    End With
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfSchSegment_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    On Error GoTo err
    With vsfSchSegment
        If Row >= 1 And Col = col_SchTimeSeg_设备 Then
            If .Cell(flexcpChecked, Row, col_SchTimeSeg_设备) = 1 Then
                '取消其他行的选择
                For i = 1 To .Rows - 1
                    If i <> Row Then
                        .Cell(flexcpChecked, i, col_SchTimeSeg_设备) = 2
                    End If
                Next i
                '更改预约时间段显示
                mstr预约开始时间 = mstr预约日期 & " " & .TextMatrix(Row, col_SchTimeSeg_开始时间)
                mstr预约结束时间 = mstr预约日期 & " " & .TextMatrix(Row, col_SchTimeSeg_结束时间)
                Call loadSchInfo
            ElseIf .Cell(flexcpChecked, Row, col_SchTimeSeg_设备) = 2 Then
                .Cell(flexcpChecked, Row, col_SchTimeSeg_设备) = 1
            End If
            .Refresh
        End If
    End With
    Exit Sub
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Private Function GetOtherMonth(ByVal blnNextMonth As Boolean, ByRef dtDate As Date) As Boolean
'-----------------------------------------------------------
'功能:切换上月或下月，切换到上月第一天，或者下月第一天
'入参:  blnNextMonth -- 是否下月，True - 下月；False - 上月
'       dtDate --【入参、出参】 需要切换的日期
'返回: True -- 成功； False -- 失败
'-----------------------------------------------------------
    Dim dtNewDate As Date
    
    On Error GoTo err
    
    If blnNextMonth = True Then '下一月的第一天
        dtNewDate = DateAdd("m", 1, dtDate)
        dtNewDate = CDate(Format(dtNewDate, "YYYY-MM") & "-01")
    Else    '上一月的最后一天
        dtNewDate = DateAdd("m", -1, dtDate)
        
        dtNewDate = CDate(Format(dtNewDate, "YYYY-MM") & "-01")
        '最后一天 dtNewDate = CDate(Format(dtNewDate, "YYYY-MM") & "-" & Day(DateSerial(Year(dtNewDate), Month(dtNewDate) + 1, 0)))
    End If
        
    dtDate = dtNewDate
    GetOtherMonth = True
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function refreshDate(dtDate As Date, Optional lngSchDeviceID As Long = 0, Optional strSelDate As String = "") As Boolean
'-----------------------------------------------------------
'功能:刷新日历
'入参:  dtDate --需要切换的日期
'       lngSchDeviceID -- 预约设备ID，如果不指定设备，传0
'       strSelDate -- 需要选中的日期，如果为空，则选中dtDate
'返回: True -- 成功； False -- 失败
'-----------------------------------------------------------
    Dim blnResult As Boolean
    
    On Error GoTo err
    
    blnResult = ChangeSchDate(dtDate, lngSchDeviceID)
    
    If blnResult = False Then
        refreshDate = False
        Exit Function
    End If
    
    Call loadSchDate(strSelDate)
    Call loadSchInfo
    refreshDate = True
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSchDate(ByVal dtDate As Date, ByRef strOutXML As String, Optional lngShcDeviceID As Long = 0) As Boolean
'-----------------------------------------------------------
'功能:根据输入日期，确定最近的预约日期
'入参:  dtDate --需要切换的日期
'       strOutXML -- 查询返回值
'       lngShcDeviceID -- 预约设备ID，如果不指定预约设备，传0
'返回: True -- 成功；False -- 失败
'-----------------------------------------------------------
    Dim strInXML As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    '查询最近的可预约设备和时间
    If lngShcDeviceID = 0 Then
        strInXML = "<IN><ADVICEID>" & mlngAdviceID & "</ADVICEID>" & _
            "<BEGINTIME>" & dtDate & "</BEGINTIME></IN>"
    Else
        strInXML = "<IN><ADVICEID>" & mlngAdviceID & "</ADVICEID>" & _
            "<BEGINTIME>" & dtDate & "</BEGINTIME><MACHINEID>" & _
            lngShcDeviceID & "</MACHINEID></IN>"
    End If
    
'    strSQL = "select zl_Test_GetSchTimes(xmltype('" & strInXML & "')) as outData from dual"
'    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "查询最早的预约日期和设备")
    
    strOutXML = gobjComLib.zlDatabase.CallProcedure("zl_影像预约_GetScheduleTimes", Me.Caption, strInXML, Empty)
    
    If strOutXML = "" Then
        GetSchDate = False
        Exit Function
    End If
    
    GetSchDate = True
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Function

Private Function DateCanSch(dtDate As Date, Optional lngDeviceID As Long = 0) As Boolean
'-----------------------------------------------------------
'功能:判断当天是否可以预约
'入参:  dtDate -- 日期
'       lngSchDeviceID -- 预约设备ID，如果不指定设备，传0
'返回: True -- 可预约；False -- 不可预约
'-----------------------------------------------------------
    Dim strOutXML As String
    Dim objXml As Object  'zl9ComLib.clsXML
    Dim strSchDate As String
    
    On Error GoTo err
    
    DateCanSch = False
    
    '如果小于今天，直接返回False不可预约
    If Format(Now, "YYYY-MM-DD") > Format(dtDate, "YYYY-MM-DD") Then
        Exit Function
    End If
    
    If GetSchDate(dtDate, strOutXML, lngDeviceID) = True Then
        Set objXml = CreateObject("zl9ComLib.clsXML")
        Call objXml.OpenXMLDocument(strOutXML)
        Call objXml.GetSingleNodeValue("SCHBEGINTIME", strSchDate)
        
        If Format(strSchDate, "YYYY-MM-DD") = Format(dtDate, "YYYY-MM-DD") Then
            DateCanSch = True
        End If
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsfSchSegment_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col_SchTimeSeg_设备 Then
        Cancel = True
    End If
End Sub
