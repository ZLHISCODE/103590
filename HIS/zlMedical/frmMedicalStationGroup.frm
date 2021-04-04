VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicalStationGroup 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame picState 
      Height          =   750
      Left            =   405
      TabIndex        =   3
      Top             =   195
      Width           =   5850
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "未结:0.00 未收:0.00"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   465
         Width           =   1710
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "合计:0.00(其中记帐:0.00 收费:0.00)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   210
         Width           =   3060
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1530
      Left            =   600
      TabIndex        =   0
      Top             =   1500
      Width           =   5430
      _cx             =   9578
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   6045
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":0000
            Key             =   "公共"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":039A
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":0734
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":0ACE
            Key             =   "单据"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":0E68
            Key             =   "附加"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":1202
            Key             =   "up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationGroup.frx":13C4
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Frame picCount 
      Height          =   555
      Left            =   3300
      TabIndex        =   1
      Top             =   3180
      Width           =   3255
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "人数统计"
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMedicalStationGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean
Private mvarParam As Variant
Private mfrmMain As Object
Private mblnDataMoved As Boolean

Public Function zlMenuClick(ByVal frmMain As Object, ByVal strMenuItem As String, Optional ByVal strParam As String = "") As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '--------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long

    On Error GoTo errHand
    
    Set mfrmMain = frmMain
    mvarParam = Split(strParam, "'")
    
    Select Case strMenuItem
    Case "刷新"
        
        Call zlClearData
        Call RefreshData(strMenuItem)
        
        Call SumCharge
    
    End Select
    
    zlMenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub zlClearData(Optional ByVal strPart As String = "所有")
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '------------------------------------------------------------------------------------------------------------------
    
    Call ResetVsf(vsf)
    Call AppendSapceRows(vsf, lnX, lnY)
        
End Sub

Public Property Get Body(Optional ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Private Sub SumCharge()
    '------------------------------------------------------------------------------------------------------------------
    '功能:费用汇总情况
    '------------------------------------------------------------------------------------------------------------------
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    Call InitSysPara
    
    lbl(0).Caption = "实收金额:0.00(记帐:0.00 收费:0.00)；应收金额:0.00(记帐:0.00 收费:0.00)；"
    lbl(1).Caption = "未结金额:0.00(记帐:0.00 收费:0.00)"
    
    '读取总的费用情况
    
    gstrSQL = GetPublicSQL(SQL.团体费用概况)
    
    '数据转储处理
    '--------------------------------------------------------------------------------------------------------------
    If DataMove(Val(mvarParam(0))) Then
        gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
        gstrSQL = Replace(gstrSQL, "体检登记记录", "H体检登记记录")
        gstrSQL = Replace(gstrSQL, "病人费用记录", "H病人费用记录")
    Else
        '此时可能费用已部份或完全转出
        strSQL = "Select a.体检时间 From 体检登记记录 a,体检人员档案 b Where a.ID=b.登记id And b.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mvarParam(0)))
        If rs.BOF = False Then
            If zlDatabase.DateMoved(Format(rs("体检时间").Value, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption) Then
                strTmp = gstrSQL
                strTmp = Replace(strTmp, "病人费用记录", "H病人费用记录")
                strSQL = gstrSQL & " Union All " & strTmp
            End If
        End If
    End If
                    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mvarParam(0)))
    If CalcCharge(rsData, rs) Then

        strTmp = ""
        
        strTmp = strTmp & "实收金额:" & Format(zlCommFun.NVL(rs("实收金额").Value, 0), gstrDec) & "(记帐:" & Format(zlCommFun.NVL(rs("记帐金额").Value, 0), gstrDec)
        strTmp = strTmp & " 收费:" & Format(zlCommFun.NVL(rs("收费金额").Value, 0), gstrDec) & ")"
        
        strTmp = strTmp & "；应收金额:" & Format(Val(zlCommFun.NVL(rs("应收金额_收").Value, 0)) + Val(zlCommFun.NVL(rs("应收金额_记").Value, 0)), gstrDec) & "(记帐:" & Format(zlCommFun.NVL(rs("应收金额_记").Value, 0), gstrDec)
        strTmp = strTmp & " 收费:" & Format(zlCommFun.NVL(rs("应收金额_收").Value, 0), gstrDec) & ")"
        
        lbl(0).Caption = strTmp
        
        If zlCommFun.NVL(rs("未结算合计").Value, 0) > 0 Then
            strTmp = ""
            strTmp = strTmp & "未结金额:" & Format(zlCommFun.NVL(rs("未结算合计").Value, 0), gstrDec) & "(记帐:" & Format(zlCommFun.NVL(rs("未结金额").Value, 0), gstrDec)
            strTmp = strTmp & " 收费:" & Format(zlCommFun.NVL(rs("未收金额").Value, 0), gstrDec) & ")"
            
            lbl(1).Caption = strTmp

        End If
        
    End If

End Sub

Private Function RefreshData(ByVal strMenu As String) As Boolean
    Dim lngLoop As Long
    Dim lngCount0 As Long
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    lbl(2).Caption = ""
    
    Select Case strMenu
    Case "刷新"
        
        gstrSQL = "SELECT A.组别名称 AS 组别,A.病人id AS ID,A.姓名,B.门诊号,a.体检编号," & _
                          "A.体检报到 AS 报到," & _
                          "DECODE(A.体检病历ID, Null, 0, 1) As 总检, " & _
                          "DECODE(A.体检状态, 5, 1, 0) As 完成,c.应收金额,c.实收金额 " & _
                     "FROM 体检人员档案 A," & _
                          "病人信息 B, " & _
                          "(Select 病人id,Sum(D.应收金额) As 应收金额,Sum(D.实收金额) As 实收金额 " & _
                             "FROM 病人费用记录 D, " & _
                                  "(SELECT C.ID " & _
                                     "FROM 体检登记记录 B, 病人医嘱记录 C " & _
                                    "WHERE C.病人来源 = 4 AND " & _
                                          "C.医嘱状态 <> 4 AND B.体检号 = C.挂号单 AND B.ID = [1]) E " & _
                            "WHERE D.记录状态 IN (0, 1) AND D.医嘱序号 = E.ID Group By d.病人id) C " & _
                    "WHERE A.病人ID = B.病人ID AND A.登记ID = [1] And a.病人id=c.病人id(+)"
                            
                    
        If Trim(mvarParam(1)) <> "" Then
            gstrSQL = gstrSQL & " And a.组别名称=[2] "
        End If
        
        gstrSQL = gstrSQL & " ORDER BY A.组别名称,B.门诊号,a.体检编号 "
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If DataMove(Val(mvarParam(0))) Then
            gstrSQL = Replace(gstrSQL, "体检人员档案", "H体检人员档案")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mvarParam(0)), CStr(mvarParam(1)))
        If rs.BOF = False Then

            Call LoadGrid(vsf, rs)
            Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
            Call AppendSapceRows(vsf, lnX, lnY)
            
            '统计已完人数、未完人数及未报到人数
            For lngLoop = 1 To vsf.Rows - 1
                
                '未报到统计
                If Abs(Val(vsf.TextMatrix(lngLoop, 4))) <> 1 Then
                    lngCount0 = lngCount0 + 1
                Else
                    '已完统计
                    If Abs(Val(vsf.TextMatrix(lngLoop, 6))) = 1 Then
                        lngCount1 = lngCount1 + 1
                    Else
                        '未完统计
                        lngCount2 = lngCount2 + 1
                    End If
                End If
            Next
            
            lbl(2).Caption = "应到:" & lngCount0 + lngCount1 + lngCount2 & "人;实到:" & lngCount1 + lngCount2 & "人(完成:" & lngCount1 & "人;未完:" & lngCount2 & "人);未到:" & lngCount0 & "人"
                        
        End If
    End Select
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    Dim strVsf As String
            
    strVsf = "组别,1500,1,1,1,;姓名,900,1,1,1,;门诊号,900,7,1,1,;体检编号,990,1,1,1,;报到,600,4,1,1,;总检,600,4,1,1,;完成,600,4,1,1,;应收金额,1080,7,1,1,;实收金额,1080,7,1,1,"
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.MergeCells = flexMergeFree
    vsf.MergeCol(0) = True
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(4) = flexDTBoolean
    vsf.ColDataType(5) = flexDTBoolean
    vsf.ColDataType(6) = flexDTBoolean
    Call AppendSapceRows(vsf, lnX, lnY)
    vsf.ColFormat(7) = "0.00"
    vsf.ColFormat(8) = "0.00"
    
    lbl(0).Caption = ""
    lbl(1).Caption = ""
    
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function



'（３）窗体及其控件的事件处理******************************************************************************************

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call InitLoad
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With picState
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth
    End With
    
    With vsf
        .Left = 0
        .Top = picState.Top + picState.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - picCount.Height + 90
    End With
                
    With picCount
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height - 90
        .Width = picState.Width
    End With
    
    Call AppendSapceRows(vsf, lnX, lnY)
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub

    On Error GoTo errHand
    Call mfrmMain.ActiveFormEnabled
    
errHand:
        
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col > 0)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 1 Then Cancel = True
End Sub



