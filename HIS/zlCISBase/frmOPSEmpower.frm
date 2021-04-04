VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmOPSEmpower 
   Caption         =   "手术授权管理"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12855
   Icon            =   "frmOPSEmpower.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   12855
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicSQ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   5640
      ScaleHeight     =   2775
      ScaleWidth      =   3855
      TabIndex        =   12
      Top             =   4440
      Width           =   3855
      Begin VSFlex8Ctl.VSFlexGrid vsSQ 
         Height          =   6420
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   7305
         _cx             =   12885
         _cy             =   11324
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOPSEmpower.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
   Begin VB.PictureBox picOPS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   5640
      ScaleHeight     =   2535
      ScaleWidth      =   3855
      TabIndex        =   11
      Top             =   1680
      Width           =   3855
      Begin VSFlex8Ctl.VSFlexGrid vsOPS 
         Height          =   6420
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   7305
         _cx             =   12885
         _cy             =   11324
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
         BackColorSel    =   16771802
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmOPSEmpower.frx":68ED
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5220
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9208
      _StockProps     =   64
   End
   Begin VB.CheckBox chkEdit 
      Caption         =   "开单权"
      Height          =   195
      Left            =   7080
      TabIndex        =   9
      ToolTipText     =   "Ctrl+勾选：单独选择"
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox chkExec 
      Caption         =   "执行权"
      Height          =   195
      Left            =   8160
      TabIndex        =   8
      ToolTipText     =   "Ctrl+勾选：单独选择"
      Top             =   120
      Value           =   1  'Checked
      Width           =   840
   End
   Begin VB.TextBox txtFindItem 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5280
      TabIndex        =   7
      ToolTipText     =   "查找病人(Ctrl+F)"
      Top             =   120
      Width           =   1155
   End
   Begin VB.Frame fraDoctor 
      Caption         =   "医生"
      ForeColor       =   &H000040C0&
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3540
      Begin XtremeReportControl.ReportControl rptDoc 
         Height          =   5295
         Left            =   70
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
         _Version        =   589884
         _ExtentX        =   5953
         _ExtentY        =   9340
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CheckBox chk待审核 
         Caption         =   "只显示待审核的医生"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1050
         Width           =   2175
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   960
         MaxLength       =   30
         TabIndex        =   3
         Top             =   667
         Width           =   1905
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "查找(&F)"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   690
         Width           =   630
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室(&D)"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8265
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   635
      SimpleText      =   $"frmOPSEmpower.frx":6988
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOPSEmpower.frx":69CF
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17595
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":7263
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":DAC5
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":14327
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":148C1
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOPSEmpower.frx":14E5B
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmOPSEmpower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mstrPrivs As String
Private mlngModul As Long
Private mlngCodeType As Long         '0-拼音,1-五笔
Private mobjBar As CommandBar
Private mlngLevel As Long
Private mblnIsUpdate As Boolean
Private mblnNotRef As Boolean
Private mBln授权审核 As Boolean

Private mlngFindNum As Long
Private mlngFindItemNum As Long    '查找项目
'手术审核暂时不启用签名功能，所以判断加了 And 1 = 0
Private mblnTmp As Boolean
Private Enum Enum_Dor
    COL_人员ID = 0
    col_选择 = 1
    COL_姓名 = 2
    col_手术等级 = 3
    COL_拼音简码 = 4
    COL_五笔简码 = 5
    COL_所属部门 = 6
    COL_所属部门ID = 7
End Enum

Private Enum Enum_Advice
    col开单 = 0
    col执行 = 1
    col编码 = 2
    col手术名称 = 3
    col手术等级 = 4
    COL手术规模 = 5
    col服务对象 = 6
    COL站点 = 7
    COL简码 = 8
    COL申请人 = 9
    COL申请时间 = 10
End Enum



Private Sub cboDept_Click()
    Call LoadDoc
End Sub

Private Sub SaveEmpower(ByVal lngType As Long)
'功能：授权
'参数：lngType：0-授开单和执行权，1-授开单权，2-授执行权
    Dim strSql As String, blnCancel As Boolean
    Dim rsTmp As Recordset, i As Long
    Dim strDocs As String, lngDoc As Long
    Dim arrSql() As Variant
    Dim strItems As String
    Dim blnTrans As Boolean
    
    For i = 0 To rptDoc.Records.Count - 1
        If rptDoc.Records(i).Tag = "1" Then
            strDocs = strDocs & "," & rptDoc.Records(i)(COL_人员ID).Value
        End If
    Next
    strDocs = Mid(strDocs, 2)
    On Error GoTo errH
    If strDocs <> "" Then
        If InStr(strDocs, ",") > 0 Then
            '如果批量授权，则检查是否已经授过权，提示重新授权
            strSql = "Select /*+Rule */" & vbNewLine & _
                " f_List2str(Cast(Collect(姓名) As t_Strlist)) As 姓名" & vbNewLine & _
                "From (Select Distinct b.姓名" & vbNewLine & _
                "       From 人员手术权限 A, 人员表 B" & vbNewLine & _
                "       Where a.人员id = b.Id " & IIf(lngType > 0, " And A.记录性质=[2]", "") & " and a.人员id In (Select Column_Value From Table(f_Num2list([1]))))"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDocs, lngType)
            If rsTmp.RecordCount > 0 Then
                If rsTmp!姓名 & "" <> "" Then
                    If MsgBox("以下医生已经授权，是否要取消这些医生的权限重新授权？" & vbCrLf & rsTmp!姓名, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            lngDoc = 0
        Else
            lngDoc = Val(strDocs)
        End If
    Else
        lngDoc = Val(rptDoc.SelectedRows(0).Record(COL_人员ID).Value)
        strDocs = lngDoc
    End If
    strSql = _
            "Select ID, 上级id, 0 As 末级, 编码, 名称, Null As 简码, Null As 手术等级,  Null As 手术规模, Null As 服务对象, Null As 站点," & vbNewLine & _
            "       Null As 已勾选check" & vbNewLine & _
            "From 诊疗分类目录" & vbNewLine & _
            "Where 类型 =5 And (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "Start With 上级id Is Null" & vbNewLine & _
            "Connect By Prior ID = 上级id" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select  B.ID,B.分类ID,1 ,b.编码, b.名称,Upper(E.简码) as 简码, d.手术类型, b.操作类型, decode(B.服务对象,1,'门诊',2,'住院',3,'门诊和住院',4,'体检','不直接应用于病人') as 服务对象, b.站点," & IIf(lngType = 0, "Decode(Count(A.记录性质), 2, 1, 0)", "Decode(Max(NVL(A.记录性质,0)),0,0,1) ") & vbNewLine & _
            "From 人员手术权限 A, 诊疗项目目录 B, 疾病诊断对照 C, 疾病编码目录 D,诊疗项目别名 E" & vbNewLine & _
            "Where a.诊疗项目id(+) = b.Id And b.Id = c.手术id(+) And c.疾病id = d.Id(+) AND E.诊疗项目ID=B.ID And a.人员id(+) = [1] and e.码类=[2] And e.性质=1 And b.类别='F' And (B.撤档时间 Is Null Or B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            IIf(cboDept.ItemData(cboDept.ListIndex) <> -1, " and (exists(Select 1 From 诊疗适用科室 F Where F.项目ID=b.ID And F.科室ID=[3])  Or Not Exists(Select 1 From 诊疗适用科室 F Where F.项目ID=b.ID))", "") & _
            IIf(lngType = 0, "", " And A.记录性质(+) = " & lngType) & _
            "Group By b.Id,B.分类ID, b.编码, b.名称, b.操作类型, b.服务对象, b.站点, d.手术类型,E.简码"
    
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSql, 2, "手术项目", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, "不显示没有子项的分类", lngDoc, mlngCodeType + 1, cboDept.ItemData(cboDept.ListIndex))
    arrSql = Array()
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有手术数据可以选择。", vbInformation, gstrSysName
        End If
    Else
        rsTmp.Filter = "已勾选check=1 And 末级=1"
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                If Len(strItems & "," & rsTmp!ID) > 4000 Then
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = "Zl_人员手术权限_Update('" & strDocs & "','" & Mid(strItems, 2) & "'," & lngType & "," & IIf(UBound(arrSql) = 0, 1, 0) & IIf(mBln授权审核, "", ",1,'" & UserInfo.姓名 & "'") & ")"
                    strItems = ""
                End If
                strItems = strItems & "," & rsTmp!ID
                
                rsTmp.MoveNext
            Loop
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "Zl_人员手术权限_Update('" & strDocs & "','" & Mid(strItems, 2) & "'," & lngType & "," & IIf(UBound(arrSql) = 0, 1, 0) & IIf(mBln授权审核, "", ",1,'" & UserInfo.姓名 & "'") & ")"
        Else
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "Zl_人员手术权限_Update('" & strDocs & "',''," & lngType & ",1" & IIf(mBln授权审核, "", ",0,1,'" & UserInfo.姓名 & "'") & ")"
        End If

        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        If tbcSub.Selected.Caption = "手术项目" Then
            Call LoadItem
            If Not mBln授权审核 Then
                Call LoadCheck
            End If
        Else
            Call LoadCheck
        End If
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveModify()
'功能：调整删除权限
    Dim i As Long
    Dim arrSql() As Variant
    Dim blnTrans As Boolean
    Dim intType As Integer
    Dim intDelete As Integer
    
    With vsOPS
        arrSql = Array()
        If chkEdit.Value = 1 And chkExec.Value = 0 Then
            intDelete = 3
        ElseIf chkEdit.Value = 0 And chkExec.Value = 1 Then
            intDelete = 4
        Else
            intDelete = 2
        End If
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, col开单) <> .Cell(flexcpData, i, col开单) And chkEdit.Value = 1 Or .Cell(flexcpChecked, i, col执行) <> .Cell(flexcpData, i, col执行) And chkExec.Value = 1 Then
                If .Cell(flexcpChecked, i, col开单) = 1 And .Cell(flexcpChecked, i, col执行) = 2 Then
                    intType = 1
                ElseIf .Cell(flexcpChecked, i, col开单) = 2 And .Cell(flexcpChecked, i, col执行) = 1 Then
                    intType = 2
                ElseIf .Cell(flexcpChecked, i, col开单) = 2 And .Cell(flexcpChecked, i, col执行) = 2 Then
                    intType = 3
                Else
                    intType = 0
                End If
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = "Zl_人员手术权限_Update('" & Val(rptDoc.SelectedRows(0).Record(COL_人员ID).Value) & "','" & .RowData(i) & "'," & intType & "," & intDelete & IIf(mBln授权审核, "", ",1,'" & UserInfo.姓名 & "'") & ")"
            End If
        Next
        On err GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        If tbcSub.Selected.Caption = "手术项目" Then
            Call LoadItem
            If Not mBln授权审核 Then
                Call LoadCheck
            End If
        Else
            Call LoadCheck
        End If
    End With
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveCheck(i As Integer)
'功能：执行授权申请
'参数：i=0 通过申请、=1拒绝申请
    Dim strSql As String
    Dim blnTrans As Boolean

    If i = 1 Then
        strSql = "Zl_人员手术权限_Update('" & Val(rptDoc.SelectedRows(0).Record(COL_人员ID).Value) & "','0',0,0,3,null,'" & UserInfo.姓名 & "')"
    Else
        strSql = "Zl_人员手术权限_Update('" & Val(rptDoc.SelectedRows(0).Record(COL_人员ID).Value) & "','0',0,0,2,null,'" & UserInfo.姓名 & "')"
    End If
    On err GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(CStr(strSql), Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    Call LoadCheck
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub LoadItem()
'功能：加载医嘱拥有的权限。

    Dim rsTmp As Recordset
    Dim strSql As String
    
    If rptDoc.SelectedRows.Count > 0 Then
        If rptDoc.SelectedRows(0).GroupRow = False Then
            strSql = "Select b.Id,Decode(Min(记录性质),1,1,2) As 开单权,Decode(Max(记录性质),2,1,2) As 执行权,b.编码,b.名称,f.编码||'-'||f.名称 as 操作类型," & _
                " decode(B.服务对象,1,'门诊',2,'住院',3,'门诊和住院',4,'体检','不直接应用于病人') as 服务对象,b.站点,d.手术类型,Upper(E.简码) as 简码" & _
                " From 人员手术权限 A,诊疗项目目录 B,疾病诊断对照 C,疾病编码目录 D,诊疗项目别名 E,诊疗手术规模 F" & _
                " Where a.诊疗项目id=b.Id And b.Id=c.手术id(+) And c.疾病id=d.Id(+) AND E.诊疗项目ID=B.ID And b.操作类型 in (f.编码,f.名称) And a.人员id=[1]" & _
                " and e.码类=[2] And (B.撤档时间 Is Null Or B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                IIf(chkEdit.Value = 0, " And a.记录性质 <>1", "") & IIf(chkExec.Value = 0, " And a.记录性质 <>2", "") & _
                " Group By b.Id,b.编码,b.名称,b.服务对象,b.站点,d.手术类型,E.简码,f.编码,f.名称"
                
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(rptDoc.SelectedRows(0).Record(COL_人员ID).Value), mlngCodeType + 1)
            With vsOPS
                .Rows = 1
                Do While Not rsTmp.EOF
                    .AddItem ""
                    .RowData(.Rows - 1) = rsTmp!ID & ""
                    .Cell(flexcpChecked, .Rows - 1, col开单) = Val(rsTmp!开单权 & "")
                    '用于取消后恢复状态
                    .Cell(flexcpData, .Rows - 1, col开单) = Val(rsTmp!开单权 & "")
                    .Cell(flexcpChecked, .Rows - 1, col执行) = Val(rsTmp!执行权 & "")
                    .Cell(flexcpData, .Rows - 1, col执行) = Val(rsTmp!执行权 & "")
                    .TextMatrix(.Rows - 1, col编码) = rsTmp!编码
                    .TextMatrix(.Rows - 1, col手术名称) = rsTmp!名称 & ""
                    .TextMatrix(.Rows - 1, col手术等级) = rsTmp!手术类型 & ""
                    .TextMatrix(.Rows - 1, COL手术规模) = rsTmp!操作类型 & ""
                    .TextMatrix(.Rows - 1, col服务对象) = rsTmp!服务对象 & ""
                    .TextMatrix(.Rows - 1, COL站点) = rsTmp!站点 & ""
                    .TextMatrix(.Rows - 1, COL简码) = rsTmp!简码 & ""
                    rsTmp.MoveNext
                Loop
                
                If .Rows = 1 Then .AddItem ""
                .Cell(flexcpBackColor, 1, col开单, .Rows - 1, col执行) = &HE1FFE1
                If chkEdit.Value = 0 Then
                    .ColHidden(col开单) = True
                Else
                    .ColHidden(col开单) = False
                End If
                If chkExec.Value = 0 Then
                    .ColHidden(col执行) = True
                Else
                    .ColHidden(col执行) = False
                End If
            End With
        Else
            vsOPS.Rows = 1: vsOPS.AddItem ""
        End If
        mlngFindItemNum = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCheck()
'功能：加载待审核的手术授权。
    Dim rsTmp As Recordset
    Dim strSql As String
    
    If rptDoc.SelectedRows.Count > 0 Then
        If rptDoc.SelectedRows(0).GroupRow = False Then
            strSql = "Select b.Id, Decode(A.权限,2, 2,3,2, 1) As 开单权, Decode(A.权限,1, 2,3,2, 1) As 执行权, b.编码, b.名称," & vbNewLine & _
                        "       f.编码 || '-' || f.名称 As 操作类型, Decode(b.服务对象, 1, '门诊', 2, '住院', 3, '门诊和住院', 4, '体检', '不直接应用于病人') As 服务对象, b.站点," & vbNewLine & _
                        "       d.手术类型, Upper(e.简码) As 简码,a.申请人,a.申请时间,a.审核状态,a.审批人,a.审批时间" & vbNewLine & _
                        "From 人员手术权限申请 A, 诊疗项目目录 B, 疾病诊断对照 C, 疾病编码目录 D, 诊疗项目别名 E, 诊疗手术规模 F" & vbNewLine & _
                        "Where a.诊疗项目id = b.Id And A.审核状态 =1 And b.Id = c.手术id(+) And c.疾病id = d.Id(+) And e.诊疗项目id = b.Id And b.操作类型 In (f.编码, f.名称) And" & vbNewLine & _
                        "      a.授权人员id = [1] And e.码类 = [2] And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) order by a.申请时间,b.编码"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(rptDoc.SelectedRows(0).Record(COL_人员ID).Value), mlngCodeType + 1)
            If Not rsTmp.EOF Then
                tbcSub.Item(1).Caption = IIf(rsTmp.RecordCount = 0, "待审核授权", "待审核授权：" & rsTmp.RecordCount & " 项")
            Else
                tbcSub.Item(1).Caption = "待审核授权"
            End If
            With vsSQ
                .Rows = 1
                Do While Not rsTmp.EOF
                    .AddItem ""
                    .RowData(.Rows - 1) = rsTmp!ID & ""
                    .Cell(flexcpChecked, .Rows - 1, col开单) = Val(rsTmp!开单权 & "")
                    '用于取消后恢复状态
                    .Cell(flexcpData, .Rows - 1, col开单) = Val(rsTmp!开单权 & "")
                    .Cell(flexcpChecked, .Rows - 1, col执行) = Val(rsTmp!执行权 & "")
                    .Cell(flexcpData, .Rows - 1, col执行) = Val(rsTmp!执行权 & "")
                    .TextMatrix(.Rows - 1, col编码) = rsTmp!编码
                    .TextMatrix(.Rows - 1, col手术名称) = rsTmp!名称 & ""
                    .TextMatrix(.Rows - 1, col手术等级) = rsTmp!手术类型 & ""
                    .TextMatrix(.Rows - 1, COL手术规模) = rsTmp!操作类型 & ""
                    .TextMatrix(.Rows - 1, col服务对象) = rsTmp!服务对象 & ""
                    .TextMatrix(.Rows - 1, COL站点) = rsTmp!站点 & ""
                    .TextMatrix(.Rows - 1, COL简码) = rsTmp!简码 & ""
                    .TextMatrix(.Rows - 1, COL申请人) = rsTmp!申请人 & ""
                    .TextMatrix(.Rows - 1, COL申请时间) = Format(rsTmp!申请时间 & "", "yyyy-MM-dd HH:mm")
                    rsTmp.MoveNext
                Loop
                
                If .Rows = 1 Then .AddItem ""
            End With
        Else
            vsSQ.Rows = 1: vsSQ.AddItem ""
        End If
        mlngFindItemNum = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub LoadDoc()
'加载权限比操作员低的医生
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long, y As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngPrssID As Long, lngSelectRow As Long, lngDept As Long
    
    
    If cboDept.ListIndex = -1 Then Exit Sub
    
    If Val(cboDept.ItemData(cboDept.ListIndex)) = -1 Then
        rptDoc.GroupsOrder.DeleteAll
    Else
        If InStr(";" & mstrPrivs & ";", ";所有部门;") > 0 And rptDoc.GroupsOrder.Count = 0 Then rptDoc.GroupsOrder.Add rptDoc.Columns(COL_所属部门)
    End If
    strSql = "Select DISTINCT a.Id, A.性别" & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", ",b.部门ID,e.名称 as 所属部门") & ",a.姓名,a.手术等级, Upper(zlSpellCode(a.姓名)) As 拼音简码, Upper(Zlwbcode(a.姓名)) As 五笔简码" & vbNewLine & _
            "From 人员表 A, 部门人员 B, 人员性质说明 D,部门表 E" & vbNewLine & _
            "Where a.Id = b.人员id And e.ID=b.部门ID And d.人员id = a.Id  And d.人员性质 = '医生' And " & vbNewLine & _
            "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null)  " & vbNewLine & _
            "   " & IIf(Val(cboDept.ItemData(cboDept.ListIndex)) = -1, "", "And b.部门id=[2]") & IIf(chk待审核.Value = 1, " And (Exists(Select 1 From 人员手术权限申请 F Where F.授权人员id = A.id And F.审核状态 = 1))", "")
            
    
    On Error GoTo errH
    
    rptDoc.Records.DeleteAll
    
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngLevel, Val(cboDept.ItemData(cboDept.ListIndex)))
    
    With rptDoc
        i = 0
        lngSelectRow = -1
        Do While Not rsTmp.EOF
            Set objRecord = .Records.Add()
            Set objItem = objRecord.AddItem(rsTmp!ID & "")
            Set objItem = objRecord.AddItem("")
            Set objItem = objRecord.AddItem(rsTmp!姓名 & "")
                objItem.Icon = img16.ListImages.Item(IIf(rsTmp!性别 & "" = "女", "feMale", "Male")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!手术等级 & "")
            Set objItem = objRecord.AddItem(rsTmp!拼音简码 & "")
            Set objItem = objRecord.AddItem(rsTmp!五笔简码 & "")
            If Val(cboDept.ItemData(cboDept.ListIndex)) <> -1 Then
                Set objItem = objRecord.AddItem(rsTmp!所属部门 & "")
                Set objItem = objRecord.AddItem(rsTmp!部门ID & "")
            End If

            
            rsTmp.MoveNext
            i = i + 1
        Loop
        .Populate
        If lngPrssID <> 0 Then
            vsOPS.Rows = 1
            vsOPS.AddItem ""
        End If
    End With
    mlngFindNum = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim objPopup As CommandBarPopup
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_Edit_Untread     '取消
        With vsOPS
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, col开单) <> .Cell(flexcpData, i, col开单) Or .Cell(flexcpChecked, i, col执行) <> .Cell(flexcpData, i, col执行) Then
                    .Cell(flexcpChecked, i, col开单) = .Cell(flexcpData, i, col开单)
                    .Cell(flexcpChecked, i, col执行) = .Cell(flexcpData, i, col执行)
                End If
            Next
        End With
        mblnIsUpdate = False
    Case conMenu_Manage_Complete '授权通过
        If MsgBox("确认要通过对" & IIf(rptDoc.SelectedRows(0).Record(COL_姓名).Value = "", "该", rptDoc.SelectedRows(0).Record(COL_姓名).Value) & "医生的授权申请？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
        Call SaveCheck(0)
    Case conMenu_Manage_UnArrange '授权拒绝
        If MsgBox("确认要拒绝对" & IIf(rptDoc.SelectedRows(0).Record(COL_姓名).Value = "", "该", rptDoc.SelectedRows(0).Record(COL_姓名).Value) & "医生的授权申请？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Exit Sub
        End If
        Call SaveCheck(1)
    Case conMenu_Edit_Save        '保存
        Call SaveModify
        mblnIsUpdate = False
    Case conMenu_Kss_Grant  '授开单和执行权
        Call SaveEmpower(0)
    Case conMenu_Kss_Grant * 100# + 1 '授开单权
        Call SaveEmpower(1)
    Case conMenu_Kss_Grant * 100# + 2 '授执行权
        Call SaveEmpower(2)
    Case conMenu_View_Find '查找
        txtFind.SetFocus '有时需要定位一下
        If txtFind.Text <> "" Then
            Call txtFind_KeyPress(vbKeyReturn)
        End If
    Case conMenu_View_FindNext '查找下一个
        If Me.ActiveControl.Name = "txtFindItem" Or Me.ActiveControl.Name = "vsOPS" Then
            If txtFindItem.Text = "" Then
                txtFindItem.SetFocus
            Else
                Call txtFindItem_KeyPress(vbKeyReturn)
            End If
        Else
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call txtFind_KeyPress(vbKeyReturn)
            End If
        End If
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
        cbsMain_Resize
    Case conMenu_View_Refresh '刷新
        Call LoadItem
        Call LoadCheck
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '退出
        Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With fraDoctor
        .Top = lngTop
        .Left = lngLeft + 100
        .Height = lngBottom - lngTop - stbThis.Height
    End With
    rptDoc.Height = fraDoctor.Height - IIf(mBln授权审核, 1540, 1250)
    
    
    With tbcSub
        .Top = fraDoctor.Top
        .Height = fraDoctor.Height
        .Width = Me.Width - fraDoctor.Left - fraDoctor.Width - 400
        picOPS.Width = .Width
        picOPS.Height = .Height - 350
        vsOPS.Top = 0: vsOPS.Left = 0: vsOPS.Width = picOPS.Width: vsOPS.Height = picOPS.Height
        PicSQ.Width = .Width
        PicSQ.Height = .Height - 350
        vsSQ.Top = 0: vsSQ.Left = 0: vsSQ.Width = PicSQ.Width: vsSQ.Height = PicSQ.Height
    End With
    
    
    Me.Refresh
End Sub

Private Sub SetControlVisible(ByRef Control As XtremeCommandBars.ICommandBarControl)
    '根据权限设置按钮可见状态
    
    Select Case Control.ID
            Case conMenu_Kss_Grant, conMenu_Kss_Grant * 100# + 1, conMenu_Kss_Grant * 100# + 2, conMenu_Edit_Save, conMenu_Edit_Untread
                Control.Visible = tbcSub.Selected.Caption = "手术项目"
            Case conMenu_Manage_Complete, conMenu_Manage_UnArrange
                Control.Visible = tbcSub.Selected.Caption <> "手术项目"
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim rptRecord As ReportRecord
        
'    '根据权限设置按钮可见状态
    If mblnIsUpdate Then
        If Control.ID = conMenu_Edit_Untread Or Control.ID = conMenu_Edit_Save Then
            Control.Enabled = True
            If Visible And fraDoctor.Enabled = True Then fraDoctor.Enabled = False
        Else
            Control.Enabled = False
        End If
        Exit Sub
    Else
        Control.Enabled = True
        If Visible And fraDoctor.Enabled = False Then fraDoctor.Enabled = True
    End If
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    Select Case Control.ID
        Case conMenu_Edit_Untread, conMenu_Edit_Save
            Control.Enabled = mblnIsUpdate
        Case conMenu_Kss_Grant  '授权
            blnEnabled = rptDoc.SelectedRows.Count > 0
            If rptDoc.SelectedRows.Count > 0 Then
                blnEnabled = rptDoc.SelectedRows(0).GroupRow = False
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_ToolBar_Button '工具栏
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbsMain.Count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_FindNext '查找下一个
            Control.Visible = False
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
        Case conMenu_Manage_Complete '审核通过
            Control.Enabled = mBln授权审核
            If Control.Enabled = True Then
                blnEnabled = rptDoc.SelectedRows.Count > 0
                If rptDoc.SelectedRows.Count > 0 Then
                    blnEnabled = rptDoc.SelectedRows(0).GroupRow = False
                End If
                Control.Enabled = blnEnabled
            End If
        Case conMenu_Manage_UnArrange '审核不通过
            Control.Enabled = mBln授权审核
            If Control.Enabled = True Then
                blnEnabled = rptDoc.SelectedRows.Count > 0
                If rptDoc.SelectedRows.Count > 0 Then
                    blnEnabled = rptDoc.SelectedRows(0).GroupRow = False
                End If
                Control.Enabled = blnEnabled
            End If
    End Select
End Sub


Private Sub chkEdit_Click()
    If chkExec.Value = 0 And chkEdit.Value = 0 Then
        mblnNotRef = True
        chkEdit.Value = 1: Exit Sub
        mblnNotRef = False
    End If
    If mblnNotRef = True Then
        mblnNotRef = False
        Exit Sub
    End If
    If tbcSub.Selected.Caption = "手术项目" Then
        Call LoadItem
    End If
End Sub

Private Sub chkExec_Click()
    If chkEdit.Value = 0 And chkExec.Value = 0 Then
        mblnNotRef = True
        chkExec.Value = 1: Exit Sub
        mblnNotRef = False
    End If
    If mblnNotRef = True Then
        mblnNotRef = False
        Exit Sub
    End If
    If tbcSub.Selected.Caption = "手术项目" Then
        Call LoadItem
    End If
End Sub

Private Sub chk待审核_Click()
    Call LoadDoc
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim strHead As String
    
    mstrPrivs = GetPrivFunc(glngSys, 1080)
    mBln授权审核 = InStr(mstrPrivs, "授权审核") > 0
    
    mlngModul = 1080
    mlngCodeType = zlDatabase.GetPara("简码方式")
    mblnIsUpdate = False
    
    rptDoc.Top = IIf(mBln授权审核, 1440, 1150)
    chk待审核.Visible = mBln授权审核
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    
    
    'TabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "手术项目", picOPS.hwnd, 0).Tag = "手术项目"
        .InsertItem(1, "待审核授权", PicSQ.hwnd, 0).Tag = "待审核授权"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    
    'vsFlexGrid
    '-----------------------------------------------------
    strHead = "开单,700,1;执行,700,1 ;编码,1000,1;手术名称,2500,1;手术等级,1000,1;手术规模,1000,1;服务对象,1000,1;院区,800,7;简码"
    Call InitTable(vsOPS, strHead)
    vsOPS.Editable = flexEDKbdMouse
    vsOPS.Cell(flexcpPictureAlignment, 0, col开单) = flexPicAlignLeftCenter
    vsOPS.Cell(flexcpPictureAlignment, 0, col执行) = flexPicAlignLeftCenter
    vsOPS.Cell(flexcpAlignment, 0, col开单) = flexPicAlignRightCenter
    vsOPS.Cell(flexcpAlignment, 0, col执行) = flexPicAlignRightCenter
    vsOPS.Cell(flexcpPicture, 0, col开单) = img16.ListImages("unCheck").Picture
    vsOPS.Cell(flexcpPicture, 0, col执行) = img16.ListImages("unCheck").Picture
    vsOPS.ColDataType(col开单) = flexDTBoolean
    vsOPS.ColDataType(col执行) = flexDTBoolean
    
    strHead = "开单,700,1;执行,700,1 ;编码,1000,1;手术名称,2500,1;手术等级,1000,1;手术规模,1000,1;服务对象,1000,1;院区,800,7;简码;申请人,700,1;申请时间,1700,1"
    Call InitTable(vsSQ, strHead)
    vsSQ.Editable = flexEDNone
    vsSQ.ColDataType(col开单) = flexDTBoolean
    vsSQ.ColDataType(col执行) = flexDTBoolean

    
    
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    
    Call RestoreWinState(Me, App.ProductName)
    
    
    Call LoadDept
End Sub

Private Sub LoadDept()
'加载操作员所属科室
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long
    
    strSql = "Select B.ID,B.编码,B.名称 " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", ",A.缺省") & vbNewLine & _
            "From " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", "部门人员 A, ") & _
            " 部门表 B, 部门性质说明 C" & vbNewLine & _
            " Where B.Id = C.部门id " & _
            IIf(InStr(";" & mstrPrivs & ";", ";所有部门;") > 0, "", " And a.部门id = B.Id And A.人员ID = [1] ") & vbNewLine & _
            "  And C.工作性质 = '临床' And C.服务对象 <> 0  And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) Order By B.编码"

    On Error GoTo errH
    cboDept.Clear
    '所有部门
    If InStr(";" & mstrPrivs & ";", ";所有部门;") > 0 Then
        cboDept.AddItem "所有部门"
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        '所属缺省
        If InStr(";" & mstrPrivs & ";", ";所有部门;") = 0 Then
            If rsTmp!缺省 = 1 Then
                Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex)
            End If
        End If
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
    End If
    Call LoadDoc
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngidx As Long, i As Long

    With rptDoc
        
        Set objCol = .Columns.Add(COL_人员ID, "人员ID", 0, False)
        Set objCol = .Columns.Add(col_选择, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("unCheck").Index - 1
        Set objCol = .Columns.Add(COL_姓名, "姓名", 70, True)
        Set objCol = .Columns.Add(col_手术等级, "手术等级", 80, True)
        Set objCol = .Columns.Add(COL_拼音简码, "拼音简码", 0, False)
        Set objCol = .Columns.Add(COL_五笔简码, "五笔简码", 0, False)
        Set objCol = .Columns.Add(COL_所属部门, "所属部门", 0, False)
        Set objCol = .Columns.Add(COL_所属部门ID, "所属部门ID", 0, False)


        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的医生..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.img16
        If InStr(";" & mstrPrivs & ";", ";所有部门;") > 0 Then .GroupsOrder.Add .Columns(COL_所属部门)
    End With
End Sub

Private Sub InitTable(vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Kss_Grant, "授开单和执行权")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant, "授开单和执行权")
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant * 100# + 1, "授开单权")
            objControl.IconId = conMenu_Kss_Grant
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant * 100# + 2, "授执行权")
            objControl.IconId = conMenu_Kss_Grant
        End With
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set mobjBar = cbsMain.Add("工具栏", xtpBarTop)
    With mobjBar.Controls

        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "授权通过"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_UnArrange, "拒绝授权"): objControl.IconId = 4114
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Kss_Grant, "授开单和执行权")
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant, "授开单和执行权")
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant * 100# + 1, "授开单权")
            objControl.IconId = conMenu_Kss_Grant
            Set objControl = .Add(xtpControlButton, conMenu_Kss_Grant * 100# + 2, "授执行权")
            objControl.IconId = conMenu_Kss_Grant
        End With

        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
    End With
    With cbsMain.ActiveMenuBar.Controls
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Show * 100# + 1, "")
        objCustom.Handle = chkEdit.hwnd
        objCustom.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Show * 100# + 2, "")
        objCustom.Handle = chkExec.hwnd
        objCustom.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlCustom, conMenu_View_Find * 100# + 1, "  查找")
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFindItem.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
End Sub

Private Sub Form_Resize()
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnIsUpdate = True Then
        If MsgBox("当前输入的内容未保存，是否要退出？", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmParent Is Nothing Then Set mfrmParent = Nothing
    mlngFindNum = 0
    mlngFindItemNum = 0
End Sub

Private Sub rptDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptDoc.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptDoc_RowDblClick(rptDoc.SelectedRows(0), rptDoc.SelectedRows(0).Record.Item(col_选择))
        End If
    End If
End Sub

Private Sub rptDoc_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '如果点击表头的图片，就选中全部
    If Button = 1 Then
        If rptDoc.HitTest(x, y).ht = xtpHitTestHeader Then
            Set objColumn = rptDoc.HitTest(x, y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_选择 Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptDoc.Columns(col_选择).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptDoc.Records.Count - 1
                            rptDoc.Records(i)(col_选择).Icon = img16.ListImages("Check").Index - 1
                            rptDoc.Records(i).Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptDoc.Columns(col_选择).Icon = img16.ListImages("unCheck").Index - 1
                        For i = 0 To rptDoc.Records.Count - 1
                            rptDoc.Records(i)(col_选择).Icon = -1
                            rptDoc.Records(i).Tag = "0"
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptDoc_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(col_选择).Icon = -1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(col_选择).Icon = img16.ListImages.Item("Check").Index - 1
        Row.Record.Tag = "1"
    End If
    rptDoc.Populate
End Sub

Private Sub rptDoc_SelectionChanged()
    If mlngFindNum <> 0 Then mlngFindNum = rptDoc.SelectedRows(0).Index + 1
    
    '加载已授权的手术项目
    If tbcSub.Selected.Caption = "手术项目" Then
        Call LoadItem
    Else
        Call LoadCheck
    End If
End Sub

Private Sub rptDoc_SortOrderChanged()
    mlngFindNum = 0
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Caption = "手术项目" Then
        Call LoadItem
    Else
        Call LoadCheck
    End If
End Sub

Private Sub txtFind_Change()
    mlngFindNum = 0
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFind)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    Dim i As Long
    Dim blnIsAllChar As Boolean
    Dim blnIsFind As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With rptDoc
        strMsg = UCase(Trim(txtFind.Text))
        If zlCommFun.IsCharAlpha(strMsg) Then blnIsAllChar = True
        
        For i = mlngFindNum To rptDoc.Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If blnIsAllChar Then
                    If .Rows(i).Record(COL_姓名).Value Like IIf(gstrMatch = "", "", "*") & strMsg & "*" Or _
                            .Rows(i).Record(IIf(mlngCodeType = 0, COL_拼音简码, COL_五笔简码)).Value Like IIf(gstrMatch = "", "", "*") & strMsg & "*" Then
                        '该行选中且显示在可见区域,并引发SelectionChanged事件
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                Else
                    If .Rows(i).Record(COL_姓名).Value Like IIf(gstrMatch = "", "", "*") & strMsg & "*" Then
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        Exit Sub
                    End If
                End If
            End If
        Next
        If mlngFindNum = 0 Then
            MsgBox "当前部门没有找到您查找的医生。", vbInformation, Me.Caption
        ElseIf mlngFindNum <> 0 And blnIsFind = False Then
            MsgBox "已经是最后一个医生了。", vbInformation, Me.Caption
            mlngFindNum = 0
        End If
    End With
End Sub

Private Sub txtFindItem_Change()
    mlngFindItemNum = 0
End Sub

Private Sub txtFindItem_GotFocus()
    If txtFind.Text <> "" Then
        Call zlControl.TxtSelAll(txtFindItem)
    End If
End Sub

Private Sub txtFindItem_KeyPress(KeyAscii As Integer)
    Dim i As Long, int性质 As Integer
    Dim strFind As String
    If KeyAscii = vbKeyReturn Then
        With vsOPS
            strFind = UCase(Trim(txtFindItem.Text))
            If zlCommFun.IsCharChinese(txtFindItem.Text) Then
                '中文的只查名称
                int性质 = 1
            ElseIf zlCommFun.IsCharAlpha(txtFindItem.Text) Then
                '英文查名称和简码
                int性质 = 2
            Else
                '否则查名称简码和编码
                int性质 = 3
            End If
            For i = mlngFindItemNum To .Rows - 1
                If int性质 = 1 Then
                    If UCase(.TextMatrix(i, col手术名称)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Then
                        .Row = i
                        .ShowCell i, col手术名称
                        mlngFindItemNum = i + 1
                        Exit Sub
                    End If
                ElseIf int性质 = 2 Then
                    If UCase(.TextMatrix(i, col手术名称)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Or UCase(.TextMatrix(i, COL简码)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Then
                        .Row = i
                        .ShowCell i, col手术名称
                        mlngFindItemNum = i + 1
                        Exit Sub
                    End If
                Else
                    If UCase(.TextMatrix(i, col手术名称)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Or UCase(.TextMatrix(i, COL简码)) Like IIf(gstrMatch = "", "", "*") & strFind & "*" Or UCase(.TextMatrix(i, col编码)) = strFind Then
                        .Row = i
                        .ShowCell i, col手术名称
                        mlngFindItemNum = i + 1
                        Exit Sub
                    End If
                End If
            Next
            If mlngFindItemNum = 0 Then
                MsgBox "当前医生不具备您查询的手术项目开单和执行权限。", vbInformation, Me.Caption
            ElseIf mlngFindItemNum <> 0 Then
                MsgBox "已经查找完最后一个项目了。", vbInformation, Me.Caption
                mlngFindItemNum = 0
            End If
        End With
    End If
End Sub

Private Sub vsOPS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col开单 Or Col = col执行 Then
        If mblnIsUpdate = False Then mblnIsUpdate = True
    End If
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mlngFindItemNum <> 0 Then mlngFindItemNum = NewRow
End Sub

Private Sub vsOPS_AfterSort(ByVal Col As Long, Order As Integer)
    mlngFindItemNum = 0
End Sub

Private Sub vsOPS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col开单 And Col <> col执行 Then
        Cancel = True
    End If
End Sub

Private Sub vsOPS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsOPS.RowData(Row) & "" = "" Then
        Cancel = True
    End If
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
'功能:记录表打印
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    
    If rptDoc.Visible = False Then Exit Sub
    If rptDoc.Records.Count > 0 Then
        If rptDoc.SelectedRows.Count = 0 Then Exit Sub
        If rptDoc.SelectedRows(0).GroupRow Then Exit Sub
        strSubhead = rptDoc.SelectedRows(0).Record(COL_姓名).Value & IIf(tbcSub.Selected.Caption = "手术项目", "手术权限清单", "待审核权限")
    Else
        Exit Sub
    End If
    
    '调用打印部件处理
    Set objPrint.Body = IIf(tbcSub.Selected.Caption = "手术项目", Me.vsOPS, Me.vsSQ)
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人:" & UserInfo.姓名)
    Call objAppRow.Add("打印时间:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

