VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChildPatient 
   BorderStyle     =   0  'None
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   1
      Left            =   630
      ScaleHeight     =   1935
      ScaleWidth      =   3000
      TabIndex        =   8
      Top             =   3465
      Width           =   3000
      Begin MSComctlLib.TreeView tvw 
         Height          =   1635
         Left            =   300
         TabIndex        =   9
         Top             =   210
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2884
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2625
      Index           =   0
      Left            =   750
      ScaleHeight     =   2625
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   510
      Width           =   3000
      Begin VB.CheckBox chkNCommit 
         Caption         =   "包含已经出院但病案未提交病人"
         Height          =   225
         Left            =   60
         TabIndex        =   13
         Top             =   660
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.ComboBox cboStatus 
         Height          =   300
         Left            =   825
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   345
         Width           =   2130
      End
      Begin VB.PictureBox picSelect 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   135
         ScaleHeight     =   435
         ScaleWidth      =   3720
         TabIndex        =   2
         Top             =   2175
         Width           =   3720
         Begin VB.CommandButton cmdStatus 
            Cancel          =   -1  'True
            Height          =   285
            Left            =   2535
            Picture         =   "frmChildPatient.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "打开下方选中的文件"
            Top             =   90
            UseMaskColor    =   -1  'True
            Width           =   300
         End
         Begin VB.Label labStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "选中      病案      份"
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   105
            Width           =   1980
         End
         Begin VB.Label labNum 
            Alignment       =   2  'Center
            Caption         =   "0"
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
            Height          =   240
            Left            =   1515
            TabIndex        =   4
            Top             =   90
            Width           =   345
         End
         Begin VB.Label labSelect 
            Alignment       =   2  'Center
            Caption         =   "审查"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   450
            TabIndex        =   3
            Top             =   90
            Width           =   570
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1200
         Left            =   150
         TabIndex        =   6
         Top             =   930
         Width           =   1845
         _cx             =   3254
         _cy             =   2117
         Appearance      =   2
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
         MergeCells      =   1
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
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   825
         TabIndex        =   1
         Top             =   30
         Width           =   2130
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病案状态"
         Height          =   180
         Left            =   45
         TabIndex        =   11
         Top             =   390
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院科室"
         Height          =   180
         Left            =   45
         TabIndex        =   7
         Top             =   75
         Width           =   720
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChildPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private Const CB_GETDROPPEDSTATE = &H157

Private mfrmMain As Object
Private mlngKey As Long
Private mlngReferKey As Long
Private mblnReading As Boolean
Private mstrSQL As String
Private mblnDataChanged As Boolean
Private mbytApplyMode   As Byte               '1-出院已提交病人;3-在院病人;4-已批准的借阅病人
Private mbytMode        As Byte
Private mrsCondition    As ADODB.Recordset
Private mclsVsf      As clsVsf
Private mlng病人ID      As Long
Private mlng主页ID      As Long
Private mstrKey         As String
Private mstrSvr科室名称 As String
Private mblnDrop        As Boolean
Private mrsDept         As ADODB.Recordset
Private mrsData         As ADODB.Recordset
Private mstrPrivs       As String
Private mstrDepts       As String
Private blnReadUsed     As Boolean
Private mblnRead病案结构 As Boolean

Public Event AfterDeptChanged()
Public Event StatusChanged()
Public Event DbClick()
Public Event AfterEdit(ByVal Row As Long, ByVal Col As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event AfterDocumentChanged(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strObject As String, ByVal strParam As String, ByVal strCaption As String, ByVal lng提交Id As Long, ByVal blnDataMove As Boolean, ByVal blnScale As Boolean)

Public Property Get Depts() As String
    Depts = mstrDepts
End Property

Public Property Let Depts(ByVal vDepts As String)
    mstrDepts = vDepts
End Property
Public Sub cboDeptRefresh(strDeptName As String)
    On Error GoTo ErrH
    If cboDept.Text <> "" Then
        cboDept.ListIndex = GetCboIndex(cboDept, zlCommFun.GetNeedName(strDeptName))
    End If
    Call cboDept_Click
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Function GetCboIndex(cbo As ComboBox, strFind As String, Optional blnKeep As Boolean, Optional blnLike As Boolean) As Long
    '功能：由字符串在ComboBox中查找索引
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '先精确查找
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), "-") > 0 Then
            If zlCommFun.GetNeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '最后模糊查找
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Public Function VsfBody() As VSFlexGrid
    Set VsfBody = vsf
End Function

Public Property Get Title() As String
    If Not tvw.SelectedItem Is Nothing Then
        Title = tvw.SelectedItem.Text
    End If
End Property

Public Function zlColumnSelect() As Boolean
    If frmTemplateColumn.ShowColumn(mfrmMain, mclsVsf) Then
        mclsVsf.AppendRows = True
    End If
End Function

Public Function zlLocationDocument(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt反馈对象 As Byte, ByVal strFileKey As String)
    Dim strKey As String
    Dim intRow As Integer
    Dim objNode As Node
    
    '1-住院医嘱;2-住院病历;3-护理病历;4-护理记录;5-首页记录;6-医嘱报告;7-疾病证明;8-知情文件
    With vsf
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("病人id"))) = lng病人ID And Val(.TextMatrix(intRow, .ColIndex("主页id"))) = lng主页ID Then
                
                .Row = intRow
                If IsNumeric(strFileKey) Then
                    strKey = "R" & byt反馈对象
                    If strFileKey <> "" And strFileKey <> "0,0,0" Then strKey = strKey & "K" & strFileKey
                End If
                
                On Error Resume Next
                mblnRead病案结构 = True
                Set objNode = tvw.Nodes(strKey)
                If Not (objNode Is Nothing) Then
                    objNode.EnsureVisible
                    objNode.Selected = True
                    If mblnRead病案结构 Then
                        mblnRead病案结构 = False
                        Call tvw_NodeClick(objNode)
                    End If
                    zlLocationDocument = True
                End If
                Exit Function
            End If
        Next
    End With
End Function

Public Function zlLocationPatient(Optional ByVal bytApplyMode As Byte = 1, Optional ByVal strFindKey As String, Optional ByVal strLocationText As String, Optional ByVal strNo As String, _
    Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long, Optional ByVal lng科室ID As Long, Optional ByVal strFindDeal As String) As Boolean
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim intCol As Integer
    Dim intRow As Integer
    Dim bytMatch As Byte
    Dim intLoop As Integer
    Dim strCols As String
    Dim i           As Integer
    
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "ID", adVarChar, 30
        .Fields.Append "姓名", adVarChar, 100
        .Fields.Append "性别", adVarChar, 10
        .Fields.Append "出院科室", adVarChar, 50
        .Fields.Append "床号", adVarChar, 30
        .Fields.Append "住院号", adVarChar, 30
        .Fields.Append "出院科室ID", adDouble, 50
        .Fields.Append "入院日期", adDBTimeStamp, 20
        .Fields.Append "出院日期", adDBTimeStamp, 20
        .Open
    End With
    
    With vsf
        Select Case bytApplyMode
             Case 0
                
                If lng病人ID > 0 And lng主页ID > 0 Then
                    For intRow = 1 To .Rows - 1
                        
                        If .TextMatrix(intRow, .ColIndex("No")) = strNo Then
                            If Val(.TextMatrix(intRow, .ColIndex("病人id"))) = lng病人ID And Val(.TextMatrix(intRow, .ColIndex("主页id"))) = lng主页ID Then
                                '找到了病人
                                .Row = intRow
                                .ShowCell .Row, .Col
                                GoTo endHand
                            End If
                        End If
                    Next
                End If
            Case 1
                
                If lng病人ID > 0 And lng主页ID > 0 Then
                    For intRow = 1 To .Rows - 1
                        If .TextMatrix(intRow, .ColIndex("No")) = strNo Then
                            If Val(.TextMatrix(intRow, .ColIndex("病人id"))) = lng病人ID And Val(.TextMatrix(intRow, .ColIndex("主页id"))) = lng主页ID Then
                                '找到了病人
                                .Row = intRow
                                .ShowCell .Row, .Col
                                GoTo endHand
                            End If
                        End If
                    Next
                End If
            Case 2
                intRow = -1
                bytMatch = 2
                intCol = mclsVsf.ColIndex(strFindKey)
                '查找出来的先放到记录集中，如果有多条时，弹出对话框进行选择确认
                If intCol > 0 Then
                    If strFindKey = "姓名" Then
                        mrsData.Filter = "姓名 like '%" & strLocationText & "%'"
                    ElseIf strFindKey = "住院号" Then
                        If IsNumeric(strLocationText) Then
                          mrsData.Filter = strFindKey & "  = '" & strLocationText & "'"
                        End If
                    Else
                        mrsData.Filter = strFindKey & "  like '%" & strLocationText & "%'"
                    End If
                    Do While Not (mrsData.EOF Or mrsData.BOF)
                        '找到了,填写
                        rsData.AddNew
                        rsData("ID").Value = mrsData!ID
                        rsData("姓名").Value = NVL(mrsData!姓名)
                        rsData("住院号").Value = NVL(mrsData!住院号)
                        rsData("出院科室").Value = NVL(mrsData!出院科室)
                        rsData("性别").Value = NVL(mrsData!性别)
                        rsData("床号").Value = NVL(mrsData!床号)
                        rsData("出院科室ID").Value = mrsData!出院科室ID
                        rsData("入院日期").Value = NVL(mrsData!入院日期, 0)
                        rsData("出院日期").Value = NVL(mrsData!出院日期, 0)
                        mrsData.MoveNext
                    Loop
                    rsData.Sort = "入院日期"
                    If rsData.RecordCount = 0 Then Exit Function
                    If rsData.RecordCount = 1 Then
                        Set rs = rsData
                    Else
                        rsData.MoveFirst
                        strCols = "姓名,900,0,;性别,500,0,;床号,600,0,;住院号,1200,0,;出院科室,1500,0,;入院日期,1600,0,;出院日期,1600,0,"
                        If ShowPubSelect(mfrmMain, mfrmMain.txtLocation, 2, strCols, "私有模块\" & App.ProductName & "\" & Me.Name & "\定位查找病人" & mbytApplyMode, "请从下表中选择您想找的病人", rsData, rs, 8790, 4500, , , , True) <> 1 Then
                            Exit Function
                        End If
                    End If
                    .SetFocus
                    DoEvents
                    '如果找不到病人，且当前科室不是所有科室，则读取当前人所在科室的数据
                    If Not (cboDept.Text = "所有科室" Or cboDept.ItemData(cboDept.ListIndex) = rs!出院科室ID) Then
                        If mbytApplyMode <> 3 Then
                            If Val(labNum.Caption) > 0 Then
                                If MsgBox("你已选择了【" & labNum.Caption & "】份病案，当前操作将重新刷新数据？" & vbCrLf & "确认执行该操作？", vbOKCancel + vbQuestion + vbDefaultButton2, ParamInfo.产品名称) = vbCancel Then
                                    Exit Function
                                End If
                            End If
                        End If
                        For i = 0 To cboDept.ListCount - 1
                            If cboDept.ItemData(i) = rs!出院科室ID Then
                                cboDept.ListIndex = i
                                Exit For
                            End If
                        Next
                    End If
                    intRow = mclsVsf.FindRow(rs("ID").Value, .ColIndex("ID"), bytMatch, .Row + 1)
                    If intRow = -1 Then
                        intRow = mclsVsf.FindRow(rs("ID").Value, .ColIndex("ID"), bytMatch)
                    End If
                    .Row = intRow
                    .ShowCell .Row, .Col
                    DoEvents
                    GoTo endHand
                End If
                

            Case 3
                If mbytApplyMode = 3 And cboDept.ListIndex >= 0 Then
                    If cboDept.ItemData(cboDept.ListIndex) <> lng科室ID And lng科室ID > 0 Then
                        zlControl.CboLocate cboDept, lng科室ID, True
                    End If
                End If
                
                If lng病人ID > 0 And lng主页ID > 0 Then
                    For intRow = 1 To .Rows - 1
                        If Val(.TextMatrix(intRow, .ColIndex("病人id"))) = lng病人ID And Val(.TextMatrix(intRow, .ColIndex("主页id"))) = lng主页ID Then
                            '找到了病人
                            .Row = intRow
                            .ShowCell .Row, .Col
                            GoTo endHand
                        End If
                    Next
                End If
            End Select
    End With
    zlLocationPatient = False
    
    Exit Function
endHand:
    With vsf
        If bytApplyMode <> 3 And .ColIndex("选择") >= 0 Then
        
            Select Case strFindDeal
            Case "查找并选中"
                
                .TextMatrix(.Row, .ColIndex("选择")) = 1
                
            Case "查找并不选"
                
                .TextMatrix(.Row, .ColIndex("选择")) = 0
                    
            Case "查找并反选"
                
                If Abs(Val(.TextMatrix(.Row, .ColIndex("选择")))) = 0 Then
                    .TextMatrix(.Row, .ColIndex("选择")) = 1
                Else
                    .TextMatrix(.Row, .ColIndex("选择")) = 0
                End If
                
            End Select
    
        End If
    End With
    
    zlLocationPatient = True
End Function

Public Function zlInitData(ByVal frmMain As Object, ByVal bytApplyMode As Byte, Optional ByVal strPrivs As String) As Boolean
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    mbytApplyMode = bytApplyMode
    If InitControl = False Then Exit Function
    If InitData = False Then Exit Function
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then '使用个性化设置
        mclsVsf.LoadStateFromString Trim(GetRegister(私有模块, Me.Name, "表格参数_20100719_" & mbytApplyMode, ""))
    End If
    zlInitData = True

End Function

Public Function zlRefreshData(Optional ByVal rsCondition As ADODB.Recordset, Optional ByVal lng病人ID As Long, Optional ByVal lng主页ID As Long) As Boolean
    Set mrsCondition = rsCondition
    mbytMode = 2
    mstrKey = ""
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    
    If ExecuteCommand("刷新数据", 0) = False Then Exit Function
    If mbytApplyMode <> 4 Then
        Call cboDept_Click
    End If
    
    zlRefreshData = True
    
End Function

Public Function zlRefreshStruct() As Boolean
    '******************************************************************************************************************
    '功能：刷新当前的档案
    '参数：
    '返回：
    '******************************************************************************************************************
     zlRefreshStruct = ExecuteCommand("读取病案结构", "Read")
    
End Function

Public Function zlShowDocument() As Boolean
    mstrKey = ""
    Call ExecuteCommand("读取病案结构", "Read")
    
End Function

'######################################################################################################################

Private Function GetParamRecord(ByVal strParam As String) As ADODB.Recordset
    Dim varTmp As Variant
    Dim varAry As Variant
    Dim rs As New ADODB.Recordset
    Dim intCount As Integer
    Dim intCol As Integer
    
    If strParam <> "" Then

        '设了审查科室范围参数
        With rs
            .Fields.Append "人员id", adBigInt
            .Fields.Append "科室id", adBigInt
            .Open
        End With
        
        varTmp = Split(strParam, ";")
                
         For intCount = 0 To UBound(varTmp)

            If varTmp(intCount) <> "" Then
                varAry = Split(varTmp(intCount), ",")
                For intCol = 1 To UBound(varAry)
                    rs.AddNew
                    rs("人员id").Value = Val(varAry(0))
                    rs("科室id").Value = Val(varAry(intCol))
                Next
                
            End If
         Next
         
         If rs.RecordCount > 0 Then rs.MoveFirst
    End If
    
    Set GetParamRecord = rs
End Function
Private Function InitControl() As Boolean

    mblnReading = True

    Set mclsVsf = New clsVsf
    With mclsVsf

        Select Case mbytApplyMode
            Case 1               '审查病人
            
                Call .Initialize(Me.Controls, vsf, True, True, frmPubResource.GetImageList(16))
                Call .ClearColumn
                Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("病案状态值", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("数据转出", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("就诊卡号", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("床号", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("封存时间", 0, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True, , , True)
                
                Call .AppendColumn("", 240, flexAlignCenterCenter, flexDTBoolean, , "[选择]", False)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "[图标]", False)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[路径]", False)
                Call .AppendColumn("姓名", 810, flexAlignLeftCenter, flexDTString, , , True)

                Call .AppendColumn("住院号", 900, flexAlignLeftCenter, flexDTDecimal, , , True)
                
                Call .AppendColumn("年龄", 500, flexAlignLeftCenter, flexDTDecimal, , , True)
                Call .AppendColumn("护理等级", 810, flexAlignLeftCenter, flexDTDecimal, , , True)
                Call .AppendColumn("住院医师", 810, flexAlignLeftCenter, flexDTDecimal, , , True)

                Call .AppendColumn("出院科室", 1080, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("病案状态", 840, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("提交人", 810, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("提交时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("接收人", 810, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("接收时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("拒审人", 810, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("拒审时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("拒审理由", 600, flexAlignLeftCenter, flexDTString, , , True)
                
                Call .AppendColumn("出院科室ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("反馈条数", 0, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("反馈完成", 0, flexAlignLeftCenter, flexDTString, "", , True) '=1完成 =0未完成
                
                Call .AppendColumn("入院日期", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("出院日期", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("性别", 450, flexAlignLeftCenter, flexDTString, "", , True)
                
                .SysHidden(.ColIndex("ID")) = True
                .SysHidden(.ColIndex("病人id")) = True
                .SysHidden(.ColIndex("主页id")) = True
                .SysHidden(.ColIndex("病案状态值")) = True
                .SysHidden(.ColIndex("就诊卡号")) = True
                .SysHidden(.ColIndex("床号")) = True
                .SysHidden(.ColIndex("封存时间")) = True
                .SysHidden(.ColIndex("数据转出")) = True
                .SysHidden(.ColIndex("出院科室ID")) = True
                .SysHidden(.ColIndex("反馈条数")) = True
                .SysHidden(.ColIndex("反馈完成")) = True
                .SysHidden(.ColIndex("性别")) = True
                .SysHidden(.ColIndex("入院日期")) = True
                .SysHidden(.ColIndex("出院日期")) = True
                
                Call .InitializeEdit(True, False, False)
                Call .InitializeEditColumn(.ColIndex("选择"), True, vbVsfEditCheck)
            Case 3                  '在院病人
                
                Call .Initialize(Me.Controls, vsf, True, False, frmPubResource.GetImageList(16))
                Call .ClearColumn
                
                Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("就诊卡号", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("病案状态值", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("数据转出", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("封存时间", 0, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd", , True, , , True)
                                
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[路径]", False)
                Call .AppendColumn("姓名", 840, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("性别", 450, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("床号", 600, flexAlignRightCenter, flexDTDecimal, "", , True)
                Call .AppendColumn("住院号", 900, flexAlignLeftCenter, flexDTDecimal, "", , True)
                
                Call .AppendColumn("年龄", 500, flexAlignLeftCenter, flexDTDecimal, , , True)
                Call .AppendColumn("护理等级", 810, flexAlignLeftCenter, flexDTDecimal, , , True)
                Call .AppendColumn("住院医师", 810, flexAlignLeftCenter, flexDTDecimal, , , True)

                Call .AppendColumn("审查状态", 840, flexAlignLeftCenter, flexDTString, , , True)
                Call .AppendColumn("入院日期", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("出院日期", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                
                Call .AppendColumn("出院科室ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
    
                
                .SysHidden(.ColIndex("ID")) = True
                .SysHidden(.ColIndex("病人id")) = True
                .SysHidden(.ColIndex("主页id")) = True
                .SysHidden(.ColIndex("病案状态值")) = True
                .SysHidden(.ColIndex("就诊卡号")) = True
                .SysHidden(.ColIndex("封存时间")) = True
                .SysHidden(.ColIndex("数据转出")) = True
                .SysHidden(.ColIndex("出院科室ID")) = True
                
                
            Case 4                  '已批准的借阅病人
            
                Call .Initialize(Me.Controls, vsf, True, False, frmPubResource.GetImageList(16))
                Call .ClearColumn
                
                Call .AppendColumn("ID", 0, flexAlignLeftCenter, flexDTString, , , True, , , True)
                Call .AppendColumn("病人id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("主页id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("就诊卡号", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("数据转出", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
                Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[路径]", False)
                Call .AppendColumn("姓名", 840, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("性别", 450, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("住院号", 900, flexAlignLeftCenter, flexDTDecimal, "", , True)
                
                Call .AppendColumn("No", 900, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("申请人", 840, flexAlignLeftCenter, flexDTString, "", , True)
                Call .AppendColumn("申请时间", 990, flexAlignLeftCenter, flexDTString, "yyyy-MM-dd HH:mm:ss", , True)
                Call .AppendColumn("申请理由", 600, flexAlignLeftCenter, flexDTString, "", , True)
                
                Call .AppendColumn("出院科室ID", 0, flexAlignLeftCenter, flexDTString, "", , True)
                
                .SysHidden(.ColIndex("ID")) = True
                .SysHidden(.ColIndex("病人id")) = True
                .SysHidden(.ColIndex("主页id")) = True
                .SysHidden(.ColIndex("就诊卡号")) = True
                .SysHidden(.ColIndex("数据转出")) = True
                .SysHidden(.ColIndex("出院科室ID")) = True
            
        End Select
        .AppendRows = True
    End With
    DoEvents
     
    '划分停靠区域

    Dim objPane As Pane
    Set objPane = dkpMain.CreatePane(1, 100, 200, DockLeftOf, Nothing): objPane.Title = "病人列表": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 200, DockBottomOf, objPane): objPane.Title = "电子病案": objPane.Options = PaneNoCaption

    Call DockPannelInit(dkpMain)
    

    Select Case mbytApplyMode
        Case 3
            lbl.Caption = "住院科室"
        Case Else
            lbl.Caption = "出院科室"
    End Select


    
    mblnReading = False
    InitControl = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function InitData(Optional ByVal btnShowDept As Boolean = False) As Boolean
Dim strTmp As String
Dim rs As New ADODB.Recordset
Dim rsParam As New ADODB.Recordset
    
    mblnReading = True
    Set mrsDept = New ADODB.Recordset
    With mrsDept
        .Fields.Append "编码", adVarChar, 20
        .Fields.Append "名称", adVarChar, 100
        .Fields.Append "简码", adVarChar, 30
        .Open
    End With
    
    cboDept.Clear
    Select Case mbytApplyMode
        Case 4
            cboDept.AddItem "所有科室"
        Case Else
            If IsPrivs(mstrPrivs, "所有科室") Then cboDept.AddItem "所有科室"
    End Select
    

    If mbytApplyMode = 4 Then
        Set rs = gclsPackage.GetDept("临床", , btnShowDept)
    Else
        Set rs = gclsPackage.GetDept("临床", , IsPrivs(mstrPrivs, "所有科室"), btnShowDept)
    End If
    
    If rs.BOF = False Then
        
        strTmp = Trim(zlDatabase.GetPara("审查科室范围", ParamInfo.系统号, mfrmMain.模块号))
        
        If strTmp = "" Then
            '没有设审查科室范围参数
            Do While Not rs.EOF
                cboDept.AddItem rs("显示名称").Value
                cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
                mstrDepts = mstrDepts & rs("ID").Value & ","
                mrsDept.AddNew
                mrsDept("编码").Value = rs("编码").Value
                mrsDept("名称").Value = rs("名称").Value
                mrsDept("简码").Value = rs("简码").Value & ""
                
                rs.MoveNext
            Loop
            mstrDepts = Left(mstrDepts, Len(mstrDepts) - 1)
        Else
            '设了审查科室范围参数
            Set rsParam = GetParamRecord(strTmp)
            rsParam.Filter = ""
            rsParam.Filter = "人员id=" & UserInfo.ID
            If rsParam.RecordCount > 0 Then
                cboDept.Clear
                
                Do While Not rs.EOF
                    rsParam.Filter = ""
                    rsParam.Filter = "人员id=" & UserInfo.ID & " And 科室id=" & Val(rs("ID").Value)
                    If rsParam.RecordCount > 0 Then
                        cboDept.AddItem rs("显示名称").Value
                        cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
                        mstrDepts = mstrDepts & rs("ID").Value & ","
                        
                        mrsDept.AddNew
                        mrsDept("编码").Value = rs("编码").Value
                        mrsDept("名称").Value = rs("名称").Value
                        mrsDept("简码").Value = rs("简码").Value & ""
                    End If
                    rs.MoveNext
                Loop
                mstrDepts = Left(mstrDepts, Len(mstrDepts) - 1)
            Else

                Do While Not rs.EOF
                    cboDept.AddItem rs("显示名称").Value
                    cboDept.ItemData(cboDept.NewIndex) = rs("ID").Value
                    mstrDepts = mstrDepts & rs("ID").Value & ","
                    mrsDept.AddNew
                    mrsDept("编码").Value = rs("编码").Value
                    mrsDept("名称").Value = rs("名称").Value
                    mrsDept("简码").Value = rs("简码").Value & ""
                    
                    rs.MoveNext
                Loop
                mstrDepts = Left(mstrDepts, Len(mstrDepts) - 1)
            End If
            
        End If
    End If
    
    If cboDept.ListCount = 0 Then
        cboDept.AddItem ""
        cboDept.ItemData(cboDept.NewIndex) = -1
    End If
    
    cboDept.ListIndex = 0
    mstrSvr科室名称 = cboDept.Text

    '添加病案状态
    cboStatus.Clear
    If mbytApplyMode = 1 Then
        cboStatus.AddItem "所有状态"
        cboStatus.ItemData(cboStatus.NewIndex) = 0
        cboStatus.AddItem "提交待收"
        cboStatus.ItemData(cboStatus.NewIndex) = 1
        cboStatus.AddItem "接收待审"
        cboStatus.ItemData(cboStatus.NewIndex) = 10
        cboStatus.AddItem "正在审查"
        cboStatus.ItemData(cboStatus.NewIndex) = 3
        cboStatus.AddItem "审查反馈"
        cboStatus.ItemData(cboStatus.NewIndex) = 4
        cboStatus.AddItem "审查整改"
        cboStatus.ItemData(cboStatus.NewIndex) = 6
        cboStatus.AddItem "审查归档"
        cboStatus.ItemData(cboStatus.NewIndex) = 5
        cboStatus.ListIndex = 0
    End If
    
    mblnReading = False
    InitData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ExecuteCommand(ByVal strCmd As String, ParamArray varParam() As Variant) As Boolean
Dim rs As New ADODB.Recordset
Dim lngLoop As Long
Dim intRow As Integer
    
    On Error GoTo errHand
    
    mblnReading = True
    Select Case strCmd
    Case "读取出院病人"
        If mlng病人ID = 0 Then
            mclsVsf.SaveKey = vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID"))
            mclsVsf.ClearGrid
            Set rs = gclsPackage.GetAduitPatient(3, ParamRead(mrsCondition, "审查开始时间"), ParamRead(mrsCondition, "审查结束时间"), _
                                                Val(ParamRead(mrsCondition, "接收待审")) & ";" & Val(ParamRead(mrsCondition, "拒绝接收")) & ";" & Val(ParamRead(mrsCondition, "正在审查")) & ";" & Val(ParamRead(mrsCondition, "审查反馈")) & ";" & Val(ParamRead(mrsCondition, "审查整改")) & ";" & Val(ParamRead(mrsCondition, "提交待收")), 0, 0, _
                                                ParamRead(mrsCondition, "出院情况"), Val(ParamRead(mrsCondition, "病人类型")), ParamRead(mrsCondition, "医保种类"), _
                                                ParamRead(mrsCondition, "住院医师"), ParamRead(mrsCondition, "疾病名称"), ParamRead(mrsCondition, "检查类型"), _
                                                ParamRead(mrsCondition, "药品信息"), ParamRead(mrsCondition, "医嘱开始时间"), ParamRead(mrsCondition, "医嘱结束时间") _
                                                )
            If rs.BOF = False Then
                Call mclsVsf.LoadDataSource(rs)
            End If
        Else
            
            Set rs = gclsPackage.GetAduitPatient(3, ParamRead(mrsCondition, "审查开始时间"), ParamRead(mrsCondition, "审查结束时间"), _
                                                Val(ParamRead(mrsCondition, "接收待审")) & ";" & Val(ParamRead(mrsCondition, "拒绝接收")) & ";" & Val(ParamRead(mrsCondition, "正在审查")) & ";" & Val(ParamRead(mrsCondition, "审查反馈")) & ";" & Val(ParamRead(mrsCondition, "审查整改")) & ";" & Val(ParamRead(mrsCondition, "提交待收")), _
                                                mlng病人ID, mlng主页ID, ParamRead(mrsCondition, "出院情况"), Val(ParamRead(mrsCondition, "病人类型")), ParamRead(mrsCondition, "医保种类"), _
                                                ParamRead(mrsCondition, "住院医师"), ParamRead(mrsCondition, "疾病名称"), ParamRead(mrsCondition, "检查类型"), _
                                                ParamRead(mrsCondition, "药品信息"), ParamRead(mrsCondition, "医嘱开始时间"), ParamRead(mrsCondition, "医嘱结束时间") _
                                                )
            If rs.BOF = False Then
                intRow = 0
                For lngLoop = 1 To vsf.Rows - 1
                    If Val(vsf.TextMatrix(lngLoop, vsf.ColIndex("病人id"))) = mlng病人ID And Val(vsf.TextMatrix(lngLoop, vsf.ColIndex("主页id"))) = mlng主页ID Then
                        intRow = lngLoop
                        Exit For
                    End If
                Next
                If intRow > 0 Then
                    '已加载
                    vsf.Row = intRow
                    Call mclsVsf.LoadGridRow(vsf.Row, rs)
                End If
            End If
        End If
        Set mrsData = rs
    Case "读取病案结构"
        
        Dim objNode As Node
        Dim strIcon As String
        Dim strKey As String
        
        If Not (tvw.SelectedItem Is Nothing) Then strKey = tvw.SelectedItem.Key
        If InStr(strKey, "K") = 0 And strKey <> "R1" And strKey <> "R5" Then strKey = ""
        
        LockWindowUpdate tvw.hWnd
        
        tvw.Nodes.Clear
                
        With vsf
            Set rs = gclsPackage.GetCISStruct(Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), Val(.TextMatrix(.Row, .ColIndex("出院科室ID"))), Val(.TextMatrix(.Row, .ColIndex("数据转出"))) = 1)
        End With
                
        If rs.BOF = False Then
            
            '报告：Decode(a.医嘱id, Null, a.名称, '<'||b.医嘱内容||'>' || a.名称 || '(' || To_Char(b.开始执行时间, 'yyyy-mm-dd') || ')') As 名称, Trim(To_Char(a.ID))||';'||Decode(a.医嘱id,Null,'0',Trim(To_Char(a.医嘱id))) As 参数
            '护理：A.名称 || '(' || B.名称 || '：' || To_Char(A.开始, 'yyyy-mm-dd hh24:mi') || ' ～ ' ||To_Char(A.截止, 'yyyy-mm-dd hh24:mi') || ')' As 名称, Trim(To_Char(B.ID))||';'||Trim(To_Char(Nvl(保留,0)))||';'||To_Char(A.开始, 'yyyy-mm-dd hh24:mi')||' ～ '||To_Char(A.截止, 'yyyy-mm-dd hh24:mi')||';'||Trim(To_Char(A.ID)) As 参数
            
            Do While Not rs.EOF
                strIcon = zlCommFun.NVL(rs("图标").Value)

                If zlCommFun.NVL(rs("上级id").Value) = "" Then
                    Set objNode = tvw.Nodes.Add(, , rs("ID").Value, rs("名称").Value, strIcon, strIcon)
                    objNode.Tag = zlCommFun.NVL(rs("参数").Value)
                Else
                    Set objNode = tvw.Nodes.Add(rs("上级id").Value, tvwChild, rs("ID").Value, rs("名称").Value, strIcon, strIcon)
                    objNode.Tag = zlCommFun.NVL(rs("参数").Value)
                End If
            
                rs.MoveNext
            Loop
        End If
        
        With vsf
            Set rs = New ADODB.Recordset
            Set rs = gclsPackage.GetEmrCISStruct(Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))))
        End With
        If Not rs Is Nothing Then
            If rs.State = ADODB.adStateOpen Then
            If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do Until rs.EOF
                Set objNode = tvw.Nodes.Add(rs!上级ID.Value, tvwChild, rs!ID.Value, rs!名称.Value, rs!图标.Value, rs!图标.Value)
                objNode.Tag = NVL(rs!参数) '文档ID[|子文档ID]
                rs.MoveNext
            Loop
            End If
            End If
        End If

        If tvw.Nodes.count > 0 Then
            
            On Error Resume Next
            Err = 0
            If strKey <> "" Then tvw.Nodes(strKey).Selected = True
            On Error GoTo errHand
            
            If Err <> 0 Or strKey = "" Or tvw.SelectedItem Is Nothing Then
                tvw.Nodes(1).Selected = True
            End If
            
            If Not (tvw.SelectedItem Is Nothing) Then
                If varParam(0) = "NoRead" Then
                    If mblnRead病案结构 Then
                        mblnRead病案结构 = False
                        With vsf
                            If mbytApplyMode = 3 Then
                                RaiseEvent AfterDocumentChanged(0, 0, "首页记录", "", "", 0, Val(.TextMatrix(.Row, .ColIndex("数据转出"))) = 1, False)
                            Else
                                RaiseEvent AfterDocumentChanged(0, 0, "首页记录", "", "", IIf(vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID")) = "", 0, vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID"))), Val(.TextMatrix(.Row, .ColIndex("数据转出"))) = 1, False)
                            End If
                        End With
                    End If
                Else
                    mblnRead病案结构 = True
                    Call tvw_NodeClick(tvw.SelectedItem)
                End If
            End If
        Else
            With vsf
                RaiseEvent AfterDocumentChanged(0, 0, "首页记录", "", "", vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID")), Val(.TextMatrix(.Row, .ColIndex("数据转出"))) = 1, False)
            End With
        End If
        
        LockWindowUpdate 0
        Call vsf_RowColChange
    Case "刷新数据"
        Select Case mbytApplyMode
            Case 1          '出院病人
                ExecuteCommand = ExecuteCommand("读取出院病人")
            Case 3          '在院病人
                ExecuteCommand = RefreshPatientIn
            Case 4          '已批准借阅的病人
                ExecuteCommand = RefreshPatientBorrow
        End Select
    End Select

    ExecuteCommand = True
    GoTo endHand
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:
    mblnReading = False
End Function
Private Function RefreshPatientBorrow() As Boolean
'读取借阅人员
Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    mclsVsf.SaveKey = vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID"))
    mclsVsf.ClearGrid
    
    Set rs = gclsPackage.GetBorrowPatient(0, ParamRead(mrsCondition, "开始单据号"), _
                                            ParamRead(mrsCondition, "结束单据号"), _
                                            ParamRead(mrsCondition, "申请人"), _
                                            ParamRead(mrsCondition, "批准人"), _
                                            ParamRead(mrsCondition, "拒绝人"), _
                                            IIf(Val(ParamRead(mrsCondition, "新登记单据")) = 1, ParamRead(mrsCondition, "登记开始日期"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "新登记单据")) = 1, ParamRead(mrsCondition, "登记结束日期"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "已批准单据")) = 1, ParamRead(mrsCondition, "批准开始日期"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "已批准单据")) = 1, ParamRead(mrsCondition, "批准结束日期"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "已拒绝单据")) = 1, ParamRead(mrsCondition, "拒绝开始日期"), ""), _
                                            IIf(Val(ParamRead(mrsCondition, "已拒绝单据")) = 1, ParamRead(mrsCondition, "拒绝结束日期"), ""))
    If rs.BOF = False Then
        Call mclsVsf.LoadDataSource(rs)
    End If
    Set mrsData = rs
    
    RefreshPatientBorrow = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function RefreshPatientIn() As Boolean
 '在院病人和出院未提交病人
    Dim l As Long, rs As New ADODB.Recordset
    
    On Error GoTo errHand
    mclsVsf.SaveKey = vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID")): mclsVsf.ClearGrid
    
    Set rs = gclsPackage.GetDeptPatient(ParamRead(mrsCondition, "出院开始时间"), ParamRead(mrsCondition, "出院结束时间"), _
                                        ParamRead(mrsCondition, "当前病况"), ParamRead(mrsCondition, "出院情况"), _
                                        Val(ParamRead(mrsCondition, "病人类型")), ParamRead(mrsCondition, "医保种类"), _
                                        ParamRead(mrsCondition, "住院医师"), ParamRead(mrsCondition, "疾病名称"), _
                                        ParamRead(mrsCondition, "检查类型"), ParamRead(mrsCondition, "药品信息"), _
                                        ParamRead(mrsCondition, "医嘱开始时间"), ParamRead(mrsCondition, "医嘱结束时间"), _
                                        chkNCommit.Value = vbChecked)
    If cboDept.ListIndex >= 0 Then
        If cboDept.Text = "所有科室" Then
            rs.Filter = ""
        Else
            rs.Filter = "出院科室ID = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
        End If
    End If
                    
    If rs.BOF = False Then
        Call mclsVsf.LoadDataSource(rs)
        vsf.Row = 1
    End If
        
    If mlng病人ID <> 0 Then
        For l = 1 To vsf.Rows - 1
            If Val(vsf.TextMatrix(l, vsf.ColIndex("病人id"))) = mlng病人ID And _
                Val(vsf.TextMatrix(l, vsf.ColIndex("主页id"))) = mlng主页ID Then
                vsf.Row = l: Exit For
            End If
        Next
        mclsVsf.AppendRows = True
    End If
    Set mrsData = rs
    Call vsf_RowColChange
    RefreshPatientIn = True
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub cboDept_Click()
    Dim rsTemp As New ADODB.Recordset
    If mblnReading Then Exit Sub
    
    mstrSvr科室名称 = cboDept.Text
    If mrsData Is Nothing Then
        Call ExecuteCommand("刷新数据")
    Else
        If mbytApplyMode = 3 Or mbytApplyMode = 4 Then
            '在院病人
             If cboDept.Text = "所有科室" Then
                mrsData.Filter = ""
             Else
                mrsData.Filter = "出院科室ID = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
             End If
        Else
            If cboDept.Text = "所有科室" Then
                If cboStatus.Text = "所有状态" Then
                    mrsData.Filter = "撤档时间 = '3000-01-01'"
                Else
                    mrsData.Filter = "撤档时间= '3000-01-01' and 病案状态值='" & cboStatus.ItemData(cboStatus.ListIndex) & "'"
                End If
                mfrmMain.AllowModify = True
            Else
                If cboStatus.Text = "所有状态" Then
                    mrsData.Filter = "出院科室ID = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
                Else
                    
                    mrsData.Filter = "出院科室ID = '" & cboDept.ItemData(cboDept.ListIndex) & "' And 病案状态值='" & cboStatus.ItemData(cboStatus.ListIndex) & "'"
                End If
            End If
        End If
        Call mclsVsf.LoadDataSource(mrsData)
        '已加载
        vsf.Row = 1
        Call mclsVsf.LoadGridRow(vsf.Row, mrsData)
        mclsVsf.AppendRows = True
    End If
    Call ExecuteCommand("读取病案结构", "Read")
    Call vsf_RowColChange
    
    If cboDept.ListIndex <> 0 And Me.cboStatus.Visible Then
        gstrSQL = "SELECT a.id FROM 部门表 a Where  a.id=[1] and ( TO_CHAR (A.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or A.撤档时间 is null) "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, cboDept.ItemData(cboDept.ListIndex))
        If rsTemp.RecordCount = 0 Then mfrmMain.AllowModify = False Else mfrmMain.AllowModify = True
    End If
    RaiseEvent AfterDeptChanged
End Sub

Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If cboDept.Locked Then Exit Sub
    cboDept.Tag = "Changed"
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cboDept.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim StrText As String
    Dim strResult As String
    
    If InStr(1, "—'|[](){}*%", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If cboDept.Locked Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        StrText = UCase(cboDept.Text)
        If cboDept.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If StrText <> cboDept.List(cboDept.ListIndex) Then Call zlControl.CboSetIndex(cboDept.hWnd, -1)
        End If
        If StrText = "" Then
            cboDept.ListIndex = 0
            cboDept.Tag = ""
        ElseIf cboDept.ListIndex = -1 Then
            intIdx = -1
            
            If mrsDept.State = adStateOpen Then
                If IsNumeric(StrText) Then                              '数字型编码
                    mrsDept.Filter = "编码 like '" & StrText & "*'"
                ElseIf zlCommFun.IsCharAlpha(StrText) Then              '字符型简码
                    mrsDept.Filter = "简码 like '*" & StrText & "*'"
                ElseIf zlCommFun.IsCharChinese(StrText) Then            '中文
                    mrsDept.Filter = "名称 like '*" & StrText & "*'"
                Else                                                    '编号支持类似N001,简码可能有ZYK01这种
                    mrsDept.Filter = "(编码 like '" & StrText & "*') OR (简码 like '*" & StrText & "*')"
                End If
                If mrsDept.RecordCount > 0 Then
                    mrsDept.MoveFirst
                    strResult = mrsDept("名称").Value     '只取第一个
                End If
            End If

            If mrsDept.State = adStateOpen Then mrsDept.Filter = ""
                        
            If strResult <> "" Then
                For i = 0 To cboDept.ListCount - 1
                    If zlCommFun.GetNeedName(cboDept.List(i)) = strResult Then
                        cboDept.ListIndex = i
                        cboDept.Tag = ""
                        Exit For
                    End If
                Next
            Else    '输入12，但只有1201,1202等;输入ZY,但只有ZYK,ZYH等
                For i = 0 To cboDept.ListCount - 1
                    If UCase(cboDept.List(i)) Like StrText & "*" Then
                        If intIdx = -1 Then
                            cboDept.ListIndex = i
                            cboDept.Tag = ""
                        End If
                        
                        intIdx = i
                    End If
                Next
                
            End If
            
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cboDept_Click
'            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cboDept.ListIndex = -1 Then
            cboDept.Text = ""
        Else
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cboDept_Click
            ElseIf intIdx <> cboDept.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cboDept.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cboDept_Click
            End If
        End If
    End If
End Sub

Private Sub cboDept_LostFocus()
    If cboDept.Tag = "Changed" Then
        cboDept.Text = mstrSvr科室名称
    End If
End Sub

Private Sub cboDept_Validate(Cancel As Boolean)
    If cboDept.Text <> "" Then
        If GetCboIndex(cboDept, zlCommFun.GetNeedName(cboDept.Text)) = -1 Then cboDept.ListIndex = -1: cboDept.Text = ""
    End If
    If cboDept.Text = "" Then Call cboDept_KeyPress(vbKeyReturn)
    
    If cboDept.ListIndex = -1 Then Cancel = True
End Sub

Private Sub cboStatus_Click()
    If mblnReading Then Exit Sub
    
    If mrsData Is Nothing Then
        Call ExecuteCommand("刷新数据")
    Else
        If cboStatus.Text = "所有状态" Then
            If cboDept.Text = "所有科室" Then
                mrsData.Filter = ""
            Else
                mrsData.Filter = "出院科室ID = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
            End If
        Else
            If cboDept.Text = "所有科室" Then
                mrsData.Filter = "病案状态值='" & cboStatus.ItemData(cboStatus.ListIndex) & "'"
            Else
                mrsData.Filter = "出院科室ID = '" & cboDept.ItemData(cboDept.ListIndex) & "' And 病案状态值='" & cboStatus.ItemData(cboStatus.ListIndex) & "'"
            End If
        End If
        Call mclsVsf.LoadDataSource(mrsData)
        '已加载
        vsf.Row = 1
        Call mclsVsf.LoadGridRow(vsf.Row, mrsData)
        mclsVsf.AppendRows = True
    End If
    Call ExecuteCommand("读取病案结构", "Read")
    Call vsf_RowColChange
    RaiseEvent StatusChanged
End Sub

Private Sub chkNCommit_Click()
    Call ExecuteCommand("刷新数据")
End Sub

Private Sub cmdStatus_Click()
    Call FileBatPrint
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Initialize()
    On Error GoTo ErrH
    DoEvents
    frmPubResource.Hide     '加载一下图标窗口
    Set tvw.ImageList = frmPubResource.ils16
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub Form_Load()
    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call SetPaneRange(dkpMain, 2, 100, 100, Me.ScaleWidth, 300)
    dkpMain.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If Not mclsVsf Is Nothing Then Call SetRegister(私有模块, Me.Name, "表格参数_20100719_" & mbytApplyMode, mclsVsf.SaveStateToString)
    Set mfrmMain = Nothing
    Set mrsCondition = Nothing
    Set mclsVsf = Nothing
    Set mrsDept = Nothing
    Set mrsData = Nothing
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
        Case 0
            cboDept.Move cboDept.Left, cboDept.Top, picPane(Index).Width - cboDept.Left - 15
            If mbytApplyMode = 1 Then '出院病人
                cboStatus.Move cboStatus.Left, cboDept.Height + 30, picPane(Index).Width - cboDept.Left - 15
                vsf.Move 0, cboStatus.Top + cboStatus.Height + 30, picPane(Index).Width, picPane(Index).Height - (cboStatus.Top + cboStatus.Height + 30) - 400
                lblStatus.Visible = True
                cboStatus.Visible = True
                chkNCommit.Visible = False
            ElseIf mbytApplyMode = 3 Then '在院
                chkNCommit.Move cboDept.Left, cboDept.Height + 30, picPane(Index).Width - chkNCommit.Left - 15
                vsf.Move 0, chkNCommit.Top + chkNCommit.Height + 30, picPane(Index).Width, picPane(Index).Height - (chkNCommit.Top + chkNCommit.Height + 30) - 400
                lblStatus.Visible = False
                cboStatus.Visible = False
                chkNCommit.Visible = True
            Else '借阅病人
                vsf.Move 0, cboDept.Top + cboDept.Height + 30, picPane(Index).Width, picPane(Index).Height - (cboDept.Top + cboDept.Height + 30) - 400
                lblStatus.Visible = False
                cboStatus.Visible = False
                chkNCommit.Visible = False
            End If
            picSelect.Move vsf.Left, vsf.Top + vsf.Height, vsf.Width, 400
            mclsVsf.AppendRows = True
            labStatus.Width = vsf.Width
            cmdStatus.Move picPane(0).Width - cmdStatus.Width - 30
        Case 1
            tvw.Move 15, 15, picPane(Index).Width - 15, picPane(Index).Height - 15
    End Select
End Sub
Public Sub FileBatPrint()
    Dim strObject As String
    Dim strParam As String
    Dim strTmp As String
    
    If ObjPtr(tvw.SelectedItem) <= 0 Then Exit Sub
    With tvw.SelectedItem
        If .Parent Is Nothing Then
            Select Case .Key
            Case "R5"
                strObject = "首页记录"
            Case "R1"
                strObject = "住院医嘱"
            Case "R9"
                strObject = "临床路径"
            Case Else
                blnReadUsed = False
                Exit Sub
            End Select
        Else
            strParam = .Tag
            Select Case .Parent.Key
            Case "R2"
                strObject = "住院病历"
            Case "R3"
                strObject = "护理病历"
            Case "R4"
                strObject = "护理记录"
            Case "R6"
                strObject = "医嘱报告"
            Case "R7"
                strObject = "疾病证明"
            Case "R8"
                strObject = "知情文件"
            End Select
        End If
    End With
    If blnReadUsed Then Exit Sub
    blnReadUsed = True
    With vsf
        strTmp = tvw.SelectedItem.Key & "," & Val(.TextMatrix(.Row, .ColIndex("病人id"))) & "," & Val(.TextMatrix(.Row, .ColIndex("主页id")))
        RaiseEvent AfterDocumentChanged(Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), strObject, strParam, .TextMatrix(.Row, .ColIndex("姓名")) & " -> " & .Text, Val(.TextMatrix(.Row, .ColIndex("ID"))), Val(.TextMatrix(.Row, .ColIndex("数据转出"))) = 1, True)
    End With
    blnReadUsed = False
    
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strObject As String
    Dim strParam As String
    Dim strTmp As String

    If Node.Parent Is Nothing Then
        Select Case Node.Key
        Case "R5"
            strObject = "首页记录"
        Case "R1"
            strObject = "住院医嘱"
        Case "R9"
            strObject = "临床路径"
        Case Else
            Select Case Node.Key
            Case "R2"
                strObject = "住院病历"
            Case "R3"
                strObject = "护理病历"
            Case "R4"
                strObject = "护理记录"
            Case "R6"
                strObject = "医嘱报告"
            Case "R7"
                strObject = "疾病证明"
            Case "R8"
                strObject = "知情文件"
            End Select
        End Select
    Else
        strParam = Node.Tag
        Select Case Node.Parent.Key
        Case "R2"
            strObject = "住院病历"
        Case "R3"
            strObject = "护理病历"
        Case "R4"
            strObject = "护理记录"
        Case "R6"
            strObject = "医嘱报告"
        Case "R7"
            strObject = "疾病证明"
        Case "R8"
            strObject = "知情文件"
        End Select
    End If
    
    With vsf
        tvw.Tag = Node.Key
        strTmp = Node.Key & "," & Val(.TextMatrix(.Row, .ColIndex("病人id"))) & "," & Val(.TextMatrix(.Row, .ColIndex("主页id")))

        mstrKey = strTmp
        RaiseEvent AfterDocumentChanged(Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), strObject, strParam, .TextMatrix(.Row, .ColIndex("姓名")) & " -> " & Node.Text, Val(.TextMatrix(.Row, .ColIndex("ID"))), Val(.TextMatrix(.Row, .ColIndex("数据转出"))) = 1, False)
    End With
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsf.AfterEdit(Row, Col)
    RaiseEvent AfterEdit(Row, Col)
End Sub

Private Sub vsf_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Call mclsVsf.AfterMoveColumn(Col, Position)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    
    If OldRow <> NewRow Then
        Call ExecuteCommand("读取病案结构", "NoRead")
    End If
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterSort(ByVal Col As Long, Order As Integer)
    Call mclsVsf.RestoreRow(mclsVsf.SaveKey)
    vsf.ShowCell vsf.Row, vsf.Col
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsf.ColIndex("选择") <> Col Then
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub vsf_RowColChange()
On Error Resume Next
    If mbytApplyMode = 3 Or mbytApplyMode = 4 Then
        labSelect.Visible = False
        labNum.Visible = False
        With vsf
            If .Rows = 1 Then
                labStatus.Caption = ""
            Else
                If .ColIndex("姓名") <> -1 Then
                    labStatus.Caption = ChkStrUniCode("姓名：" & .TextMatrix(.Row, .ColIndex("姓名")) & "                    ", 20) & "   住院号：" & .TextMatrix(.Row, .ColIndex("住院号"))
                End If
            End If
        End With
    End If
End Sub

Private Sub vsf_BeforeSort(ByVal Col As Long, Order As Integer)
    mclsVsf.SaveKey = Val(vsf.TextMatrix(vsf.Row, vsf.ColIndex("ID")))
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick()
    Call mclsVsf.DbClick
    Call ExecuteCommand("读取病案结构", "Read")
    RaiseEvent DbClick
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsf_DblClick
    End If
End Sub

Private Sub vsf_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Call SendLMouseButton(vsf.hWnd, x, y)
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub


