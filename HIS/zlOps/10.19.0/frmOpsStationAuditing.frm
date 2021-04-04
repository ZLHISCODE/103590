VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmOpsStationAuditing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "手术审核"
   ClientHeight    =   6045
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7275
   Icon            =   "frmOpsStationAuditing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   6045
      Left            =   45
      TabIndex        =   22
      Top             =   -45
      Width           =   5850
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   3
         Left            =   5445
         Picture         =   "frmOpsStationAuditing.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5550
         Width           =   300
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   2
         Left            =   2460
         Picture         =   "frmOpsStationAuditing.frx":685E
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   5550
         Width           =   300
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   3825
         TabIndex        =   17
         Top             =   5565
         Width           =   1620
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1140
         TabIndex        =   14
         Top             =   5550
         Width           =   1305
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1140
         TabIndex        =   8
         Top             =   2610
         Width           =   4575
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1905
         Width           =   4575
      End
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   0
         Left            =   5400
         Picture         =   "frmOpsStationAuditing.frx":D0B0
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2235
         Width           =   300
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1140
         TabIndex        =   5
         Top             =   2250
         Width           =   4260
      End
      Begin VB.TextBox txt 
         Height          =   735
         Index           =   1
         Left            =   1140
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   4725
         Width           =   4575
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfAdvice 
         Height          =   1665
         Left            =   1140
         TabIndex        =   1
         Top             =   180
         Width           =   4575
         _cx             =   8070
         _cy             =   2937
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
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDiagnose 
         Height          =   1680
         Left            =   1140
         TabIndex        =   10
         Top             =   2985
         Width           =   4575
         _cx             =   8070
         _cy             =   2963
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
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "申请科室(&K)"
         Height          =   180
         Index           =   7
         Left            =   2805
         TabIndex        =   16
         Top             =   5610
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "申请医生(&D)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   5595
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "麻醉类型(&T)"
         Height          =   180
         Index           =   6
         Left            =   105
         TabIndex        =   7
         Top             =   2655
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手术规模(&G)"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   2
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "麻醉方式(&M)"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   4
         Top             =   2295
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "附加说明(&N)"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   11
         Top             =   4725
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "术前诊断(&B)"
         Height          =   180
         Index           =   3
         Left            =   105
         TabIndex        =   9
         Top             =   2985
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "拟行手术(&B)"
         Height          =   180
         Index           =   4
         Left            =   105
         TabIndex        =   0
         Top             =   195
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6015
      TabIndex        =   21
      Top             =   1155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6015
      TabIndex        =   19
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6015
      TabIndex        =   20
      Top             =   525
      Width           =   1100
   End
End
Attribute VB_Name = "frmOpsStationAuditing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'（１）窗体级变量定义

Private WithEvents mclsVsfAdvice As clsVsf
Attribute mclsVsfAdvice.VB_VarHelpID = -1
Private WithEvents mclsVsfDiagnose As clsVsf
Attribute mclsVsfDiagnose.VB_VarHelpID = -1

Private mblnOK As Boolean
Private mlngKey As Long
Private mlng麻醉id As Long
Private mlng麻醉医嘱id As Long
Private mlng病人科室id As Long
Private mfrmMain As Form
Private mblnDataChanged As Boolean
Private mbytMode As Byte
Private mblnAllowModify As Boolean
Private mblnReading As Boolean

Private Type Items
    麻醉方式 As String
    申请医生 As String
    申请科室 As String
End Type

Private usrSaveItem As Items

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Form, ByVal lngKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnOK = False
    mlngKey = lngKey
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    
    Set mfrmMain = frmMain
    
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    Call ExecuteCommand("读取数据")
    
    DataChanged = False

    Me.Show 1, mfrmMain
    
    ShowEdit = mblnOK
    
End Function


Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        Set mclsVsfAdvice = New clsVsf
        With mclsVsfAdvice
            Call .Initialize(Me.Controls, vsfAdvice, True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)

            Call .AppendColumn("主要手术", 900, flexAlignCenterCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("手术名称", 1200, flexAlignLeftCenter, flexDTString, "", , True)

            Call .InitializeEdit(True, True, True)
            Call .InitializeEditColumn(.ColIndex("手术名称"), True, vbVsfEditCommand)
            Call .InitializeEditColumn(.ColIndex("主要手术"), True, vbVsfEditCheck)
            .IndicatorCol = 0
            Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            
            .AppendRows = True
        End With
        
        '术前诊断
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfDiagnose = New clsVsf
        With mclsVsfDiagnose
            Call .Initialize(Me.Controls, vsfDiagnose, True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
            Call .AppendColumn("疾病编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("疾病id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("诊断描述", 1500, flexAlignLeftCenter, flexDTString, "", , True)
                        
            Call .InitializeEdit(True, True, True)
            Call .InitializeEditColumn(.ColIndex("疾病编码"), True, vbVsfEditCommand)
            Call .InitializeEditColumn(.ColIndex("诊断编码"), True, vbVsfEditCommand)
            Call .InitializeEditColumn(.ColIndex("诊断描述"), True, vbVsfEditText)
                                    
            .IndicatorCol = 0
            Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture

            .AppendRows = True
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
                 
                 
        '麻醉类型
        '--------------------------------------------------------------------------------------------------------------
        cbo(1).Clear
        gstrSQL = "SELECT 名称,0 FROM 诊疗麻醉类型"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then Call AddComboData(cbo(1), rs, False)

        '诊疗手术规模
        '--------------------------------------------------------------------------------------------------------------
        cbo(0).Clear
        gstrSQL = "SELECT 名称,0 FROM 诊疗手术规模"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
        
        mblnReading = True
        
        mclsVsfAdvice.ClearGrid
        mclsVsfDiagnose.ClearGrid
        
        txt(0).Text = ""
        txt(1).Text = ""
        cmd(0).Tag = ""
        txt(0).Tag = ""
        txt(2).Tag = ""
        txt(3).Tag = ""
    
        mblnReading = False
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
        
        mclsVsfAdvice.ClearGrid
        
        cbo(0).ListIndex = -1
        
        gstrSQL = "Select a.诊疗项目id As ID,Decode(a.相关id,Null,1,0) As 主要手术,Decode(a.诊疗项目id,Null,a.医嘱内容,b.名称) As 手术名称,b.操作类型,a.开嘱医生,a.开嘱科室id,a.病人科室id,c.名称 As 开嘱科室 From 病人医嘱记录 a,诊疗项目目录 b,部门表 c Where c.ID(+)=a.开嘱科室id And a.诊疗项目id=b.ID(+) And [1] In (a.ID,a.相关id) And Nvl(a.诊疗类别,'F')<>'G' Order By Decode(a.相关id,Null,1,0) Desc"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            
            txt(2).Text = zlCommFun.NVL(rs("开嘱医生").Value)
            txt(3).Text = zlCommFun.NVL(rs("开嘱科室").Value)
            cmd(3).Tag = zlCommFun.NVL(rs("开嘱科室id").Value)
            txt(2).Tag = ""
            txt(3).Tag = ""
        
            mlng病人科室id = zlCommFun.NVL(rs("病人科室id").Value, 0)
            
            Call zlControl.CboLocate(cbo(0), zlCommFun.NVL(rs("操作类型").Value), True)
            
            Call mclsVsfAdvice.LoadGrid(rs)
        End If
    
        gstrSQL = "Select a.医嘱内容,a.诊疗项目id,b.操作类型 From 病人医嘱记录 a,诊疗项目目录 b Where a.相关id=[1] And a.诊疗类别='G' And a.诊疗项目id=b.ID "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            txt(0).Text = zlCommFun.NVL(rs("医嘱内容").Value)
            cbo(1).Text = zlCommFun.NVL(rs("操作类型").Value)
            cmd(0).Tag = zlCommFun.NVL(rs("诊疗项目id").Value, 0)
        End If
        
        usrSaveItem.麻醉方式 = txt(0).Text
        usrSaveItem.申请医生 = txt(2).Text
        usrSaveItem.申请科室 = txt(3).Text
    End Select

    ExecuteCommand = True

    Exit Function
    
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function


Private Function SaveData(ByRef lngOrderKey As Long) As Boolean
    '******************************************************************************************************************
    '功能：对新增、修改后的数据进行保存/更新处理
    '参数：返回参lngOrderKey，表示更新记录的关键字
    '返回：保存成功返回True，否则返回False
    '******************************************************************************************************************
    Dim lngNo As Long
    Dim lngLoop As Long
    Dim strSQL As String
    Dim lngCount As Long
    Dim intNo As Integer
    Dim lng记录ID As Long
    Dim lng病人id As Long
    Dim lng主页id As Long
    Dim lng挂号id As Long
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strKeys As String
    
    Call SQLRecord(rsSQL)
    
    '初始数据
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "SELECT b.ID FROM 病人医嘱记录 a,病人挂号记录 b WHERE a.挂号单=b.NO And a.病人来源=1 and a.ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngOrderKey)
    If rs.BOF = False Then lng挂号id = zlCommFun.NVL(rs("ID").Value, 0)
    
    gstrSQL = "SELECT a.病人id,a.主页id,b.ID From 病人医嘱记录 a,病人手术记录 b Where a.ID=[1] And a.ID=b.医嘱id(+)"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngOrderKey)
    If rs.BOF Then Exit Function
    
    lng病人id = zlCommFun.NVL(rs("病人id").Value, 0)
    lng主页id = zlCommFun.NVL(rs("主页id").Value, 0)
    lng记录ID = zlCommFun.NVL(rs("ID").Value, 0)
    
            
    '审核手术处理
    '------------------------------------------------------------------------------------------------------------------
    If lng记录ID = 0 Then lng记录ID = zlDatabase.GetNextId("病人手术记录")
    strSQL = "zl_病人手术记录_Aduit(" & lng记录ID & "," & lngOrderKey & ",'" & txt(0).Text & "','" & cbo(1).Text & "','" & cbo(0).Text & "'," & Val(cmd(0).Tag) & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    
    '手术情况记录
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Zl_病人手术情况_Delete(" & lng记录ID & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsfAdvice
        For lngLoop = 1 To .Rows - 1
            If Val(.RowData(lngLoop)) > 0 Then
                If Abs(Val(.TextMatrix(lngLoop, .ColIndex("主要手术")))) = 1 Then
                    strSQL = "Zl_病人手术情况_Insert(" & lng记录ID & ",1,1,'" & .TextMatrix(lngLoop, .ColIndex("手术名称")) & "',Null," & Val(.RowData(lngLoop)) & ")"
                    Call SQLRecordAdd(rsSQL, strSQL)
                    strSQL = "Zl_病人手术情况_Insert(" & lng记录ID & ",2,1,'" & .TextMatrix(lngLoop, .ColIndex("手术名称")) & "',Null," & Val(.RowData(lngLoop)) & ")"
                    Call SQLRecordAdd(rsSQL, strSQL)
                Else
                    strSQL = "Zl_病人手术情况_Insert(" & lng记录ID & ",1,0,'" & .TextMatrix(lngLoop, .ColIndex("手术名称")) & "',Null," & Val(.RowData(lngLoop)) & ")"
                    Call SQLRecordAdd(rsSQL, strSQL)
                    strSQL = "Zl_病人手术情况_Insert(" & lng记录ID & ",2,0,'" & .TextMatrix(lngLoop, .ColIndex("手术名称")) & "',Null," & Val(.RowData(lngLoop)) & ")"
                    Call SQLRecordAdd(rsSQL, strSQL)
                End If
            End If
        Next
    End With
    
    '术前诊断记录
    '------------------------------------------------------------------------------------------------------------------
    intNo = 0
    strSQL = "ZL_病人诊断记录_DELETE2(" & lngOrderKey & ",8)"
    Call SQLRecordAdd(rsSQL, strSQL)
    strSQL = "ZL_病人诊断记录_DELETE2(" & lngOrderKey & ",9)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsfDiagnose
        For lngLoop = 1 To .Rows - 1
            If Val(.RowData(lngLoop)) > 0 And (Val(.TextMatrix(lngLoop, .ColIndex("疾病id"))) > 0 Or Val(.TextMatrix(lngLoop, .ColIndex("诊断id"))) > 0) Then
                intNo = intNo + 1
                strSQL = "ZL_病人诊断记录_INSERT(" & lng病人id & "," & IIf(lng主页id > 0, lng主页id, ZVal(lng挂号id)) & ",1,NULL,8," & Val(.TextMatrix(lngLoop, .ColIndex("疾病id"))) & "," & Val(.TextMatrix(lngLoop, .ColIndex("诊断id"))) & ",NULL,'" & .TextMatrix(lngLoop, .ColIndex("诊断描述")) & "',NULL,NULL,NULL,SYSDATE," & lngOrderKey & "," & intNo & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                strSQL = "ZL_病人诊断记录_INSERT(" & lng病人id & "," & IIf(lng主页id > 0, lng主页id, ZVal(lng挂号id)) & ",1,NULL,9," & Val(.TextMatrix(lngLoop, .ColIndex("疾病id"))) & "," & Val(.TextMatrix(lngLoop, .ColIndex("诊断id"))) & ",NULL,'" & .TextMatrix(lngLoop, .ColIndex("诊断描述")) & "',NULL,NULL,NULL,SYSDATE," & lngOrderKey & "," & intNo & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    '更新申请说明
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "zl_病人手术记录_UpdateAdvice(" & lng记录ID & ")"
    Call SQLRecordAdd(rsSQL, strSQL)

    '提交到数据库中
    '------------------------------------------------------------------------------------------------------------------
    SaveData = SQLRecordExecute(rsSQL, Me.Caption)
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：对新增、修改的数据进行合法性校验
    '返回：校验合法返回True，否则返回False
    '******************************************************************************************************************

    Dim lngLoop As Long
    Dim blnDefault As Boolean
    
    
    '------------------------------------------------------------------------------------------------------------------
    If StrIsValid(txt(1).Text, txt(1).MaxLength) = False Then
        zlControl.TxtSelAll txt(1)
        txt(1).SetFocus
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    blnDefault = False
    With vsfAdvice
        For lngLoop = 1 To .Rows - 1
            If Val(.RowData(lngLoop)) > 0 Then
                If Abs(Val(.TextMatrix(lngLoop, .ColIndex("主要手术")))) = 1 Then
                    blnDefault = True
                End If
            End If
            
            
            If Val(.RowData(lngLoop)) <= 0 And lngLoop <> .Rows - 1 Then
                ShowSimpleMsg "必须指定具体的手术项目！"
                Call LocationGrid(vsfAdvice, lngLoop, .ColIndex("手术名称"))
                Exit Function
            End If
        Next
        If blnDefault = False Then
            ShowSimpleMsg "必须指定主要手术，即指定主手术！"
            Call LocationGrid(vsfAdvice, lngLoop, .ColIndex("手术名称"))
            Exit Function
        End If
    
    End With
    

    
    '------------------------------------------------------------------------------------------------------------------
    With vsfDiagnose
        For lngLoop = 1 To .Rows - 1
    
            If StrIsValid(.TextMatrix(lngLoop, .ColIndex("诊断描述")), 100) = False Then
                Call LocationGrid(vsfDiagnose, lngLoop, .ColIndex("诊断描述"))
                Exit Function
            End If
            
            If Val(.TextMatrix(lngLoop, .ColIndex("疾病id"))) = 0 And Val(.TextMatrix(lngLoop, .ColIndex("诊断id"))) = 0 And lngLoop <> .Rows - 1 Then
                ShowSimpleMsg "不正确的诊断输入！"
                Call LocationGrid(vsfDiagnose, lngLoop, .ColIndex("疾病id"))
                Exit Function
            End If
        Next
    End With
    
    '校验处理
    '------------------------------------------------------------------------------------------------------------------
    If CheckHaveOrder(mlngKey) = False Then
        MsgBox "手术医嘱记录已经不存在或者已作废！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    If CheckAllowAudit(mlngKey) = False Then Exit Function
    
'    If txt(2).Text = "" Then
'        MsgBox "手术医嘱记录的申请医生不能为空！", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    If Val(cmd(3).Tag) = 0 Then
'        MsgBox "手术医嘱记录的申请科室不能为空！", vbInformation, gstrSysName
'        Exit Function
'    End If
    
    ValidData = True

End Function


Private Sub cbo_Click(Index As Integer)
    Select Case Index
    Case 2, 3
        DataChanged = True
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 0      '诊疗项目
        gstrSQL = GetPublicSQL(SQL.麻醉方式选择)
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
 
        If ShowPubSelect(Me, txt(0), 2, "编码,900,0,;名称,2400,0,;麻醉类型,900,0,", Me.Name & "\麻醉方式选择", "请从下表中选择一个麻醉方式", rsData, rs, 8790, 4500, , Val(cmd(0).Tag)) = 1 Then
            If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then

                txt(0).Text = AppendCode(zlCommFun.NVL(rs("名称")), zlCommFun.NVL(rs("编码")))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                txt(0).Tag = ""
                
                cbo(1).Text = zlCommFun.NVL(rs("麻醉类型"))
                
                usrSaveItem.麻醉方式 = txt(0).Text
                
                DataChanged = True

            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '开单医生
        
        gstrSQL = GetPublicSQL(SQL.人员信息选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "医生", 0)
        If ShowPubSelect(Me, txt(2), 2, "编号,900,0,;姓名,900,0,;简码,900,0,", Me.Name & "\申请医生选择", "请从下表中选择一个申请医生", rsData, rs, 3900, 4500, , Val(cmd(2).Tag)) = 1 Then
            If Val(cmd(2).Tag) <> zlCommFun.NVL(rs("ID")) Then

                txt(2).Text = zlCommFun.NVL(rs("姓名").Value)
                cmd(2).Tag = zlCommFun.NVL(rs("ID"))
                txt(2).Tag = ""
                usrSaveItem.申请医生 = txt(2).Text
                DataChanged = True

            End If
        End If

    '------------------------------------------------------------------------------------------------------------------
    Case 3          '开单科室
        
        gstrSQL = GetPublicSQL(SQL.科室信息选择)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "临床")
        If ShowPubSelect(Me, txt(3), 2, "编码,900,0,;名称,1500,0,;简码,900,0,", Me.Name & "\申请医生选择", "请从下表中选择一个申请医生", rsData, rs, 3900, 4500, , Val(cmd(3).Tag)) = 1 Then
            If Val(cmd(2).Tag) <> zlCommFun.NVL(rs("ID")) Then

                txt(3).Text = zlCommFun.NVL(rs("名称").Value)
                cmd(3).Tag = zlCommFun.NVL(rs("ID"))
                txt(3).Tag = ""
                usrSaveItem.申请科室 = txt(3).Text
                DataChanged = True

            End If
        End If
    End Select
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((ParamInfo.系统号) / 100))
End Sub

Private Sub cmdOK_Click()

    If ValidData = False Then Exit Sub
    If SaveData(mlngKey) = False Then Exit Sub
    
    mblnOK = True
    DataChanged = False

    '退出编辑窗体
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("更新的数据必须保存才能生效，真的不保存就退出吗？", vbQuestion + vbDefaultButton2 + vbYesNo, ParamInfo.系统名称) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Set mclsVsfAdvice = Nothing
    Set mclsVsfDiagnose = Nothing
End Sub

Private Sub mclsVsfAdvice_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsfAdvice.RowData(Row)) = 0)
End Sub

Private Sub mclsVsfDiagnose_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsfDiagnose.TextMatrix(Row, vsfDiagnose.ColIndex("诊断id"))) = 0 And Val(vsfDiagnose.TextMatrix(Row, vsfDiagnose.ColIndex("疾病id"))) = 0)
End Sub

Private Sub txt_Change(Index As Integer)
    If mblnReading Then Exit Sub
    
    DataChanged = True
    
    Select Case Index
    Case 0, 2, 3
        txt(Index).Tag = "Changed"
    End Select

End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            txt(Index).Text = ""
            cmd(0).Tag = ""
            txt(Index).Tag = ""
            usrSaveItem.麻醉方式 = ""
        End If
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytMode As Byte
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""
                
                strText = UCase(txt(Index).Text)
                bytMode = GetApplyMode(strText)

                strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, "", "%") & strText & "%"
                
                gstrSQL = GetPublicSQL(SQL.麻醉方式过滤, bytMode)
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                If ShowPubSelect(Me, txt(Index), 2, "编码,990,0,1;名称,1500,0,0;麻醉类型,900,0,0", Me.Name & "\麻醉方式过滤", "请从下面选择一个麻醉方式", rsData, rs, , , , Val(cmd(0).Tag)) = 1 Then
                    If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
            
                        txt(Index).Text = AppendCode(zlCommFun.NVL(rs("名称")), zlCommFun.NVL(rs("编码")))
                        cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                        
                        DataChanged = True
                        
                        usrSaveItem.麻醉方式 = txt(Index).Text
                        
                    End If
                Else
                    txt(Index).Text = usrSaveItem.麻醉方式
                    txt(Index).Tag = ""
                    Exit Sub
                End If
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 2
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""
                
                strText = UCase(txt(Index).Text)
                bytMode = GetApplyMode(strText)

                strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, "", "%") & strText & "%"

                gstrSQL = GetPublicSQL(SQL.人员信息过滤, bytMode)
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "医生", 0, strText, strTmp)
                If ShowPubSelect(Me, txt(Index), 2, "编号,900,0,;姓名,900,0,;简码,900,0,", Me.Name & "\申请医生过滤", "请从下面选择一个申请医生", rsData, rs, 3900, 4500, , Val(cmd(2).Tag)) = 1 Then
                    If Val(cmd(2).Tag) <> zlCommFun.NVL(rs("ID")) Then
            
                        txt(Index).Text = zlCommFun.NVL(rs("姓名").Value)
                        cmd(2).Tag = zlCommFun.NVL(rs("ID"))
                        txt(Index).Tag = ""
                        
                        DataChanged = True
                        
                        usrSaveItem.申请医生 = txt(Index).Text
                        
                    End If
                Else
                    txt(Index).Text = usrSaveItem.申请医生
                    txt(Index).Tag = ""
                    Exit Sub
                End If
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 3
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""
                
                strText = UCase(txt(Index).Text)
                bytMode = GetApplyMode(strText)

                strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, "", "%") & strText & "%"

                gstrSQL = GetPublicSQL(SQL.科室信息过滤, bytMode)
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "临床", strText, strTmp)
                If ShowPubSelect(Me, txt(Index), 2, "编码,900,0,;名称,1500,0,;简码,900,0,", Me.Name & "\申请医生过滤", "请从下面选择一个申请医生", rsData, rs, 3900, 4500, , Val(cmd(3).Tag)) = 1 Then
                    If Val(cmd(3).Tag) <> zlCommFun.NVL(rs("ID")) Then
            
                        txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                        cmd(3).Tag = zlCommFun.NVL(rs("ID"))
                        txt(Index).Tag = ""
                        
                        DataChanged = True
                        
                        usrSaveItem.申请科室 = txt(Index).Text
                        
                    End If
                Else
                    txt(Index).Text = usrSaveItem.申请科室
                    txt(Index).Tag = ""
                    Exit Sub
                End If
            Else
                zlCommFun.PressKey vbKeyTab
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 0, 1
        zlCommFun.OpenIme False
    End Select
    
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub

    Select Case Index
    Case 0
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.麻醉方式
            txt(Index).Tag = ""
        End If
    Case 2
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.申请医生
            txt(Index).Tag = ""
        End If
    Case 3
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.申请科室
            txt(Index).Tag = ""
        End If
    End Select
    
End Sub


Private Sub vsfAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsfAdvice.AfterEdit(Row, Col)
    
    With vsfAdvice
        Select Case Col
        Case .ColIndex("主要手术")
            If Abs(Val(.Cell(flexcpText, Row, Col, Row, Col))) = 1 Then
                .Cell(flexcpText, 1, Col, .Rows - 1, Col) = 0
                .Cell(flexcpText, Row, Col, Row, Col) = 1
            End If
        End Select
    End With
    
    DataChanged = True
End Sub

Private Sub vsfAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsfAdvice.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfAdvice_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsfAdvice.AppendRows = True
End Sub

Private Sub vsfAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsfAdvice.AppendRows = True
End Sub

Private Sub vsfAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfAdvice.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsfAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    With vsfAdvice
        If Col = .ColIndex("手术名称") Then

            gstrSQL = GetPublicSQL(SQL.手术项目选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)

            If ShowPubSelect(Me, vsfAdvice, 3, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目选择", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then
                If mclsVsfAdvice.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
    
                .EditText = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(Row, mclsVsfAdvice.ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                
                DataChanged = True
            End If
        End If
    End With
End Sub

Private Sub vsfAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsfAdvice.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfAdvice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    
    With vsfAdvice
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("手术名称") Then
                
                If InStr(.EditText, "'") > 0 Then
                    KeyCode = 0
                    .EditText = ""
                    Exit Sub
                End If

                strText = UCase(.EditText)
                bytMode = GetApplyMode(strText)

                gstrSQL = GetPublicSQL(SQL.手术项目过滤, bytMode)

                strText = strText & "%"
                If ParamInfo.项目输入匹配方式 = 1 Then
                    strTmp = strText
                Else
                    strTmp = "%" & strText
                End If
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)

                If ShowPubSelect(Me, vsfAdvice, 2, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目过滤", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsfAdvice.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                        Exit Sub
                    End If

                    .EditText = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
                    
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)

                    DataChanged = True

                Else
                    KeyCode = 0

                    .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                    .EditText = .Cell(flexcpData, Row, Col)
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)

                End If
            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsfAdvice_KeyPress(KeyAscii As Integer)
    Call mclsVsfAdvice.KeyPress(KeyAscii)
End Sub

Private Sub vsfAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsfAdvice.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfAdvice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsfAdvice.AutoAddRow(vsfAdvice.MouseRow, vsfAdvice.MouseCol)
    End Select
End Sub

Private Sub vsfAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsfAdvice.EditSelAll
End Sub

Private Sub vsfAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfAdvice.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsfAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfAdvice.ValidateEdit(Col, Cancel)
End Sub

Private Sub vsfDiagnose_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call mclsVsfDiagnose.AfterEdit(Row, Col)
    DataChanged = True
End Sub

Private Sub vsfDiagnose_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsfDiagnose.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfDiagnose_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsfDiagnose.AppendRows = True
End Sub

Private Sub vsfDiagnose_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsfDiagnose.AppendRows = True
End Sub

Private Sub vsfDiagnose_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfDiagnose.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsfDiagnose_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    
    With vsfDiagnose
        
        If Col = .ColIndex("疾病编码") Or Col = .ColIndex("诊断编码") Then
            Select Case Col
            '----------------------------------------------------------------------------------------------------------
            Case .ColIndex("疾病编码")
            
                gstrSQL = GetPublicSQL(SQL.疾病编码选择)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "D")
    
                bytRet = ShowPubSelect(Me, vsfDiagnose, 3, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码选择", "请从下表中选择一个疾病编码项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                If bytRet = 1 Then
                    If mclsVsfDiagnose.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                        Exit Sub
                    End If
        
                    .EditText = zlCommFun.NVL(rs("编码").Value)
                    .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("编码").Value)
                    .TextMatrix(Row, .ColIndex("诊断描述")) = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(Row)), 0)
                    
                    DataChanged = True
                End If
            '----------------------------------------------------------------------------------------------------------
            Case .ColIndex("诊断编码")
            
                gstrSQL = GetPublicSQL(SQL.疾病诊断选择)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
    
                bytRet = ShowPubSelect(Me, vsfDiagnose, 3, "编码,1200,0,;名称,2700,0,", Me.Name & "\疾病诊断选择", "请从下表中选择一个疾病诊断项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                If bytRet = 1 Then
                    If mclsVsfDiagnose.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                        Exit Sub
                    End If
        
                    .EditText = zlCommFun.NVL(rs("编码").Value)
                    .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("编码").Value)
                    .TextMatrix(Row, .ColIndex("诊断描述")) = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, 0, Val(.RowData(Row)))
            
                    DataChanged = True
                End If
            End Select
            
            '----------------------------------------------------------------------------------------------------------
            If bytRet = 1 Then
                If rsData.BOF = False Then
                    .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("疾病编码").Value)
                    .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("诊断编码").Value)
                    .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("疾病id").Value, 0)
                    .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("诊断id").Value, 0)
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfDiagnose_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsfDiagnose.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfDiagnose_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    Dim bytRet As Byte
    
    With vsfDiagnose
        If KeyCode = vbKeyReturn Then
            If Col = .ColIndex("疾病编码") Or Col = .ColIndex("诊断编码") Then
            
                If InStr(.EditText, "'") > 0 Then
                    KeyCode = 0
                    .EditText = ""
                    Exit Sub
                End If

                strText = UCase(.EditText)
                bytMode = GetApplyMode(strText)
                    
                strText = strText & "%"
                strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, strText, "%" & strText)
                    
                Select Case Col
                '--------------------------------------------------------------------------------------------------
                Case .ColIndex("疾病编码")

                    gstrSQL = GetPublicSQL(SQL.疾病编码过滤, bytMode)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp, "D")
    
                    bytRet = ShowPubSelect(Me, vsfDiagnose, 2, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码过滤", "请从下表中选择一个疾病编码项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    If bytRet = 1 Then
                        If mclsVsfDiagnose.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
    
                        .EditText = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断描述")) = zlCommFun.NVL(rs("名称").Value)
                        .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(Row)), 0)
                    
                        DataChanged = True
                    End If
                '--------------------------------------------------------------------------------------------------
                Case .ColIndex("诊断编码")
                    gstrSQL = GetPublicSQL(SQL.疾病诊断过滤, bytMode)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
    
                    bytRet = ShowPubSelect(Me, vsfDiagnose, 2, "编码,1200,0,;名称,2700,0,", Me.Name & "\疾病诊断过滤", "请从下表中选择一个疾病诊断项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    If bytRet = 1 Then
                        If mclsVsfDiagnose.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
    
                        .EditText = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断描述")) = zlCommFun.NVL(rs("名称").Value)
                        .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                        gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, 0, Val(.RowData(Row)))
                    
                        DataChanged = True
                    End If
                End Select
                
                If bytRet = 1 Then
                
                    '--------------------------------------------------------------------------------------------------
                    If rsData.BOF = False Then
                        
                        .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("疾病编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("诊断编码").Value)
                        
                        .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("疾病id").Value, 0)
                        .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("诊断id").Value, 0)
                    End If
            
                Else
                    KeyCode = 0
                    .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                    .EditText = .Cell(flexcpData, Row, Col)
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                End If
                
            End If
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsfDiagnose_KeyPress(KeyAscii As Integer)
    Call mclsVsfDiagnose.KeyPress(KeyAscii)
End Sub

Private Sub vsfDiagnose_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call mclsVsfDiagnose.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfDiagnose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsfDiagnose.AutoAddRow(vsfDiagnose.MouseRow, vsfDiagnose.MouseCol)
    End Select
End Sub

Private Sub vsfDiagnose_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call mclsVsfDiagnose.EditSelAll
End Sub

Private Sub vsfDiagnose_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfDiagnose.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsfDiagnose_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsfDiagnose.ValidateEdit(Col, Cancel)
End Sub


