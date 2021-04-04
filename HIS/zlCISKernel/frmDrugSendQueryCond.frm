VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrugSendQueryCond 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   11565
   Icon            =   "frmDrugSendQueryCond.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   6435
      ScaleHeight     =   450
      ScaleWidth      =   2565
      TabIndex        =   36
      Top             =   4170
      Width           =   2565
      Begin VB.ComboBox cboReqDruDep 
         Height          =   300
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   60
         Width           =   2460
      End
   End
   Begin VB.PictureBox picCond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4000
      Left            =   3435
      ScaleHeight     =   4005
      ScaleWidth      =   2460
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   2460
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   630
         TabIndex        =   8
         Top             =   1770
         Width           =   1800
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H80000005&
         Caption         =   "以药房发药为准"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   4
         Top             =   560
         Width           =   1560
      End
      Begin VB.OptionButton optDate 
         BackColor       =   &H80000005&
         Caption         =   "以医嘱发送为准"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   3
         Top             =   350
         Value           =   -1  'True
         Width           =   1560
      End
      Begin VB.ComboBox cbo药房 
         Height          =   300
         Left            =   465
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   15
         Width           =   1845
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查询(&O)"
         Height          =   350
         Left            =   1200
         TabIndex        =   16
         Top             =   3600
         Width           =   1100
      End
      Begin VB.CheckBox chk退药 
         Alignment       =   1  'Right Justify
         Caption         =   "退药申请时间"
         Height          =   180
         Left            =   45
         TabIndex        =   13
         Top             =   2620
         Width           =   1380
      End
      Begin VB.CheckBox chk期效 
         Caption         =   "长嘱"
         Height          =   195
         Index           =   0
         Left            =   465
         TabIndex        =   9
         Top             =   2115
         Width           =   675
      End
      Begin VB.CheckBox chk期效 
         Caption         =   "临嘱"
         Height          =   195
         Index           =   1
         Left            =   465
         TabIndex        =   10
         Top             =   2340
         Width           =   660
      End
      Begin VB.CheckBox chk发药 
         Caption         =   "未发药"
         Height          =   195
         Index           =   0
         Left            =   1350
         TabIndex        =   11
         Top             =   2115
         Width           =   885
      End
      Begin VB.CheckBox chk发药 
         Caption         =   "已发药"
         Height          =   195
         Index           =   1
         Left            =   1350
         TabIndex        =   12
         Top             =   2340
         Width           =   885
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   630
         TabIndex        =   7
         Top             =   1440
         Width           =   1800
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   1
         Left            =   465
         TabIndex        =   6
         Top             =   1095
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   122617859
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   0
         Left            =   465
         TabIndex        =   5
         Top             =   780
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   122617859
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   3
         Left            =   465
         TabIndex        =   15
         Top             =   3240
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   122617859
         CurrentDate     =   37953
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Index           =   2
         Left            =   465
         TabIndex        =   14
         Top             =   2900
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   122617859
         CurrentDate     =   37953
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发药号"
         Height          =   180
         Index           =   1
         Left            =   45
         TabIndex        =   35
         Top             =   1830
         Width           =   540
      End
      Begin VB.Label lbl药房 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药房"
         Height          =   180
         Left            =   45
         TabIndex        =   33
         Top             =   75
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Index           =   5
         Left            =   225
         TabIndex        =   32
         Top             =   3315
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   31
         Top             =   2975
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   30
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "从"
         Height          =   180
         Index           =   0
         Left            =   225
         TabIndex        =   29
         Top             =   840
         Width           =   180
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间"
         Height          =   180
         Left            =   45
         TabIndex        =   28
         Top             =   350
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单  据"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   27
         Top             =   1500
         Width           =   540
      End
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3525
      Left            =   6480
      ScaleHeight     =   3525
      ScaleWidth      =   2475
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   2475
      Begin VB.CheckBox chkPreOut 
         Caption         =   "预出院(&P)"
         Height          =   195
         Left            =   0
         TabIndex        =   34
         Top             =   2970
         Width           =   1200
      End
      Begin VB.ComboBox cboUnit 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   0
         Width           =   2490
      End
      Begin VB.CommandButton cmdAllPati 
         Caption         =   "全选"
         Height          =   330
         Left            =   1620
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   3210
         Width           =   870
      End
      Begin VB.CheckBox chkOut 
         Caption         =   "最近出院(&A)"
         Height          =   195
         Left            =   1140
         TabIndex        =   20
         Top             =   2970
         Width           =   1320
      End
      Begin MSComctlLib.ListView lvwPati 
         Height          =   2610
         Left            =   0
         TabIndex        =   19
         Top             =   330
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   4604
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "姓名"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "住院号"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "床号"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "费别"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "科室"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "入院日期"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "出院日期"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "病人类型"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdNoPati 
         Caption         =   "全清"
         Height          =   330
         Left            =   765
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   3210
         Width           =   870
      End
   End
   Begin VB.PictureBox picWay 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   8925
      ScaleHeight     =   2790
      ScaleWidth      =   2430
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   390
      Width           =   2430
      Begin VB.CommandButton cmdAllWay 
         Caption         =   "全选"
         Height          =   330
         Left            =   1575
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + A"
         Top             =   2475
         Width           =   870
      End
      Begin MSComctlLib.ListView lvwWay 
         Height          =   2445
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   4313
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "给药途径"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton cmdNoWay 
         Caption         =   "全清"
         Height          =   330
         Left            =   720
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Ctrl + R"
         Top             =   2475
         Width           =   870
      End
   End
   Begin XtremeSuiteControls.TaskPanel tkpMain 
      Height          =   6600
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3285
      _Version        =   589884
      _ExtentX        =   5794
      _ExtentY        =   11642
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmDrugSendQueryCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event DoQuery(ByVal 药房ID As Long, ByVal Mode As Byte, ByVal DateBegin As Date, ByVal DateEnd As Date, ByVal 退药DateB As Date, ByVal 退药DateE As Date, _
    ByVal NO As String, ByVal 发药号 As String, ByVal 期效 As Integer, ByVal 状态 As String, ByVal 病区ID As Long, ByVal 病人IDs As String, ByVal 给药途径 As String, ByVal 领药部门ID As Long)

Private mMainPrivs As String 'IN
Private mlng病区ID As Long 'IN
Private mlng病人ID As Long 'IN
Private mblnOnePati As Boolean 'IN，单病人模式

Private Type QUERY_COND
    DateBegin As Date
    DateEnd As Date
    退药DateB As Date
    退药DateE As Date
    给药途径 As String
    NO As String
    发药号 As String
    药房ID As Long
    病人IDs As String
    病区ID As Long
    领药部门ID As Long
    期效 As Integer '2-全部
    状态 As String
End Type
Private mvQuery As QUERY_COND

Private Enum tkpItemIndex
    Item_查询内容 = 1
    Item_病区与病人 = 2
    Item_给药途径 = 3
    Item_领药部门 = 4
End Enum


Public Sub InitParameter(ByVal strMainPrivs As String, ByVal lng病区ID As Long, ByVal lng病人ID As Long, ByVal blnOnePati As Boolean)
    mMainPrivs = strMainPrivs
    mlng病区ID = lng病区ID
    mlng病人ID = lng病人ID
    mblnOnePati = blnOnePati
End Sub

Private Sub cboUnit_Click()
'功能：读取指定范围内的病人列表
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long
    Dim lngUnitID, lngColor As Long
        
    On Error GoTo errH
    lvwPati.ListItems.Clear
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    
    
    str病人IDs = zlDatabase.GetPara("发送病人", glngSys, p住院医嘱发送)
    If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
        lng病区ID = Val(Split(str病人IDs, ":")(0))
        str病人IDs = Split(str病人IDs, ":")(1)
    End If
        
    Set rsTmp = GetPatiRsByUnit(lngUnitID, mlng病人ID, False, False, chkOut.value, chkPreOut.value)
  
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!主页ID, rsTmp!姓名)
        objItem.SubItems(1) = Nvl(rsTmp!住院号)
        objItem.SubItems(2) = Nvl(rsTmp!床号)
        objItem.SubItems(3) = Nvl(rsTmp!费别)
        objItem.SubItems(4) = Nvl(rsTmp!科室)
        objItem.SubItems(5) = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
        objItem.SubItems(6) = Format(Nvl(rsTmp!出院日期), "yyyy-MM-dd HH:mm")
        objItem.SubItems(7) = Nvl(rsTmp!病人类型)
        
        '病人颜色
        lngColor = zlDatabase.GetPatiColor(Nvl(rsTmp!病人类型))
        objItem.ListSubItems(1).ForeColor = lngColor
        objItem.ListSubItems(7).ForeColor = lngColor
        
        '上次是否选择
        If lngUnitID = lng病区ID And str病人IDs <> "" Then
            If str病人IDs = "ALL" _
                Or Left(str病人IDs, 1) <> "-" And InStr("," & str病人IDs & ",", "," & rsTmp!病人ID & ",") > 0 _
                Or Left(str病人IDs, 1) = "-" And InStr("," & Mid(str病人IDs, 2) & ",", "," & rsTmp!病人ID & ",") = 0 Then
                objItem.Checked = True
                If k = 0 Then '为了看到有选择的
                    objItem.EnsureVisible
                    objItem.Selected = True
                    k = 1
                End If
            End If
        ElseIf rsTmp!病人ID = mlng病人ID Then
            objItem.Checked = True '缺省只选择当前病人
            objItem.EnsureVisible
            objItem.Selected = True
        End If
       
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkOut_Click()
    If Visible Then Call cboUnit_Click
End Sub

Private Sub chkPreOut_Click()
    If Visible Then Call cboUnit_Click
End Sub

Private Sub chk发药_Click(index As Integer)
    If chk发药(0).value = 0 And chk发药(1).value = 0 Then
        chk发药(index).value = 1: Exit Sub
    End If
    
    chk退药.Enabled = chk发药(1).value = 0
    If Not chk退药.Enabled Then chk退药.value = 0
End Sub

Private Sub chk期效_Click(index As Integer)
    If chk期效(0).value = 0 And chk期效(1).value = 0 Then
        chk期效(index).value = 1: Exit Sub
    End If
End Sub

Private Sub chk退药_Click()
    dtpDate(2).Enabled = chk退药.value = 1 And dtpDate(2).Tag = ""
    dtpDate(3).Enabled = chk退药.value = 1 And dtpDate(3).Tag = ""
    
    If dtpDate(2).Enabled Then dtpDate(2).SetFocus
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(lvwPati, True)
    lvwPati.SetFocus
End Sub

Private Sub cmdAllWay_Click()
    Call SelectLVW(lvwWay, True)
    lvwWay.SetFocus
End Sub

Private Sub cmdNoWay_Click()
    Call SelectLVW(lvwWay, False)
    lvwWay.SetFocus
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(lvwPati, False)
    lvwPati.SetFocus
End Sub

Private Sub cmdQuery_Click()
    Dim str病人IDs As String, strUn病人IDs As String
    Dim str给药IDs As String, i As Long
        
    If cbo药房.ListIndex = -1 Then
        MsgBox "请选择一个药房。", vbInformation, gstrSysName
        tkpMain.Groups(Item_查询内容).Expanded = True: cbo药房.SetFocus: Exit Sub
    End If
    If dtpDate(0).value >= dtpDate(1).value Then
        MsgBox "开始时间应小于结束时间。", vbInformation, gstrSysName
        tkpMain.Groups(Item_查询内容).Expanded = True: dtpDate(0).SetFocus: Exit Sub
    End If
    If chk退药.value = 1 Then
        If dtpDate(2).value >= dtpDate(3).value Then
            MsgBox "开始时间应小于结束时间。", vbInformation, gstrSysName
            tkpMain.Groups(Item_查询内容).Expanded = True: dtpDate(2).SetFocus: Exit Sub
        End If
    End If
    If cboUnit.ListIndex = -1 Then
        MsgBox "请选择一个病区。", vbInformation, gstrSysName
        tkpMain.Groups(Item_病区与病人).Expanded = True: cboUnit.SetFocus: Exit Sub
    End If
    
    '病人
    str病人IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            str病人IDs = str病人IDs & "," & Split(lvwPati.ListItems(i).Key, "_")(1)
        Else
            strUn病人IDs = strUn病人IDs & "," & Split(lvwPati.ListItems(i).Key, "_")(1)
        End If
    Next
    str病人IDs = Mid(str病人IDs, 2)
    strUn病人IDs = Mid(strUn病人IDs, 2)
    If str病人IDs = "" Or (UBound(Split(str病人IDs, ",")) = 0 And Val(str病人IDs) = mlng病人ID) Then
        str病人IDs = ""
    Else
        If strUn病人IDs = "" Then
            str病人IDs = cboUnit.ItemData(cboUnit.ListIndex) & ":ALL"
        ElseIf UBound(Split(str病人IDs, ",")) > UBound(Split(strUn病人IDs, ",")) Then
            str病人IDs = cboUnit.ItemData(cboUnit.ListIndex) & ":-" & strUn病人IDs
        Else
            str病人IDs = cboUnit.ItemData(cboUnit.ListIndex) & ":" & str病人IDs
        End If
    End If
    
    '给药途径
    mvQuery.给药途径 = "": str给药IDs = ""
    For i = 1 To lvwWay.ListItems.Count
        If lvwWay.ListItems(i).Checked Then
            str给药IDs = str给药IDs & "," & Mid(lvwWay.ListItems(i).Key, 2)
            mvQuery.给药途径 = mvQuery.给药途径 & "," & lvwWay.ListItems(i).Text
        End If
    Next
    str给药IDs = Mid(str给药IDs, 2)
    mvQuery.给药途径 = Mid(mvQuery.给药途径, 2)
    If str给药IDs = "" Then
        MsgBox "请至少选择一种给药途径。", vbInformation, gstrSysName
        tkpMain.Groups(Item_给药途径).Expanded = True: lvwWay.SetFocus: Exit Sub
    End If
    If UBound(Split(str给药IDs, ",")) + 1 = lvwWay.ListItems.Count Then
        str给药IDs = "": mvQuery.给药途径 = ""
    End If
        
    '保存参数到注册表中
    '---------------------------------------------------------------
    '药房
    Call zlDatabase.SetPara("药疗查询药房", cbo药房.ItemData(cbo药房.ListIndex), glngSys, p住院医嘱发送)
    
    '领药部门
    Call zlDatabase.SetPara("药疗查询领药部门", cboReqDruDep.ItemData(cboReqDruDep.ListIndex), glngSys, p住院医嘱发送)
    
    '时间
    Call zlDatabase.SetPara("药疗查询间隔", DateDiff("d", dtpDate(0).value, dtpDate(1).value), glngSys, p住院医嘱发送)
    If chk退药.value = 1 Then
        Call zlDatabase.SetPara("退药查询间隔", DateDiff("d", dtpDate(2).value, dtpDate(3).value), glngSys, p住院医嘱发送)
    End If
    
    '期效
    If chk期效(0).value = 1 And chk期效(1).value = 1 Then
        i = 2
    ElseIf chk期效(0).value = 1 Then
        i = 0
    Else
        i = 1
    End If
    Call zlDatabase.SetPara("药疗查询期效", i, glngSys, p住院医嘱发送)
    
    '状态
    If chk发药(0).value = 1 And chk发药(1).value = 1 Then
        i = 2
    ElseIf chk发药(0).value = 1 Then
        i = 0
    Else
        i = 1
    End If
    Call zlDatabase.SetPara("药疗查询状态", i, glngSys, p住院医嘱发送)
        
    '病人
    Call zlDatabase.SetPara("发送病人", str病人IDs, glngSys, p住院医嘱发送)
    
    '包含出院病人
    Call zlDatabase.SetPara("药疗查询出院病人", chkOut.value, glngSys, p住院医嘱发送)
    '包含预出院病人
    Call zlDatabase.SetPara("药疗查询预出院病人", chkPreOut.value, glngSys, p住院医嘱发送)
    
    '给药途径
    Call zlDatabase.SetPara("药疗查询给药途径", str给药IDs, glngSys, p住院医嘱发送)
    
    '收集条件
    '---------------------------------------------------------------------
    '病区
    mvQuery.病区ID = cboUnit.ItemData(cboUnit.ListIndex)

    '病人
    mvQuery.病人IDs = ""
    For i = 1 To lvwPati.ListItems.Count
        If lvwPati.ListItems(i).Checked Then
            mvQuery.病人IDs = mvQuery.病人IDs & "," & Split(lvwPati.ListItems(i).Key, "_")(1)
        End If
    Next
    mvQuery.病人IDs = Mid(mvQuery.病人IDs, 2)

    '时间
    mvQuery.DateBegin = Format(dtpDate(0).value, "yyyy-MM-dd HH:mm:00")
    mvQuery.DateEnd = Format(dtpDate(1).value, "yyyy-MM-dd HH:mm:59")
    If chk退药.value = 1 Then
        mvQuery.退药DateB = Format(dtpDate(2).value, "yyyy-MM-dd HH:mm:00")
        mvQuery.退药DateE = Format(dtpDate(3).value, "yyyy-MM-dd HH:mm:59")
    Else
        mvQuery.退药DateB = Empty
        mvQuery.退药DateE = Empty
    End If
    
    'NO
    mvQuery.NO = txtNO(0).Text
    '发药号
    mvQuery.发药号 = Trim(txtNO(1).Text)
    
    '期效
    If chk期效(0).value = 1 And chk期效(1).value = 1 Then
        mvQuery.期效 = 2
    ElseIf chk期效(0).value = 1 Then
        mvQuery.期效 = 0
    ElseIf chk期效(1).value = 1 Then
        mvQuery.期效 = 1
    End If

    '状态
    mvQuery.状态 = chk发药(0).value & chk发药(1).value

    '药房
    mvQuery.药房ID = cbo药房.ItemData(cbo药房.ListIndex)
    
    '领药部门
    mvQuery.领药部门ID = cboReqDruDep.ItemData(cboReqDruDep.ListIndex)
    '激活事件
    '------------------------------------------------------------------------
    With mvQuery
        RaiseEvent DoQuery(.药房ID, IIF(optDate(0).value, 0, 1), .DateBegin, .DateEnd, .退药DateB, .退药DateE, .NO, .发药号, .期效, .状态, .病区ID, .病人IDs, .给药途径, .领药部门ID)
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwPati Then
            Call cmdAllPati_Click
        ElseIf Me.ActiveControl Is lvwWay Then
            Call cmdAllWay_Click
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If Me.ActiveControl Is lvwPati Then
            Call cmdNoPati_Click
        ElseIf Me.ActiveControl Is lvwWay Then
            Call cmdNoWay_Click
        End If
    ElseIf KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim objGroup As TaskPanelGroup
    Dim objItem As TaskPanelGroupItem
    Dim curDate As Date, i As Long
    
    Call RestoreListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
    Me.Width = tkpMain.Width: Me.Height = tkpMain.Height
    
    '分组控件------------------------------------------
    Call tkpMain.SetMargins(8, 8, 8, 8, 8)

    Set objGroup = tkpMain.Groups.Add(Item_查询内容, "查询内容")
    objGroup.Expandable = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picCond
    picCond.BackColor = objItem.BackColor
    optDate(0).BackColor = objItem.BackColor
    optDate(1).BackColor = objItem.BackColor
    chk期效(0).BackColor = objItem.BackColor
    chk期效(1).BackColor = objItem.BackColor
    chk发药(0).BackColor = objItem.BackColor
    chk发药(1).BackColor = objItem.BackColor
    chk退药.BackColor = objItem.BackColor
    
    If mblnOnePati Then
        picPati.Visible = False
    Else
        Set objGroup = tkpMain.Groups.Add(Item_病区与病人, "病区与病人")
        Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
        Set objItem.Control = picPati
        picPati.BackColor = objItem.BackColor
        chkOut.BackColor = objItem.BackColor
        chkPreOut.BackColor = objItem.BackColor
    End If
    
    Set objGroup = tkpMain.Groups.Add(Item_领药部门, "领药部门")
    objGroup.Expanded = True
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picDept
    picDept.BackColor = objItem.BackColor
    
    Set objGroup = tkpMain.Groups.Add(Item_给药途径, "给药途径")
    objGroup.Expanded = False
    Set objItem = objGroup.Items.Add(0, "", xtpTaskItemTypeControl)
    Set objItem.Control = picWay
    picWay.BackColor = objItem.BackColor
    
    '-------------------------------------------------
    '设置缺省查询时间
    i = Val(zlDatabase.GetPara("药疗查询间隔", glngSys, p住院医嘱发送, "0", Array(lblDate, dtpDate(0), dtpDate(1))))
    curDate = zlDatabase.Currentdate
    dtpDate(0).value = Format(DateAdd("d", -1 * i, curDate), "yyyy-MM-dd 00:00")
    dtpDate(1).value = Format(curDate, "yyyy-MM-dd HH:mm")
    dtpDate(0).MaxDate = dtpDate(1).value
    dtpDate(1).MaxDate = dtpDate(1).value
    
    i = Val(zlDatabase.GetPara("退药查询间隔", glngSys, p住院医嘱发送, "0", Array(dtpDate(2), dtpDate(3))))
    If Not dtpDate(2).Enabled Then
        dtpDate(2).Tag = "1": dtpDate(3).Tag = "1" '表示固定不可用
    Else
        dtpDate(2).Enabled = False: dtpDate(3).Enabled = False '先正常根据状态初始为不可用
    End If
    curDate = zlDatabase.Currentdate
    dtpDate(2).value = Format(DateAdd("d", -1 * i, curDate), "yyyy-MM-dd 00:00")
    dtpDate(3).value = Format(curDate, "yyyy-MM-dd HH:mm")
    dtpDate(2).MaxDate = dtpDate(3).value
    dtpDate(3).MaxDate = dtpDate(3).value
    
    '缺省查询期效
    i = Val(zlDatabase.GetPara("药疗查询期效", glngSys, p住院医嘱发送, "2", Array(chk期效(0), chk期效(1))))
    If i = 2 Then
        chk期效(0).value = 1
        chk期效(1).value = 1
    Else
        chk期效(i).value = 1
    End If
    
    '缺省查询状态
    i = Val(zlDatabase.GetPara("药疗查询状态", glngSys, p住院医嘱发送, "2", Array(chk发药(0), chk发药(1))))
    If i = 2 Then
        chk发药(0).value = 1
        chk发药(1).value = 1
    Else
        chk发药(i).value = 1
    End If
    
    '缺省是否包含出院病人
    chkOut.value = Val(zlDatabase.GetPara("药疗查询出院病人", glngSys, p住院医嘱发送, "0", Array(chkOut)))
    '缺省是否包含预出院病人
    chkPreOut.value = Val(zlDatabase.GetPara("药疗查询预出院病人", glngSys, p住院医嘱发送, "0", Array(chkPreOut)))
    '病区/病人
    'Call zlControl.LvwFlatColumnHeader(lvwPati)
    Call InitUnits
    
    '药房
    Call Load药房
    
    '领药部门
    Call LoadReqDruDep
    
    '给药途径
    Call Load给药途径
    
End Sub

Private Function Load给药途径() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objItem As ListItem, str给药IDs As String
    
    On Error GoTo errH
    
    str给药IDs = zlDatabase.GetPara("药疗查询给药途径", glngSys, p住院医嘱发送, "", Array(lvwWay))

    strSQL = "Select ID,编码,名称 From 诊疗项目目录" & _
        " Where 类别='E' And 操作类型 in ('2', '4') And 服务对象 IN(2,3) And (站点='" & gstrNodeNo & "' Or 站点 is Null) Order by 操作类型, 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwWay.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称)
        
        If str给药IDs <> "" Then
            If InStr("," & str给药IDs & ",", "," & rsTmp!ID & ",") > 0 Then
                objItem.Checked = True
            End If
        Else
            objItem.Checked = True
        End If
        rsTmp.MoveNext
    Next
    Load给药途径 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
    
    '包含门诊观察室
    If InStr(mMainPrivs, "全院病人") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSQL = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = strSQL & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
    End If
    
    cboUnit.Clear
    If InStr(mMainPrivs, "全院病人") > 0 Then
        cboUnit.AddItem "所有病区"
        If mlng病区ID = 0 Then cboUnit.ListIndex = cboUnit.NewIndex
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng病区ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Load药房() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng药房 As Long
    
    On Error GoTo errH
    
    lng药房 = Val(zlDatabase.GetPara("药疗查询药房", glngSys, p住院医嘱发送, "0", Array(lbl药房, cbo药房)))

    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " AND B.部门ID=A.ID And B.服务对象 IN(2,3) and B.工作性质 in('中药房','西药房','成药房')" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        cbo药房.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cbo药房.ItemData(cbo药房.NewIndex) = rsTmp!ID
        If rsTmp!ID = lng药房 Then
            cbo药房.ListIndex = cbo药房.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cbo药房.ListCount > 0 And cbo药房.ListIndex = -1 Then cbo药房.ListIndex = 0
    Load药房 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadReqDruDep() As Boolean
'功能：加载领药部门
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng部门 As Long
    
    On Error GoTo errH
    
    lng部门 = Val(zlDatabase.GetPara("药疗查询领药部门", glngSys, p住院医嘱发送, "0", Array(cboReqDruDep)))

    strSQL = "Select a.Id, a.编码, a.名称" & _
        " From 部门表 A, 部门性质说明 B" & _
        " Where a.Id = b.部门id And b.工作性质 = '领药部门' And (a.撤档时间 Is Null Or Trunc(a.撤档时间) = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        " Order By 编码"

    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        
    With cboReqDruDep
        .Clear
        .AddItem "所有部门"
        .ItemData(.NewIndex) = 0
        
        Do While Not rsTmp.EOF
            .AddItem rsTmp!编码 & "-" & rsTmp!名称
            .ItemData(.NewIndex) = rsTmp!ID
            If rsTmp!ID = lng部门 Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        
        If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = 0
    End With
    LoadReqDruDep = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SelectLVW(objLVW As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objLVW.ListItems.Count
        objLVW.ListItems(i).Checked = blnCheck
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tkpMain.Left = 0
    tkpMain.Top = 0
    tkpMain.Width = Me.ScaleWidth
    tkpMain.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveListViewState(Me.lvwPati, App.ProductName & Me.Name, "")
End Sub


Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.index)
End Sub

Private Sub optDate_Click(index As Integer)
    If index = 1 And optDate(index).value Then
        chk发药(0).Enabled = False
        chk发药(1).value = 1 '如果值改变，自动触发Click；两个必勾一个，因互先勾后消
        chk发药(0).value = 0 '如果值改变，自动触发Click
    Else
        chk发药(0).Enabled = True
    End If
End Sub

Private Sub picCond_Resize()
    On Error Resume Next
    
    cbo药房.Width = picCond.ScaleWidth - cbo药房.Left
    
    dtpDate(0).Width = picCond.ScaleWidth - dtpDate(0).Left
    dtpDate(1).Width = dtpDate(0).Width
    dtpDate(2).Width = dtpDate(0).Width
    dtpDate(3).Width = dtpDate(0).Width
    
    txtNO(0).Width = picCond.ScaleWidth - txtNO(0).Left
    txtNO(1).Width = picCond.ScaleWidth - txtNO(1).Left
    
    cmdQuery.Left = picCond.ScaleWidth - cmdQuery.Width - 30
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    
    cboUnit.Left = 0
    cboUnit.Width = picPati.ScaleWidth
    
    lvwPati.Left = 0
    lvwPati.Width = picPati.ScaleWidth
    
    chkOut.Left = lvwPati.Left + lvwPati.Width - chkOut.Width - 15
    chkPreOut.Left = chkOut.Left - chkPreOut.Width - 15
    cmdAllPati.Left = picPati.ScaleWidth - cmdAllPati.Width + 15
    cmdNoPati.Left = cmdAllPati.Left - cmdNoPati.Width + 15
End Sub

Private Sub picDept_Resize()
    On Error Resume Next
    
    cboReqDruDep.Left = 0
    cboReqDruDep.Width = picPati.ScaleWidth
    
End Sub

Private Sub picWay_Resize()
    On Error Resume Next

    lvwWay.Left = 0
    lvwWay.Width = picWay.ScaleWidth
    
    cmdAllWay.Left = picWay.ScaleWidth - cmdAllWay.Width + 15
    cmdNoWay.Left = cmdAllWay.Left - cmdNoWay.Width + 15
End Sub

Private Sub txtNO_GotFocus(index As Integer)
    Call zlControl.TxtSelAll(txtNO(index))
End Sub


Private Sub txtNO_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtNO_Validate(index, False)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Len(Trim(txtNO(index).Text)) > 18 And index = 1 Then
            txtNO(index).Text = Mid(Trim(txtNO(index).Text), 18)
            MsgBox "发药号长度不能大于18位。", vbInformation, gstrSysName
        End If
        If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtNO_Validate(index As Integer, Cancel As Boolean)
    If txtNO(index).Text <> "" Then
        txtNO(index).Text = IIF(index = 0, GetFullNO(txtNO(index), 14), Trim(txtNO(index).Text))
    End If
End Sub
