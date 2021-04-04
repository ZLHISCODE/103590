VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDistPara 
   Caption         =   "分诊科室设置"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   ControlBox      =   0   'False
   Icon            =   "frmDistPara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBarcodeSet 
      Caption         =   "条码打印设置"
      Height          =   375
      Left            =   5220
      TabIndex        =   9
      Top             =   4680
      Width           =   1620
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "排队单打印设置"
      Height          =   375
      Left            =   5220
      TabIndex        =   8
      Top             =   4215
      Width           =   1620
   End
   Begin VB.Frame fra呼叫对象 
      Caption         =   "分诊台分诊呼叫或医生主动呼叫"
      Height          =   930
      Left            =   120
      TabIndex        =   5
      Top             =   4125
      Width           =   5010
      Begin VB.OptionButton opt呼叫对象 
         Caption         =   "分诊台分诊呼叫"
         Height          =   240
         Index           =   0
         Left            =   375
         TabIndex        =   7
         Top             =   420
         Width           =   1620
      End
      Begin VB.OptionButton opt呼叫对象 
         Caption         =   "医生主动呼叫"
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   420
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7320
      TabIndex        =   2
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   1
      Top             =   270
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   5790
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistPara.frx":058A
            Key             =   "bm"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   3360
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "Ctrl+A全选,Ctrl+C全消"
      Top             =   585
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   5927
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imglst"
      SmallIcons      =   "imglst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编码"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.Frame fraSplit 
      Caption         =   "Frame1"
      Height          =   7995
      Left            =   7050
      TabIndex        =   4
      Top             =   -120
      Width           =   45
   End
   Begin VB.CheckBox chk分诊呼叫 
      Caption         =   "分诊后立即呼叫"
      Height          =   300
      Left            =   150
      TabIndex        =   10
      Top             =   5175
      Width           =   1740
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   285
      Picture         =   "frmDistPara.frx":0B24
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    一个分诊台可以同时管理多个门诊临床科室挂号病人，进行分诊相关处理；请选择由本分诊台进行分诊的临床科室(Ctrl+A全选,Ctrl+C全消)"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   930
      TabIndex        =   3
      Top             =   150
      Width           =   5805
   End
End
Attribute VB_Name = "frmDistPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Public mlngModul As Long
Private mblnNotClick As Boolean
 

Private Sub cmdBarcodeSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & Int(glngSys \ 100) & "_BILL_1113_1", Me)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim ObjItem As ListItem, strTmp As String
    
    For Each ObjItem In Me.lvwMain.ListItems
        If ObjItem.Checked Then
            strTmp = strTmp & "," & Mid(ObjItem.Key, 2)
        End If
    Next
    If strTmp = "" Then
        If MsgBox("你没有设置对任何科室分诊，该分诊台将不能进行分诊操作。" & vbCrLf & "真的暂时不设置吗？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
        strTmp = "0"
    Else
        strTmp = Mid(strTmp, 2)
        If UBound(Split(strTmp, ",")) + 1 = lvwMain.ListItems.Count Then strTmp = ""
    End If
    zlDatabase.SetPara "分诊科室", strTmp, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0 '空表示全部科室
    zlDatabase.SetPara "排队呼叫站点", IIf(opt呼叫对象(0).Value, 0, 1), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0  '空表示全部科室
    zlDatabase.SetPara "分诊后立即呼叫", IIf(chk分诊呼叫.Enabled = False, 0, chk分诊呼叫.Value), glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0 '空表示全部科室

    Unload Me
End Sub

 
Private Sub cmdPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & Int(glngSys \ 100) & "_BILL_1113", Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Dim i As Integer
        If UCase(Chr(KeyCode)) = "A" Then
            For i = 1 To lvwMain.ListItems.Count
                lvwMain.ListItems(i).Checked = True
            Next
        ElseIf UCase(Chr(KeyCode)) = "C" Then
            For i = 1 To lvwMain.ListItems.Count
                lvwMain.ListItems(i).Checked = False
            Next
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim ObjItem As ListItem
    Dim blnEnabled As Boolean
    
    Call RestoreWinState(Me, App.ProductName)
    mblnNotClick = True
    '1.代表分诊台分诊呼叫或医生主动呼叫;2-先分诊呼叫,再医生呼叫就诊.0-不排队叫号
    blnEnabled = Val(zlDatabase.GetPara("排队叫号模式", glngSys, mlngModul)) = 1
    Select Case Val(zlDatabase.GetPara("排队呼叫站点", glngSys, mlngModul, , Array(opt呼叫对象(0), opt呼叫对象(1)), InStr(1, mstrPrivs, ";参数设置;") > 0))
    Case 0  '0-代表分诊台分诊呼叫;1-代表医生主动呼叫
        opt呼叫对象(0).Value = True
    Case Else
        opt呼叫对象(1).Value = True
    End Select
 
    opt呼叫对象(1).Enabled = blnEnabled
    opt呼叫对象(0).Enabled = blnEnabled
    fra呼叫对象.Enabled = blnEnabled
    
    
    chk分诊呼叫.Value = IIf(Val(zlDatabase.GetPara("分诊后立即呼叫", glngSys, mlngModul, , Array(chk分诊呼叫), InStr(1, mstrPrivs, ";参数设置;") > 0)) = 1, 1, 0)
    chk分诊呼叫.Tag = IIf(chk分诊呼叫.Enabled, 1, 0)
    
    mblnNotClick = False
    
    '先得到以前设置的分诊科室ID,空表示所有诊室
    strTmp = zlDatabase.GetPara("分诊科室", glngSys, mlngModul, , Array(lvwMain), InStr(1, mstrPrivs, ";参数设置;") > 0)
    Me.lvwMain.ListItems.Clear
    On Error GoTo errH
    
    '143274:李南春,2019/7/26，如果操作员不具有“所有科室”权限，只显示操作员所属科室
    Set rsTmp = GetDepartments("'临床'", "1,3", InStr(mstrPrivs, "所有科室") = 0)
    
    With rsTmp
        Do While Not .EOF
            Set ObjItem = Me.lvwMain.ListItems.Add(, "K" & !id, !编码, "bm", "bm")
            ObjItem.SubItems(1) = Nvl(!名称)
            If InStr("," & strTmp & ",", "," & !id & ",") > 0 Or strTmp = "" Then ObjItem.Checked = True
            .MoveNext
        Loop
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim i As Long
    Me.cmdCancel.Left = ScaleWidth - (Me.cmdCancel.Width + 90)
    Me.cmdOK.Left = Me.cmdCancel.Left 'Me.cmdCancel.Left - (Me.cmdOK.Width + 20)
    Me.fraSplit.Left = Me.cmdOK.Left - fraSplit.Width - 50
    Me.fraSplit.Height = ScaleHeight + 100
    
    chk分诊呼叫.Top = Me.Height - chk分诊呼叫.Height - 650
    fra呼叫对象.Top = chk分诊呼叫.Top - fra呼叫对象.Height - 50
    cmdPrintSet.Top = fra呼叫对象.Top + 90
    cmdBarcodeSet.Top = cmdPrintSet.Top + cmdPrintSet.Height + 50
    lvwMain.Width = fraSplit.Left - lvwMain.Left - 50
    lvwMain.Height = fra呼叫对象.Top - 50 - lvwMain.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwMain.Sorted = True
    If lvwMain.SortKey = ColumnHeader.index - 1 Then
        If lvwMain.SortOrder = lvwAscending Then
            lvwMain.SortOrder = lvwDescending
        Else
            lvwMain.SortOrder = lvwAscending
        End If
    Else
        lvwMain.SortKey = ColumnHeader.index - 1
    End If
End Sub
