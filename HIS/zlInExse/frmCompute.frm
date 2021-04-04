VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "自动记帐计算"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmCompute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5655
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkForceCacl 
      Caption         =   "强制自动记帐计算(&F)"
      Height          =   195
      Left            =   255
      TabIndex        =   14
      Top             =   1905
      Width           =   2745
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   360
      TabIndex        =   11
      Top             =   6015
      Width           =   1050
   End
   Begin MSComctlLib.StatusBar stbLvw 
      Height          =   300
      Left            =   300
      TabIndex        =   9
      Top             =   5490
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8837
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   -135
      TabIndex        =   8
      Top             =   1035
      Width           =   7170
   End
   Begin VB.CheckBox chkPati 
      Caption         =   "仅计算指定病人(&P)"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   1215
      Width           =   1905
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   3870
      Top             =   3195
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
            Picture         =   "frmCompute.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   3885
      Top             =   2385
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompute.frx":0624
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   -120
      TabIndex        =   6
      Top             =   5895
      Width           =   7170
   End
   Begin VB.CommandButton cmdCpt 
      Caption         =   "计算(&C)"
      Height          =   350
      Left            =   2985
      TabIndex        =   4
      Top             =   6015
      Width           =   1050
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   4245
      TabIndex        =   5
      Top             =   6015
      Width           =   1050
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3285
      Left            =   270
      TabIndex        =   3
      Top             =   2160
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      SmallIcons      =   "img16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "姓名"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "性别"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "年龄"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "住院号"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "入院日期"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "床号"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.ComboBox cboWard 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Text            =   "cboWard"
      Top             =   1170
      Width           =   2325
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "yyyy-mm-dd hh:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   13
      ToolTipText     =   "有效期为max(病人入院时间，重算期间的开始时间)～当前时间,如果要修正上一期间费用,则重算期间等于上一期间,否则为本期间"
      Top             =   1560
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   117899267
      CurrentDate     =   38713
   End
   Begin VB.CheckBox chkReCalc 
      Caption         =   "从                        开始重新计算(&L)"
      Height          =   195
      Left            =   255
      TabIndex        =   12
      ToolTipText     =   "有效期为max(病人入院时间，重算期间的开始时间)～当前时间,如果要修正上一期间费用,则重算期间等于上一期间,否则为本期间"
      Top             =   1620
      Width           =   4200
   End
   Begin VB.Label lblNote 
      ForeColor       =   &H00000080&
      Height          =   210
      Index           =   1
      Left            =   840
      TabIndex        =   10
      Top             =   750
      Width           =   4350
   End
   Begin VB.Label lblNote 
      Caption         =   "    本程序根据用户对各个病区自动计算的设置情况，自动对指定范围内的在院病人，计算其费用，并完成记帐处理。"
      ForeColor       =   &H00000080&
      Height          =   525
      Index           =   0
      Left            =   840
      TabIndex        =   7
      Top             =   150
      Width           =   4350
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmCompute.frx":093E
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblWard 
      AutoSize        =   -1  'True
      Caption         =   "指定病区"
      Height          =   180
      Left            =   255
      TabIndex        =   0
      Top             =   1230
      Width           =   720
   End
End
Attribute VB_Name = "frmCompute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsPati As New ADODB.Recordset
Private mobjItem As ListItem
Private mstrPrivs As String
Private mlngModul As Long
Private mlngUnitID  As Long
Private mrsWard As New ADODB.Recordset
Private Sub cboWard_Click()
    mlngUnitID = cboWard.ItemData(cboWard.ListIndex)
    Me.lvwPati.ListItems.Clear
    Me.chkPati.Value = 0
    If Me.cboWard.ItemData(Me.cboWard.ListIndex) = 0 Then
        Me.chkPati.Enabled = False
        chkReCalc.Enabled = False
        Call chkReCalc_Click
    Else
        Me.chkPati.Enabled = True
        chkReCalc.Enabled = True
        Call chkReCalc_Click
    End If
    Me.stbLvw.Panels(1).Text = "计算" & Me.cboWard.Text & "的在院病人"

End Sub

Private Sub cboWard_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lng医生ID As Long
    Dim strRootCaption As String
    
    If KeyAscii <> 13 Then Exit Sub
    If cboWard.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsWard Is Nothing Then Exit Sub
    strRootCaption = ""
    If InStr(";" & mstrPrivs, ";所有病区") > 0 Then strRootCaption = "所有病区"
    
    If zlSelectDept(Me, mlngModul, cboWard, mrsWard, cboWard.Text, True, strRootCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cboWard_Validate(Cancel As Boolean)
    Dim lngID As Long
    If cboWard.ListIndex >= 0 Then Exit Sub
    lngID = mlngUnitID
   zlControl.CboLocate cboWard, lngID, True
   If cboWard.ListIndex < 0 And cboWard.ListCount <> 0 Then cboWard.ListIndex = 0
End Sub

Private Sub chkPati_Click()
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    Me.lvwPati.Enabled = (Me.chkPati.Value = 1)
    Me.lvwPati.ListItems.Clear
    Me.stbLvw.Panels(1).Text = "计算" & Me.cboWard.Text & "的在院病人"
    If Me.lvwPati.Enabled = False Then Exit Sub
    '问题:
    '       and P.主页ID <>0
    strSQL = "Select I.住院号,nvl(p.姓名,i.姓名) as 姓名,nvl(p.性别,i.性别) as 性别,nvl(p.年龄,i.年龄) as 年龄,P.病人id,P.主页id,P.入院日期,P.出院病床" & _
            " From 病人信息 I,病案主页 P ,在院病人 C " & _
            " Where I.病人id=P.病人id  And P.出院日期 is null And I.当前病区id=[1] And I.病人ID=C.病人ID  And I.当前病区ID=C.病区ID and P.主页ID <>0"
            
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboWard.ItemData(Me.cboWard.ListIndex)))
    With mrsPati
        If .EOF Or .BOF Then Me.chkPati.Enabled = False
        Do While Not .EOF
            Set mobjItem = Me.lvwPati.ListItems.Add(, "_" & !病人ID, !姓名, 1, 1)
            mobjItem.SubItems(1) = IIf(IsNull(!性别), "", !性别)
            mobjItem.SubItems(2) = IIf(IsNull(!年龄), "", !年龄)
            mobjItem.SubItems(3) = IIf(IsNull(!住院号), "", !住院号)
            mobjItem.SubItems(4) = Format(!入院日期, "yyyy-mm-dd; ; ")
            mobjItem.SubItems(5) = "" & !出院病床
            mobjItem.Tag = !主页ID
            .MoveNext
        Loop
    End With
    lvwPati.SortKey = lvwPati.ColumnHeaders.Count - 1   '缺省按床号排序
    Me.stbLvw.Panels(1).Text = "请选择需要计算的病人"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub chkReCalc_Click()
    dtpBegin.Enabled = (chkReCalc.Value = 1 And chkReCalc.Enabled)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Activate()
    If Me.cboWard.ListCount = 0 Then
        MsgBox "目前没有数据，无法计算", vbExclamation, gstrSysName
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim str病区IDs As String
    
    On Error GoTo errHandle
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    If Val(zlDatabase.GetPara(7, glngSys, , 0)) = 1 Then
        Me.lblNote(1).Caption = "     根据参数设置，本次计算将修正上期计算结果。"
    Else
        Me.lblNote(1).Caption = "     根据参数设置，本次计算仅修改本期计算结果。"
    End If
    strSQL = ""
     If InStr(1, mstrPrivs, ";所有病区;") = 0 Then
        '问题:43133
        str病区IDs = "," & GetUserUnits & ","
        strSQL = " And Instr(','||[1]||',',','||D.ID||',')>0"
     End If
    strSQL = "Select distinct D.id,D.名称,D.简码,D.编码 From 部门表 D,病人信息 L,在院病人 C" & _
            " Where D.id=L.当前病区id And L.当前病区ID=C.病区ID And  L.病人ID=C.病人ID " & strSQL
    Set mrsWard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str病区IDs)
    With mrsWard
        If .EOF Or .BOF Then
            Me.cboWard.Enabled = False
            Me.cmdCpt.Enabled = False
            Me.chkPati.Enabled = False
            Exit Sub
        End If
        If InStr(1, mstrPrivs, ";所有病区;") > 0 Then
            Me.cboWard.AddItem "所有病区"
            Me.cboWard.ItemData(Me.cboWard.NewIndex) = 0
            Me.cboWard.ListIndex = 0
        End If
        Do While Not .EOF
            Me.cboWard.AddItem !名称
            Me.cboWard.ItemData(Me.cboWard.NewIndex) = !ID
            .MoveNext
        Loop
        If cboWard.ListIndex < 0 And cboWard.ListCount <> 0 Then cboWard.ListIndex = 0
    End With
    Me.lvwPati.Tag = "0"
    
    dtpBegin.MinDate = CDate("1900-01-01 00:00:00")
    dtpBegin.MaxDate = CDate("3000-01-01 00:00:00")
    dtpBegin.Value = zlDatabase.Currentdate
    dtpBegin.Enabled = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwPati
        If .Tag = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .Tag = ColumnHeader.Index - 1
            .SortKey = .Tag
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub cmdCpt_Click()
    Dim strSQL As String
    
    On Error GoTo errH
    
    stbLvw.Panels(1).Text = "正在自动计算……"
    
    If chkPati.Value = 1 Then        '单个病人计算
        For Each mobjItem In lvwPati.ListItems
            If mobjItem.Selected Then
                stbLvw.Panels(1).Text = "正在自动计算" & mobjItem.Text & "……"
                strSQL = "zl1_AutoCptPati(" & CLng(Mid(mobjItem.Key, 2)) & "," & CLng(mobjItem.Tag) & _
                    "," & IIf(dtpBegin.Enabled, "To_Date('" & Format(dtpBegin.Value, dtpBegin.CustomFormat) & "','YYYY-MM-DD HH24:MI:SS')", "Null") & "," & _
                    IIf(chkForceCacl.Value = 1, 1, 0) & ")"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        Next
    Else
        If cboWard.ItemData(cboWard.ListIndex) <> 0 Then '计算一个病区
            strSQL = "zl1_AutoCptWard(" & cboWard.ItemData(cboWard.ListIndex) & _
            "," & IIf(dtpBegin.Enabled, "To_Date('" & Format(dtpBegin.Value, dtpBegin.CustomFormat) & "','YYYY-MM-DD HH24:MI:SS')", "Null") & "," & _
            IIf(chkForceCacl.Value = 1, 1, 0) & ")"
        Else '计算所有病区
            strSQL = "zl1_AutoCptAll(" & IIf(chkForceCacl.Value = 1, 1, 0) & ")"
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    stbLvw.Panels(1).Text = "计算完毕！"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
