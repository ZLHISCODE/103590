VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistoryDataMgr 
   BackColor       =   &H80000005&
   Caption         =   "数据转移空间管理"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmHistoryDataMgr.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFunction 
      Caption         =   "转移(&T)"
      Height          =   350
      Index           =   6
      Left            =   6960
      TabIndex        =   14
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "合并(&L)"
      Height          =   350
      Index           =   5
      Left            =   5850
      TabIndex        =   12
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CheckBox chkAll只读 
      BackColor       =   &H00FFFFFF&
      Caption         =   "全部只读(&A)"
      Height          =   285
      Left            =   4335
      TabIndex        =   11
      ToolTipText     =   "只对本机的历史数据空间有效"
      Top             =   1260
      Width           =   1485
   End
   Begin VB.CheckBox chk只读 
      BackColor       =   &H00FFFFFF&
      Caption         =   "当前项只读(&Z)"
      Height          =   240
      Left            =   2430
      TabIndex        =   10
      ToolTipText     =   "只对本机的历史数据空间有效"
      Top             =   1275
      Width           =   1695
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "复制(&Z)"
      Height          =   350
      Index           =   3
      Left            =   3600
      TabIndex        =   9
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "切换(&Q)"
      Height          =   350
      Index           =   4
      Left            =   4730
      TabIndex        =   8
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "再植(&R)"
      Height          =   350
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "拆卸(&M)"
      Height          =   350
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "创建(&C)"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1100
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   615
      Width           =   3570
   End
   Begin MSComctlLib.ImageList imgSys 
      Left            =   5280
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoryDataMgr.frx":04F9
            Key             =   "Other"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoryDataMgr.frx":158B
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoryDataMgr.frx":497D
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoryDataMgr.frx":7D6F
            Key             =   "LockAndRun"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwHistory 
      Height          =   2460
      Left            =   150
      TabIndex        =   3
      Top             =   1560
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   4339
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgSys"
      SmallIcons      =   "imgSys"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "编号"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "当前"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "编号"
         Text            =   "只读"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "所有者"
         Text            =   "所有者"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "版本号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "最后转储日期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "最后复制日期"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "DB连接"
         Text            =   "DB连接"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblMain 
      BackColor       =   &H8000000E&
      Height          =   15
      Left            =   120
      TabIndex        =   13
      Top             =   5250
      Width           =   6360
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "已创建的历史数据空间"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用系统"
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据转移空间管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   1920
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmHistoryDataMgr.frx":B161
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmHistoryDataMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim intCount As Integer
Dim mblnHavingSpace As Boolean '存在历史数据表空间
Dim mblnSel As Boolean
Dim mintColumn As Integer
Dim mblnChk As Boolean

Private Enum ENUFT
    F0创建 = 0
    F1拆卸 = 1
    F2再植 = 2
    F3复制 = 3
    F4切换 = 4
    F5合并 = 5
    F6转移 = 6
End Enum
'0-创建历名数据空间,1-拆卸历史数据空间,2-再植历史数据空间,3-复制非转储数据,4－切换在当前的历史数据空间,5-合并历史表空间,6-将异地历史表空间转移到当前库

Private Enum ENUCOL
    C0编号 = 0
    C1名称 = 1
    C2当前 = 2
    C3只读 = 3
    C4所有者 = 4
    C5版本号 = 5
    C6最后转储日期 = 6
    C7最后复制日期 = 7
    C8DBLink = 8
End Enum

Private Sub chkAll只读_Click()
    Dim lng系统 As Long
    If mblnChk = True Then Exit Sub
    If cmbSystem.ListIndex < 0 Then Exit Sub
    lng系统 = cmbSystem.ItemData(cmbSystem.ListIndex)
    
    If lvwHistory.ListItems.Count = 0 Then Exit Sub
    
    Call SetHistoreReadPro(chkAll只读.value = 1, 0, lng系统, True)
    Call cmbSystem_Click
End Sub

Private Sub chk只读_Click()
    Dim lng系统 As Long
    Dim str编号 As String
    Dim lng空间编号 As Long
    Dim bln只读 As Boolean
    Dim strImgKey As String
    If mblnSel = True Then Exit Sub
    If cmbSystem.ListIndex < 0 Then Exit Sub
    lng系统 = cmbSystem.ItemData(cmbSystem.ListIndex)
    
    err = 0: On Error Resume Next
    
    If lvwHistory.SelectedItem Is Nothing Then Exit Sub
    
    '不能对远程的历史空间进行设置只读
    If Mid(lvwHistory.SelectedItem.Tag, 3, 1) <> "1" Then Exit Sub
    
    lng空间编号 = Val(Mid(lvwHistory.SelectedItem.Key, 2))
    If lng空间编号 < 0 Then Exit Sub
    
    If SetHistoreReadPro(chk只读.value = 1, lng空间编号, lng系统, False) = False Then
        Exit Sub
    End If
    
    
    lvwHistory.SelectedItem.SubItems(C3只读) = IIf(chk只读.value = 1, "√", "")
    lvwHistory.SelectedItem.Tag = Mid(lvwHistory.SelectedItem.Tag, 1, 1) & IIf(chk只读.value = 1, "1", "0") & Mid(lvwHistory.SelectedItem.Tag, 3, 1)
            
    If Val(Mid(lvwHistory.SelectedItem.Tag, 1, 1)) = 1 Then
        If chk只读.value = 1 Then
            strImgKey = "LockAndRun"
        Else
            strImgKey = "Run"
        End If
    Else
        If chk只读.value = 1 Then
            strImgKey = "Lock"
        Else
            strImgKey = "Other"
        End If
    End If
    lvwHistory.SelectedItem.SmallIcon = strImgKey
    lvwHistory.SelectedItem.Icon = strImgKey
    mblnChk = True
    chkAll只读.value = 2
    mblnChk = False
End Sub
Private Function SetHistoreReadPro(ByVal bln只读 As Boolean, ByVal lng编号 As String, lngSys As Long, ByVal blnAll As Boolean) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能:设置相关的历史数据空间的只读属性
    '参数:bln只读-设置为只读
    '     lng编号-空间编号
    '     lngSys-系统号
    '     blnAll-所有项目
    '返回:成功返回true,否则False
    '--------------------------------------------------------------------------------------------------------------
        
    err = 0: On Error GoTo errHand:
    gstrSQL = "Update zltools.zlbakspaces set 只读=" & IIf(bln只读, 1, 0) & " where  DB连接 is null and  系统=" & lngSys & IIf(blnAll, "", " and 编号=" & lng编号)
    gcnOracle.Execute gstrSQL
    SetHistoreReadPro = True
    Exit Function
errHand:
    MsgBox "设置失败,详细错误信息如下:" & vbCrLf & "(" & err.Number & ")" & err.Description
End Function

Private Sub cmbSystem_Click()
    Dim lng系统 As Long
    Dim rsTemp As New ADODB.Recordset
    
    lng系统 = Val(cmbSystem.ItemData(cmbSystem.ListIndex))
    cmbSystem.Tag = GetOwnerName(lng系统, gcnOracle)
    
    gstrSQL = "Select  1 From zlBakTables where rownum<=1 and  系统=" & lng系统
    rsTemp.Open gstrSQL, gcnOracle
    If rsTemp.EOF Then
        mblnHavingSpace = False
    Else
        mblnHavingSpace = True
    End If
    Call LoadHistorySpace(lng系统)
    '设置控件属性
    Call SetCtlEnable
End Sub
Private Function LoadHistorySpace(ByVal lng系统 As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:加载历史数据空间
    '参数:lng系统-系统编号
    '返回:加载成功,返回true,否则返回false
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsbakspaces As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strImgKey As String
    Dim objItem As ListItem
    Dim lngMaxLen As Long
    
    LoadHistorySpace = False
    
    gstrSQL = "Select max(length(编号)) as MaxLen From zltools.zlbakspaces where 系统=" & lng系统
    OpenRecordset rsbakspaces, gstrSQL, "获取历史数据空间"
    lngMaxLen = Val(Nvl(rsbakspaces!MaxLen))
    
    gstrSQL = "Select 系统, 编号, 名称, 所有者, db连接, 当前, 只读 From zltools.zlbakspaces where 系统=" & lng系统 & " Order by 编号"
    OpenRecordset rsbakspaces, gstrSQL, "获取历史数据空间"
    
    mblnSel = True
    err = 0: On Error Resume Next
    
    strImgKey = ""
    
    With lvwHistory
        .ListItems.Clear
        Do While Not rsbakspaces.EOF
            Set objItem = .ListItems.Add(, "K" & Nvl(rsbakspaces!编号), Lpad(Nvl(rsbakspaces!编号), lngMaxLen), 0, 0)
            objItem.SubItems(C1名称) = Nvl(rsbakspaces!名称)
            objItem.SubItems(C2当前) = IIf(Val(Nvl(rsbakspaces!当前)) = 1, "√", "")
            objItem.SubItems(C3只读) = IIf(Val(Nvl(rsbakspaces!只读)) = 1, "√", "")
            
            objItem.Tag = Val(Nvl(rsbakspaces!当前)) & Val(Nvl(rsbakspaces!只读)) & IIf(Nvl(rsbakspaces!db连接) <> "", "0", "1")
            objItem.SubItems(C4所有者) = Nvl(rsbakspaces!所有者)
            objItem.SubItems(C8DBLink) = Nvl(rsbakspaces!db连接)
                        
            If Val(Nvl(rsbakspaces!当前)) = 1 Then
                If Val(Nvl(rsbakspaces!只读)) = 1 Then
                    strImgKey = "LockAndRun"
                Else
                    strImgKey = "Run"
                End If
            Else
                If Val(Nvl(rsbakspaces!只读)) = 1 Then
                    strImgKey = "Lock"
                Else
                    strImgKey = "Other"
                End If
            End If
            objItem.SmallIcon = strImgKey
            objItem.Icon = strImgKey
            
            On Error Resume Next
            gstrSQL = "select 系统,版本号,更新日期,最后转储日期,最后复制日期 from " & rsbakspaces!所有者 & ".ZLBAKINFO" & IIf(IsNull(rsbakspaces!db连接), "", "@" & rsbakspaces!db连接) & " where 系统=" & lng系统
            If rsTmp.State = 1 Then rsTmp.Close
            Set rsTmp = New ADODB.Recordset
            Call OpenRecordset(rsTmp, gstrSQL, gstrSysName, , , gcnOldOra) '创建历史空间后，所有权限都授给了应用系统的所有者的，所以应该能访问
            If err <> 0 Or gcnOldOra.Errors.Count > 0 Then
                MsgBox "警告:" & vbCrLf & "  历史数据空间" & rsbakspaces!名称 & "不能正常连接,请检查" & _
                    IIf(IsNull(rsbakspaces!db连接), "权限", "DB连接""" & rsbakspaces!db连接 & """") & "是否正常? ", vbInformation + vbDefaultButton1
            Else
                If Not rsTmp.EOF Then
                    objItem.SubItems(C5版本号) = Nvl(rsTmp!版本号)
                    objItem.SubItems(C6最后转储日期) = Format(rsTmp!最后转储日期, "yyyy-mm-dd")
                    objItem.SubItems(C7最后复制日期) = Format(rsTmp!最后复制日期, "yyyy-mm-dd")
                End If
            End If
            err.Clear: err = 0

            If Nvl(rsbakspaces!当前) = 1 Then
                objItem.Selected = True
                objItem.EnsureVisible
            End If
            rsbakspaces.MoveNext
        Loop
    End With
    mblnSel = False
    
    LoadHistorySpace = True
End Function

Private Sub SetCtlEnable(Optional blnCtlEnable As Boolean = True)
    '------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enable属性
    '参数:
    '------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim blnSys As Boolean
    Dim blnOwner As Boolean
    Dim blnCurSys As Boolean
    Dim blnSel As Boolean
    Dim bln只读 As Boolean
    Dim bln本地 As Boolean
    
    blnSys = cmbSystem.ListIndex >= 0
    blnOwner = cmbSystem.Tag = gstrUserName
    
    If Me.lvwHistory.SelectedItem Is Nothing Then
        blnCurSys = False
        blnSel = False
        bln只读 = True
        bln本地 = False
    Else
        blnSel = True
        blnCurSys = Mid(Me.lvwHistory.SelectedItem.Tag, 1, 1) = "1"
        bln只读 = Mid(Me.lvwHistory.SelectedItem.Tag, 2, 1) = "1"
        bln本地 = Mid(Me.lvwHistory.SelectedItem.Tag, 3, 1) = "1"
    End If
    
    cmdFunction(F0创建).Enabled = blnCtlEnable And blnSys And blnOwner And mblnHavingSpace
    cmdFunction(F2再植).Enabled = cmdFunction(F0创建).Enabled
    
    cmdFunction(F1拆卸).Enabled = blnSel And blnCtlEnable And blnSys And blnOwner And mblnHavingSpace
    cmdFunction(F4切换).Enabled = cmdFunction(F1拆卸).Enabled
    cmdFunction(F3复制).Enabled = blnSel And blnCtlEnable And blnSys And blnOwner And mblnHavingSpace And bln本地
    
    cmdFunction(5).Enabled = blnSel And blnCtlEnable And blnSys And blnOwner And mblnHavingSpace And bln本地
    
    Me.chk只读.Enabled = blnSel
    mblnSel = True
    If bln只读 Then
        Me.chk只读.value = 1
    Else
        Me.chk只读.value = 0
    End If
    chk只读.Enabled = bln本地
        
    mblnSel = False
   ' cmbSystem.Enabled = blnCtlEnable
End Sub
 
Private Sub cmdFunction_Click(Index As Integer)
    '--------------------------------------------------------------------------------------------------------------------------
    '创建历史数据空间
    '--------------------------------------------------------------------------------------------------------------------------
    Dim blnSucced As Boolean
    Dim lng系统 As Long
    Dim str空间名称 As String, str合并空间编号 As String
    Dim lng空间编号 As Long
    Dim bln当前 As Boolean
    Dim bln只读 As Boolean
    Dim strNote As String

    Dim lngSelNum As Long, i As Long
    
    If cmbSystem.ListIndex < 0 Then Exit Sub
    lng系统 = cmbSystem.ItemData(cmbSystem.ListIndex)
    
    Call SetCtlEnable(False)
    If Index <> 0 Then
        If lvwHistory.SelectedItem Is Nothing Then Exit Sub
        lng空间编号 = Val(Mid(lvwHistory.SelectedItem.Key, 2))
        bln当前 = Val(Mid(lvwHistory.SelectedItem.Tag, 1, 1)) = 1
        bln只读 = Val(Mid(lvwHistory.SelectedItem.Tag, 2, 1)) = 1
        
    Else
        lng空间编号 = 0
    End If
    
    '//todo
    Select Case Index
    Case F0创建
        '以odbc连接传入，因为oledb连接在创建dblink时，即使没有错误，也会出现cn.Errors.Count > 0,并且vb的err对象捕获不到错误
        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lng系统, 0, 0)
    Case F1拆卸
        If bln当前 Then
            MsgBox "该历史数据空间为当前历史数据空间，不能拆卸，请先对其他历史空间使用切换功能!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        
        lngSelNum = 0
        For i = 1 To lvwHistory.ListItems.Count
            If lvwHistory.ListItems(i).Checked Then
                If Val(Mid(lvwHistory.ListItems(i).Tag, 1, 1)) = 1 Then
                    MsgBox "选择的历史数据空间" & lvwHistory.ListItems(i).SubItems(C1名称) & "为当前历史数据空间，不能拆卸!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Sub
                End If
                lngSelNum = lngSelNum + 1
            End If
        Next
        
        If lngSelNum > 0 Then
            If MsgBox("你选择了" & lngSelNum & "个要拆卸的历史数据空间，你确定要继续吗？", vbOKCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbCancel Then
                Exit Sub
            End If
            For i = 1 To lvwHistory.ListItems.Count
                If lvwHistory.ListItems(i).Checked Then
                    lng空间编号 = Val(Mid(lvwHistory.ListItems(i).Key, 2))
                    blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lng系统, 1, lng空间编号)
                End If
            Next
        Else
            blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lng系统, 1, lng空间编号)
        End If
    Case F2再植
        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lng系统, 2, lng空间编号)
    Case F3复制
        
        If gstrServer = "" Then
            MsgBox "当前登录的服务名为空，复制操作要求必须指定服务名，请重新登录。", vbInformation, gstrSysName
            Exit Sub
        End If
        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lng系统, 3, lng空间编号)
    Case F4切换
        If bln当前 Then
            MsgBox "该历史数据空间为当前历史数据空间，不能切换!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        Else
            If MsgBox("即将重新创建H表视图指向当前选择的历史数据空间，以便查询历史数据。" & vbCrLf & "你确定要将当前历史数据空间切换" & lvwHistory.SelectedItem.SubItems(C1名称) & "为吗？", vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lng系统, 4, lng空间编号)
        If blnSucced Then
            '插入重要操作日志
            Call SaveAuditLog(2, "切换", "将历史数据空间切换为" & lvwHistory.SelectedItem.SubItems(C1名称))
        End If
    Case F5合并
        lngSelNum = 0
        lng空间编号 = 0
        For i = 1 To lvwHistory.ListItems.Count
            If lvwHistory.ListItems(i).Checked Then
                lngSelNum = lngSelNum + 1
                '取最小空间编号为保留空间编号
                If Val(Mid(lvwHistory.ListItems(i).Key, 2)) < lng空间编号 Or lng空间编号 = 0 Then
                    lng空间编号 = Val(Mid(lvwHistory.ListItems(i).Key, 2))
                    str空间名称 = lvwHistory.ListItems(i).SubItems(C1名称)
                    strNote = str空间名称
                End If
            End If
        Next
        
        If lngSelNum < 2 Then
            MsgBox "请至少勾选2个要合并的历史数据空间。", vbInformation + vbDefaultButton1, gstrSysName
        Else
            If MsgBox("勾选的" & lngSelNum & "个空间的数据将会被合并到编号最小的空间【" & str空间名称 & "】中。" & vbCrLf & _
                    "完成后，被合并空间及数据文件将会被删除,请确保已进行有效备份。" & vbCrLf & "你确定要继续吗？", vbOKCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbOK Then
                
                For i = 1 To lvwHistory.ListItems.Count
                    If lvwHistory.ListItems(i).Checked Then
                        If lng空间编号 <> Mid(lvwHistory.ListItems(i).Key, 2) Then
                            str合并空间编号 = str合并空间编号 & "," & Mid(lvwHistory.ListItems(i).Key, 2)
                            strNote = strNote & "," & lvwHistory.ListItems(i).SubItems(C1名称)
                        End If
                    End If
                Next
                blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lng系统, 5, lng空间编号, Mid(str合并空间编号, 2))
                If blnSucced Then
                    '插入重要操作日志
                    Call SaveAuditLog(2, "合并", "将历史数据空间" & strNote & "合并为" & str空间名称)
                End If
            End If
        End If
    Case F6转移

        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lng系统, 6, 0)
    End Select
    
    Call SetCtlEnable(True)
    If blnSucced = True Then
        Call cmbSystem_Click
    End If
    
End Sub
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    If gblnDBA Then
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", gstrUserName)
    End If
    
    With rsTemp
        .Filter = "编号=100 or 编号=2100 or 编号=2400"
        Do While Not .EOF
            cmbSystem.addItem !名称 & " v" & !版本号 & "（" & !编号 & "）"
            cmbSystem.ItemData(cmbSystem.NewIndex) = !编号
            .MoveNext
        Loop
        If cmbSystem.ListCount = 0 Then
            Call SetCtlEnable
        End If
        If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
        If cmbSystem.ListCount = 1 Then cmbSystem.Locked = True
    End With
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub


Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    With cmbSystem
        .Width = ScaleWidth - .Left - 50
    End With
    
    With lvwHistory
        .Width = ScaleWidth - .Left - 50
    End With
    With chkAll只读
        .Left = lvwHistory.Left + lvwHistory.Width - .Width - 50
        chk只读.Left = .Left - chk只读.Width - 50
    End With
    
    With cmdFunction(F0创建)
        .Left = lvwHistory.Left
        .Top = ScaleHeight - .Height - 100
        
        cmdFunction(F1拆卸).Top = .Top
        cmdFunction(F2再植).Top = .Top
        cmdFunction(F3复制).Top = .Top
        cmdFunction(F4切换).Top = .Top
        cmdFunction(F5合并).Top = .Top
        cmdFunction(F6转移).Top = .Top
        
        cmdFunction(F1拆卸).Left = .Left + .Width + 15
        cmdFunction(F2再植).Left = cmdFunction(F1拆卸).Left + cmdFunction(F1拆卸).Width + 15
        
        cmdFunction(F6转移).Left = ScaleWidth - cmdFunction(F6转移).Width - 60
        cmdFunction(F5合并).Left = cmdFunction(F6转移).Left - cmdFunction(F5合并).Width - 15
        cmdFunction(F4切换).Left = cmdFunction(F5合并).Left - cmdFunction(F4切换).Width - 15
        cmdFunction(F3复制).Left = cmdFunction(F4切换).Left - cmdFunction(F3复制).Width - 15
        
    End With
    lvwHistory.Height = cmdFunction(F0创建).Top - lvwHistory.Top - 10
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As zlPrintLvw
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "已安装的历名数据空间"
    Set objPrint.Body.objData = lvwHistory
    objPrint.BelowAppItems.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

 

Private Sub lvwHistory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '仍是刚才那列
        lvwHistory.SortOrder = IIf(lvwHistory.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwHistory.SortKey = mintColumn
        lvwHistory.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwHistory_DblClick()

    
    If lvwHistory.SelectedItem Is Nothing Then Exit Sub
    mblnSel = True
    If lvwHistory.SelectedItem.SubItems(C3只读) = "√" Then
        chk只读.value = 0
    Else
        chk只读.value = 1
    End If
    mblnSel = False
   Call chk只读_Click
   
    '加载时click隐式调用不执行
   If Me.Visible Then Call SetControlAll只读
End Sub

Private Sub SetControlAll只读()
    Dim i As Long, lngSel As Long

    lngSel = 0
    For i = 1 To lvwHistory.ListItems.Count
        If lvwHistory.ListItems(i).SubItems(C3只读) = "√" Then lngSel = lngSel + 1
    Next
    
    mblnChk = True
    chkAll只读.value = IIf(lngSel = 0, 0, IIf(lngSel = lvwHistory.ListItems.Count, 1, 2))
    mblnChk = False
End Sub

Private Sub lvwHistory_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SetCtlEnable
End Sub


