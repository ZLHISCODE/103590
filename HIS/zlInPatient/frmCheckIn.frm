VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人入病区"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmCheckIn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraInfo 
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   5940
      Begin VB.CheckBox chk包房 
         Caption         =   "包床"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3255
         TabIndex        =   1
         Top             =   1155
         Width           =   660
      End
      Begin VB.CheckBox chk陪伴 
         Caption         =   "是否陪伴"
         Height          =   195
         Left            =   4710
         TabIndex        =   7
         Top             =   1155
         Width           =   1020
      End
      Begin VB.ComboBox cbo床号 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1102
         Width           =   1530
      End
      Begin VB.ComboBox cbo病况 
         Height          =   300
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   671
         Width           =   1170
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.ComboBox cbo责任护士 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1965
         Width           =   1830
      End
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   671
         Width           =   1635
      End
      Begin VB.ComboBox cbo护理等级 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1533
         Width           =   4770
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3930
         TabIndex        =   5
         Top             =   1965
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   570
         TabIndex        =   26
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "责任护士"
         Height          =   180
         Left            =   210
         TabIndex        =   25
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lbl床位 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位"
         Height          =   180
         Left            =   570
         TabIndex        =   24
         Top             =   1162
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   570
         TabIndex        =   23
         Top             =   731
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理等级"
         Height          =   180
         Left            =   210
         TabIndex        =   22
         Top             =   1593
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病况"
         Height          =   180
         Left            =   4200
         TabIndex        =   21
         Top             =   731
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2910
         TabIndex        =   20
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4200
         TabIndex        =   19
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入病区时间"
         Height          =   180
         Left            =   3000
         TabIndex        =   15
         Top             =   2025
         Width           =   900
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6165
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4350
      Width           =   6165
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4935
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3750
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   105
         TabIndex        =   10
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fraLvw 
      Caption         =   "包房病床"
      Height          =   1830
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   5940
      Begin MSComctlLib.ListView lvw 
         Height          =   1425
         Left            =   150
         TabIndex        =   6
         Top             =   255
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   2514
         View            =   2
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "床位"
            Object.Width           =   5292
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mlng病人ID As Long
Public mlng主页ID As Long
Public mlngUnit As Long

Public mstr床号 As String '入:缺省定位的床号,表示家庭病床,出:入住的床号,可能多张床,用,号分隔
Public mlng床位科室ID As Long
Public mstrPrivs As String

Private mstrIDs As String
Private mstrText As String
Private mrsPatiInfo As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo病况_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo病况.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo病况.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo病况.ListIndex = lngIdx
    ElseIf cbo病况.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo床号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo护理等级_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo护理等级.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo护理等级.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo护理等级.ListIndex = lngIdx
    ElseIf cbo护理等级.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo责任护士_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo责任护士.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("护士", "", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo责任护士.ListCount - 1
                If cbo责任护士.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo责任护士.ListIndex = i: Exit Sub
                End If
            Next
            cbo责任护士.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo责任护士.ListCount - 1
            cbo责任护士.ListIndex = cbo责任护士.NewIndex
            cbo责任护士.ItemData(cbo责任护士.NewIndex) = rsTmp!上级ID
        Else
            cbo责任护士.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo责任护士_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo责任护士.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo责任护士.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo责任护士.ListIndex = lngIdx
    ElseIf cbo责任护士.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chk包房_Click()
    If chk包房.Value = 1 Then
        lbl床位.Caption = "主要床位"
        Call LoadMainBed
        lvw.Visible = True
        Me.Height = Me.Height + fraLvw.Height ' + 100
        If Visible Then lvw.SetFocus
    Else
        lbl床位.Caption = "床位"
        Call ShowBeds
        lvw.Visible = False
        Me.Height = Me.Height - fraLvw.Height '- 100
        If Visible Then cmdOK.SetFocus
    End If
End Sub

Private Sub chk包房_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk陪伴_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Integer
    Dim strSQL医疗小组 As String
    Dim strIDs As String, strID As String, strCode As String
    Dim strTmp As String
    
    On Error GoTo errH
    gblnOK = False

    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID, 2)
    '初始化数据
    With mrsPatiInfo
        txt科室.Enabled = False
        
        '所选病床的科室与病人科室不同时,不允许换科.
        If mlng床位科室ID <> 0 Then
            If mlng床位科室ID <> !入住科室id Then
                MsgBox "病人当前科室【" & !当前科室 & "】与选择的床位所属科室【" & GetDeptName(mlng床位科室ID) & "】不同,不能入住该床位,请选择其它床位!", vbInformation, gstrSysName
                Unload Me: Exit Sub
            End If
        End If

        txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")

        '病人信息
        txt姓名.Text = !姓名
        txt性别.Text = "" & !性别
        txt年龄.Text = "" & !年龄
        txt科室.Text = "" & !当前科室
        'txt住院号.Text = "" & !住院号
        txt科室.Tag = "" & !入住科室id
        
        
        '确定病区的服务对象
        strSql = "Select 服务对象 From 部门性质说明 Where 工作性质='护理' And 部门ID=[1]" '
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit)
            
        '有床位的临床科室
        If rsTmp!服务对象 = 1 Then
            strTmp = "1,3"
        ElseIf rsTmp!服务对象 = 2 Then
            strTmp = "2,3"
        ElseIf rsTmp!服务对象 = 3 Then
            If Val("" & !病人性质) = 1 Then
                strTmp = "1,3"
            Else
                strTmp = "2,3"
            End If
        End If
        Set rsTmp = GetDeptOrUnit(0, mlngUnit, strTmp)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                strIDs = strIDs & "," & rsTmp!ID
                rsTmp.MoveNext
            Next
        Else
            '没有对应的床位科室
            MsgBox "在当前病区没有设置对应科室,病人不能入住！" & vbCrLf, vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        
        '病况
        strSql = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 病情 Order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo病况.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 And cbo病况.ListIndex = -1 Then cbo病况.ListIndex = cbo病况.NewIndex
                If rsTmp!名称 = "" & !当前病况 Then cbo病况.ListIndex = cbo病况.NewIndex
                rsTmp.MoveNext
            Next
        End If
    
        '护理等级
        cbo护理等级.Enabled = InStr(mstrPrivs, ";" & "调整护理等级" & ";") > 0
        Set rsTmp = GetNurseGrade
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo护理等级.AddItem rsTmp!编码 & "-" & rsTmp!名称
                cbo护理等级.ItemData(cbo护理等级.NewIndex) = rsTmp!ID
                If rsTmp!ID = !护理等级ID Then cbo护理等级.ListIndex = cbo护理等级.NewIndex
                rsTmp.MoveNext
            Next
        End If
        
        '住院护士
        Set rsTmp = GetDoctorOrNurse(1, strIDs & "," & mlngUnit & ",")
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo责任护士.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                cbo责任护士.ItemData(i - 1) = rsTmp!ID
                rsTmp.MoveNext
            Next
            Call SeekDoctor(cbo责任护士, "" & mrsPatiInfo!责任护士)
        End If
        cbo责任护士.AddItem "其它..."
        
        '显示该科室的床位
        Call ShowBeds
        If Not Visible Then chk包房_Click
        
    End With
    
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngUnit = 0
    mstrText = ""
    '卸载消息对象
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadMainBed
End Sub

Private Sub LoadMainBed()
    Dim i As Integer, strBed As String
    
    If cbo床号.ListIndex <> -1 Then strBed = cbo床号.Text
    cbo床号.Clear
    For i = 1 To lvw.ListItems.Count
        If lvw.ListItems(i).Checked Then
            cbo床号.AddItem lvw.ListItems(i).Text
            If lvw.ListItems(i).Text = strBed Then cbo床号.ListIndex = cbo床号.NewIndex
            If cbo床号.ListIndex = -1 And mstr床号 <> "" Then
                If lvw.ListItems(i).Text = mstr床号 Then cbo床号.ListIndex = cbo床号.NewIndex
            End If
        End If
    Next
    If cbo床号.ListIndex = -1 And cbo床号.ListCount = 1 Then cbo床号.ListIndex = 0
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If IsDate(txtDate.Text) And KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub ShowBeds()
'功能：显示当前病区当前科室可用的病床
    Dim i As Integer, objItem As ListItem
    Dim lng科室ID As Long
    Dim rsBeds As ADODB.Recordset
    
    lvw.ListItems.Clear
    cbo床号.Clear
    If InStr(1, mstrPrivs, "家庭病床") > 0 Then
        cbo床号.AddItem "家庭病床"
        If mstr床号 = "家庭病床" Then cbo床号.ListIndex = 0
    End If
    lng科室ID = txt科室.Tag
    Set rsBeds = GetFreeBeds(mlngUnit, lng科室ID, mrsPatiInfo!性别, mlng病人ID)
    
    With rsBeds
        For i = 1 To rsBeds.RecordCount
            Set objItem = lvw.ListItems.Add(, "_" & !床号, !床号 & IIf(IsNull(!房间号), "", " 房间:" & !房间号))
            objItem.Tag = !等级ID
            cbo床号.AddItem objItem.Text
            
            If !床号 = mstr床号 Then
                objItem.Checked = True: objItem.Selected = True: objItem.EnsureVisible
                cbo床号.ListIndex = cbo床号.NewIndex
            End If
            
            .MoveNext
        Next
    End With
    
    If cbo床号.ListIndex = -1 And cbo床号.ListCount > 0 Then cbo床号.ListIndex = 0
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim dMax As Date, i As Integer, Curdate As Date
    Dim strPreRoom As String, intRoom As Integer, intCheck As Integer, lngNurseGrade As Long
    Dim strSql As String, strBed As String, strTmp As String
    Dim str床号 As String, str房间号 As String, blnTrans As Boolean, strMainBed As String
    Dim rsTmp As New ADODB.Recordset
        Dim colSQL As New Collection, strSQLtmp As String, rsPati As Recordset

    If cbo病况.ListIndex = -1 Then
        MsgBox "请指定病人的当前病况！", vbInformation, gstrSysName
        cbo病况.SetFocus: Exit Sub
    End If
    
    If cbo护理等级.ListIndex = -1 And gbln入科确定护理等级 Then
        MsgBox "请指定病人的当前护理等级！", vbInformation, gstrSysName
        cbo护理等级.SetFocus: Exit Sub
    End If
    
    If cbo护理等级.ListIndex <> -1 Then
        lngNurseGrade = cbo护理等级.ItemData(cbo护理等级.ListIndex)
    End If
    
    '时间不能超过当前时间太长(一个月)
    Curdate = zlDatabase.Currentdate
    If InStr(Trim(cbo床号.Text), " 房间") <> 0 Then
        str床号 = Mid(Trim(cbo床号.Text), 1, InStr(Trim(cbo床号.Text), " 房间") - 1)
        str房间号 = Mid(Trim(cbo床号.Text), InStr(Trim(cbo床号.Text), "房间:") + 3)
    ElseIf InStr(cbo床号.Text, "家庭病床") > 0 Then
        str床号 = ""
    Else
        str床号 = Trim(cbo床号.Text)
    End If
    If CDate(txtDate.Text) > Curdate Then
        MsgBox "入病区时间大于了当前系统时间,请检查！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    '可能与入院时间相同
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)

    If Format(txtDate.Text, "yyyyMMddhhmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "时间必须大于该病人的上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If

    
    If chk包房.Value = 1 Then
        strPreRoom = "一定不同"
        For i = 1 To lvw.ListItems.Count
            If lvw.ListItems(i).Checked Then
                intCheck = intCheck + 1
                strTmp = lvw.ListItems(i).Text
                If InStr(1, strTmp, ":") > 0 Then   '冒号后是房间号
                    strTmp = Mid(strTmp, InStr(1, strTmp, ":") + 1)
                    If strTmp <> strPreRoom Then
                        intRoom = intRoom + 1
                        strPreRoom = strTmp
                    End If
                End If
            End If
        Next
        If intCheck < 2 Then
            MsgBox "包床病人必须分配两张以上的床位！", vbInformation, gstrSysName
            lvw.SetFocus: Exit Sub
        End If
        If intRoom > 1 Then
            MsgBox "包房病人所分配的床位必须在一个房间内！", vbInformation, gstrSysName
            lvw.SetFocus: Exit Sub
        End If
    End If
    
    If chk包房.Value = 0 Then
        strBed = str床号
        strMainBed = str床号
    Else
        strMainBed = str床号
        For i = 1 To lvw.ListItems.Count
            If lvw.ListItems(i).Checked Then
                strBed = strBed & "," & Mid(lvw.ListItems(i).Key, 2)
            End If
        Next
        strBed = Mid(strBed, 2)
    End If

    strSql = "zl_病人变动记录_InUnit(" & mlng病人ID & "," & mlng主页ID & ",'" & strBed & "'," & _
            mlngUnit & "," & lngNurseGrade & ",'" & zlCommFun.GetNeedName(cbo病况.Text) & "'," & chk陪伴.Value & ",'" & zlCommFun.GetNeedName(cbo责任护士.Text) & "'," & _
            "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "','" & strMainBed & "')"

    '转病区费用检查
    If CreatePublicExpenseBillOperation() And gbln转病区转费用 Then
        strSQLtmp = "Select ID, 病区id" & vbNewLine & _
                    "From 病人变动记录" & vbNewLine & _
                    "Where 病人id = [1] And 主页id = [2] And 开始时间 Is Not Null And 终止时间 Is Null And NVL(附加床位,0) = 0"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQLtmp, Me.Caption, mlng病人ID, mlng主页ID)
        If rsPati.RecordCount > 0 Then
            If gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(Me, 0, mlng病人ID, mlng主页ID, Val(rsPati!ID & ""), Val(rsPati!病区ID & ""), mlngUnit, colSQL) = False Then Exit Sub
        End If
    End If
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSql, Me.Caption
        For i = 1 To colSQL.Count
            zlDatabase.ExecuteProcedure colSQL(i), Me.Caption
        Next

    If Val("" & mrsPatiInfo!险类) <> 0 Then
        If Not gclsInsure.ModiPatiSwap(mlng病人ID, mlng主页ID, Val("" & mrsPatiInfo!险类), "1") Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    '新网96847、118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng病人ID, mlng主页ID) <> 1 Then
            MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISModPati)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "当前启用了影像信息系统接口，但由于RIS接口创建失败未调用(HISModPati)接口，请与系统管理员联系。", vbInformation, gstrSysName
    End If
    mstr床号 = strBed
    gblnOK = True
    
    On Error Resume Next
    '入病区成功后触发消息
    If mclsMipModule.IsConnect = True Then
        mclsXML.ClearXmlText '清除缓存中的XML
        '--进行消息组装
        '病人信息
        mclsXML.AppendNode "in_patient"
        'patient_id      病人id  1   N
        mclsXML.appendData "patient_id", mlng病人ID, xsNumber  '病人ID
        'page_id     主页id  1   N
        mclsXML.appendData "page_id", mlng主页ID, xsNumber '主页ID
        'patient_name        姓名    1   S
        mclsXML.appendData "patient_name", txt姓名.Text, xsString '姓名
        'patient_sex     性别    0..1    S
        mclsXML.appendData "patient_sex", txt性别.Text, xsString '性别
        'in_number       住院号  1   S
        mclsXML.appendData "in_number", Nvl(mrsPatiInfo!住院号), xsString '住院号
        mclsXML.AppendNode "in_patient", True
        
        strSql = " Select A.ID,B.名称 床位等级,C.名称 病区名称  From  病人变动记录 A,收费项目目录 B,部门表 C" & _
            " Where NVl(A.附加床位,0)=0 And A.床位等级id=B.id(+) And A.病区Id=C.id(+) And A.病人ID=[1] And A.主页ID=[2] And A.开始原因=[3] And 开始时间+0=[4]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病人变动记录", mlng病人ID, mlng主页ID, 15, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        
        '住院信息
        mclsXML.AppendNode "in_hospital"
        'in_date     入院时间    1   s
        mclsXML.appendData "in_date", Format(mrsPatiInfo!入院时间, "yyyy-MM-dd HH:mm:ss"), xsString
        'out_area_id     转出病区id  0..1    N
        mclsXML.appendData "out_area_id", Val(Nvl(mrsPatiInfo!当前病区ID)), xsNumber
        'out_area_title      转出病区    0..1    S
        mclsXML.appendData "out_area_title", Nvl(mrsPatiInfo!当前病区), xsString
        'out_dept_id     转出科室id  1   N
        mclsXML.appendData "out_dept_id", Val(Nvl(mrsPatiInfo!出院科室id, 0)), xsNumber
        'out_dept_title      转出科室    1   S
        mclsXML.appendData "out_dept_title", Nvl(mrsPatiInfo!当前科室), xsString
        'in_area_id      转入病区id  0..1    N
        mclsXML.appendData "in_area_id", mlngUnit, xsNumber
        'in_area_title       转入病区    0..1    S
        mclsXML.appendData "in_area_title", Nvl(rsTmp!病区名称), xsString
        'in_dept_id      转入科室id  1   N
        mclsXML.appendData "in_dept_id", Val(txt科室.Tag), xsNumber
        'in_dept_title       转入科室    1   S
        mclsXML.appendData "in_dept_title", txt科室.Text, xsString
        mclsXML.AppendNode "in_hospital", True
        '转科入科
        mclsXML.AppendNode "change_dept_arrange"
        'change_id       变动id  1   N
        mclsXML.appendData "change_id", rsTmp!ID, xsNumber '变动ID
        'in_room     入住病房    0..1    S
        mclsXML.appendData "in_room", str房间号, xsString
        'in_bed      入住病床    1   S
        mclsXML.appendData "in_bed", strMainBed, xsString
        'in_tendgrade        护理等级    0..1    S
        If cbo护理等级.ListIndex <> -1 Then
            mclsXML.appendData "in_tendgrade", zlCommFun.GetNeedName(cbo护理等级.Text), xsString
        Else
            mclsXML.appendData "in_tendgrade", "", xsString
        End If
        'in_bedgrade     床位等级    0..1    S
        mclsXML.appendData "in_bedgrade", Nvl(rsTmp!床位等级), xsString
        'in_doctor       住院医师    0..1    S
        mclsXML.appendData "in_doctor", Nvl(mrsPatiInfo!住院医师), xsString
        'duty_nurse      责任护士    0..1    S
        mclsXML.appendData "duty_nurse", zlCommFun.GetNeedName(cbo责任护士.Text), xsString
        'change_operator         操作员      1   S
        mclsXML.appendData "change_operator", UserInfo.姓名, xsString
        mclsXML.AppendNode "change_dept_arrange", True
        mclsMipModule.CommitMessage "ZLHIS_PATIENT_012", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'问题28432 by lesfeng 2010-03-10
Private Function GetDeptName(ByVal lngID As Long) As String

    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select 名称 From 部门表 Where ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    
    If Not rsTmp.EOF Then
        GetDeptName = IIf(IsNull(rsTmp!名称), "", rsTmp!名称)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SeekDoctor(cbo As ComboBox, Optional strPre As String)
    Dim strIDs As String, i As Integer
    
    If strPre <> "" Then
        For i = 0 To cbo.ListCount - 1
            If zlCommFun.GetNeedName(cbo.List(i)) = strPre Then cbo.ListIndex = i: Exit Sub
        Next
    End If
    
    strIDs = GetDeptDoctors(txt科室.Tag)
    For i = 0 To cbo.ListCount - 1
        If InStr("," & strIDs & ",", "," & cbo.ItemData(i) & ",") > 0 Then cbo.ListIndex = i: Exit Sub
    Next
    
    strIDs = GetDeptDoctors(mlngUnit)
    For i = 0 To cbo.ListCount - 1
        If InStr("," & strIDs & ",", "," & cbo.ItemData(i) & ",") > 0 Then cbo.ListIndex = i: Exit Sub
    Next
End Sub


Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
        ByRef str床号 As String, ByVal lng床位科室ID As Long, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr床号 = str床号
    mlng床位科室ID = lng床位科室ID
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    str床号 = mstr床号
    ShowMe = gblnOK
End Function


