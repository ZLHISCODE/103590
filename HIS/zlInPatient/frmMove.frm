VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMove 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人换床"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmMove.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   5805
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3750
      Width           =   5805
      Begin VB.CheckBox chk包房 
         Caption         =   "包床(&M)"
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   1455
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4440
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3240
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fraBed 
      Height          =   1950
      Left            =   120
      TabIndex        =   12
      Top             =   75
      Width           =   5565
      Begin VB.TextBox txtUnit 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1065
         Width           =   1800
      End
      Begin VB.ComboBox cboNew 
         Height          =   300
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1065
         Width           =   1845
      End
      Begin VB.TextBox txt科室 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   660
         Width           =   1800
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1800
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3510
         TabIndex        =   5
         Top             =   1500
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前病区"
         Height          =   180
         Left            =   105
         TabIndex        =   20
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   465
         TabIndex        =   18
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   180
         Left            =   465
         TabIndex        =   17
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         Height          =   180
         Left            =   2910
         TabIndex        =   16
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "换床时间"
         Height          =   180
         Left            =   2730
         TabIndex        =   15
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label lblNew 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新病床"
         Height          =   180
         Left            =   2925
         TabIndex        =   14
         Top             =   1125
         Width           =   525
      End
      Begin VB.Label lblPre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原病床"
         Height          =   180
         Left            =   2910
         TabIndex        =   13
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.Frame fraLvw 
      Caption         =   "选择病床"
      Height          =   1545
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   5565
      Begin MSComctlLib.ListView lvw 
         Height          =   1140
         Left            =   165
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   255
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   2011
         View            =   2
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mbytInFun As Byte '0-换床,1-包房增减床,2-撤消出院时换新床
Public mstr目标床号 As String '入：当mbytInFun=0时表示拖动到指定的床位上，=2时表示床位对应的等级ID，出：分配的新床号(可能多张)
Public mstr床号 As String   '当mbytInFun=2时才有值,表示出院前的床位

Public mstrPrivs As String
Public mlngUnit As Long
Public mlng病人ID As Long, mlng主页ID As Long
Private mfrmParent As Object

Private mrsPatiInfo As ADODB.Recordset
Private mrsBeds As ADODB.Recordset '可选床位集

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub chk包房_Click()
        
    If chk包房.Value = 1 Then
        lvw.Visible = True
        lvw.TabStop = True
        lblNew.Caption = "主床位"
        If Visible Then Call LoadMainBed
        
        fraLvw.Top = fraBed.Top + fraBed.Height + 200
        Me.Height = fraLvw.Top + fraLvw.Height + Picture1.Height + 400
        If lvw.Visible Then lvw.SetFocus
        
    Else
        lvw.Visible = False
        lvw.TabStop = False
        lblNew.Caption = "新病床"
        If Visible Then Call InitBed(mlngUnit)
        
        Me.Height = fraBed.Top + fraBed.Height + Picture1.Height + 400
        If cboNew.Visible Then cboNew.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Activate()
    If lvw.Visible And lvw.Enabled Then
        lvw.SetFocus
    ElseIf cboNew.Visible And cboNew.Enabled Then
        cboNew.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    gblnOK = False
    
    '界面调整
    Me.Height = fraBed.Top + fraBed.Height + Picture1.Height + 400
    
    Call InitData
    
    Select Case mbytInFun
        Case 0  '0-换床
            If cboNew.ListCount = 0 Then
                MsgBox "病人所在科室的病区已没有合适的床位可供换床！", vbInformation, gstrSysName
                Unload Me: Exit Sub
            End If
        Case 1 '1-包房增减床
            If lvw.ListItems.Count = 0 Then
                MsgBox "病人当前病房没有其它空床！", vbInformation, gstrSysName
                Unload Me: Exit Sub
            End If
        Case 2 '2-撤消出院时换新床
            If UBound(Split(txtPre.Text, ";")) > 0 Then '出院之前多张床位
                If lvw.ListItems.Count = 0 Then
                    MsgBox "病人当前病房没有其它空床！", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            Else
                If cboNew.ListCount = 0 Then
                    MsgBox "病人所在科室的病区已没有合适的床位可供换床！", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
    End Select
    
    '创建消息对象
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Sub InitData()
    Dim i As Integer, rsTmp As ADODB.Recordset, str床号 As String, str房间号 As String
    
    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With mrsPatiInfo
        txt姓名.Text = !姓名
        txt姓名.Tag = "" & !性别
        txt住院号.Text = "" & !住院号
        txt科室.Text = !当前科室
    End With
    
    str房间号 = ""
    If mbytInFun = 2 Then
        txtPre.Text = mstr床号
    Else
        Set rsTmp = GetPatiBeds(mlng病人ID)
        If rsTmp.RecordCount = 0 Then
            str床号 = "家庭病床"
        Else
            Do While Not rsTmp.EOF
                str床号 = str床号 & "," & rsTmp!床号
                If Nvl(rsTmp!床号) = Nvl(mrsPatiInfo!主要床号) And Nvl(rsTmp!科室ID) = Nvl(mrsPatiInfo!入住科室id) Then
                    str房间号 = Nvl(rsTmp!房间号)
                End If
                rsTmp.MoveNext
            Loop
            str床号 = Mid(str床号, 2)
        End If
        txtPre.Text = str床号
        txtPre.Tag = str房间号
    End If
                                
    If UBound(Split(txtPre.Text, ",")) > 0 And mstr目标床号 <> "家庭病床" Or mbytInFun = 1 Then
        chk包房.Value = 1   '触发click事件
    Else
        Call chk包房_Click  '缺省值为0
    End If
    If mbytInFun = 1 Or mbytInFun = 2 Then chk包房.Visible = False
    If mbytInFun <> 0 Then txtUnit.Visible = False: lblUnit.Visible = False
    If mbytInFun = 2 Then lblDate.Visible = False: txtDate.Visible = False
    
    Select Case mbytInFun
        Case 0 '换床
            Me.Caption = "病人换床"
            '目前包含门诊观察室
            Set rsTmp = GetDeptOrUnit(1, mrsPatiInfo!出院科室id, "1,2,3")
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    If rsTmp!ID = mlngUnit Then txtUnit.Text = rsTmp!名称 '当前病区优先
                    rsTmp.MoveNext
                Next
            End If
            Call InitBed(mlngUnit)
            
        Case 1 '1-包房增减床
            Me.Caption = "病人包房加减床位"
            lblDate.Caption = "变动时间"
            
            lblDate.Top = lblNew.Top: txtDate.Top = cboNew.Top
            lblNew.Left = lblUnit.Left: cboNew.Left = txtUnit.Left
            
            fraBed.Height = fraBed.Height - cboNew.Height - 100
            fraLvw.Top = fraLvw.Top - cboNew.Height - 100
            Me.Height = Me.Height - cboNew.Height - 100
            
            Call InitBed(0)
            If lvw.ListItems.Count = 0 Then Exit Sub
            
        Case 2 '2-撤消出院时换新床
            Me.Caption = "撤消出院安排新床位"
            fraBed.Height = fraBed.Height - cboNew.Height - 100
            fraLvw.Top = fraLvw.Top - cboNew.Height - 100
            Me.Height = Me.Height - cboNew.Height - 100
            
            Call InitBed(mlngUnit)
     End Select
    
End Sub

Private Sub InitBed(ByVal lng病区ID As Long)
'功能：初始化床位,此时取该病区及科室对应的所有空床
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Integer
    Dim bytLen As Byte
    Dim strTmp As String
    
    On Error GoTo errH
        
    If InStr(mrsPatiInfo!性别, "男") > 0 Then
        strTmp = "男床,不限床"
    ElseIf InStr(mrsPatiInfo!性别, "女") > 0 Then
        strTmp = "女床,不限床"
    Else
        strTmp = "不限床"
    End If
        
    lvw.ListItems.Clear
    cboNew.Clear
    
    Select Case mbytInFun
        Case 0, 2 '0-换床 '2-撤消出院时换新床
            If mbytInFun = 0 Then
                If InStr(1, mstrPrivs, "家庭病床") > 0 And txtPre.Text <> "家庭病床" And chk包房.Value = 0 Then
                    cboNew.AddItem "家庭病床", 0
                    If mstr目标床号 = "家庭病床" Then cboNew.ListIndex = 0
                End If
            End If
            
            bytLen = GetMaxBedLen(lng病区ID)
            '当前病区的共用空床+当前病区当前科室的空床
            strSql = "Select 床号,性别分类,房间号,等级ID From 床位状况记录" & vbNewLine & _
                    " Where 状态='空床'" & vbNewLine & _
                    " And instr([1],性别分类)>0 And (科室ID is Null Or 科室ID=[2]) And 病区ID=[3] " & vbNewLine & _
                    " Order by  LPad(NVL(房间号,0), 10, ' '),LPad(床号, 10, ' ')"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp, Val(mrsPatiInfo!出院科室id), lng病区ID)
            Set mrsBeds = rsTmp.Clone
            
            For i = 1 To rsTmp.RecordCount
                cboNew.AddItem Space(bytLen - Len(rsTmp!床号)) & rsTmp!床号 & IIf(IsNull(rsTmp!房间号), "", " 房间:" & rsTmp!房间号)
                If mbytInFun = 0 Then
                    If rsTmp!床号 = mstr目标床号 Then cboNew.ListIndex = cboNew.NewIndex
                ElseIf mbytInFun = 2 Then
                    If rsTmp!床号 = "" & mrsPatiInfo!主要床号 Then cboNew.ListIndex = cboNew.NewIndex
                End If
                                
                If mbytInFun = 0 Or UBound(Split(txtPre.Text, ",")) > 0 Then
                    lvw.ListItems.Add , "_" & rsTmp!床号, rsTmp!床号 & IIf(IsNull(rsTmp!房间号), "", " 房间:" & rsTmp!房间号)
                    lvw.ListItems(lvw.ListItems.Count).Tag = "" & rsTmp!房间号
                    If mbytInFun = 2 Then
                        If InStr(1, "," & txtPre.Text & ",", "," & rsTmp!床号 & ",") > 0 Then
                            lvw.ListItems(lvw.ListItems.Count).Checked = True
                        End If
                    End If
                End If
                
                rsTmp.MoveNext
            Next
            
            If chk包房.Value = 1 Then
                If Not lvw.ListItems.Count > 0 Then Exit Sub
                lvw.ListItems(1).Selected = True
                lvw.SelectedItem.EnsureVisible
            Else
                If cboNew.ListIndex = -1 And cboNew.ListCount > 0 Then cboNew.ListIndex = 0
            End If
                             
        Case 1 '1-包房增减床
            
            strSql = "Select A.床号,A.状态 From 床位状况记录 A" & vbNewLine & _
                    " Where (A.科室ID,A.病区ID,A.房间号) In (Select Distinct B.科室ID,B.病区ID,B.房间号 From 床位状况记录 B Where 病人ID = [1]) " & vbNewLine & _
                    " And (A.状态 = '占用' And 病人ID = [1] Or A.状态 = '空床') And instr([2],性别分类)>0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID, strTmp)
            If rsTmp.RecordCount < 2 Then Exit Sub
            
            For i = 1 To rsTmp.RecordCount
                lvw.ListItems.Add , "_" & rsTmp!床号, rsTmp!床号
                
                If rsTmp!状态 = "占用" Then
                    lvw.ListItems(lvw.ListItems.Count).Checked = True
                    lvw.ListItems(lvw.ListItems.Count).Selected = True
                    
                    cboNew.AddItem rsTmp!床号 '可用主要床号
                    If rsTmp!床号 = "" & mrsPatiInfo!主要床号 Then cboNew.ListIndex = cboNew.NewIndex
                End If
                
                rsTmp.MoveNext
            Next
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
    mlngUnit = 0
    
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

Private Function StringDelItem(ByVal strAll As String, ByVal strItem As String) As String
'功能：从指定的字符串列表中删除一项(如果有多个匹配的,只移除第一个)
    Dim i As Long, arrTmp As Variant
    
    arrTmp = Split(strAll, ",")
    For i = 0 To UBound(arrTmp)
        If arrTmp(i) = strItem Then
            strItem = ""
        Else
            StringDelItem = StringDelItem & "," & arrTmp(i)
        End If
    Next
    StringDelItem = Mid(StringDelItem, 2)
End Function

Private Sub cmdOK_Click()
    Dim strBeds As String, strBed As String, strSql As String, strUnitID As String, strMainBed As String
    Dim dMax As Date, i As Integer, j As Integer, blnTrans As Boolean
    Dim strRoom As String, Curdate As Date, strBedGrids As String, strBedGridsNew As String
    Dim rsTmp As New ADODB.Recordset
    
    '时间不能超过当前时间太长(一个月)
    If mbytInFun <> 2 Then
        Curdate = zlDatabase.Currentdate
        If CDate(txtDate.Text) > Curdate Then
            If CDate(txtDate.Text) - Curdate > 30 Then
                MsgBox "换床时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
            If MsgBox("换床时间大于了当前系统时间,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        dMax = GetMaxDate(mlng病人ID, mlng主页ID)
        If CDate(txtDate.Text) <= dMax Then
            MsgBox "病人换床时间必须大于上次变动的时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    If chk包房.Value = 1 Then
        If Trim(cboNew.Text) = "" Then
            MsgBox "请为病人指定主床位！", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(Trim(cboNew.Text), "家庭病床") > 0 Then
            strMainBed = ""
        ElseIf InStr(Trim(cboNew.Text), " 房间") > 0 Then
            strMainBed = Mid(Trim(cboNew.Text), 1, InStr(Trim(cboNew.Text), " 房间") - 1)
        Else
            strMainBed = Trim(cboNew.Text)
        End If
    Else
        If cboNew.ListIndex = -1 Then
            MsgBox "请选择要换入的床位！", vbInformation, gstrSysName
            cboNew.SetFocus: Exit Sub
        End If
        strMainBed = Trim(Split(cboNew.Text, "房间:")(0))
    End If
    
    Select Case mbytInFun
        Case 0
            '取床位
            If chk包房.Value = 0 Then
                If InStr(Trim(cboNew.Text), "家庭病床") > 0 Then
                    strBeds = ""
                Else
                    If InStr(Trim(cboNew.Text), " 房间") > 0 Then
                        strBeds = Mid(Trim(cboNew.Text), 1, InStr(Trim(cboNew.Text), " 房间") - 1)
        
                        strRoom = Mid(Trim(cboNew.Text), InStr(Trim(cboNew.Text), "房间:") + 3)
                        
                        strSql = "Select 性别 From 病人信息 A,床位状况记录 B  Where A.病人ID = b.病人id And b.病人ID Is Not Null And 病区ID = [1] And 房间号 =[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit, strRoom)
                        
                        Do While Not rsTmp.EOF
                         
                            If Trim(txt姓名.Tag) <> rsTmp!性别 Then
                                If (MsgBox("指定床位所在房间存在男女混住情况，是否继续入住？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                                    Exit Do
                                Else
                                    Exit Sub
                                    cboNew.SetFocus
                                End If
                            End If
                            rsTmp.MoveNext
                        Loop
                    Else
                        strBeds = Trim(cboNew.Text)
                    End If
                End If
                
            Else
                For i = 1 To lvw.ListItems.Count
                    If lvw.ListItems(i).Checked Then
                        j = j + 1
                        If j = 1 Then
                            strRoom = lvw.ListItems(i).Tag
                        ElseIf lvw.ListItems(i).Tag <> strRoom Then
                            MsgBox "病人包房时必须选择同一个房间内的床位！", vbInformation, gstrSysName
                            lvw.SetFocus: Exit Sub
                        End If
                        strBeds = strBeds & "," & Mid(lvw.ListItems(i).Key, 2)
                    End If
                Next
                If strBeds = "" Then
                    MsgBox "请选择要换入的床位！", vbInformation, gstrSysName
                    lvw.SetFocus: Exit Sub
                End If
                strBeds = Mid(strBeds, 2)
                If UBound(Split(strBeds, ",")) = 0 Then
                    MsgBox "病人包床时至少应选择两张以上的床位！", vbInformation, gstrSysName
                    lvw.SetFocus: Exit Sub
                End If
                
                If strBeds = txtPre.Text Then
                    MsgBox "病人新换入的包房床位与原床位相同,请重新选择！", vbInformation, gstrSysName
                    lvw.SetFocus: Exit Sub
                End If
            End If
            
            strUnitID = mlngUnit
        Case 1
            For i = 1 To lvw.ListItems.Count
                If lvw.ListItems(i).Checked Then Exit For
            Next
            If i = lvw.ListItems.Count + 1 Then
                MsgBox "请至少选择一张病床！", vbInformation, gstrSysName
                lvw.SetFocus: Exit Sub
            End If
            For i = 1 To lvw.ListItems.Count
                If lvw.ListItems(i).Checked Then
                    strBeds = strBeds & "," & Mid(lvw.ListItems(i).Key, 2)
                End If
            Next
            strBeds = Mid(strBeds, 2)
            strUnitID = mlngUnit
        Case 2
            If UBound(Split(txtPre.Text, ",")) > 0 Then '多张床
                j = 0
                For i = 1 To lvw.ListItems.Count
                    If lvw.ListItems(i).Checked Then j = j + 1
                Next
                If j <> UBound(Split(txtPre.Text, ",")) + 1 Then
                    MsgBox "新安排的床位数量与原入住床位数量不一至，请检查！", vbExclamation, gstrSysName
                    Exit Sub
                End If
                
                strBeds = strMainBed  '主床号放到第一个，用于传出
                strBedGrids = mstr目标床号
                For i = 1 To lvw.ListItems.Count
                    If lvw.ListItems(i).Checked Then
                        strBed = Trim(lvw.ListItems(i).Text)
                        If UBound(Split(strBed, "房间:")) = 0 Then '无房间号
                            mrsBeds.Filter = "床号='" & Trim(strBed) & "'"
                        Else
                            mrsBeds.Filter = "床号='" & Trim(Split(strBed, "房间:")(0)) & "' and 房间号='" & Split(strBed, "房间:")(1) & "'"
                        End If
                        
                        strBedGridsNew = StringDelItem(strBedGrids, mrsBeds!等级ID)
                        If strBedGridsNew = strBedGrids Then
                            MsgBox "新安排的床位等级与原入住床位等级不一至，请检查！", vbExclamation, gstrSysName
                            Exit Sub
                        Else
                            strBedGrids = strBedGridsNew
                        End If
                        If InStr(1, strBeds, Trim(Split(strBed, "房间:")(0))) = 0 Then    '主床号已放到第一个
                            strBeds = strBeds & "," & Trim(Split(strBed, "房间:")(0))
                        End If
                    End If
                Next
                If strBedGrids <> "" Then
                    MsgBox "新安排的床位等级与原入住床位等级不一至，请检查！", vbExclamation, gstrSysName
                    Exit Sub
                End If
            Else '单张床
                '不会有家庭病床,因为撤销出院时必须重新指定床位
                strBed = Trim(cboNew.Text)
                If UBound(Split(strBed, "房间:")) = 0 Then
                    mrsBeds.Filter = "床号='" & strBed & "'"
                Else
                    mrsBeds.Filter = "床号='" & Trim(Split(strBed, "房间:")(0)) & "' and 房间号='" & Split(strBed, "房间:")(1) & "'"
                End If
                '等级比较
                If mstr目标床号 <> CStr(mrsBeds!等级ID) Then
                    MsgBox "新安排的床位等级与原入住床位等级不一至，请检查！", vbExclamation, gstrSysName
                    Exit Sub
                End If
                strBeds = Trim(Split(strBed, "房间:")(0))
            End If
    End Select
    
    If mbytInFun <> 2 Then
        strSql = "zl_病人变动记录_MOVE(" & mlng病人ID & "," & mlng主页ID & "," & _
            "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),'" & strBeds & "'," & _
            "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & strUnitID & ",'" & strMainBed & "')"
            
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        
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
    End If
    
    mstr目标床号 = strBeds
    gblnOK = True
    
    On Error Resume Next
    '换床成功后触发消息
    If mclsMipModule.IsConnect = True And mbytInFun <> 2 Then
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
        mclsXML.appendData "patient_sex", txt姓名.Tag, xsString '性别
        'in_number       住院号  1   S
        mclsXML.appendData "in_number", txt住院号.Text, xsString  '住院号
        mclsXML.AppendNode "in_patient", True
        
        '当前情况
        'current_state       当前情况    1
        mclsXML.AppendNode "current_state"
        'current_area_id     当前病区id  0..1    N
        mclsXML.appendData "current_area_id", Val(Nvl(mrsPatiInfo!当前病区ID)), xsNumber
        'current_area_title      当前病区    0..1    S
        mclsXML.appendData "current_area_title", Nvl(mrsPatiInfo!当前病区), xsString
        'current_dept_id     当前科室id  1   N
        mclsXML.appendData "current_dept_id", Val(Nvl(mrsPatiInfo!出院科室id, 0)), xsNumber
        'current_dept_title      当前科室    1   S
        mclsXML.appendData "current_dept_title", Nvl(mrsPatiInfo!当前科室), xsString
        'current_room        当前病房    0..1    S
        mclsXML.appendData "current_room", txtPre.Tag, xsString
        'current_bed     当前病床    1   S
        mclsXML.appendData "current_bed", Nvl(mrsPatiInfo!主要床号), xsString
        mclsXML.AppendNode "current_state", True
        
        strSql = " Select ID 变动id,开始时间 变动时间 From 病人变动记录 Where 病人ID=[1] And 主页Id=[2] And 开始原因=[3] And 开始时间+0=[4] And NVL(附加床位,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病人变动记录", mlng病人ID, mlng主页ID, 4, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        '转入信息
        'change_state        转入信息    1
        mclsXML.AppendNode "change_state"
        'change_id       转科变更id  1   N
        mclsXML.appendData "change_id", rsTmp!变动ID, xsNumber
        'change_date     变更时间    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!变动时间), "YYYY-MM-DD HH:mm:ss"), xsString
        'change_room     入住病房    0..1    S
        mclsXML.appendData "change_room", strRoom, xsString
        'change_bed      入住病床    1   S
        mclsXML.appendData "change_bed", strMainBed, xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_004", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadMainBed
End Sub

Private Sub LoadMainBed()
    Dim i As Integer, strBed As String
    
    If cboNew.ListIndex <> -1 Then strBed = cboNew.Text
    cboNew.Clear
    For i = 1 To lvw.ListItems.Count
        If lvw.ListItems(i).Checked Then
            cboNew.AddItem lvw.ListItems(i).Text
            If lvw.ListItems(i).Text = strBed Then cboNew.ListIndex = cboNew.NewIndex
            If cboNew.ListIndex = -1 Then
                If lvw.ListItems(i).Text = mrsPatiInfo!主要床号 Then cboNew.ListIndex = cboNew.NewIndex
            End If
        End If
    Next
    If cboNew.ListIndex = -1 And cboNew.ListCount = 1 Then cboNew.ListIndex = 0
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    lvw.ToolTipText = lvw.ListItems(Item.Index).Text
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
        ByVal bytInFun As Byte, ByRef str目标床号 As String, ByVal str床号 As String, ByVal strPrivs As String) As Boolean
'#########################################################################################################
'### 参数：bytInFun :'0-换床,1-包房增减床,2-撤消出院时换新床
'###       str目标床号 :当mbytInFun=0时表示拖动到指定的床位上，=2时表示病人当前床位对应的等级ID
'###       str床号 :当mbytInFun=2时才有值,表示出院前的床位
'### 返回：目标床号
'#########################################################################################################
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mbytInFun = bytInFun
    mstr目标床号 = str目标床号
    mstr床号 = str床号
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    str目标床号 = mstr目标床号
    
    ShowMe = gblnOK
End Function
