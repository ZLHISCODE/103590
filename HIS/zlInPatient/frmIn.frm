VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIn 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人入住"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraGroup 
      Height          =   1335
      Index           =   1
      Left            =   3150
      TabIndex        =   38
      Top             =   1920
      Width           =   3480
      Begin VB.ComboBox cbo责任护士 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   180
         Width           =   1830
      End
      Begin VB.ComboBox cbo主任医师 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   945
         Width           =   1830
      End
      Begin VB.ComboBox cbo门诊医师 
         Height          =   300
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   565
         Width           =   1830
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主任(副主任)医师"
         Height          =   180
         Left            =   75
         TabIndex        =   24
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "责任护士"
         Height          =   180
         Left            =   795
         TabIndex        =   22
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊医师"
         Height          =   180
         Left            =   795
         TabIndex        =   23
         Top             =   625
         Width           =   720
      End
   End
   Begin VB.Frame fraGroup 
      Height          =   1335
      Index           =   0
      Left            =   105
      TabIndex        =   37
      Top             =   1920
      Width           =   2970
      Begin VB.ComboBox cbo主治医师 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   950
         Width           =   1890
      End
      Begin VB.ComboBox cbo医疗小组 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   180
         Width           =   1890
      End
      Begin VB.ComboBox cbo住院医师 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   565
         Width           =   1890
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗小组"
         Height          =   180
         Left            =   210
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主治医师"
         Height          =   180
         Left            =   210
         TabIndex        =   21
         Top             =   1010
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院医师"
         Height          =   180
         Left            =   210
         TabIndex        =   20
         Top             =   625
         Width           =   720
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
      ScaleWidth      =   6705
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5610
      Width           =   6705
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   105
         TabIndex        =   18
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4260
         TabIndex        =   16
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5445
         TabIndex        =   17
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fraLvw 
      Caption         =   "包房病床"
      Height          =   1830
      Left            =   105
      TabIndex        =   28
      Top             =   3720
      Width           =   6525
      Begin MSComctlLib.ListView lvw 
         Height          =   1425
         Left            =   150
         TabIndex        =   25
         Top             =   255
         Width           =   6255
         _ExtentX        =   11033
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
   Begin VB.Frame fraInfo 
      Height          =   1935
      Left            =   105
      TabIndex        =   27
      Top             =   0
      Width           =   6525
      Begin VB.ComboBox cbo床号 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1150
         Width           =   1890
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   765
         Width           =   1890
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4650
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   1170
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   780
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
      End
      Begin VB.ComboBox cbo病况 
         Height          =   300
         Left            =   4650
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   765
         Width           =   1170
      End
      Begin VB.ComboBox cbo护理等级 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1560
         Width           =   5505
      End
      Begin VB.CheckBox chk陪伴 
         Caption         =   "是否陪伴"
         Height          =   195
         Left            =   4650
         TabIndex        =   15
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CheckBox chk包房 
         Caption         =   "包床"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3300
         TabIndex        =   4
         Top             =   1203
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4230
         TabIndex        =   35
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2850
         TabIndex        =   34
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   570
         TabIndex        =   33
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病况"
         Height          =   180
         Left            =   4230
         TabIndex        =   32
         Top             =   825
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理等级"
         Height          =   180
         Left            =   210
         TabIndex        =   31
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   570
         TabIndex        =   30
         Top             =   825
         Width           =   360
      End
      Begin VB.Label lbl床位 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位"
         Height          =   180
         Left            =   570
         TabIndex        =   29
         Top             =   1210
         Width           =   360
      End
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   300
      Left            =   4740
      TabIndex        =   14
      Top             =   3375
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
   Begin MSMask.MaskEdBox txtIn 
      Height          =   300
      Left            =   1065
      TabIndex        =   40
      Top             =   3375
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   16
      Format          =   "yyyy-MM-dd hh:mm"
      Mask            =   "####-##-## ##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblIn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "入院时间"
      Height          =   180
      Left            =   315
      TabIndex        =   39
      Top             =   3435
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "入住时间"
      Height          =   180
      Left            =   3945
      TabIndex        =   36
      Top             =   3435
      Width           =   720
   End
End
Attribute VB_Name = "frmIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mlng病人ID As Long
Public mlng主页ID As Long
Public mlngUnit As Long
Public mbyt入住方式 As Byte '0-入院入住，1-转科入住

Public mstr床号 As String '入:缺省定位的床号,表示家庭病床,出:入住的床号,可能多张床,用,号分隔
Public mlng床位科室ID As Long
Public mstrPrivs As String
Private mfrmParent As Object
Private mblnAppoint As Boolean      'T-预约中心病人;False-非预约中心病人
Private mstrAppointBed As String    '预约中心安排床位
Private mstrIDs As String
Private mstrText As String
Private mrsPatiInfo As ADODB.Recordset
Private mint调整科室 As Integer

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

Private Sub cbo床号_Click()
    cbo.SetListWidth cbo床号.hWnd, cbo床号.width * 1.8
    If mblnAppoint Then
        cbo床号.Tag = Trim(Split(cbo床号.Text, " ")(0))
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

Private Sub cbo科室_Click()
    Dim rsTmp As ADODB.Recordset
    If mstrText = cbo科室.Text Then Exit Sub
    If cbo科室.Text = "" Then Exit Sub
    mstrText = cbo科室.Text
    '显示该科室的床位
    Call ShowBeds
    If Not Visible Then chk包房_Click
    
    On Error GoTo errHandle
    
     '医疗小组
    gstrSQL = "Select ID,名称,说明,建档时间,撤档时间 From 临床医疗小组 Where 科室id=[1] " & _
            " And (撤档时间 Is NULL Or Trunc(撤档时间) = To_Date('3000-01-01','YYYY-MM-DD')) Order By Id "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val("" & cbo科室.ItemData(cbo科室.ListIndex)))
    
    cbo医疗小组.Clear
    Do Until rsTmp.EOF
        cbo医疗小组.AddItem rsTmp!ID & "-" & rsTmp!名称
        cbo医疗小组.ItemData(cbo医疗小组.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    cbo医疗小组.AddItem "": cbo医疗小组.ItemData(cbo医疗小组.NewIndex) = 0: cbo医疗小组.ListIndex = cbo医疗小组.ListCount - 1
    If cbo医疗小组.ListCount = 1 Then cbo医疗小组.Enabled = False
    
    '缺省定位该科室的医生、护士
    
    Call SeekDoctor(cbo责任护士, NVL(mrsPatiInfo!责任护士))
    Call SeekDoctor(cbo门诊医师, NVL(mrsPatiInfo!门诊医师))
    Call SeekDoctor(cbo住院医师, NVL(mrsPatiInfo!住院医师))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo医疗小组_Click()
    Dim strSQL As String, strSQL医疗小组 As String
    Dim rsTmp As ADODB.Recordset
    Dim lng医师 As Long
    
    If cbo医疗小组.ListCount = 1 Then Exit Sub
    On Error GoTo errHandle
    '如果为病人指定了医疗小组，则"住院医师、主治医师"都从对应医疗小组中的医生中选择
    strSQL医疗小组 = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
                        " From 人员表 A, 人员性质说明 B, 部门人员 C, 医疗小组人员 D" & vbNewLine & _
                        " Where A.ID = B.人员id And A.ID = C.人员id And a.id = d.人员id And B.人员性质 = '医生' And d.小组id = [1] And" & vbNewLine & _
                        "   (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                        "   Instr(',' || [2] || ',', ',' || C.部门id || ',') > 0 And Instr(',' || [3] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
                        "   And (A.站点=[4] Or A.站点 is Null)" & vbNewLine & _
                        " Order By A.简码"
    strSQL = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
                        " From 人员表 A, 人员性质说明 B, 部门人员 C" & vbNewLine & _
                        " Where A.ID = B.人员id And A.ID = C.人员id And B.人员性质 = '医生' And" & vbNewLine & _
                        "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                        "      Instr(',' || [1] || ',', ',' || C.部门id || ',') > 0 And Instr(',' || [2] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
                        "      And (A.站点=[3] Or A.站点 is Null)" & _
                        " Order By A.简码"
    
    If cbo医疗小组.ListIndex <> -1 And cbo医疗小组.ListIndex <> cbo医疗小组.ListCount - 1 Then
        If Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)) > 0 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL医疗小组, Me.Caption, Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)), mstrIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
            If Not rsTmp.RecordCount > 0 Then
                '如果小组未设置医生，则保持以前的科内选择范围
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
            End If
            If cbo住院医师.ListIndex <> -1 Then
                lng医师 = cbo住院医师.ItemData(cbo住院医师.ListIndex)
            Else
                lng医师 = 0
            End If
            cbo住院医师.Clear
            Do Until rsTmp.EOF
                cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
            '105133:当住院医师在所选医疗小组时不改变住院医师
            If lng医师 <> 0 Then Call cbo.SetIndex(cbo住院医师.hWnd, cbo.FindIndex(cbo住院医师, lng医师))
        
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL医疗小组, Me.Caption, Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)), mstrIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师", gstrNodeNo)
            
            If Not rsTmp.RecordCount > 0 Then
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师", gstrNodeNo)
            End If
            If cbo主治医师.ListIndex <> -1 Then
                lng医师 = cbo主治医师.ItemData(cbo主治医师.ListIndex)
            Else
                lng医师 = 0
            End If
            cbo主治医师.Clear
            Do Until rsTmp.EOF
                cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
             '105133:当主治医师在所选医疗小组时不改变主治医师
            If lng医师 <> 0 Then Call cbo.SetIndex(cbo主治医师.hWnd, cbo.FindIndex(cbo主治医师, lng医师))
        End If
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
        cbo住院医师.Clear
        Do Until rsTmp.EOF
            cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师", gstrNodeNo)
        cbo主治医师.Clear
        Do Until rsTmp.EOF
            cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
    End If
    
    cbo住院医师.AddItem "其它..."
    cbo主治医师.AddItem "其它..."
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo医疗小组_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo医疗小组.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo医疗小组.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo医疗小组.ListIndex = lngIdx
    Else
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo主任医师_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo主任医师.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo主任医师.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo主任医师.ListIndex = lngIdx
    ElseIf cbo主治医师.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo主治医师_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo主治医师.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo主治医师.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo主治医师.ListIndex = lngIdx
    ElseIf cbo主治医师.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo住院医师_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    If cbo住院医师.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("医生", "主任医师,副主任医师,主治医师,医师,医士", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo住院医师.ListCount - 1
                If cbo住院医师.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo住院医师.ListIndex = i: Exit Sub
                End If
            Next
            cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo住院医师.ListCount - 1
            cbo住院医师.ListIndex = cbo住院医师.NewIndex
            cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!上级ID
        Else
            cbo住院医师.ListIndex = -1
        End If
    Else
        If cbo医疗小组.ListCount <= 1 Or Not Me.Visible Then Exit Sub
        strSQL = "Select ID,名称,说明 From 临床医疗小组 A, 医疗小组人员 B " & _
                "Where a.id=b.小组id And b.人员id=[1] And a.科室id=[2] And (撤档时间 Is NULL Or Trunc(撤档时间)=To_Date('3000-01-01','YYYY-MM-DD')) Order By ID"
        If cbo住院医师.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo住院医师.ItemData(cbo住院医师.ListIndex)), Val(cbo科室.ItemData(cbo科室.ListIndex)))
            Do While Not rsTmp.EOF
                If cbo医疗小组.Text = NVL(rsTmp!ID) & "-" & NVL(rsTmp!名称) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo.FindIndex(cbo医疗小组, NVL(rsTmp!名称), True))
                Exit Sub
            End If
        End If
        If cbo主治医师.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo主治医师.ItemData(cbo主治医师.ListIndex)), Val(cbo科室.ItemData(cbo科室.ListIndex)))
            Do While Not rsTmp.EOF
                If cbo医疗小组.Text = NVL(rsTmp!ID) & "-" & NVL(rsTmp!名称) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo.FindIndex(cbo医疗小组, NVL(rsTmp!名称), True))
            Else
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo医疗小组.ListCount - 1)
            End If
        Else
            Call cbo.SetIndex(cbo医疗小组.hWnd, cbo医疗小组.ListCount - 1)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo主治医师_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim strSQL As String
    On Error GoTo errHandle
    
    If cbo主治医师.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("医生", "主任医师,副主任医师,主治医师", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo主治医师.ListCount - 1
                If cbo主治医师.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo主治医师.ListIndex = i: Exit Sub
                End If
            Next
            cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo主治医师.ListCount - 1
            cbo主治医师.ListIndex = cbo主治医师.NewIndex
            cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
        Else
            cbo主治医师.ListIndex = -1
        End If
    Else
        '主治医师选择时医疗小组以住院医师为先
        If cbo医疗小组.ListCount <= 1 Or Not Me.Visible Then Exit Sub
        strSQL = "Select ID,名称,说明 From 临床医疗小组 A, 医疗小组人员 B " & _
                "Where a.id=b.小组id And b.人员id=[1] And a.科室id=[2] And (撤档时间 Is NULL Or Trunc(撤档时间)=To_Date('3000-01-01','YYYY-MM-DD')) Order By ID"
        If cbo住院医师.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo住院医师.ItemData(cbo住院医师.ListIndex)), Val(cbo科室.ItemData(cbo科室.ListIndex)))
            Do While Not rsTmp.EOF
                If cbo医疗小组.Text = NVL(rsTmp!ID) & "-" & NVL(rsTmp!名称) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo.FindIndex(cbo医疗小组, NVL(rsTmp!名称), True))
                Exit Sub
            End If
        End If
        If cbo主治医师.ListIndex <> -1 Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo主治医师.ItemData(cbo主治医师.ListIndex)), Val(cbo科室.ItemData(cbo科室.ListIndex)))
            Do While Not rsTmp.EOF
                If cbo医疗小组.Text = NVL(rsTmp!ID) & "-" & NVL(rsTmp!名称) Then Exit Sub
            rsTmp.MoveNext
            Loop
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo.FindIndex(cbo医疗小组, NVL(rsTmp!名称), True))
            Else
                Call cbo.SetIndex(cbo医疗小组.hWnd, cbo医疗小组.ListCount - 1)
            End If
        Else
            Call cbo.SetIndex(cbo医疗小组.hWnd, cbo医疗小组.ListCount - 1)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo主任医师_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo主任医师.Text = "其它..." Then
        Set rsTmp = GetSelectPersonal("医生", "主任医师,副主任医师", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo主任医师.ListCount - 1
                If cbo主任医师.List(i) = rsTmp!简码 & "-" & rsTmp!名称 Then
                    cbo主任医师.ListIndex = i: Exit Sub
                End If
            Next
            cbo主任医师.AddItem rsTmp!简码 & "-" & rsTmp!名称, cbo主任医师.ListCount - 1
            cbo主任医师.ListIndex = cbo主任医师.NewIndex
            cbo主任医师.ItemData(cbo主任医师.NewIndex) = rsTmp!ID
        Else
            cbo主任医师.ListIndex = -1
        End If
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


Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If cbo科室.Locked Then Exit Sub
        If SendMessage(cbo科室.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo科室.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo科室.ListIndex = lngIdx
    ElseIf cbo科室.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo门诊医师_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo门诊医师.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo门诊医师.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo门诊医师.ListIndex = lngIdx
    ElseIf cbo门诊医师.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
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

Private Sub cbo住院医师_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo住院医师.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo住院医师.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo住院医师.ListIndex = lngIdx
    ElseIf cbo住院医师.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chk包房_Click()
    If chk包房.Value = 1 Then
        lbl床位.Caption = "主要床位"
        Call LoadMainBed
        lvw.Visible = True
        Me.Height = Me.Height + fraLvw.Height + 80
        If Visible Then lvw.SetFocus
    Else
        lbl床位.Caption = "床位"
        Call ShowBeds
        lvw.Visible = False
        Me.Height = Me.Height - fraLvw.Height - 80
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
    Dim strSQL As String, i As Integer
    Dim strSQL医疗小组 As String
    Dim strIDs As String, strID As String, strCode As String
    Dim strTmp As String
    Dim blnNurseGrade As Boolean    '护理等级默认为空 ?
    Dim strInfo As String, blnHeav As Boolean
    
    On Error GoTo errH
    gblnOK = False
    
    '50194:刘鹏飞,2012-09-21,转科接收入住时检查：
    '如果已进行病案首页的住院医生签名，或主治医生签名，或主任医生签名，则禁止接收入住，并且提示应该先由该医生取消签名再进行。
    If mbyt入住方式 <> 0 Then '转科入住
        '获取首页已经签名最高级别
        strInfo = "该病人的首页已经由以下医生进行了签名："
        blnHeav = False
        strSQL = "Select 信息名,信息值 From 病案主页从表 Where 病人ID=[1] And 主页ID=[2] And 信息值 is Not Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        rsTmp.Filter = "信息名='住院医师签名'"
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!信息值) Then
                strInfo = strInfo & vbCrLf & "住院医师签名【" & NVL(rsTmp!信息值) & "】"
                blnHeav = True
            End If
        End If
        rsTmp.Filter = "信息名='主治医师签名'"
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!信息值) Then
                strInfo = strInfo & vbCrLf & "主治医师签名【" & NVL(rsTmp!信息值) & "】"
                blnHeav = True
            End If
        End If
        rsTmp.Filter = "信息名='主任医师签名'"
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!信息值) Then
                strInfo = strInfo & vbCrLf & "主任医师签名【" & NVL(rsTmp!信息值) & "】"
                blnHeav = True
            End If
        End If
        strInfo = strInfo & vbCrLf & "请先由上述医生取消首页签名在进行转科入住操作！"
        strSQL = ""
        
        If blnHeav = True Then
            MsgBox strInfo, vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    Set mrsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID, mbyt入住方式)
    '问题28432 by lesfeng 2010-03-10
    mint调整科室 = Val(zlDatabase.GetPara("允许调整科室", glngSys, glngModul, 0))
    '初始化数据
    With mrsPatiInfo
        cbo科室.Enabled = mlng床位科室ID = 0
        If mint调整科室 = 0 And cbo科室.Enabled = True Then
            cbo科室.Enabled = False
        End If
        If mbyt入住方式 = 0 Then
            mstrAppointBed = ""
            mblnAppoint = IsAppointPati(Val(!挂号ID & ""), mstrAppointBed) 'T-预约中心病人
        End If
        '所选病床的科室与病人科室不同时,不允许换科.
        If mlng床位科室ID <> 0 Then
            If mbyt入住方式 = 0 Then      '新入病人
                If mlng床位科室ID <> !出院科室id Then
                    '问题28432 by lesfeng 2010-03-10
                    If mint调整科室 = 1 Then
                        cbo科室.Enabled = True
                    Else
                        MsgBox "病人登记的科室【" & !当前科室 & "】与选择的床位所属科室【" & GetDeptName(mlng床位科室ID) & "】不同,不能入住该床位,请选择其它床位!", vbInformation, gstrSysName
                        Unload Me: Exit Sub
                    End If
                End If
            Else                           '转科病人
                If mlng床位科室ID <> !入住科室id Then
                    MsgBox "病人转入的科室【" & !当前科室 & "】与选择的床位所属科室【" & GetDeptName(mlng床位科室ID) & "】不同,不能入住该床位,请选择其它床位!", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
        End If
        
        If mbyt入住方式 = 0 And gbyt入科时间 = 0 Then
            txtDate.Text = Format(!入院时间, "yyyy-MM-dd HH:mm:ss")
        Else
            txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
        End If
        
        If mbyt入住方式 = 0 And Val(zlDatabase.GetPara("允许修改入院时间", glngSys, 1132)) = 1 Then
            lblIn.Visible = True
            txtIn.Visible = True
            txtIn.Text = Format(!入院时间, "yyyy-MM-dd HH:mm")
        Else
            lblIn.Visible = False
            txtIn.Visible = False
        End If
        
        '病人信息
        txt姓名.Text = !姓名
        txt性别.Text = "" & !性别
        txt年龄.Text = "" & !年龄
                
        
        '确定病区的服务对象
        strSQL = "Select 服务对象 From 部门性质说明 Where 工作性质='护理' And 部门ID=[1]" '
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUnit)
            
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
                cbo科室.AddItem rsTmp!编码 & "-" & rsTmp!名称
                cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
                If mlng床位科室ID = 0 Then '是不是拖到指定的床上
                    If mbyt入住方式 = 0 Then '新入病人缺省取登记科室
                        If rsTmp!ID = !出院科室id Then cbo科室.ListIndex = cbo科室.NewIndex     '调用click事件加载床位
                    Else
                        '转科病人缺省取转入科室
                        If rsTmp!ID = !入住科室id Then cbo科室.ListIndex = cbo科室.NewIndex
                    End If
                Else
                    '入住病区病床的病人科室已由床位定了
                    If rsTmp!ID = mlng床位科室ID Then cbo科室.ListIndex = cbo科室.NewIndex
                End If
                strIDs = strIDs & "," & rsTmp!ID
                rsTmp.MoveNext
            Next
        Else
            '没有对应的床位科室
            MsgBox "在当前病区没有设置对应科室,病人不能入住！" & vbCrLf, vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        mstrIDs = strIDs
        '预约床位被占用
        If mbyt入住方式 = 0 And mblnAppoint And mstrAppointBed <> cbo床号.Tag Then
            If MsgBox("在当前可用床位中没有找到病人预约的床位【" & mstrAppointBed & "】，是否继续入住？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Unload Me: Exit Sub
            End If
        End If
        '病况
        strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 病情 Order by 编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo病况.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 And cbo病况.ListIndex = -1 Then cbo病况.ListIndex = cbo病况.NewIndex
                If rsTmp!名称 = "" & !当前病况 Then cbo病况.ListIndex = cbo病况.NewIndex
                rsTmp.MoveNext
            Next
        End If
    
        '护理等级
        If mbyt入住方式 = 1 Then cbo护理等级.Enabled = InStr(mstrPrivs, ";" & "调整护理等级" & ";") > 0
        Set rsTmp = GetNurseGrade
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo护理等级.AddItem rsTmp!编码 & "-" & rsTmp!名称
                cbo护理等级.ItemData(cbo护理等级.NewIndex) = rsTmp!ID
                If rsTmp!ID = !护理等级ID Then cbo护理等级.ListIndex = cbo护理等级.NewIndex
                rsTmp.MoveNext
            Next
        End If
        
        blnNurseGrade = zlDatabase.GetPara("护理等级默认为空", glngSys, 1132, 0)
        If blnNurseGrade And mbyt入住方式 = 1 Then cbo护理等级.ListIndex = -1
        
        cbo医疗小组.Clear
        If Not cbo科室.ListIndex = -1 Then
            '医疗小组
            strSQL = "Select ID,名称,说明,建档时间,撤档时间 From 临床医疗小组 Where 科室id=[1] " & _
                    " And (撤档时间 Is NULL Or Trunc(撤档时间) = To_Date('3000-01-01','YYYY-MM-DD')) Order By Id "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val("" & cbo科室.ItemData(cbo科室.ListIndex)))
            
            cbo医疗小组.Clear
            Do Until rsTmp.EOF
                cbo医疗小组.AddItem rsTmp!ID & "-" & rsTmp!名称
                cbo医疗小组.ItemData(cbo医疗小组.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Loop
        End If
        cbo医疗小组.AddItem "": cbo医疗小组.ItemData(cbo医疗小组.NewIndex) = 0: cbo医疗小组.ListIndex = cbo医疗小组.ListCount - 1
        If cbo医疗小组.ListCount = 1 Then cbo医疗小组.Enabled = False
        'by lesfeng 2010-01-12 性能优化
        '住院医师,主治医师,主任医师
        '如果为病人指定了医疗小组，则"住院医师、主治医师"都从对应医疗小组中的医生中选择
        strSQL医疗小组 = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
                        " From 人员表 A, 人员性质说明 B, 部门人员 C, 医疗小组人员 D" & vbNewLine & _
                        " Where A.ID = B.人员id And A.ID = C.人员id And a.id = d.人员id And B.人员性质 = '医生' And d.小组id = [1] And" & vbNewLine & _
                        "   (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                        "   Instr(',' || [2] || ',', ',' || C.部门id || ',') > 0 And Instr(',' || [3] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
                        "   And (A.站点=[4] Or A.站点 is Null)" & vbNewLine & _
                        " Order By A.简码"
        strSQL = "Select Distinct A.ID, A.编号, A.简码, A.姓名" & vbNewLine & _
                            " From 人员表 A, 人员性质说明 B, 部门人员 C" & vbNewLine & _
                            " Where A.ID = B.人员id And A.ID = C.人员id And B.人员性质 = '医生' And" & vbNewLine & _
                            "      (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & vbNewLine & _
                            "      Instr(',' || [1] || ',', ',' || C.部门id || ',') > 0 And Instr(',' || [2] || ',', ',' || A.专业技术职务 || ',') > 0" & vbNewLine & _
                            "      And (A.站点=[3] Or A.站点 is Null)" & _
                            " Order By A.简码"
        If cbo医疗小组.ListCount = 1 Then
            If cbo医疗小组.ListIndex <> -1 And cbo医疗小组.ListIndex <> cbo医疗小组.ListCount - 1 Then
                If Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)) > 0 Then
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL医疗小组, Me.Caption, Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)), strIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
                    If Not rsTmp.RecordCount > 0 Then
                        '如果小组未设置医生，则保持以前的科内选择范围
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
                    End If
                    cbo住院医师.Clear
                    Do Until rsTmp.EOF
                        cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                        cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!ID
                        rsTmp.MoveNext
                    Loop
    
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL医疗小组, Me.Caption, Val(cbo医疗小组.ItemData(cbo医疗小组.ListIndex)), strIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师", gstrNodeNo)
    
                    If Not rsTmp.RecordCount > 0 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师", gstrNodeNo)
                    End If
                    cbo主治医师.Clear
                    Do Until rsTmp.EOF
                        cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                        cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
                        rsTmp.MoveNext
                    Loop
                End If
            Else
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师,医师,医士", gstrNodeNo)
                cbo住院医师.Clear
                Do Until rsTmp.EOF
                    cbo住院医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                    cbo住院医师.ItemData(cbo住院医师.NewIndex) = rsTmp!ID
                    rsTmp.MoveNext
                Loop
    
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "主任医师,副主任医师,主治医师", gstrNodeNo)
                cbo主治医师.Clear
                Do Until rsTmp.EOF
                    cbo主治医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                    cbo主治医师.ItemData(cbo主治医师.NewIndex) = rsTmp!ID
                    rsTmp.MoveNext
                Loop
            End If
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs & "," & mlngUnit, "主任医师,副主任医师", gstrNodeNo)
        Do Until rsTmp.EOF
            cbo主任医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            cbo主任医师.ItemData(cbo主任医师.NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        
        '转科入住
        If mbyt入住方式 = 1 Then
            If Not cbo.Locate(cbo住院医师, "" & !住院医师) Then
                Call GetPersonnelIDCode("" & !住院医师, strID, strCode)
                cbo住院医师.AddItem strCode & "-" & !住院医师
                cbo住院医师.ItemData(cbo住院医师.NewIndex) = Val(strID)
                cbo住院医师.ListIndex = cbo住院医师.NewIndex
                strID = "": strCode = ""
            End If
            
            strSQL = " Select 信息名,信息值 From 病案主页从表 Where (信息名='主治医师' Or 信息名='主任医师') And 病人ID=[1] And 主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
            
            rsTmp.Filter = "信息名='主治医师'"
            If Not rsTmp.EOF Then
                If Not cbo.Locate(cbo主治医师, "" & rsTmp!信息值) Then
                    Call GetPersonnelIDCode("" & rsTmp!信息值, strID, strCode)
                    cbo主治医师.AddItem strCode & "-" & rsTmp!信息值
                    cbo主治医师.ItemData(cbo主治医师.NewIndex) = Val(strID)
                    cbo主治医师.ListIndex = cbo主治医师.NewIndex
                    strID = "": strCode = ""
                End If
            End If
            
            rsTmp.Filter = "信息名='主任医师'"
            If Not rsTmp.EOF Then
                If Not cbo.Locate(cbo主任医师, "" & rsTmp!信息值) Then
                    Call GetPersonnelIDCode("" & rsTmp!信息值, strID, strCode)
                    cbo主任医师.AddItem strCode & "-" & rsTmp!信息值
                    cbo主任医师.ItemData(cbo主任医师.NewIndex) = Val(strID)
                    cbo主任医师.ListIndex = cbo主任医师.NewIndex
                    strID = "": strCode = ""
                End If
            End If
        '入住
        Else
            Call SeekDoctor(cbo住院医师, "" & !住院医师)
            Call SeekDoctor(cbo主治医师, "" & !住院医师)
            '主任医师,一般无法确定缺省
        End If
    
        '门诊医师(所有)
        strSQL = "SELECT DISTINCT a.Id, a.编号, a.简码, a.姓名 " & vbNewLine & _
                " FROM 人员表 a, 人员性质说明 b, 部门人员 c, 部门性质说明 d " & vbNewLine & _
                " WHERE a.Id = b.人员id AND a.Id = c.人员id AND c.部门id = d.部门id AND b.人员性质 = '医生' AND d.服务对象 IN (1, 2, 3) AND " & vbNewLine & _
                "      (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') OR a.撤档时间 IS NULL) " & vbNewLine & _
                " ORDER BY 简码 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo门诊医师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                cbo门诊医师.ItemData(cbo门诊医师.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Next
            Call SeekDoctor(cbo门诊医师, "" & !门诊医师)
        End If
    
        '住院护士
        Set rsTmp = GetDoctorOrNurse(1, strIDs & "," & mlngUnit & ",")
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo责任护士.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                cbo责任护士.ItemData(i - 1) = rsTmp!ID
                rsTmp.MoveNext
            Next
            Call SeekDoctor(cbo责任护士, "" & !责任护士)
        End If
        
        cbo主任医师.AddItem "其它..."
        cbo责任护士.AddItem "其它..."
        If InStr(mstrPrivs, "调整门诊医师") = 0 Then
            cbo门诊医师.Enabled = False
        End If
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
    Dim strBed  As String
    
    lvw.ListItems.Clear
    cbo床号.Clear: cbo床号.Tag = ""
    If InStr(1, mstrPrivs, "家庭病床") > 0 Then
        cbo床号.AddItem "家庭病床"
        If mstr床号 = "家庭病床" Then cbo床号.ListIndex = 0
    End If
    If cbo科室.ListIndex <> -1 Then lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    Set rsBeds = GetFreeBeds(mlngUnit, lng科室ID, mrsPatiInfo!性别, mlng病人ID)
    If mstrAppointBed <> "" Then
        strBed = mstrAppointBed
    Else
        strBed = mstr床号
    End If
    With rsBeds
        For i = 1 To rsBeds.RecordCount
            Set objItem = lvw.ListItems.Add(, "_" & !床号, !床号 & IIf(IsNull(!房间号), "", " 房间:" & !房间号 & "|") & _
                            IIf(IsNull(!房间号) Or ((Not IsNull(!房间号)) And Trim(NVL(!性别) = "")), "", "(" & NVL(!性别) & ")"))
            objItem.Tag = !等级ID
            cbo床号.AddItem objItem.Text
            If !床号 = strBed Then
                objItem.Checked = True: objItem.Selected = True: objItem.EnsureVisible
                cbo床号.ListIndex = cbo床号.NewIndex
                cbo床号.Tag = !床号
            End If
            .MoveNext
        Next
    End With
    
    If cbo床号.ListIndex = -1 And cbo床号.ListCount > 0 Then cbo床号.ListIndex = 0
End Sub

Private Sub SeekDoctor(cbo As ComboBox, Optional strPre As String)
    Dim strIDs As String, i As Integer
    
    If cbo科室.ListIndex = -1 Then Exit Sub
    
    If strPre <> "" Then
        For i = 0 To cbo.ListCount - 1
            If zlCommFun.GetNeedName(cbo.List(i)) = strPre Then cbo.ListIndex = i: Exit Sub
        Next
    End If
    
    strIDs = GetDeptDoctors(cbo科室.ItemData(cbo科室.ListIndex))
    For i = 0 To cbo.ListCount - 1
        If InStr("," & strIDs & ",", "," & cbo.ItemData(i) & ",") > 0 Then cbo.ListIndex = i: Exit Sub
    Next
    
    strIDs = GetDeptDoctors(mlngUnit)
    For i = 0 To cbo.ListCount - 1
        If InStr("," & strIDs & ",", "," & cbo.ItemData(i) & ",") > 0 Then cbo.ListIndex = i: Exit Sub
    Next
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim dMax As Date, i As Integer, Curdate As Date
    Dim strPreRoom As String, intRoom As Integer, intCheck As Integer, lngNurseGrade As Long
    Dim strSQL As String, strBed As String, strTmp As String, strNewSql As String
    Dim str床号 As String, blnTrans As Boolean, strMainBed As String
    Dim rsTmp As ADODB.Recordset
    Dim str房间号 As String
    Dim blnTrue As Boolean
    Dim int母婴转科标志 As Integer
    
    If cbo科室.ListIndex = -1 Then
        MsgBox "请确定病人要入住的科室！", vbInformation, gstrSysName
        cbo科室.SetFocus: Exit Sub
    End If
    
    If cbo病况.ListIndex = -1 Then
        MsgBox "请指定病人的当前病况！", vbInformation, gstrSysName
        cbo病况.SetFocus: Exit Sub
    End If
    
    If cbo护理等级.ListIndex = -1 And gbln入科确定护理等级 Then
        MsgBox "请指定病人的当前护理等级！", vbInformation, gstrSysName
        cbo护理等级.SetFocus: Exit Sub
    End If
    
    '72433:刘鹏飞,2014-08-02
    blnTrue = (Val(zlDatabase.GetPara("入住指定医疗小组", glngSys, glngModul, 0)) = 1) And cbo医疗小组.Enabled
    If cbo医疗小组.ItemData(cbo医疗小组.ListIndex) = 0 And blnTrue = True Then
        MsgBox "请指定病人的当前医疗小组！", vbInformation, gstrSysName
        If cbo医疗小组.Enabled And cbo医疗小组.Visible Then cbo医疗小组.SetFocus
        Exit Sub
    End If
    blnTrue = False
    
    If cbo护理等级.ListIndex <> -1 Then
        lngNurseGrade = cbo护理等级.ItemData(cbo护理等级.ListIndex)
    End If

    '78877:出生时间不能大于入院时间
    If txtIn.Enabled And txtIn.Visible Then
        If CDate(mrsPatiInfo!出生日期 & "") > CDate(txtIn.Text) Then
            MsgBox "病人的入院时间[" & Format(txtIn.Text, "YYYY-MM-DD HH:MM:SS") & "]必须大于病人的出生日期[" & mrsPatiInfo!出生日期 & "]！", vbInformation, gstrSysName
            txtIn.SetFocus
            Exit Sub
        End If
    End If

    '时间不能超过当前时间太长(一个月)
    Curdate = zlDatabase.Currentdate
    If InStr(Trim(cbo床号.Text), " 房间") <> 0 Then
        str床号 = Mid(Trim(cbo床号.Text), 1, InStr(Trim(cbo床号.Text), " 房间") - 1)
        
        If InStr(Trim(cbo床号.Text), "|") - InStr(Trim(cbo床号.Text), "房间:") - 3 > 0 Then
            str房间号 = Mid(Trim(cbo床号.Text), InStr(Trim(cbo床号.Text), "房间:") + 3, InStr(Trim(cbo床号.Text), "|") - InStr(Trim(cbo床号.Text), "房间:") - 3)
        End If
        strSQL = "Select 性别 From 病人信息 A,床位状况记录 B  Where A.病人ID = b.病人id And b.病人ID Is Not Null And 病区ID = [1] And 房间号 =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUnit, str房间号)
        
        Do While Not rsTmp.EOF
            If Trim(txt性别.Text) <> rsTmp!性别 Then
                If (MsgBox("指定床位所在房间存在男女混住情况，是否继续入住？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                    Exit Do
                Else
                    Exit Sub
                    cbo床号.SetFocus
                End If
            End If
            rsTmp.MoveNext
        Loop
    ElseIf InStr(cbo床号.Text, "家庭病床") > 0 Then
        str床号 = ""
    Else
        str床号 = Trim(cbo床号.Text)
    End If
    
    If CDate(txtDate.Text) > Curdate Then
        MsgBox "入住时间大于了当前系统时间,请检查！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    If mbyt入住方式 <> 1 And txtIn.Visible Then
        If IsDate(txtIn.Text) Then
            If CDate(txtIn.Text) > Curdate Then
                MsgBox "入院时间大于了当前系统时间，请检查！", vbInformation, gstrSysName
                txtIn.SetFocus: Exit Sub
            End If
            If CDate(txtIn.Text) > CDate(txtDate.Text) Then
                MsgBox "入住时间不能小于入院时间，请检查！", vbInformation, gstrSysName
                txtIn.SetFocus: Exit Sub
            End If
        Else
            MsgBox "入院时间输入错误，请检查！", vbInformation, gstrSysName
            txtIn.SetFocus: Exit Sub
        End If
    End If
    
    If mbyt入住方式 = 0 Then
        If Format(txtDate.Text, "yyyyMMddhhmmss") < Format(mrsPatiInfo!入院时间, "yyyyMMddHHmmss") Then
            MsgBox "入住时间不能小于该病人的入院时间[" & Format(mrsPatiInfo!入院时间, "yyyy-MM-dd HH:mm:ss") & "]！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    '可能与入院时间相同
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)
    If mbyt入住方式 = 1 Then
        If Format(txtDate.Text, "yyyyMMddhhmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
            MsgBox "入住时间必须大于该病人的上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    Else
        If Format(txtDate.Text, "yyyyMMddhhmmss") < Format(dMax, "yyyyMMddHHmmss") Then
            MsgBox "入住时间不能小于该病人的上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
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
    
    
    If cbo科室.ItemData(cbo科室.ListIndex) <> mrsPatiInfo!入住科室id Then
        '问题28432 by lesfeng 2010-03-10
        If mint调整科室 = 0 And mbyt入住方式 = 0 Then
            MsgBox "当前选择的科室【" & zlCommFun.GetNeedName(cbo科室.Text) & "】不是病人原先登记的科室【" & mrsPatiInfo!当前科室 & "】，不能操作！", vbInformation, gstrSysName
            If cbo科室.Enabled Then cbo科室.SetFocus
            Exit Sub
        Else
            If MsgBox("当前选择的科室【" & zlCommFun.GetNeedName(cbo科室.Text) & "】不是病人原先登记的科室【" & mrsPatiInfo!当前科室 & "】,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
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
    On Error GoTo errH
    int母婴转科标志 = 1
    '先禁用母婴分离，等床位信息分析完成后再启用
    If mbyt入住方式 <> 0 And 1 = 0 Then
        strSQL = "Select Count(1) as 婴儿数 From 病人新生儿记录 Where 病人id = [1] And 主页id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If rsTmp!婴儿数 > 0 Then
            '有婴儿的转科时提示是否要转出
            strSQL = "Select 婴儿科室id, 婴儿病区id From 病案主页 Where 病人id = [1] And 主页id = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
            If rsTmp!婴儿科室ID & "" = "" Then
                '为null表示母婴还未分离
                If MsgBox("当前病人有新生儿记录，新生儿是否一起入住？", vbQuestion + vbDefaultButton1 + vbYesNo) = vbYes Then
                    int母婴转科标志 = 1
                Else
                    int母婴转科标志 = 0
                End If
            End If
        End If
    End If
    
    strSQL = "zl_病人变动记录_InDept(" & mlng病人ID & "," & mlng主页ID & ",'" & strBed & "'," & _
            mlngUnit & "," & cbo科室.ItemData(cbo科室.ListIndex) & "," & cbo医疗小组.ItemData(cbo医疗小组.ListIndex) & "," & _
            lngNurseGrade & ",'" & zlCommFun.GetNeedName(cbo病况.Text) & "'," & _
            "'" & zlCommFun.GetNeedName(cbo责任护士.Text) & "','" & zlCommFun.GetNeedName(cbo门诊医师.Text) & "'," & _
            "'" & zlCommFun.GetNeedName(cbo住院医师.Text) & "'," & chk陪伴.Value & "," & _
            "To_Date('" & IIf(txtIn.Text = "____-__-__ __:__", mrsPatiInfo!入院时间, txtIn.Text) & "','YYYY-MM-DD HH24:MI:SS')," & _
            "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(mbyt入住方式 = 0, 1, 0) & ",'" & _
            zlCommFun.GetNeedName(cbo主治医师.Text) & "','" & zlCommFun.GetNeedName(cbo主任医师.Text) & "','" & strMainBed & "','" & int母婴转科标志 & "')"
    
    
    
    
    blnTrue = False
    strNewSql = " Select Count(*) 记录" & vbNewLine & _
        "  From 住院费用记录" & vbNewLine & _
        "  Where 病人id = [1] And 主页id = [2] And 记录性质 = 3 And 记录状态 = 1 And 附加标志 = 8"
    Set rsTmp = zlDatabase.OpenSQLRecord(strNewSql, "检查病人是否计算过计算一次的一次费用", mlng病人ID, mlng主页ID)
    blnTrue = (Val(NVL(rsTmp!记录, 0)) > 0)
    
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '36454,刘鹏飞,2012-09-06,没有计算过计算的一次项目入住时进行计算
    If mbyt入住方式 <> 1 And blnTrue = False Then
         '如果入院登记时没有确定病区,此时mbyt入住方式才确定病区,需要计算入院一次费用
         '过程中会自动判断是否已计算过(附加标志=8,记录性质=3)
        strSQL = "ZL_住院一次费用_Insert(" & mlng病人ID & "," & mlng主页ID & ")"

        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
     
    If Val("" & mrsPatiInfo!险类) <> 0 Then
        If Not gclsInsure.ModiPatiSwap(mlng病人ID, mlng主页ID, Val("" & mrsPatiInfo!险类), "1") Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    '新网96847
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
    '入科成功还是触发消息
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
        mclsXML.appendData "in_number", NVL(mrsPatiInfo!住院号), xsString '住院号
        mclsXML.AppendNode "in_patient", True
        If mbyt入住方式 = 0 Then '正常入科
            strSQL = " Select A.ID,B.名称 床位等级,C.名称 病区名称  From  病人变动记录 A,收费项目目录 B,部门表 C" & _
                " Where NVl(A.附加床位,0)=0 And A.床位等级id=B.id(+) And A.病区Id=C.id(+) And A.病人ID=[1] And A.主页ID=[2] And A.开始原因=[3]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, 2)
            
            '住院信息
            mclsXML.AppendNode "in_hospital"
            'in_date     入院时间    1   s
            mclsXML.appendData "in_date", Format(IIf(txtIn.Text = "____-__-__ __:__", mrsPatiInfo!入院时间, txtIn.Text), "yyyy-MM-dd HH:mm:ss"), xsString '入院日期
            'in_area_id      入院病区id  0..1    N
            mclsXML.appendData "in_area_id", mlngUnit, xsNumber '入院病区ID
            'in_area_title       入院病区    0..1    S
            mclsXML.appendData "in_area_title", NVL(rsTmp!病区名称), xsString  '入院病区
            'in_dept_id      入院科室id  1   N
            mclsXML.appendData "in_dept_id", cbo科室.ItemData(cbo科室.ListIndex), xsNumber '入院科室id
            'in_dept_title       入院科室    1   S
            mclsXML.appendData "in_dept_title", zlCommFun.GetNeedName(cbo科室.Text), xsString  '入院科室
            mclsXML.appendData "in_again", Val(NVL(mrsPatiInfo!再入院, 0)), xsNumber
            mclsXML.AppendNode "in_hospital", True
            '入住情况
            mclsXML.AppendNode "dept_arrange"
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
            mclsXML.appendData "in_bedgrade", NVL(rsTmp!床位等级), xsString
            'in_doctor       住院医师    0..1    S
            mclsXML.appendData "in_doctor", zlCommFun.GetNeedName(cbo住院医师.Text), xsString
            'duty_nurse      责任护士    0..1    S
            mclsXML.appendData "duty_nurse", zlCommFun.GetNeedName(cbo责任护士.Text), xsString
            mclsXML.AppendNode "dept_arrange", True
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_002", mclsXML.XmlText
        Else '转科入科
            strSQL = " Select A.ID,B.名称 床位等级,C.名称 病区名称  From  病人变动记录 A,收费项目目录 B,部门表 C" & _
                " Where NVl(A.附加床位,0)=0 And A.床位等级id=B.id(+) And A.病区Id=C.id(+) And A.病人ID=[1] And A.主页ID=[2] And A.开始原因=[3] And 开始时间+0=[4]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人变动记录", mlng病人ID, mlng主页ID, 3, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
            
            '住院信息
            mclsXML.AppendNode "in_hospital"
            'in_date     入院时间    1   s
            mclsXML.appendData "in_date", Format(mrsPatiInfo!入院时间, "yyyy-MM-dd HH:mm:ss"), xsString
            'out_area_id     转出病区id  0..1    N
            mclsXML.appendData "out_area_id", Val(NVL(mrsPatiInfo!当前病区ID)), xsNumber
            'out_area_title      转出病区    0..1    S
            mclsXML.appendData "out_area_title", NVL(mrsPatiInfo!当前病区), xsString
            'out_dept_id     转出科室id  1   N
            mclsXML.appendData "out_dept_id", Val(NVL(mrsPatiInfo!出院科室id, 0)), xsNumber
            'out_dept_title      转出科室    1   S
            mclsXML.appendData "out_dept_title", NVL(mrsPatiInfo!当前科室), xsString
            'in_area_id      转入病区id  0..1    N
            mclsXML.appendData "in_area_id", mlngUnit, xsNumber
            'in_area_title       转入病区    0..1    S
            mclsXML.appendData "in_area_title", NVL(rsTmp!病区名称), xsString
            'in_dept_id      转入科室id  1   N
            mclsXML.appendData "in_dept_id", cbo科室.ItemData(cbo科室.ListIndex), xsNumber
            'in_dept_title       转入科室    1   S
            mclsXML.appendData "in_dept_title", zlCommFun.GetNeedName(cbo科室.Text), xsString
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
            mclsXML.appendData "in_bedgrade", NVL(rsTmp!床位等级), xsString
            'in_doctor       住院医师    0..1    S
            mclsXML.appendData "in_doctor", zlCommFun.GetNeedName(cbo住院医师.Text), xsString
            'duty_nurse      责任护士    0..1    S
            mclsXML.appendData "duty_nurse", zlCommFun.GetNeedName(cbo责任护士.Text), xsString
            'change_operator         操作员      1   S
            mclsXML.appendData "change_operator", UserInfo.姓名, xsString
            mclsXML.AppendNode "change_dept_arrange", True
            mclsMipModule.CommitMessage "ZLHIS_PATIENT_012", mclsXML.XmlText
        End If
    End If
    If Err <> 0 Then Err.Clear
    
    '调用外挂接口
    If CreatePlugInOK(glngModul) Then
        On Error Resume Next
        Call gobjPlugIn.InPatiCheckInBranchAfter(mlng病人ID, mlng主页ID)
        Call zlPlugInErrH(Err, "InPatiCheckInBranchAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errH
    '49854:刘鹏飞,2013-10-31,病人腕带打印
    '只有新入住的病人才打印腕带
    If InStr(mstrPrivs, ";腕带打印;") And mbyt入住方式 <> 1 Then
        blnTrue = True
        If gbytCourseWristletPrint = 0 Then
            blnTrue = False
        Else
            If gbytCourseWristletPrint = 2 Then
                If MsgBox("是否打印病人腕带？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnTrue = False
                End If
            End If
        End If
        
        If blnTrue = True Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1132_4", Me, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, 2)
        End If
    End If
    
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
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 名称 From 部门表 Where ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If Not rsTmp.EOF Then
        GetDeptName = IIf(IsNull(rsTmp!名称), "", rsTmp!名称)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
        ByRef str床号 As String, ByVal lng床位科室ID As Long, ByVal byt入住方式 As Byte, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstr床号 = str床号
    mlng床位科室ID = lng床位科室ID
    mbyt入住方式 = byt入住方式
    mstrPrivs = strPrivs
    mstrAppointBed = ""
    mblnAppoint = False
    
    Me.Show 1, frmParent
    str床号 = mstr床号
    ShowMe = gblnOK
End Function

