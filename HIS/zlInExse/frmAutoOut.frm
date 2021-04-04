VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoOut 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "出院登记"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmAutoOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraInfo 
      Height          =   2625
      Left            =   105
      TabIndex        =   15
      Top             =   45
      Width           =   6195
      Begin VB.ComboBox cbo随诊 
         Height          =   300
         ItemData        =   "frmAutoOut.frx":020A
         Left            =   4850
         List            =   "frmAutoOut.frx":021D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2130
         Width           =   1215
      End
      Begin VB.ComboBox cbo出院情况 
         Height          =   300
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   4230
      End
      Begin VB.CheckBox chk随诊 
         Alignment       =   1  'Right Justify
         Caption         =   "随诊"
         Height          =   195
         Left            =   2835
         TabIndex        =   8
         Top             =   2190
         Width           =   660
      End
      Begin VB.TextBox txt随诊 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4170
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2130
         Width           =   405
      End
      Begin VB.CheckBox chk尸检 
         Alignment       =   1  'Right Justify
         Caption         =   "尸检"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5430
         TabIndex        =   6
         Top             =   1800
         Width           =   660
      End
      Begin VB.TextBox txt出院诊断 
         Height          =   300
         Left            =   960
         MaxLength       =   100
         TabIndex        =   0
         Top             =   240
         Width           =   5100
      End
      Begin VB.CheckBox chk疑诊 
         Alignment       =   1  'Right Justify
         Caption         =   "确诊"
         Height          =   195
         Left            =   2835
         TabIndex        =   5
         Top             =   1800
         Width           =   660
      End
      Begin VB.TextBox txt中医诊断 
         Height          =   300
         Left            =   960
         MaxLength       =   100
         TabIndex        =   2
         Top             =   960
         Width           =   5100
      End
      Begin VB.ComboBox cbo中医出院情况 
         Height          =   300
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   4230
      End
      Begin VB.ComboBox cbo出院方式 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1740
         Width           =   1830
      End
      Begin MSComCtl2.UpDown UD随诊 
         Height          =   300
         Left            =   4590
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2130
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txt随诊"
         BuddyDispid     =   196613
         OrigLeft        =   3945
         OrigTop         =   645
         OrigRight       =   4185
         OrigBottom      =   930
         Max             =   99999
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   960
         TabIndex        =   7
         Top             =   2130
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
      Begin MSMask.MaskEdBox txtOkDate 
         Height          =   300
         Left            =   3600
         TabIndex        =   23
         Top             =   1760
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况"
         Height          =   180
         Left            =   1035
         TabIndex        =   22
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   180
         TabIndex        =   21
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "期限"
         Height          =   180
         Left            =   3780
         TabIndex        =   20
         Top             =   2190
         Width           =   360
      End
      Begin VB.Label lbl出院诊断 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl中医诊断 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "中医诊断"
         Height          =   180
         Left            =   180
         TabIndex        =   18
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院情况"
         Height          =   180
         Left            =   1035
         TabIndex        =   17
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label lbl出院方式 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院方式"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   1800
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   225
      TabIndex        =   14
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3870
      TabIndex        =   12
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5070
      TabIndex        =   13
      Top             =   2850
      Width           =   1100
   End
End
Attribute VB_Name = "frmAutoOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlng病人ID As Long
Public mlng主页ID As Long
Public mstr性别 As String
Public mint险类 As Integer
Public mlngDepID As Long '出院科室ID
Private mdteDeathDate As Date
Private mintDeath As Integer

Private Sub cbo出院方式_Click()
    If InStr(cbo出院方式.Text, "死亡") > 0 Then
        txt随诊.Text = ""
        chk随诊.Value = 0

        txt随诊.Enabled = (chk随诊.Value = 1)
        UD随诊.Enabled = txt随诊.Enabled

        chk随诊.Enabled = False
    
        chk尸检.Enabled = True
    Else
        chk随诊.Enabled = True
        
        chk尸检.Value = 0
        chk尸检.Enabled = False
    End If
End Sub

Private Sub cbo出院情况_Click()
    Dim i As Integer
    If InStr(cbo出院情况.Text, "死亡") > 0 Then
        i = cbo.FindIndex(cbo出院方式, "死亡", True)
        If i <> -1 Then cbo出院方式.ListIndex = i
    End If
End Sub

Private Sub cbo随诊_Click()
    txt随诊.Enabled = (cbo随诊.ItemData(cbo随诊.ListIndex) <> 9)
    UD随诊.Enabled = txt随诊.Enabled
End Sub

Private Sub cbo中医出院情况_Click()
    Dim i As Integer
    If InStr(cbo中医出院情况.Text, "死亡") > 0 Then
        i = cbo.FindIndex(cbo出院方式, "死亡", True)
        If i <> -1 Then cbo出院方式.ListIndex = i
    End If
End Sub

Private Sub chk随诊_Click()
    txt随诊.Enabled = (chk随诊.Value = 1)
    UD随诊.Enabled = txt随诊.Enabled
    cbo随诊.Enabled = txt随诊.Enabled
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
'问题28982 by lesfeng 2010-06-09
Private Sub chk疑诊_Click()
    If chk疑诊.Value = 1 Then
        txtOkDate.Enabled = True
    Else
        txtOkDate.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp "zl9InPatient", Me.hWnd, "frmOut"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Not Me.ActiveControl Is txt出院诊断 _
            And Not Me.ActiveControl Is txt中医诊断 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") And Not (Me.ActiveControl Is txt出院诊断 Or Me.ActiveControl Is txt中医诊断) Then KeyAscii = 0      '诊断内容中可能有'号
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, rsDiagnosis As ADODB.Recordset
    Dim rsPatiInfo As ADODB.Recordset
    Dim i As Long, strSQL As String
    Dim dMax As Date, int原因 As Integer
    Dim lng疾病ID As Long, str诊断 As String, str出院情况 As String, str中医出院情况 As String
        
    On Error GoTo errH
    '问题28612 by lesfeng 2010-07-05
    
    mintDeath = 0
    mdteDeathDate = GetdeathTime(mlng病人ID, mlng主页ID)
    '问题31652 by lesfeng 从病案主页直接提取确诊日期
    Set rsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
    
    txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    dMax = GetMaxDate(mlng病人ID, mlng主页ID, int原因)
    If int原因 = 10 Then
        txtDate.Text = Format(dMax + 1 / 24 / 60 / 60, "yyyy-MM-dd HH:mm:ss")
    Else
        If dMax > CDate(txtDate.Text) Then
            txtDate.Text = Format(dMax + 1 / 24 / 60 / 60, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    '问题28612 by lesfeng 2010-07-05
    If mintDeath = 1 Then
        txtDate.Text = Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If mlngDepID <> 0 Then
        txt中医诊断.Enabled = (InStr(1, "," & GetDepCharacter(mlngDepID) & ",", ",中医科,") > 0)
        txt中医诊断.ToolTipText = "只有当病人所在科室的性质为中医科时才允许输入中医诊断!"
        cbo中医出院情况.Enabled = txt中医诊断.Enabled
    End If
    
     '显示病人诊断记录
    Set rsDiagnosis = GetDiagnosticInfo(mlng病人ID, mlng主页ID, "1,11,2,12,3,13", "2,3")
    If Not rsDiagnosis Is Nothing Then
        'a.西医诊断
        rsDiagnosis.Filter = "诊断类型=3 and 记录来源=3"            '先取首页整理的出院诊断
        If Not rsDiagnosis.EOF Then
            txt出院诊断.Text = NVL(rsDiagnosis!诊断描述): txt出院诊断.Tag = NVL(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl出院诊断.Tag = txt出院诊断.Text
            str出院情况 = "" & rsDiagnosis!出院情况
            '问题28982 by lesfeng 2010-06-09
            chk疑诊.Value = IIf(Val("" & rsDiagnosis!是否疑诊) = 1, 0, 1)
        Else
            rsDiagnosis.Filter = "诊断类型=2 and 记录来源=2"        '再取入院登记的入院诊断
            If Not rsDiagnosis.EOF Then
                txt出院诊断.Text = NVL(rsDiagnosis!诊断描述): txt出院诊断.Tag = NVL(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl出院诊断.Tag = txt出院诊断.Text
            Else
                rsDiagnosis.Filter = "诊断类型=1 and 记录来源=2"    '最后取入院登记的门诊诊断
                If Not rsDiagnosis.EOF Then
                    txt出院诊断.Text = NVL(rsDiagnosis!诊断描述): txt出院诊断.Tag = NVL(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl出院诊断.Tag = txt出院诊断.Text
                End If
            End If
        End If
        
        'b.中医诊断
        If txt中医诊断.Enabled Then
            rsDiagnosis.Filter = "诊断类型=13 and 记录来源=3"            '先取首页整理的出院诊断
            If Not rsDiagnosis.EOF Then
                txt中医诊断.Text = NVL(rsDiagnosis!诊断描述): txt中医诊断.Tag = NVL(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                str中医出院情况 = "" & rsDiagnosis!出院情况
            Else
                rsDiagnosis.Filter = "诊断类型=12 and 记录来源=2"        '再取入院登记的入院诊断
                If Not rsDiagnosis.EOF Then
                    txt中医诊断.Text = NVL(rsDiagnosis!诊断描述): txt中医诊断.Tag = NVL(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                Else
                    rsDiagnosis.Filter = "诊断类型=11 and 记录来源=2"    '最后取入院登记的门诊诊断
                    If Not rsDiagnosis.EOF Then
                        txt中医诊断.Text = NVL(rsDiagnosis!诊断描述): txt中医诊断.Tag = NVL(rsDiagnosis!疾病ID, rsDiagnosis!诊断ID & ";"): lbl中医诊断.Tag = txt中医诊断.Text
                    End If
                End If
            End If
        End If
    End If
    '问题28982 by lesfeng 2010-06-09
    '问题31652 by lesfeng 从病案主页直接提取确诊日期
    If Not IsNull(rsPatiInfo!确诊日期) Then
        txtOkDate.Text = Format(rsPatiInfo!确诊日期, "yyyy-MM-dd HH:mm:ss")
        chk疑诊.Value = IIf(Val("" & rsPatiInfo!是否确诊) = 1, 1, 0)
        If chk疑诊.Value = 0 Then chk疑诊.Value = 1
        chk疑诊.Enabled = False
        txtOkDate.Enabled = False
    End If
        
    '出院情况
    cbo出院情况.AddItem "": cbo出院情况.ListIndex = cbo出院情况.NewIndex
    If cbo中医出院情况.Enabled Then cbo中医出院情况.AddItem "": cbo中医出院情况.ListIndex = cbo中医出院情况.NewIndex
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 治疗结果 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo出院情况.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then
                If txt出院诊断.Text <> "" Then cbo出院情况.ListIndex = cbo出院情况.NewIndex
                cbo出院情况.ItemData(cbo出院情况.NewIndex) = 1
            End If
        
            If cbo中医出院情况.Enabled Then
                cbo中医出院情况.AddItem rsTmp!编码 & "-" & rsTmp!名称
                If rsTmp!缺省 = 1 Then
                    If txt中医诊断.Text <> "" Then cbo中医出院情况.ListIndex = cbo中医出院情况.NewIndex
                    cbo中医出院情况.ItemData(cbo中医出院情况.NewIndex) = 1
                End If
            End If
            
            rsTmp.MoveNext
        Next
    End If
    Call zlControl.CboLocate(cbo出院情况, str出院情况)
    Call zlControl.CboLocate(cbo中医出院情况, str中医出院情况)
    
    
    '出院方式
    strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 出院方式 Order by 编码"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo出院方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then cbo出院方式.ListIndex = cbo出院方式.NewIndex
            rsTmp.MoveNext
        Next
    End If
        
    cbo随诊.ListIndex = 1
    Call chk随诊_Click
    '问题28982 by lesfeng 2010-06-09
    If chk疑诊.Enabled Then Call chk疑诊_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, dMax As Date, strSQL As String, Curdate As Date, blnTrans As Boolean
    Dim lng西医疾病ID As Long, lng中医疾病ID As Long
    Dim lng西医诊断ID As Long, lng中医诊断ID As Long
    Dim int随诊 As Integer
    Dim str确诊日期  As String
    Dim str入院时间 As String
    Dim strInfo As String
    Dim rsPatiInfo As ADODB.Recordset
    
    On Error GoTo errH
    
    '出院诊断
    If Not zlControl.TxtCheckInput(txt出院诊断, "出院诊断", txt出院诊断.MaxLength) Then Exit Sub
    If Not zlControl.TxtCheckInput(txt中医诊断, "中医诊断", txt中医诊断.MaxLength, True) Then Exit Sub
    If mint险类 <> 0 Then
        If gclsInsure.GetCapability(support必须录入入出诊断, mlng病人ID, mint险类) Then
            If txt出院诊断.Text = "" Then
                MsgBox "请填写该病人的出院诊断！", vbInformation, gstrSysName
                txt出院诊断.SetFocus: Exit Sub
            End If
        End If
    End If
    If txt出院诊断.Text <> "" And cbo出院情况.Text = "" Then
        MsgBox "请选择出院诊断的出院情况。", vbInformation, gstrSysName
        cbo出院情况.SetFocus: Exit Sub
    End If
    If txt中医诊断.Text <> "" And cbo中医出院情况.Text = "" And cbo中医出院情况.Enabled Then
        MsgBox "请选择中医出院诊断的出院情况。", vbInformation, gstrSysName
        cbo中医出院情况.SetFocus: Exit Sub
    End If
    
    If Not IsDate(txtDate.Text) Then
        MsgBox "请输入正确的病人出院时间！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    '时间不能超过当前时间太长(一周)
    Curdate = zlDatabase.Currentdate
    If CDate(txtDate.Text) > Curdate Then
        If CDate(txtDate.Text) - Curdate > 7 Then
            MsgBox "出院时间比当前时间大得过多,请检查！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If MsgBox("出院时间大于了当前系统时间,确实要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    dMax = GetMaxDate(mlng病人ID, mlng主页ID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "病人出院时间必须大于该病人上次变动时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    dMax = GetLastAdviceTime(mlng病人ID, mlng主页ID)
    If Format(txtDate.Text, "yyyyMMddHHmmss") < Format(dMax, "yyyyMMddHHmmss") Then
        If MsgBox("出院时间小于该病人最后有效医嘱的时间 " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & ",确实要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txtDate.SetFocus: Exit Sub
        End If
    End If
    '问题28612 by lesfeng 2010-07-05
    If InStr(cbo出院方式.Text, "死亡") = 0 And mintDeath = 1 Then
        If MsgBox("该病人存在有效临床死亡医嘱,其死亡医嘱的时间 " & Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss") & ",但出院方式不为死亡,确实要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            cbo出院方式.SetFocus: Exit Sub
        End If
    End If
    
    If InStr(cbo出院方式.Text, "死亡") > 0 And mintDeath = 1 Then
        If Format(txtDate.Text, "yyyyMMddHHmmss") <> Format(mdteDeathDate, "yyyyMMddHHmmss") Then
            If MsgBox("出院时间不等于该病人有效临床死亡医嘱的时间 " & Format(mdteDeathDate, "yyyy-MM-dd HH:mm:ss") & ",确实要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txtDate.SetFocus: Exit Sub
            End If
        End If
    End If
    '问题32764 by lesfeng 2010-09-13 撤分参数22及32 新增154、155
    '检查病人是否有未执行完成的诊疗项目及未发药品
    If gbyt检查未执行 <> 0 Then
        strInfo = ExistWaitExe(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt检查未执行 = 1 Then
                If MsgBox("发现病人存在尚未执行完成的内容：" & _
                    vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "确定要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "发现病人存在尚未执行完成的内容：" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "不允许出院.", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If

    If gbyt检查未发药 <> 0 Then
        strInfo = ExistWaitDrug(mlng病人ID, mlng主页ID)
        If strInfo <> "" Then
            If gbyt检查未发药 = 1 Then
                If MsgBox("发现病人" & strInfo & vbCrLf & vbCrLf & "确定要出院吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            Else
                MsgBox "发现病人" & strInfo & vbCrLf & vbCrLf & "不允许出院。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    '问题28982 by lesfeng 2010-06-09
    str确诊日期 = ""
    If chk疑诊.Value = 1 Then
        Set rsPatiInfo = GetPatiInfo(mlng病人ID, mlng主页ID)
        str入院时间 = Format(rsPatiInfo!入院时间, "yyyy-MM-dd HH:mm:ss")
        If Not IsDate(txtOkDate.Text) Then
            MsgBox "请输入正确的病人确诊时间！", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        If Format(txtOkDate.Text, "yyyyMMddHHmmss") >= Format(txtDate.Text, "yyyyMMddHHmmss") Then
            MsgBox "确诊时间必须小于病人出院时间 " & Format(txtDate.Text, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        
        If Format(str入院时间, "yyyyMMddHHmmss") > Format(txtOkDate.Text, "yyyyMMddHHmmss") Then
            MsgBox "确诊时间必须大于等于病人入院时间 " & Format(str入院时间, "yyyy-MM-dd HH:mm:ss") & " ！", vbInformation, gstrSysName
            If txtOkDate.Enabled Then txtOkDate.SetFocus: Exit Sub
        End If
        
        str确诊日期 = Format(txtOkDate.Text, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If InStr(1, txt出院诊断.Tag, ";") <= 0 Then
        lng西医疾病ID = Val(txt出院诊断.Tag)
    Else
        lng西医诊断ID = Val(txt出院诊断.Tag)
    End If
    If InStr(1, txt中医诊断.Tag, ";") <= 0 Then
        lng中医疾病ID = Val(txt中医诊断.Tag)
    Else
        lng中医诊断ID = Val(txt中医诊断.Tag)
    End If
    
    If cbo随诊.ListIndex <> -1 Then int随诊 = cbo随诊.ItemData(cbo随诊.ListIndex)
    '问题28982 by lesfeng 2010-06-09
    strSQL = "zl_病人变动记录_Out(" & mlng病人ID & "," & mlng主页ID & "," & _
            ZVal(lng西医疾病ID) & "," & ZVal(lng西医诊断ID) & ",'" & Replace(txt出院诊断.Text, "'", "''") & "','" & zlStr.NeedName(cbo出院情况.Text) & "'," & _
            ZVal(lng中医疾病ID) & "," & ZVal(lng中医诊断ID) & ",'" & Replace(txt中医诊断.Text, "'", "''") & "','" & zlStr.NeedName(cbo中医出院情况.Text) & "'," & _
            chk疑诊.Value & ",'" & zlStr.NeedName(cbo出院方式.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            IIf(chk随诊.Value = 1, int随诊, 0) & "," & IIf(chk随诊.Value = 1 And int随诊 <> 9, Val(txt随诊.Text), "Null") & "," & IIf(chk尸检.Enabled, chk尸检.Value, "NULL") & "," & _
            "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & IIf(str确诊日期 = "", "NULL", "To_Date('" & str确诊日期 & "','YYYY-MM-DD HH24:MI:SS')") & ")"
    
    gcnOracle.BeginTrans: blnTrans = True
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        If mint险类 <> 0 Then '医保改动
            If Not gclsInsure.LeaveSwap(mlng病人ID, mlng主页ID, mint险类) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '刘兴洪:24662
    Dim strOutPut As String
    'If Not mobjICCard Is Nothing Then
        Call zlExcuteUploadSwap(mlng病人ID, strOutPut)
    'End If
    
    
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng病人ID = 0
    mlng主页ID = 0
    mint险类 = 0
    mstr性别 = ""
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub txt随诊_GotFocus()
    zlControl.TxtSelAll txt随诊
End Sub

Private Sub txt随诊_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt出院诊断_GotFocus()
    zlControl.TxtSelAll txt出院诊断
End Sub

Private Sub txt中医诊断_GotFocus()
    zlControl.TxtSelAll txt中医诊断
End Sub

Private Sub txt出院诊断_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    
If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt出院诊断.Text = lbl出院诊断.Tag And txt出院诊断.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt出院诊断.Text = "" Then
            txt出院诊断.Tag = "": lbl出院诊断.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt出院诊断.Left, txt出院诊断.Top)
            strInput = UCase(txt出院诊断.Text)
            strSex = mstr性别
            lngTxtHeight = txt出院诊断.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "D", vPoint.X, vPoint.Y, lngTxtHeight)
            If Not rsTmp Is Nothing Then
                txt出院诊断.Tag = rsTmp!ID
                txt出院诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl出院诊断.Tag = txt出院诊断.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If lbl出院诊断.Tag <> "" Then txt出院诊断.Text = lbl出院诊断.Tag
                Call txt出院诊断_GotFocus
                txt出院诊断.SetFocus
            End If
        End If
    Else
        CheckInputLen txt出院诊断, KeyAscii
    End If
End Sub

Private Sub txt中医诊断_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, strInput As String, strSex As String, lngTxtHeight As Long, vPoint As POINTAPI
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not RequestCode Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt中医诊断.Text = lbl中医诊断.Tag And txt中医诊断.Text <> "" Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf txt中医诊断.Text = "" Then
            txt中医诊断.Tag = "": lbl中医诊断.Tag = ""
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            vPoint = zlControl.GetCoordPos(fraInfo.hWnd, txt中医诊断.Left, txt中医诊断.Top)
            strInput = UCase(txt中医诊断.Text)
            strSex = mstr性别
            lngTxtHeight = txt中医诊断.Height
            Set rsTmp = GetDiseaseCode(Me, blnCancel, strInput, strSex, "B", vPoint.X, vPoint.Y, lngTxtHeight)
            
            If Not rsTmp Is Nothing Then
                txt中医诊断.Tag = rsTmp!ID
                txt中医诊断.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                lbl中医诊断.Tag = txt中医诊断.Text '用于恢复显示
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                End If
                If lbl中医诊断.Tag <> "" Then txt中医诊断.Text = lbl中医诊断.Tag
                Call txt中医诊断_GotFocus
                txt中医诊断.SetFocus
                
            End If
        End If
    Else
        CheckInputLen txt中医诊断, KeyAscii
    End If
End Sub

Private Sub txt出院诊断_Validate(Cancel As Boolean)
    If Val(txt出院诊断.Tag) > 0 And txt出院诊断.Text <> lbl出院诊断.Tag Then
        txt出院诊断.Text = lbl出院诊断.Tag
    ElseIf Val(txt出院诊断.Tag) = 0 And RequestCode Then
        txt出院诊断.Text = ""
    End If
    
    If txt出院诊断.Text <> "" And cbo出院情况.Text = "" Then
        cbo出院情况.ListIndex = cbo.FindIndex(cbo出院情况, 1)
        If cbo出院情况.ListIndex = -1 Then cbo出院情况.ListIndex = 0
    ElseIf txt出院诊断.Text = "" And cbo出院情况.Text <> "" Then
        cbo出院情况.ListIndex = 0
    End If
End Sub

Private Sub txt中医诊断_Validate(Cancel As Boolean)
    If Val(txt中医诊断.Tag) > 0 And txt中医诊断.Text <> lbl中医诊断.Tag Then
        txt中医诊断.Text = lbl中医诊断.Tag
    ElseIf Val(txt中医诊断.Tag) = 0 And RequestCode Then
        txt中医诊断.Text = ""
    End If
    
    If txt中医诊断.Text <> "" And cbo中医出院情况.Text = "" Then
        cbo中医出院情况.ListIndex = cbo.FindIndex(cbo中医出院情况, 1)
        If cbo中医出院情况.ListIndex = -1 Then cbo中医出院情况.ListIndex = 0
    ElseIf txt中医诊断.Text = "" And cbo中医出院情况.Text <> "" Then
        cbo中医出院情况.ListIndex = 0
    End If
End Sub

Private Function RequestCode() As Boolean
    RequestCode = gint诊断输入 = 2 Or (gint诊断输入 = 3 And mint险类 <> 0)
End Function

Public Function GetFindIllWhere(ByVal strInput As String, ByVal str别名 As String) As String
'功能:获得查找疾病编码目录的条件
    Dim strWhere As String
    
    If zlCommFun.IsCharChinese(strInput) Then
        strWhere = str别名 & ".名称 like '" & gstrLike & strInput & "%'"
    Else
        strWhere = "(" & str别名 & ".编码 Like '" & strInput & "%'" & _
                " Or " & str别名 & ".名称 Like '" & gstrLike & strInput & "%'" & _
                " Or " & str别名 & ".简码 Like '" & gstrLike & strInput & "%')"
    End If
    GetFindIllWhere = strWhere
End Function
'问题28982 by lesfeng 2010-06-09
Private Function GetPatiInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：获取病人相关信息
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    '问题31652 by lesfeng 从病案主页直接提取确诊日期，增加是否确诊及确诊日期
    strSQL = "" & _
        "   Select  nvl(B.姓名,A.姓名) as 姓名, nvl(B.性别,Nvl(A.性别,'未知')) as  性别, B.年龄, B.险类, B.病人性质, B.当前病况, B.护理等级id, B.住院医师, B.门诊医师, B.责任护士, B.出院科室id, B.出院科室id 入住科室id," & vbNewLine & _
        "           To_Char(B.入院日期, 'YYYY-MM-DD HH24:MI:SS') As 入院时间, B.当前病区id, B.住院号, D.名称 as 当前科室, B.出院病床 as 主要床号,B.是否确诊,B.确诊日期" & vbNewLine & _
        "   From 病人信息 A, 病案主页 B, 部门表 D" & vbNewLine & _
        "   Where A.病人id = B.病人id And B.病人id = [1] And B.主页id = [2] And B.出院科室id = D.id"
   
    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'问题28612 by lesfeng 2010-07-05
Private Function GetdeathTime(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Date
'功能：获取指定病人是否存在死亡医嘱，存在出院时间为死亡时间加1秒
'说明：用于获取病人死亡时间为出院时间
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    GetdeathTime = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    
    On Error GoTo errH
    
    strSQL = "Select Max(Nvl(A.执行终止时间, Nvl(A.上次执行时间, A.开始执行时间)) + 1 / 24 / 60 ) As 时间 " & _
             "  From 病人医嘱记录 A, 诊疗项目目录 B " & _
             " Where A.诊疗类别 = 'Z' And A.诊疗项目id = B.ID And B.操作类型 = 11 And B.类别 = 'Z' And A.医嘱状态 In (3, 8, 9) And " & _
             "       A.病人ID = [1] And A.主页ID = [2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!时间) Then
            GetdeathTime = rsTmp!时间
            mintDeath = 1
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


