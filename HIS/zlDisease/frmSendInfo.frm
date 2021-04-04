VERSION 5.00
Begin VB.Form frmSendInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "报告报送备注"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6585
   Icon            =   "frmSendInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5070
      TabIndex        =   14
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3840
      TabIndex        =   13
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "申报信息"
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6300
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   650
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   300
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1000
         Width           =   1725
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1725
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "报送单位："
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "报送时间："
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "报送人员："
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   645
         Width           =   900
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "病    人:"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   300
         Width           =   810
      End
   End
   Begin VB.TextBox txtContent 
      Height          =   3420
      Left            =   3360
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2160
      Width           =   3060
   End
   Begin VB.TextBox txtSuggestion 
      Height          =   3420
      Left            =   120
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2160
      Width           =   3060
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "处理情况说明"
      Height          =   180
      Left            =   3480
      TabIndex        =   3
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "反馈信息"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   720
   End
End
Attribute VB_Name = "frmSendInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer '1-新增;2-修改
Private mlng申报ID As Long
Private mdatTime As Date
Private mblnChange As Boolean

Private mstrOld反馈 As String
Private mstrOld处理说明 As String

Public Enum TextInfoEnum
    txt_病人 = 0
    txt_报送人员 = 1
    txt_报送时间 = 2
    txt_报送单位 = 3
End Enum

Public Function ShowMe(ByVal frmParent As Object, ByVal intType As Integer, ByVal lngID As Long, Optional ByVal datTime As Date) As Boolean
    mintType = intType
    mlng申报ID = lngID
    mdatTime = datTime
    
    Me.Show 1, frmParent
End Function

Private Sub SetFormState()
    Dim blnLock As Boolean
    
    blnLock = IIf(mintType = 3, True, False)
    txtSuggestion.Locked = blnLock
    txtContent.Locked = blnLock
End Sub

Private Sub loadData()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
     
    On Error GoTo errH
    If mintType = 2 Or mintType = 3 Then
        strSQL = "Select a.反馈信息, a.登记人, a.处理情况说明, b.姓名, b.报送人, b.报送时间, b.报送单位" & vbNewLine & _
                "From 疾病申报反馈 A, 疾病申报记录 B" & vbNewLine & _
                "Where a.申报id = b.文件id And a.申报id = [1] And a.登记时间 = [2]"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng申报ID, mdatTime)
        
        If rsTmp.RecordCount > 0 Then
            mstrOld反馈 = NVL(rsTmp!反馈信息)
            mstrOld处理说明 = NVL(rsTmp!处理情况说明)
            txtSuggestion.Text = mstrOld反馈
            txtContent.Text = mstrOld处理说明
            
            txtInfo(txt_病人) = NVL(rsTmp!姓名)
            txtInfo(txt_报送人员) = NVL(rsTmp!报送人)
            txtInfo(txt_报送时间) = NVL(rsTmp!报送时间)
            txtInfo(txt_报送单位) = NVL(rsTmp!报送单位)
        End If
    ElseIf mintType = 1 Then
        strSQL = "Select  姓名, 报送人, 报送时间, 报送单位 From 疾病申报记录 A Where 文件id = [1] "
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng申报ID)

        If rsTmp.RecordCount > 0 Then
            txtInfo(txt_病人) = NVL(rsTmp!姓名)
            txtInfo(txt_报送人员) = NVL(rsTmp!报送人)
            txtInfo(txt_报送时间) = NVL(rsTmp!报送时间)
            txtInfo(txt_报送单位) = NVL(rsTmp!报送单位)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CheckData(False)
    If mintType = 1 Then
        If mblnChange Then
            If MsgBox("已经填写了内容，是否放弃？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    ElseIf mintType = 2 Then
        If mblnChange Then
            If MsgBox("内容已经发生改变，是否放弃修改？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End Sub

Private Function CheckData(ByVal blnOK As Boolean) As Boolean
    Dim strSuggestion As String, strContent As String
    
    strSuggestion = Trim(txtSuggestion.Text)
    strContent = Trim(txtContent.Text)
    
    If strSuggestion <> mstrOld反馈 Or strContent <> mstrOld处理说明 Then
        mblnChange = True
    Else
        mblnChange = False
    End If
    
    If blnOK Then
        If strSuggestion = "" Then
            MsgBox "反馈信息不能够为空，请先填写反馈信息！", vbInformation, gstrSysName
            txtSuggestion.SetFocus
            Exit Function
        ElseIf strContent = "" Then
            MsgBox "处理情况说明不能够为空，请先填写处理情况说明！", vbInformation, gstrSysName
            txtSuggestion.SetFocus
            Exit Function
        ElseIf LenB(StrConv(strSuggestion, vbFromUnicode)) > txtSuggestion.MaxLength Then
            MsgBox "反馈信息超长（最多" & txtSuggestion.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName
            txtSuggestion.SetFocus: Exit Function
        ElseIf LenB(StrConv(strContent, vbFromUnicode)) > txtContent.MaxLength Then
            MsgBox "处理情况说明超长（最多" & txtContent.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName
            txtContent.SetFocus: Exit Function
        End If
    End If
    CheckData = True
End Function

Private Sub cmdOK_Click()
    Dim strSuggestion As String, strContent As String, str登记时间 As String
    Dim str登记人 As String
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not CheckData(True) Then Exit Sub
    
    strSuggestion = "'" & Trim(txtSuggestion.Text) & "'"
    strContent = "'" & Trim(txtContent.Text) & "'"
    str登记人 = "'" & UserInfo.姓名 & "'"
    
    If mintType = 1 Then
        str登记时间 = "to_date('" & Format(gobjComlib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
        strSQL = "zl_疾病申报反馈_insert(" & mlng申报ID & "," & strSuggestion & "," & str登记人 & "," & str登记时间 & "," & strContent & ")"
    ElseIf mintType = 2 Then
        str登记时间 = "to_date('" & Format(mdatTime, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS') "
        strSQL = "zl_疾病申报反馈_Update(" & mlng申报ID & "," & strSuggestion & "," & str登记人 & "," & str登记时间 & "," & strContent & ")"
    End If

    Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, Me.Caption)

    mstrOld反馈 = Trim(txtSuggestion.Text)
    mstrOld处理说明 = Trim(txtContent.Text)
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mblnChange = False
    mstrOld反馈 = ""
    mstrOld处理说明 = ""
    Call loadData
End Sub

Private Sub txtContent_GotFocus()
    Me.txtContent.SelStart = 0: Me.txtContent.SelLength = 500
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call gobjComlib.ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtContent_LostFocus()
    Me.txtContent.Text = Replace(Me.txtContent, Chr(vbKeyReturn), "")
End Sub

Private Sub txtSuggestion_GotFocus()
    Me.txtSuggestion.SelStart = 0: Me.txtSuggestion.SelLength = 500
    Call gobjComlib.ZLCommFun.OpenIme(True)
End Sub

Private Sub txtSuggestion_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call gobjComlib.ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtSuggestion_LostFocus()
    Me.txtSuggestion.Text = Replace(Me.txtSuggestion, Chr(vbKeyReturn), "")
End Sub
