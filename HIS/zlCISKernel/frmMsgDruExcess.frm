VERSION 5.00
Begin VB.Form frmMsgDruExcess 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   Icon            =   "frmMsgDruExcess.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdYes 
      Caption         =   "确定(&Y)"
      Height          =   350
      Left            =   3165
      TabIndex        =   6
      Top             =   1170
      Width           =   1100
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "取消(&N)"
      Height          =   350
      Left            =   4305
      TabIndex        =   5
      Top             =   1170
      Width           =   1100
   End
   Begin VB.CommandButton cmdComExcReason 
      Height          =   300
      Left            =   5070
      Picture         =   "frmMsgDruExcess.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "将当前嘱托设置为常用说明"
      Top             =   690
      Width           =   315
   End
   Begin VB.CommandButton cmdExcReason 
      Caption         =   "…"
      Height          =   265
      Left            =   4740
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   285
   End
   Begin VB.TextBox txtExcessInfo 
      Height          =   300
      IMEMode         =   1  'ON
      Left            =   1110
      MaxLength       =   500
      TabIndex        =   2
      Top             =   690
      Width           =   3945
   End
   Begin VB.TextBox txtPSYX 
      Height          =   300
      IMEMode         =   1  'ON
      Left            =   1110
      MaxLength       =   500
      TabIndex        =   7
      Top             =   690
      Width           =   3945
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "一行25个汉字"
      Height          =   180
      Left            =   1110
      TabIndex        =   4
      Top             =   285
      Width           =   1080
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "超量说明"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   765
      Width           =   720
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   300
      Picture         =   "frmMsgDruExcess.frx":0596
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgDruExcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvResult As VbMsgBoxResult
Private mstrInfo As String
Private mstrResult As String
Private mintType As Integer ' 0-超量说明时，1-皮试阳性结果用药说明

Public Function ShowMe(frmParent As Object, ByVal intType As Integer, ByVal strInfo As String, ByRef strResult As String) As VbMsgBoxResult
'参数：strInfo=提示信息
'      strResult 传出参数，表示所填写的超量说明
'      intType 0-超量说明时，1-皮试阳性时
'返回：
'      vbYes=是,vbNo=否
    Dim intMouse As Integer
    strResult = ""
    mstrInfo = strInfo
    intMouse = Screen.MousePointer
    mintType = intType
    Screen.MousePointer = 0
    Me.Show 1, frmParent
    Screen.MousePointer = intMouse
    strResult = IIF(mstrResult = "", "*NULL*", mstrResult) '什么都没填写，则记录一个特殊字符串，用于外部判断
    ShowMe = mvResult
End Function

Private Sub cmdComExcReason_Click()
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    If txtExcessInfo.Text = "" Then MsgBox "请输入您需要收藏的超量说明。", vbInformation, Me.Caption: txtExcessInfo.SetFocus: Exit Sub
    
    On Error GoTo errH
    strSQL = "Select 1 From 医嘱常用原因 Where 名称=[1] And 性质=1 And 人员=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtExcessInfo.Text, UserInfo.姓名)
    '如果已经有了，提示用户是否继续。
    If rsTmp.RecordCount > 0 Then
        MsgBox "已经存在相同的超量说明。", vbInformation, Me.Caption
        Exit Sub
    End If
    strSQL = "zl_医嘱常用原因_Update(0,Null,'" & txtExcessInfo.Text & "',Null ,1,'" & UserInfo.姓名 & "'" & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    MsgBox "超量说明收藏成功。", vbInformation, Me.Caption
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReasonSelect(Optional ByVal strFind As String) As Boolean
'常用超量说明选择器
    Dim strRetrun As String
    Dim blnCancle As Boolean
    Dim lngLeft As Long, lngTop As Long
    
    lngLeft = txtExcessInfo.Left + Me.Left
    lngTop = txtExcessInfo.Top + Me.Top - 2600
    
    strRetrun = frmKssReasonSelect.ShowMe(Me, strFind, blnCancle, lngLeft, lngTop, 3)
    If Not blnCancle Then
        If strRetrun = "" Then
            If strFind = "" Then
                MsgBox "没有找到可用的超量说明。", vbInformation, Me.Caption
            End If
        Else
            txtExcessInfo.Text = strRetrun
        End If
    End If
    ReasonSelect = blnCancle
End Function

Private Sub cmdExcReason_Click()
    Call ReasonSelect
End Sub

Private Sub cmdYes_Click()
    If mintType = 0 Then
        mstrResult = txtExcessInfo.Text
    ElseIf mintType = 1 Then
        mstrResult = txtPSYX.Text
    End If
    mvResult = vbYes
    Unload Me
End Sub

Private Sub cmdNo_Click()
    mstrResult = ""
    mvResult = vbCancel
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '点击窗体关闭按钮
    If UnloadMode = vbFormControlMenu Then
        mstrResult = ""
        mvResult = vbCancel
    End If
End Sub

Private Sub Form_Activate()
    If mintType = 0 Then
        txtExcessInfo.SetFocus
        Beep
    ElseIf mintType = 1 Then
        txtPSYX.SetFocus
        Beep
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyY Then
        Call cmdYes_Click
    ElseIf KeyCode = vbKeyN Then
        Call cmdNo_Click
    End If
End Sub

Private Sub Form_Load()
    Dim strTmp As String
    Dim strLoop As String
    Dim strR As String
    
    Caption = gstrSysName
    strLoop = mstrInfo
    
    Do While strLoop <> ""
        If Len(strLoop) > 25 Then
            strTmp = Mid(strLoop, 1, 25)
        Else
            strTmp = strLoop
        End If
        strR = strR & strTmp & vbCrLf
        strLoop = Mid(strLoop, 26)
    Loop
    lblInfo.Caption = strR
    lblInfo.Top = 200
    Me.Height = Me.Height + lblInfo.Height - 550
    
    
    If mintType = 1 Then
       lblName.Caption = "阳性用药"
       cmdComExcReason.Visible = False
       cmdExcReason.Visible = False
       txtExcessInfo.Visible = False
    Else
        txtPSYX.Visible = False
    End If
    
    lblName.Top = lblInfo.Height + lblInfo.Top + 60
    imgIcon.Top = lblInfo.Top + lblInfo.Height / 2 - imgIcon.Height / 2
    
    txtExcessInfo.Top = lblName.Top - 40
    txtPSYX.Top = lblName.Top - 40
    cmdComExcReason.Top = txtExcessInfo.Top - 10
    cmdExcReason.Top = txtExcessInfo.Top + 10
    
    cmdYes.Top = txtExcessInfo.Top + txtExcessInfo.Height + 100
    cmdNo.Top = cmdYes.Top
End Sub

Private Sub txtExcessInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtExcessInfo.Text <> "" Then
            If ReasonSelect(txtExcessInfo.Text) Then Exit Sub
        End If
    End If
End Sub

