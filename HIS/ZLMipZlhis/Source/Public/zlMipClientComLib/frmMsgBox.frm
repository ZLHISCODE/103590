VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "frmMsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picVisible 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   720
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   15
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   580
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   115802115
      CurrentDate     =   40829.9738657407
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "###"
      Height          =   350
      Index           =   0
      Left            =   1695
      TabIndex        =   0
      Top             =   900
      Width           =   1100
   End
   Begin VB.Label lblDateCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "日期标题(&D)"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   630
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   270
      Picture         =   "frmMsgBox.frx":000C
      Top             =   210
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   270
      Picture         =   "frmMsgBox.frx":08D6
      Top             =   210
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   270
      Picture         =   "frmMsgBox.frx":11A0
      Top             =   210
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMsgBox.frx":1A6A
      Height          =   360
      Left            =   960
      TabIndex        =   1
      Top             =   210
      Width           =   3150
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   270
      Picture         =   "frmMsgBox.frx":1AB6
      Top             =   210
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrInfo As String
Private mstrCaption As String
Private mstrCmds As String
Private mvStyle As VbMsgBoxStyle
Private mstrDateCaption As String
Private mDateInput As Date
Private mstrDateFormat As String

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, _
    frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion, Optional ByVal strDateCaption As String, _
    Optional ByRef DateInput As Date, Optional ByVal strDateFormat As String) As String
'参数：strCaption=消息窗体标题
'      strInfo=具体提示内容,可用"^"表示换行,">"表示缩进。
'      strCmds=按钮描述,如"重试(&R),!忽略(&A),?取消(&C)"
'              至少要有两个按钮,"!"表示缺省按钮,"?"表示取消按钮
'              每个按钮文字最多支持4个汉字
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'      strDateCaption=传入的日期的标题，如果<>""则显示日期控件，供用户输入日期，将日期DateInput返回。
'      strDateFormat=时间格式 格式""yyyy-MM-dd HH:mm:ss" 其中HH为大写是24小时制"
'返回：按钮文字,如"按钮2"(不包含()和&),如果按关闭或取消则返回""
    Dim intMouse As Integer
    
    mstrCaption = strCaption
    mstrInfo = strInfo
    mstrCmds = strCmds
    mvStyle = vStyle
    mstrDateCaption = strDateCaption
    mDateInput = DateInput
    mstrDateFormat = strDateFormat
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    Me.Show 1, frmParent
    DateInput = mDateInput
    Screen.MousePointer = intMouse
    
    ShowMsgBox = mstrCmds
End Function

Private Sub cmdDo_Click(Index As Integer)
    mstrCmds = Replace(Split(cmdDo(Index).Caption, "(")(0), "&", "")
    If cmdDo(Index).Cancel Then mstrCmds = ""
    mDateInput = CDate(dtpDate.value)
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If Me.ActiveControl.Name = "dtpDate" And KeyCode = vbKeyReturn Then picVisible.SetFocus: Exit Sub
    If Me.ActiveControl.Name = "dtpDate" And KeyCode <> vbKeyEscape Then Exit Sub
    '直接按单键热键
    If (KeyCode >= vbKey0 And KeyCode <= vbKey9 _
        Or KeyCode >= vbKeyA And KeyCode <= vbKeyZ) And Shift = 0 Then
        For i = 0 To cmdDo.UBound
            If InStr(cmdDo(i).Caption, "&") > 0 Then
                If Mid(cmdDo(i).Caption, InStr(cmdDo(i).Caption, "&") + 1, 1) = Chr(KeyCode) Then
                    Call cmdDo_Click(i): Exit Sub
                End If
            End If
        Next
        
        '没有定义快捷时，也可以用数字1-X为快捷
        If KeyCode >= vbKey1 And KeyCode <= vbKey9 Then
            For i = 0 To cmdDo.UBound
                If i + 1 = Val(Chr(KeyCode)) Then
                    Call cmdDo_Click(i): Exit Sub
                End If
            Next
        End If
    ElseIf KeyCode = vbKeyAdd Or KeyCode = 187 Then '(+)
        For i = 0 To cmdDo.UBound
            If InStr(cmdDo(i).Caption, "(+)") > 0 Then
                Call cmdDo_Click(i): Exit Sub
            End If
        Next
    ElseIf KeyCode = vbKeySubtract Or KeyCode = 189 Then '(-)
        For i = 0 To cmdDo.UBound
            If InStr(cmdDo(i).Caption, "(-)") > 0 Then
                Call cmdDo_Click(i): Exit Sub
            End If
        Next
    ElseIf KeyCode = vbKeyEscape Then
        mstrCmds = "": Unload Me
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.ActiveControl.Name = "dtpDate" Then picVisible.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '点击窗体关闭按钮
    If UnloadMode = vbFormControlMenu Then mstrCmds = ""
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    If i = 0 And (mvStyle And vbDefaultButton1) <> 0 Then i = 1
    If i = 0 And (mvStyle And vbDefaultButton2) <> 0 Then i = 2
    If i = 0 And (mvStyle And vbDefaultButton3) <> 0 Then i = 3
    If i = 0 And (mvStyle And vbDefaultButton4) <> 0 Then i = 4
    If i <> 0 Then
        cmdDo(i - 1).SetFocus
    Else
        '缺省定位到缺省按钮上
        For i = 0 To cmdDo.UBound
            If cmdDo(i).Default Then cmdDo(i).SetFocus: Exit For
        Next
        '没有缺省，没有指定定位按钮，则定位到最后一个上面
        If i > cmdDo.UBound Then
            cmdDo(cmdDo.UBound).SetFocus
        End If
    End If
    VBA.Beep
End Sub

Private Sub Form_Load()
    Dim arrCmds As Variant, i As Integer
    Dim lngCmdW As Long, lngCmdL As Long
    Dim lngLen As Long, Y As Long, z As Long
    
    Me.Caption = mstrCaption
    lblInfo.Caption = Replace(Replace(mstrInfo, "^", vbCrLf), ">", "　　")
    arrCmds = Split(mstrCmds, ","): mstrCmds = ""
    If (mvStyle And vbInformation) <> 0 Then
        imgIcon(0).Visible = True
    ElseIf (mvStyle And vbQuestion) <> 0 Then
        imgIcon(1).Visible = True
    ElseIf (mvStyle And vbExclamation) <> 0 Then
        imgIcon(2).Visible = True
    ElseIf (mvStyle And vbCritical) <> 0 Then
        imgIcon(3).Visible = True
    End If
    
    Me.Height = lblInfo.Top + lblInfo.Height + 1150
    If Me.Height < 1800 Then Me.Height = 1800
    
    '加载按钮
    For i = 0 To UBound(arrCmds)
        If i > 0 Then Load cmdDo(i)
        cmdDo(i).Caption = arrCmds(i)
        cmdDo(i).Top = Me.ScaleHeight - cmdDo(i).Height - 180
        cmdDo(i).Visible = True
    Next
    For i = 0 To UBound(arrCmds)
        If Left(cmdDo(i).Caption, 1) = "?" Then
            cmdDo(i).Caption = Mid(cmdDo(i).Caption, 2)
            cmdDo(i).Cancel = True
        ElseIf Left(cmdDo(i).Caption, 1) = "!" Then
            cmdDo(i).Caption = Mid(cmdDo(i).Caption, 2)
            cmdDo(i).Default = True
        End If
    Next
    
    '根据按钮确定按钮宽度
    For i = 0 To UBound(arrCmds)
        If LenB(StrConv(Replace(Split(cmdDo(i).Caption, "(")(0), "&", ""), vbFromUnicode)) > 8 Then
            Me.cmdDo(0).Width = 1500
        ElseIf LenB(StrConv(Replace(Split(cmdDo(i).Caption, "(")(0), "&", ""), vbFromUnicode)) > 4 Then
            Me.cmdDo(0).Width = 1300
        End If
    Next
    lngCmdW = (UBound(arrCmds) + 1) * (cmdDo(0).Width + 100)
    
    '时间控件
    If Trim(mstrDateCaption) <> "" Then
        lblDateCaption.Visible = True
        dtpDate.Visible = True
        dtpDate.CustomFormat = mstrDateFormat
        '如果标题太长，则插入换行符
        lngLen = LenB(StrConv(mstrDateCaption, vbFromUnicode))
        mstrDateCaption = mstrDateCaption & "(&D)"
        If lngLen > 12 Then
            Y = 0
            z = 1     'z=行数
            For i = 1 To Len(mstrDateCaption)
                
                Y = Y + LenB(StrConv(Mid(mstrDateCaption, i, 1), vbFromUnicode))
                If Y + z - 1 >= 12 * z Then
                    mstrDateCaption = Mid(mstrDateCaption, 1, i + z - 1) & Chr(vbKeyReturn) & Mid(mstrDateCaption, i + z)
                    z = z + 1
                End If
                    
            Next
        End If
        lblDateCaption.Caption = mstrDateCaption
        dtpDate.value = gobjComLib.zlDatabase.Currentdate
        '确定位置
        lblDateCaption.Top = lblInfo.Top + lblInfo.Height + 150
        dtpDate.Top = lblDateCaption.Top - 20
        dtpDate.Left = lblDateCaption.Left + lblDateCaption.Width + 50
    End If
    
     '确定窗体宽度和按钮整体位置
    Me.Width = lblInfo.Left + lblInfo.Width + 500
    If Me.Width < lblInfo.Left + lngCmdW + 500 Then
        Me.Width = lblInfo.Left + lngCmdW + 500
    End If
    If Me.Width < 4500 Then Me.Width = 4500
    Me.Height = Me.Height + IIf(lblDateCaption.Visible, lblDateCaption.Height + 150, 0)
    lngCmdL = (Me.ScaleWidth - lngCmdW) / 2 + 200
    For i = 0 To UBound(arrCmds)
        cmdDo(i).Width = cmdDo(0).Width
        cmdDo(i).Left = lngCmdL + (cmdDo(0).Width + 100) * i
        cmdDo(i).Top = cmdDo(i).Top + IIf(lblDateCaption.Visible, lblDateCaption.Height + 150, 0)
    Next
End Sub
