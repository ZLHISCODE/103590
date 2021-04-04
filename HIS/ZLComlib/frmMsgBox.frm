VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   0
   ClientWidth     =   4410
   ControlBox      =   0   'False
   Icon            =   "frmMsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtInfo 
      Height          =   270
      IMEMode         =   1  'ON
      Left            =   2040
      TabIndex        =   7
      Top             =   1235
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton opt 
      Caption         =   "ѡ��1"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   923
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picVisible 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   720
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   15
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   580
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   158990339
      CurrentDate     =   40829.9738657407
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "###"
      Height          =   350
      Index           =   0
      Left            =   1695
      TabIndex        =   8
      Top             =   1620
      Width           =   1100
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ı�����(&D)"
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   1280
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lblOpt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ѡ�����(&D)"
      Height          =   180
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lblDateCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ڱ���(&D)"
      Height          =   180
      Left            =   960
      TabIndex        =   1
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
      TabIndex        =   0
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
Private mstrSelectCaption As String
Private mstrSelectInput As String
Private mstrTextCaption As String
Private mlngTextLength As Long
Private mstrTextInput As String
Private mstrSort As String
Private mblnSelectMust As Boolean

Public Function ShowMsgBox(ByVal strCaption As String, ByVal strInfo As String, ByVal strCmds As String, _
    frmParent As Object, Optional vStyle As VbMsgBoxStyle = vbQuestion, Optional ByVal strDateCaption As String, _
    Optional ByRef DateInput As Date, Optional ByVal strDateFormat As String, Optional ByVal strSelectCaption As String, _
    Optional ByRef strSelectInput As String, Optional ByVal strTextCaption As String, _
    Optional ByVal lngTextLength As Long, Optional ByRef strTextInput As String, Optional ByVal strSort As String = "1,2,3", _
    Optional ByVal blnSelectMust As Boolean) As String
'������strCaption=��Ϣ�������
'      strInfo=������ʾ����,����"^"��ʾ����,">"��ʾ������
'      strCmds=��ť����,��"����(&R),!����(&A),?ȡ��(&C)"
'              ����Ҫ��������ť,"!"��ʾȱʡ��ť,"?"��ʾȡ����ť
'              ÿ����ť�������֧��4������
'      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
'      strDateCaption=��������ڵı��⣬���<>""����ʾ���ڿؼ������û��������ڣ�������DateInput���ء�
'      strDateFormat=ʱ���ʽ ��ʽ""yyyy-MM-dd HH:mm:ss" ����HHΪ��д��24Сʱ��"
'      strSelectCaption=ѡ��ı���:ѡ��1|1(1��Ϊȱʡ),ѡ��2|0|1(ѡ��ѡ��ʱ��������д��1�������ڣ�2�����ı���0��������)������
'      strSelectInput=ѡ��ѡ��ķ���ֵ(����ѡ�������)
'      strTextCaption=�ı������
'      lngTextLength=�ı������¼�볤��
'      strTextInput=�ı���ķ���ֵ
'      strSort=���ڡ�ѡ��ı������������=1��ѡ��=2���ı�=3��Ĭ������"1,2,3"
'      blnSelectMust=����е�ѡ������ѡ��һ����������ʾ��
'���أ���ť����,��"��ť2"(������()��&),������رջ�ȡ���򷵻�""
    Dim intMouse As Integer
    
    mstrCaption = strCaption
    mstrInfo = strInfo
    mstrCmds = strCmds
    mvStyle = vStyle
    mstrDateCaption = strDateCaption
    mDateInput = DateInput
    mstrDateFormat = strDateFormat
    mstrSelectCaption = strSelectCaption
    mstrSelectInput = strSelectInput
    mstrTextCaption = strTextCaption
    mlngTextLength = lngTextLength
    mstrTextInput = strTextInput
    mstrSort = strSort
    mblnSelectMust = blnSelectMust
    
    intMouse = Screen.MousePointer
    Screen.MousePointer = 0
    Me.Show 1, frmParent
    DateInput = mDateInput
    strSelectInput = mstrSelectInput
    strTextInput = mstrTextInput
    Screen.MousePointer = intMouse
    
    ShowMsgBox = mstrCmds
End Function

Private Sub cmdDo_Click(Index As Integer)
    mstrCmds = Replace(Split(cmdDo(Index).Caption, "(")(0), "&", "")
    If cmdDo(Index).Cancel Then mstrCmds = ""
    If mstrCmds <> "" And mblnSelectMust And mstrSelectCaption <> "" And mstrSelectInput = "" Then
        MsgBox "����ѡ��һ��ѡ�", vbInformation, Me.Caption
        Exit Sub
    End If
    mDateInput = CDate(dtpDate.value)
    mstrTextInput = IIf(txtInfo.Enabled, txtInfo.Text, "")
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If Me.ActiveControl.Name = "dtpDate" And KeyCode = vbKeyReturn Then picVisible.SetFocus: Exit Sub
    If Me.ActiveControl.Name = "dtpDate" And KeyCode <> vbKeyEscape Then Exit Sub
    If Me.ActiveControl.Name = "Opt" Or Me.ActiveControl.Name = "txtInfo" Then Exit Sub
    'ֱ�Ӱ������ȼ�
    If (KeyCode >= vbKey0 And KeyCode <= vbKey9 _
        Or KeyCode >= vbKeyA And KeyCode <= vbKeyZ) And Shift = 0 Then
        For i = 0 To cmdDo.UBound
            If InStr(cmdDo(i).Caption, "&") > 0 Then
                If Mid(cmdDo(i).Caption, InStr(cmdDo(i).Caption, "&") + 1, 1) = Chr(KeyCode) Then
                    Call cmdDo_Click(i): Exit Sub
                End If
            End If
        Next
        
        'û�ж�����ʱ��Ҳ����������1-XΪ���
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
    '�������رհ�ť
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
        'ȱʡ��λ��ȱʡ��ť��
        For i = 0 To cmdDo.UBound
            If cmdDo(i).Default Then cmdDo(i).SetFocus: Exit For
        Next
        'û��ȱʡ��û��ָ����λ��ť����λ�����һ������
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
    Dim lngTop As Long, j As Long, arrOpt As Variant
    Dim lngIndex As Long
    
    Me.Caption = IIf(mstrCaption = "", " ", mstrCaption)
    lblInfo.Caption = Replace(Replace(mstrInfo, "^", vbCrLf), ">", "����")
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
    
    '���ذ�ť
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
    
    '���ݰ�ťȷ����ť���
    For i = 0 To UBound(arrCmds)
        If LenB(StrConv(Replace(Split(cmdDo(i).Caption, "(")(0), "&", ""), vbFromUnicode)) > 8 Then
            Me.cmdDo(0).Width = 1500
        ElseIf LenB(StrConv(Replace(Split(cmdDo(i).Caption, "(")(0), "&", ""), vbFromUnicode)) > 4 Then
            Me.cmdDo(0).Width = 1300
        End If
    Next
    lngCmdW = (UBound(arrCmds) + 1) * (cmdDo(0).Width + 100)
    
    lngTop = lblInfo.Top + lblInfo.Height + 150
    lngIndex = 1
    For i = 0 To UBound(Split(mstrSort, ","))
        Select Case Split(mstrSort, ",")(i)
        Case "1"
            'ʱ��ؼ�
            If Trim(mstrDateCaption) <> "" Then
                lblDateCaption.Visible = True
                dtpDate.Visible = True
                dtpDate.CustomFormat = mstrDateFormat
                '�������̫��������뻻�з�
                lngLen = LenB(StrConv(mstrDateCaption, vbFromUnicode))
                mstrDateCaption = mstrDateCaption & "(&D)"
                If lngLen > 12 Then
                    Y = 0
                    z = 1     'z=����
                    For j = 1 To Len(mstrDateCaption)
                        
                        Y = Y + LenB(StrConv(Mid(mstrDateCaption, j, 1), vbFromUnicode))
                        If Y + z - 1 >= 12 * z Then
                            mstrDateCaption = Mid(mstrDateCaption, 1, j + z - 1) & Chr(vbKeyReturn) & Mid(mstrDateCaption, j + z)
                            z = z + 1
                        End If
                            
                    Next
                End If
                lblDateCaption.Caption = mstrDateCaption
                lblDateCaption.Width = Me.TextWidth(lblDateCaption.Caption)
                dtpDate.value = gobjComLib.zlDatabase.Currentdate
                'ȷ��λ��
                lblDateCaption.Top = lngTop
                lngTop = lngTop + lblDateCaption.Height + 150
                dtpDate.Top = lblDateCaption.Top - 20
                dtpDate.Left = lblDateCaption.Left + lblDateCaption.Width + 50
                dtpDate.TabIndex = lngIndex
                lngIndex = lngIndex + 1
            End If
        Case "2"
            If Trim(mstrSelectCaption) <> "" Then
                arrOpt = Split(mstrSelectCaption, ":")
                lblOpt.Caption = arrOpt(0)
                lblOpt.Top = lngTop
                lblOpt.Visible = True
                lngTop = lngTop + lblOpt.Height + 150
                lblOpt.Width = Me.TextWidth(lblOpt.Caption)
                opt(0).Left = lblOpt.Left + lblOpt.Width + 50
                If UBound(arrOpt) > 0 Then
                    arrOpt = Split(Split(mstrSelectCaption, ":")(1), ",")
                    For j = 0 To UBound(arrOpt)
                        If j > 0 Then Load opt(j)
                        opt(j).Caption = Split(arrOpt(j), "|")(0)
                        opt(j).Top = lblOpt.Top - 20
                        opt(j).Width = Me.TextWidth(opt(j).Caption & "����")
                        If j > 0 Then opt(j).Left = opt(j - 1).Left + opt(j - 1).Width
                        opt(j).Visible = True
                        If UBound(Split(arrOpt(j), "|")) > 0 Then
                            If Val(Split(arrOpt(j), "|")(1)) = 1 Then
                                opt(0).Tag = j
                            End If
                        End If
                        If UBound(Split(arrOpt(j), "|")) > 1 Then
                            If Split(arrOpt(j), "|")(2) <> "" Then opt(j).Tag = Split(arrOpt(j), "|")(2)
                        End If
                        opt(j).TabIndex = lngIndex
                        lngIndex = lngIndex + 1
                    Next
                End If
                opt(Val(opt(0).Tag)).value = True
            End If
        Case "3"
            If Trim(mstrTextCaption) <> "" Then
                lblText.Caption = mstrTextCaption
                lblText.Top = lngTop
                lblText.Visible = True
                lngTop = lngTop + lblText.Height + 150
                lblText.Width = Me.TextWidth(lblText.Caption)
                txtInfo.Left = lblText.Left + lblText.Width + 50
                txtInfo.TabIndex = lngIndex
                txtInfo.Visible = True
                txtInfo.Top = lblText.Top - 20
                lngIndex = lngIndex + 1
                txtInfo.Text = mstrTextInput
                txtInfo.MaxLength = mlngTextLength
            End If
        End Select
    Next
    
     'ȷ�������ȺͰ�ť����λ��
    Me.Width = lblInfo.Left + lblInfo.Width + 500
    If Me.Width < lblInfo.Left + lngCmdW + 500 Then
        Me.Width = lblInfo.Left + lngCmdW + 500
    End If
    If Me.Width < opt(opt.count - 1).Left + opt(opt.count - 1).Width + 400 Then
        Me.Width = opt(opt.count - 1).Left + opt(opt.count - 1).Width + 400
    End If
    If Me.Width < 4500 Then Me.Width = 4500
    txtInfo.Width = Me.Width - txtInfo.Left - 400
    Me.Height = lngTop + cmdDo(0).Height + 500
    lngCmdL = (Me.ScaleWidth - lngCmdW) / 2 + 200
    For i = 0 To UBound(arrCmds)
        cmdDo(i).Width = cmdDo(0).Width
        cmdDo(i).Left = lngCmdL + (cmdDo(0).Width + 100) * i
        cmdDo(i).Top = lngTop
    Next
End Sub

Private Sub opt_Click(Index As Integer)
    mstrSelectInput = Replace(Split(opt(Index).Caption, "(")(0), "&", "")
    If opt(Index).Tag <> "" Then
        Select Case opt(Index).Tag
            Case "0"
                dtpDate.Enabled = False
                txtInfo.Enabled = False
                txtInfo.BackColor = vbButtonFace
            Case "1"
                dtpDate.Enabled = False
            Case "2"
                txtInfo.Enabled = False
                txtInfo.BackColor = vbButtonFace
        End Select
    Else
        dtpDate.Enabled = True
        txtInfo.Enabled = True
        txtInfo.BackColor = vbWindowBackground
    End If
End Sub

Private Sub opt_DblClick(Index As Integer)
    mstrSelectInput = Replace(Split(opt(Index).Caption, "(")(0), "&", "")
    cmdDo_Click 0
End Sub

Private Sub txtInfo_KeyPress(KeyAscii As Integer)
    If InStr(";()',", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
