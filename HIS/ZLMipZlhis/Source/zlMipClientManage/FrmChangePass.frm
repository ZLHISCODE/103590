VERSION 5.00
Begin VB.Form FrmChangePass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�޸�����"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4860
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CDMȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3510
      TabIndex        =   3
      Top             =   240
      Width           =   1230
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3510
      TabIndex        =   4
      Top             =   690
      Width           =   1230
   End
   Begin VB.Frame Fra���� 
      Caption         =   "��������"
      Height          =   1455
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   3165
      Begin VB.TextBox TXTȷ������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1005
         Width           =   1590
      End
      Begin VB.TextBox TXT������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   645
         Width           =   1590
      End
      Begin VB.TextBox TXTԭ���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1110
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label Lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   450
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Lbl������ 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   705
         Width           =   540
      End
      Begin VB.Label Lbl������֤ 
         AutoSize        =   -1  'True
         Caption         =   "������֤"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   1065
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'���
Private mfrmParent As Object '������
Private mstrUserName As String 'ԭʼ�û���
Private mstrPwd As String 'ԭʼ����
Private mstrServer As String 'ԭʼ������
Private mblnת�� As Boolean '�Ƿ�����Ҫת��
'ģ�����

Private mblnOk As Boolean
Public Function ShowMe(ByVal frmParent As Object, ByVal strUserName As String, ByRef strPWD As String, ByRef strServer As String, Optional ByVal blnTrans As Boolean) As Boolean
'���ܣ��޸�����
'������frmParent=������
'          strUserName=�û���
'          strPwd=����
'          strServer=������
    Set mfrmParent = frmParent
    mstrUserName = strUserName
    mstrPwd = strPWD
    mstrServer = strServer
    mblnת�� = blnTrans
    mblnOk = False
    Me.Show vbModal
    strUserName = mstrUserName
    strPWD = mstrPwd
    strServer = mstrServer
    ShowMe = mblnOk
End Function

Private Sub CDMȷ��_Click()
    Dim strPassword As String
    Dim strServer As String, strError As String
    Dim intPos As Integer
    Dim strSQL As String, rsData As New ADODB.Recordset
    Dim arrTmp As Variant, lngLen As Long, i As Long, intChr As Integer
    Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
    Dim blnPwdLen As Boolean, intPwdMin As Integer, intPwdMax As Integer
    Dim blnComplex As Boolean, strOterChrs As String
    Dim blnTransPassword As Boolean

    If Trim(TXTԭ����.Text) = "" Then
        MsgBox "����������룡", vbInformation, gstrSysName
        TXTԭ����.SetFocus
        Exit Sub
    End If
    If Trim(TXT������.Text) = "" Then
        MsgBox "�����������룡", vbInformation, gstrSysName
        TXT������.SetFocus
        Exit Sub
    End If
    If Trim(TXTȷ������.Text) = "" Then
        MsgBox "������������֤��", vbInformation, gstrSysName
        TXTȷ������.SetFocus
        Exit Sub
    End If
    If TXT������.Text <> TXTȷ������.Text Then
        MsgBox "����������������������룡", vbInformation, gstrSysName
        TXT������.SetFocus
        Exit Sub
    End If
    
    If TXT������.Text = Trim(TXTԭ����.Text) Then
        MsgBox "������;�������ȫһ�������������룡", vbInformation, gstrSysName
        TXT������.SetFocus
        Exit Sub
    End If
    
    strPassword = Trim(TXTԭ����.Text)
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If TXTԭ����.Enabled Then TXTԭ����.SetFocus
            MsgBox "���������", vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    '�����ַ���
    intPos = InStr(strPassword, "@")
    If intPos > 0 Then
        strServer = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    If strServer = "" Then
        strServer = mstrServer
    End If
    
    If Not gclsMsgOracle.OraDataOpen(strServer, mstrUserName, strPassword, True) Then
        Exit Sub
    Else
        gstrDbUser = UCase(mstrUserName)
        Call gclsBusiness.InitBusiness(gclsMsgOracle, "", gstrDbUser)
        
        strSQL = "Select ������,Nvl(����ֵ,ȱʡֵ) ����ֵ From zlOptions Where ������ in (20,21,22,23)"
        Set rsData = gclsMsgOracle.OpenSQLRecord(strSQL, Me.Caption)
        blnPwdLen = False: intPwdMin = 0: intPwdMax = 0
        blnComplex = False: strOterChrs = ""
        Do While Not rsData.EOF
            Select Case rsData!������
                Case 20 '�Ƿ�������볤��
                    blnPwdLen = Val(rsData!����ֵ & "") = 1
                Case 21 '���볤������
                    intPwdMin = Val(rsData!����ֵ & "")
                Case 22 '���볤������
                    intPwdMax = Val(rsData!����ֵ & "")
                Case 23 '�Ƿ�������븴�Ӷ�
                    blnComplex = Val(rsData!����ֵ & "") = 1
            End Select
            rsData.MoveNext
        Loop
        '����������ʾ
        If blnPwdLen Then
            If intPwdMin = intPwdMax Then
                TXT������.ToolTipText = "�������Ϊ" & intPwdMax & " λ�ַ���"
            Else
                TXT������.ToolTipText = "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���"
            End If
         End If
         If blnComplex Then
            If TXT������.ToolTipText <> "" Then
                TXT������.ToolTipText = TXT������.ToolTipText & vbNewLine & "���ٰ���һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
            Else
                TXT������.ToolTipText = "������һ�����֡�һ����ĸ��һ�������ַ���ɡ�"
            End If
         End If
         TXTȷ������.ToolTipText = TXT������.ToolTipText
         strPassword = Trim(TXT������.Text)
        '���ȼ��
        lngLen = ActualLen(strPassword)
        If lngLen <> Len(strPassword) Then
            MsgBox "���������˫�ֽ��ַ������飡", vbInformation, gstrSysName
            Exit Sub
        End If
        If blnPwdLen Then
            If Not (lngLen >= intPwdMin And lngLen <= intPwdMax) Then
                If intPwdMin = intPwdMax Then
                    MsgBox "�������Ϊ" & intPwdMax & " λ�ַ���", vbInformation, gstrSysName
                    Exit Sub
                Else
                    MsgBox "�������Ϊ" & intPwdMin & "��" & intPwdMax & " λ�ַ���", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        For i = 1 To Len(strPassword)
            intChr = Asc(UCase(Mid(strPassword, i, 1)))
            If intChr >= 32 And intChr < 127 Then
                'Dim blnHaveNum As Boolean, blnAlpha As Boolean, blnChar As Boolean
                Select Case intChr
                    Case 48 To 57 '����
                        blnHaveNum = True
                    Case 65 To 90 '��ĸ
                        blnAlpha = True
                    Case 32, 34, 47, 64  '�ո�,˫����,/,@
                        strOterChrs = strOterChrs & Chr(intChr)
                    Case Is < 48, 58 To 64, 91 To 96, Is > 122
                        blnChar = True
                End Select
            Else
                strOterChrs = strOterChrs & Chr(intChr)
            End If
        Next
        If strOterChrs <> "" Then
            MsgBox "���벻�����������ַ���" & strOterChrs, vbInformation, gstrSysName
            Exit Sub
        ElseIf Not (blnHaveNum And blnAlpha And blnChar) And blnComplex Then
            MsgBox "����������һ�����֡�һ����ĸ��һ�������ַ���ɡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If gobjRegister.UpdateUserPassword(mstrUserName, strPassword, blnTransPassword, strError) Then
            MsgBox "�����޸ĳɹ�", vbInformation + vbOKOnly, "��ʾ"
            mstrServer = strServer
            mstrPwd = strPassword
            mblnOk = True
        Else
            If strError <> "" Then
                MsgBox "�����޸�ʧ�ܣ�" & vbCrLf & strError, vbExclamation, "��ʾ"
            End If
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub CMD����_Click()
    mstrUserName = ""
    mstrPwd = ""
    mstrServer = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Height / 15, Me.Width / 15, SWP_NOSIZE + SWP_SHOWWINDOW)
    If mstrPwd <> "" And mstrUserName = mstrPwd Then
        TXTԭ����.Enabled = False
    ElseIf TXTԭ����.Text = "" Then
        TXTԭ����.SetFocus
    Else
        TXT������.SetFocus
    End If
End Sub

Private Sub Form_Load()
    TXTԭ����.Text = mstrPwd
End Sub

Private Sub TXTȷ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call CDMȷ��_Click
End Sub

Private Sub TXT������_GotFocus()
    GetFocus TXT������
End Sub

Private Sub TXT������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub TXTԭ����_GotFocus()
    GetFocus TXTԭ����
End Sub

Private Sub TXTȷ������_GotFocus()
    GetFocus TXTȷ������
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub TXTԭ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub
