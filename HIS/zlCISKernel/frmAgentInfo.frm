VERSION 5.00
Begin VB.Form frmAgentInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������Ϣ"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "frmAgentInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraAgent 
      Caption         =   "������"
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   3615
      Begin VB.TextBox txtReason 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   2160
         Width           =   2130
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   200
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1425
         Width           =   600
      End
      Begin VB.PictureBox picAge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1770
         ScaleHeight     =   255
         ScaleWidth      =   765
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1380
         Width           =   765
         Begin VB.ComboBox cboAge 
            Height          =   300
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   -30
            Width           =   700
         End
      End
      Begin VB.TextBox txtAgentPhone 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   200
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1800
         Width           =   2130
      End
      Begin VB.PictureBox picSex 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   1425
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1425
         Begin VB.ComboBox cboSex 
            Height          =   300
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   -30
            Width           =   1395
         End
      End
      Begin VB.TextBox txtAgentName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   200
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   1
         Top             =   300
         Width           =   2130
      End
      Begin VB.TextBox txtAgentIDNO 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   200
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   2
         Top             =   660
         Width           =   2130
      End
      Begin VB.Line line 
         Index           =   8
         X1              =   1080
         X2              =   3240
         Y1              =   2900
         Y2              =   2900
      End
      Begin VB.Label lblReason 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ����"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   720
      End
      Begin VB.Line line 
         Index           =   7
         X1              =   1065
         X2              =   3225
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Line line 
         Index           =   6
         X1              =   2040
         X2              =   3225
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line line 
         Index           =   5
         X1              =   1065
         X2              =   2115
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Line line 
         Index           =   4
         X1              =   1065
         X2              =   3225
         Y1              =   1305
         Y2              =   1305
      End
      Begin VB.Label lblAgentPhone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�绰"
         Height          =   180
         Left            =   600
         TabIndex        =   17
         Top             =   1800
         Width           =   360
      End
      Begin VB.Label lblAgentAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   600
         TabIndex        =   16
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label lblAgentSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   600
         TabIndex        =   15
         Top             =   1065
         Width           =   360
      End
      Begin VB.Line line 
         Index           =   3
         X1              =   1065
         X2              =   3225
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line line 
         Index           =   2
         X1              =   1065
         X2              =   3225
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label lblAgentName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   600
         TabIndex        =   12
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblAgentIDNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   660
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   14
      Top             =   4455
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1440
      TabIndex        =   13
      Top             =   4455
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Caption         =   "����"
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtPatiIDNO 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   200
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   0
         Top             =   660
         Width           =   2130
      End
      Begin VB.Line line 
         Index           =   0
         X1              =   3225
         X2              =   1065
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line line 
         Index           =   1
         X1              =   1065
         X2              =   3225
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label lblPatiName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����  ���"
         Height          =   180
         Left            =   585
         TabIndex        =   9
         Top             =   285
         Width           =   1080
      End
      Begin VB.Label lblPatiIDNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   660
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmAgentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mblnOK As Boolean
Private mlng����ID As Long
Private mlng����ID As Long
Private mstr���� As String
Private mstr�Ա� As String
Private mstr���� As String
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mlng���� As Long '0-���1-סԺ

Public Function ShowMe(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal str�������� As String, _
                ByVal str�������֤�� As String, ByVal str���������� As String, ByVal str���������֤�� As String, ByVal str�������Ա� As String, ByVal str���������� As String, ByVal str�����˵绰 As String, ByVal str��ҩ���� As String) As Boolean
    Screen.MousePointer = 0
    mblnOK = False
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mstr���� = str����������
    mstr�Ա� = str�������Ա�
    mstr���� = str��������
    
    If Not frmParent Is Nothing Then
        mlng���� = IIF(frmParent.Name = "frmInAdviceEdit", 1, 0)
    End If
    
    lblPatiName.Caption = "����  " & str��������
    txtPatiIDNO.Text = str�������֤��
    txtAgentName.Text = str����������
    txtAgentIDNO.Text = str���������֤��
    txtAgentPhone.Text = str�����˵绰
    txtReason.Text = str��ҩ����
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strInfo As String
    Dim lngTmp As Long
    Dim strMask As String
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    If Trim(txtPatiIDNO.Text) = "" Then
        MsgBox "�����벡�����֤�ţ�", vbInformation, gstrSysName
        txtPatiIDNO.SetFocus: Exit Sub
    End If
    
    If Len(txtPatiIDNO.Text) <> 15 And Len(txtPatiIDNO.Text) <> 18 Then
        MsgBox "���֤�ų��Ȳ���ȷ��������15��18λ���֤�ţ�", vbInformation, gstrSysName
        txtPatiIDNO.SetFocus: Exit Sub
    End If
    
    'ֻ����д�˴�����������ż����Ϊ������Ҳ�������жϵ�
    If txtAgentName.Text <> "" Then
        If txtAgentIDNO.Text <> "" And Len(txtAgentIDNO.Text) <> 15 And Len(txtAgentIDNO.Text) <> 18 Then
            MsgBox "���֤�ų��Ȳ���ȷ��������15��18λ���֤�ţ�", vbInformation, gstrSysName
            txtAgentIDNO.SetFocus: Exit Sub
        End If
        
        If Trim(txtAgentIDNO.Text) = Trim(txtPatiIDNO.Text) Then
            MsgBox "���������֤���벡�����֤����ͬ�����������룡", vbInformation, gstrSysName
            txtAgentIDNO.SetFocus: Exit Sub
        End If
        
        '���䣬��̨�ֶγ�100
        strInfo = Trim(txtAge.Text)
        If strInfo <> "" Then
            If Not IsNumeric(strInfo) Then
                MsgBox "�������Ϊ���֣���������ȷ�����䣡", vbInformation, gstrSysName
                txtAge.SetFocus: Exit Sub
            ElseIf Len(strInfo) > 6 Then
                MsgBox "���䳤����಻�ܳ���6λ����������ȷ����������䵥��", vbInformation, gstrSysName
                txtAge.SetFocus: Exit Sub
            End If
        End If
        '�绰
        strInfo = Trim(txtAgentPhone.Text)
        If strInfo <> "" And Len(strInfo) > 20 Then
            MsgBox "�绰���Ȳ���ȷ��������಻�ܳ���20λ����������ȷ�ĵ绰�ţ�", vbInformation, gstrSysName
            txtAgentPhone.SetFocus: Exit Sub
        End If
        If strInfo <> "" Then
            strMask = "1234567890-()"
            lngTmp = Len(strInfo)
            strTmp = strInfo
            For i = 1 To lngTmp
                If InStr(strMask, Mid(strTmp, i, 1)) = 0 Then
                    MsgBox "�绰-�����а����Ƿ��ַ�(����¼�������ַ�����" & strMask & "��)��", vbInformation, gstrSysName
                    txtAgentPhone.SetFocus: Exit Sub
                End If
            Next
        End If
        
        strInfo = Trim(txtAge.Text) & cboAge.Text
        gstrSQL = "Zl_��������Ϣ_Insert(" & mlng����ID & ",'" & Trim(txtPatiIDNO.Text) & "','" & Trim(txtAgentName.Text) & "','" & _
                Trim(txtAgentIDNO.Text) & "'," & mlng����ID & ",'" & Split(cboSex.Text, "-")(1) & "','" & strInfo & "','" & Trim(txtAgentPhone.Text) & "','" & Trim(txtReason.Text) & "')"
    Else
        gstrSQL = "Zl_��������Ϣ_Insert(" & mlng����ID & ",'" & Trim(txtPatiIDNO.Text) & "',null,null," & mlng����ID & ")"
    End If
    Screen.MousePointer = 11
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Screen.MousePointer = 0
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If txtPatiIDNO.Text <> "" Then
        txtAgentName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim str���� As String
    
    On Error GoTo errH

    If InStr(GetInsidePrivs(IIF(mlng���� = 0, p����ҽ��վ, pסԺҽ���´�)), "��������Ϣ��������¼��") = 0 Then
        txtPatiIDNO.Locked = True
        txtAgentName.Locked = True
        txtAgentIDNO.Locked = True
    End If
    Me.Caption = Me.Caption & "  (���֤ˢ��¼��)"
    strSQL = "select ����||'-'||���� as ����, rownum as id,ȱʡ��־,���� from �Ա� order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Call Cbo.AddData(cboSex, rsTmp, True)
    If mstr�Ա� = "" Then
        rsTmp.Filter = "ȱʡ��־=1"
        If Not rsTmp.EOF Then mstr�Ա� = rsTmp!���� & ""
    End If
    Call Cbo.Locate(cboSex, mstr�Ա�)
    
    With cboAge
        .Clear
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .AddItem "Сʱ"
        .AddItem "����"
    End With
    
    If InStr(mstr����, "��") > 0 Then
        str���� = Replace(mstr����, "��", "")
        cboAge.ListIndex = 0
    ElseIf InStr(mstr����, "��") > 0 Then
        str���� = Replace(mstr����, "��", "")
        cboAge.ListIndex = 1
    ElseIf InStr(mstr����, "��") > 0 Then
        str���� = Replace(mstr����, "��", "")
        cboAge.ListIndex = 2
    ElseIf InStr(mstr����, "Сʱ") > 0 Then
        str���� = Replace(mstr����, "Сʱ", "")
        cboAge.ListIndex = 3
    ElseIf InStr(mstr����, "����") > 0 Then
        str���� = Replace(mstr����, "����", "")
        cboAge.ListIndex = 4
    Else
        cboAge.ListIndex = 0
    End If
     
    txtAge.Text = str����
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Screen.MousePointer = 11
End Sub



Private Sub txtAgentIDNO_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtAgentIDNO)
End Sub

Private Sub txtAgentIDNO_GotFocus()
    zlControl.TxtSelAll txtAgentIDNO
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtAgentIDNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789X" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtAgentIDNO_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtAgentIDNO.Text = Trim(txtAgentIDNO.Text)
End Sub

Private Sub txtAgentName_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtAgentName)
End Sub

Private Sub txtAgentName_GotFocus()
    zlControl.TxtSelAll txtAgentName
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtAgentName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtAgentName_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtAgentName.Text = Trim(txtAgentName.Text)
End Sub

Private Sub txtPatiIDNO_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtPatiIDNO)
End Sub

Private Sub txtPatiIDNO_GotFocus()
    zlControl.TxtSelAll txtPatiIDNO
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtPatiIDNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789X" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPatiIDNO_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtPatiIDNO.Text = Trim(txtPatiIDNO.Text)
End Sub


Private Sub cboSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtAgentPhone_GotFocus()
    zlControl.TxtSelAll txtAgentPhone
End Sub

Private Sub txtAgentPhone_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890-()" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtAge_GotFocus()
    zlControl.TxtSelAll txtAge
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub



Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
On Error GoTo errH
    If Me.ActiveControl Is txtPatiIDNO Then
        If mstr���� = strName Then
            txtPatiIDNO.Text = strID
        Else
            MsgBox "�����Ϣ¼��ʧ��,��ʹ�õ�ǰ���˵����֤ˢ����", vbInformation, gstrSysName
        End If
    ElseIf Me.ActiveControl Is txtAgentName Or Me.ActiveControl Is txtAgentIDNO Then
        txtAgentName.Text = strName
        txtAgentIDNO.Text = strID
        txtAge.Text = GetOldAcademic(datBirthDay, "��")
        Call Cbo.Locate(cboSex, strSex)
        cboAge.Text = "��"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetOldAcademic(ByVal DateBir As Date, ByVal str���䵥λ As String) As Long
'���ܣ����ݵ�ǰ�ĳ������ں����䵥λ�����������ϵ�����ֵ
'���أ�����
    Dim datCur As Date, lngOld As Long, strInterval As String
    If DateBir = CDate(0) Or InStr(" ������", str���䵥λ) < 2 Then Exit Function
    
    datCur = zlDatabase.Currentdate
    
    strInterval = Switch(str���䵥λ = "��", "yyyy", str���䵥λ = "��", "m", str���䵥λ = "��", "d")
    lngOld = DateDiff(strInterval, DateBir, datCur)
    If DateAdd(strInterval, lngOld, DateBir) > datCur Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function


Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����֤������
    '����:����
    '����:2012-08-31 16:28:23
    '�����:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��ʼ���Կ�����
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hwnd)
    End If
    '�򿪶�����
    mobjIDCard.SetEnabled (blnEnabled)
End Sub
