VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmModiPatiPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����޸�"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9135
   Icon            =   "frmModiPatiPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPass 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   180
      ScaleHeight     =   3495
      ScaleWidth      =   6720
      TabIndex        =   13
      Top             =   540
      Width           =   6720
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1110
         Width           =   4845
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   0
         TabIndex        =   14
         Top             =   750
         Width           =   6555
      End
      Begin VB.TextBox txtPati 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1650
         Width           =   4845
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1815
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   4125
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2175
         Width           =   1815
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2700
         Width           =   1815
      End
      Begin VB.TextBox txtAudi 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   4125
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   330
         TabIndex        =   17
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label lblNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "�뽫[XX]��ˢ���������Ữ����  Ȼ����������������ͬ�����룡"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1110
         TabIndex        =   15
         Top             =   120
         Width           =   5325
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   390
         TabIndex        =   1
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   390
         TabIndex        =   3
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   2235
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   105
         TabIndex        =   7
         Top             =   2760
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3390
         TabIndex        =   9
         Top             =   2760
         Width           =   630
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   120
         Picture         =   "frmModiPatiPass.frx":06EA
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7410
      TabIndex        =   12
      Top             =   825
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7410
      TabIndex        =   11
      Top             =   315
      Width           =   1500
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   4545
      Left            =   60
      TabIndex        =   16
      Top             =   150
      Width           =   7275
      _Version        =   589884
      _ExtentX        =   12832
      _ExtentY        =   8017
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmModiPatiPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'------------------------------------------------------
'���
Private mlngModule As Long, mlngCardTypeID As Long
Private mstrCardNo As String, mlng����ID As Long
'-------------------------------------------------------
Private mblnDO As Boolean
Private mobjKeyboard As Object
Private mblnOk As Boolean
Private mrsInfo As ADODB.Recordset
Private mobjCardObject As clsCardObject
Private mobjICCard As Object
Private mblnFirst As Boolean
Private mblnCheckOldPass As Boolean
Private WithEvents mobjIDCard As zlIDCard.clsIDCard '�����:54278
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents '�����:56597
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mobjSquare As Object '�����:56597

Public Function zlModifyPass(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional lng����ID As Long, Optional strCardNo As String, _
    Optional blnCheckOldPass As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���
    '���:frmMain-���õ�������
    '       lngModule -ģ���
    '       lngCardTypeId-�����ID
    '       lng����ID-����ID
    '       strCardNo-����
    '����:�޸ĳɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-07-29 11:08:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngCardTypeID = lngCardTypeID: mlngModule = lngModule: mlng����ID = lng����ID
    mstrCardNo = strCardNo: mblnOk = False
    mblnCheckOldPass = blnCheckOldPass
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlModifyPass = mblnOk
End Function
Private Sub InitTaskPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��InitTaskPancel
    '����:���˺�
    '����:2011-06-30 18:20:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    Set tkpGroup = wndTaskPanel.Groups.Add(1, "��ˢ���������޸�����")
    Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
   Set Item.Control = picPass
    tkpGroup.CaptionVisible = False
   ' Call Item.SetMargins(0, -19, 0, -4)
    picPass.BackColor = Item.BackColor
    Me.BackColor = Item.BackColor
    cmdOK.BackColor = Item.BackColor
    cmdCancel.BackColor = Item.BackColor
    tkpGroup.Expandable = False
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
End Sub
Private Function InitCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ƭ��Ϣ
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-29 14:25:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
   Set mobjCardObject = zlGetClsCardObject(mlngCardTypeID, False)
   If Err <> 0 Then Err = 0: Exit Function
   If mobjCardObject Is Nothing Then Exit Function
   If mobjCardObject.CardPreporty.���� = "���￨" And mobjCardObject.CardPreporty.ϵͳ Then
            lbl����.BorderStyle = 1: lbl����.Tag = "1"
   Else
        If mobjCardObject.CardPreporty.�Ƿ�Ӵ�ʽ���� Then
            lbl����.BorderStyle = 1: lbl����.Tag = "1"
        Else
            lbl����.BorderStyle = 0: lbl����.Tag = "0"
        End If
    End If
    '108779�����ϴ�,2017/5/8,���벻�̶��Ͳ�Ӧ������10λ
    If mobjCardObject.CardPreporty.���볤������ <> 0 Then
        txtPass.MaxLength = mobjCardObject.CardPreporty.���볤��
        txtAudi.MaxLength = mobjCardObject.CardPreporty.���볤��
    End If
    lblNotes.Caption = Replace(lblNotes.Caption, "[XX]", "[" & mobjCardObject.CardPreporty.���� & "]")
    InitCardInfor = True
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������봴��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:��ɳɹ�,����true,����False
    '����:���˺�
    '����:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub cmdCancel_Click()
    If txtPati.Text <> "" And Val(txtPati.Tag) <> 0 Then
        Call ClearFace:
        If txt����.Enabled Then txt����.SetFocus
        Exit Sub
    End If
    mstrCardNo = ""
    Unload Me
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ƿ���Ч
    '����:������Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-29 11:15:42
    '---------------------------------------------------------------------------------------------------------------------------------------------\
    Dim str�ƺ� As String
    str�ƺ� = IIf(glngSys Like "8??", "�ͻ�", "����")
    
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then
        MsgBox "���ܶ�ȡ" & str�ƺ� & "��Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
        Call ClearFace: txt����.SetFocus: Exit Function
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "���ܶ�ȡ" & str�ƺ� & "��Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
        Call ClearFace: txt����.SetFocus: Exit Function
    End If
    If txtPass.Text <> txtAudi.Text Then
        MsgBox "������������벻һ�£����������룡", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
        Exit Function
    End If
    If txtPass.Text = "" Then
        Select Case mobjCardObject.CardPreporty.������������
            Case 0 '������
            Case 1 'δ��������
                If MsgBox("δ�������뽫��Ӱ���ʻ���ʹ�ð�ȫ,�Ƿ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
                    Exit Function
                End If
            Case 2 'Ϊ�����ֹ
                MsgBox "δ���뿨����,���ܽ��з�����", vbExclamation, gstrSysName
                If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
                Exit Function
        End Select
    Else
        '108779:���ϴ�,2017/5/8,������볤��
        If txtPass.Visible Then
            Select Case mobjCardObject.CardPreporty.���볤������
            Case 0
            Case 1
                If Len(txtPass.Text) <> mobjCardObject.CardPreporty.���볤�� Then
                    MsgBox "ע��:" & vbCrLf & "�����������" & mobjCardObject.CardPreporty.���볤�� & "λ", vbOKOnly + vbInformation
                    txtPass.Text = "": txtAudi.Text = ""
                    If txtPass.Enabled Then txtPass.SetFocus
                    Exit Function
                 End If
            Case Else
                If Len(txtPass.Text) < Abs(mobjCardObject.CardPreporty.���볤������) Then
                    MsgBox "ע��:" & vbCrLf & "�����������" & Abs(mobjCardObject.CardPreporty.���볤������) & "λ����.", vbOKOnly + vbInformation
                    txtPass.Text = "": txtAudi.Text = ""
                    If txtPass.Enabled Then txtPass.SetFocus
                    Exit Function
                 End If
            End Select
        End If
    End If
        
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ModifPatiPass() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ĳ��˵�����
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-29 11:18:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim lng����ID As Long, Curdate As Date, cllPro As Collection
   Dim strSQL As String, strPassWord As String
   
    On Error GoTo errHandle
    strPassWord = zlCommFun.zlStringEncode(txtPass.Text)     '�������
    lng����ID = Val(Nvl(mrsInfo!����ID))
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
      'Zl_ҽ�ƿ��䶯_Insert
       strSQL = "Zl_ҽ�ƿ��䶯_Insert("
      '      �䶯����_In   Number,
      '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
      strSQL = strSQL & "" & 5 & ","
      '      ����id_In     סԺ���ü�¼.����id%Type,
      strSQL = strSQL & "" & lng����ID & ","
      '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
      strSQL = strSQL & "" & mlngCardTypeID & ","
      '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
      strSQL = strSQL & "'" & mstrCardNo & "',"
      '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
      strSQL = strSQL & "'" & mstrCardNo & "',"
      '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
      '      --�䶯ԭ��_In:�������������䶯ԭ��Ϊ����.���ܵ�
      strSQL = strSQL & "'" & "�������" & "',"
      '      ����_In       ������Ϣ.����֤��%Type,
      strSQL = strSQL & "'" & strPassWord & "',"
      '      ����Ա����_In סԺ���ü�¼.����Ա����%Type,
      strSQL = strSQL & "'" & UserInfo.���� & "',"
      '      �䶯ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
      strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
      '      Ic����_In     ������Ϣ.Ic����%Type := Null,
      strSQL = strSQL & "NULL,"
      '      ��ʧ��ʽ_In   ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null
      strSQL = strSQL & "NULL)"
     Call zlAddArray(cllPro, strSQL)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ModifPatiPass = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
ErrSaveRollTo:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If ModifPatiPass = False Then Exit Sub
    MsgBox "�����޸ĳɹ�!", vbOKOnly + vbInformation, gstrSysName
    mblnOk = True
    mstrCardNo = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitCardInfor = False Then Unload Me: Exit Sub
    Call ClearFace
    If mstrCardNo <> "" Then
        If GetPatient(mstrCardNo) = False Then
            Call ClearFace: If txt����.Enabled Then txt����.SetFocus
            Exit Sub
        End If
        If txtPass.Enabled Then txtPass.SetFocus
    Else
        If txt����.Enabled Then txt����.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    If glngSys Like "8??" Then lbl����.Caption = "�ͻ�"
    
    Call CreateObjectKeyboard
    Call InitTaskPancel
    '�����:56597
    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Set mobjCommEvents = Nothing
End Sub

Private Sub lbl����_Click()
    Dim strCardNo As String, strOutXml As String, strExpand As String
  
    If mlngCardTypeID = 0 Then Exit Sub
    If mobjCardObject.CardObject Is Nothing Then Exit Sub
    If Not mobjCardObject.CardPreporty.�Ƿ�Ӵ�ʽ���� Then Exit Sub
    
    If mobjCardObject.CardPreporty.���� Like "IC��*" And mobjCardObject.CardPreporty.ϵͳ = True Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txt����.Text = mobjICCard.Read_Card()
            If txt����.Text <> "" Then
                If Not GetPatient(txt����.Text) Then
                    Call ClearFace
                    txt����.SetFocus: Exit Sub
                End If
            End If
        End If
        Exit Sub
    End If
    
    '�����:54278
    If mobjCardObject.CardPreporty.���� Like "*���֤*" And mobjCardObject.CardPreporty.�ӿڳ����� = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled True
        Exit Sub
    End If
    If mobjCardObject.CardObject.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    
    txt����.Text = Trim(strCardNo)
    If Trim(txt����.Text) = "" Then Exit Sub
    If Not GetPatient(txt����.Text) Then
        Call ClearFace: If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Sub
    End If
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
'�����:56597
    If strCardType <> "" Then mlngCardTypeID = Val(strCardType)
    If strCardNo = "" Or strCardType = "" Then Exit Sub
    If Not GetPatient(strCardNo) Then
        Call ClearFace: If txt����.Enabled And txt����.Visible Then txt����.SetFocus
        Exit Sub
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
        '�����:54278
        txt����.Text = Trim(strID)
        If Trim(txt����.Text) = "" Then Exit Sub
        If Not GetPatient(txt����.Text) Then
            Call ClearFace: If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Exit Sub
        End If
End Sub

Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    OpenPassKeyboard txtAudi, True
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    '108779�����ϴ�,2017/5/8,���벻�̶��Ͳ�Ӧ������10λ
    Call CheckInputPassWord(KeyAscii, mobjCardObject.CardPreporty.������� = 1)
    If KeyAscii = 13 Then
        KeyAscii = 0: cmdOK_Click
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAudi_LostFocus()
   ClosePassKeyboard txtPass
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    OpenPassKeyboard txtPass, False
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    '108779�����ϴ�,2017/5/8,���벻�̶��Ͳ�Ӧ������10λ
    Call CheckInputPassWord(KeyAscii, mobjCardObject.CardPreporty.������� = 1)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            cmdOK.SetFocus
        Else
            txtAudi.SetFocus
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub

Private Sub txtPati_GotFocus()
    zlControl.TxtSelAll txtPati
End Sub
Private Function GetPatient(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '����:���˺�
    '����:2011-07-29 11:34:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, strWhere As String
    
    Set mrsInfo = Nothing
    On Error GoTo errH
    '�������ĺ���
    If GetPatiID(mlngCardTypeID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
    If lng����ID = 0 Then GoTo NotFoundPati:
    mstrCardNo = strInput
    
    If lng����ID <= 0 Then GoTo NotFoundPati:
    strSQL = "" & _
    "   Select ����ID,�����,סԺ��,���￨��,����,�Ա�,����" & _
    "   From ������Ϣ " & _
    "   Where ����ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    txtPati.Text = Nvl(mrsInfo!����)
    txtPati.Tag = Val(mrsInfo!����ID)
    txtSex.Text = Nvl(mrsInfo!�Ա�)
    txtAge.Text = Nvl(mrsInfo!����)
    txtPass.Text = "": txtAudi.Text = ""
    txtPass.Tag = strPassWord
    If mblnCheckOldPass Then
        If zlCommFun.VerifyPassWord(Me, strPassWord, txtPati.Text, txtSex.Text, txtAge.Text, True) = False Then
            Call ClearFace
            Exit Function
        End If
    End If
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = Nothing
    Exit Function
NotFoundPati:
    If strErrMsg = "" Then
        MsgBox "���ܶ�ȡ" & IIf(glngSys Like "8??", "�ͻ�", "����") & "��Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
    End If
    Set mrsInfo = Nothing
End Function
Private Sub ClearFace()
    txt����.PasswordChar = IIf(mobjCardObject.CardPreporty.�������Ĺ��� <> "", "*", "")
    txt����.Text = ""
    txtPass.Text = "": txtPati.Text = ""
    txtSex.Text = "": txtAge.Text = ""
    txtPass.Text = "": txtAudi.Text = ""
End Sub

Private Sub txt����_Change()
    txtPass.Enabled = Trim(txt����.Text) <> ""
    txtPass.BackColor = IIf(txtPass.Enabled = False, txtPati.BackColor, txt����.BackColor)
    txtAudi.Enabled = Trim(txt����.Text) <> ""
    txtAudi.BackColor = IIf(txtAudi.Enabled = False, txtPati.BackColor, txt����.BackColor)
End Sub

Private Sub txt����_GotFocus()
    Dim strExpend As String
    
    On Error GoTo Errhand
    zlControl.TxtSelAll txt����
    txt����.PasswordChar = IIf(mobjCardObject.CardPreporty.�������Ĺ��� <> "", "*", "")
    '�����:56597
    '��ʼ��IC��
    If mobjCardObject.CardPreporty.���� Like "IC��*" And mobjCardObject.CardPreporty.ϵͳ = True Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        Exit Sub
    End If
    '��ʼ���������֤
    If mobjCardObject.CardPreporty.���� Like "*���֤*" And mobjCardObject.CardPreporty.�ӿڳ����� = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled True
        Exit Sub
    End If
    If mobjSquare Is Nothing Then Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
    '��ʼ����Ƶ������
    '86152:���ϴ�,2015/7/6,��ʼ������
    If Err <> 0 Then Exit Sub
    mobjSquare.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    mobjSquare.zlInitEvents Me.hWnd, mobjCommEvents
    mobjSquare.SetEnabled True
    
    '85565:���ϴ�,2015/7/21,����ˢ���ӿ�
    Err = 0: On Error Resume Next
    If mobjCardObject.CardPreporty.�ӿ���� = 0 Or mobjCardObject.CardPreporty.�ӿڳ����� = "" Then Exit Sub
    If Not (mobjCardObject.CardPreporty.�Ƿ�ˢ�� Or mobjCardObject.CardPreporty.�Ƿ�ɨ��) Then Exit Sub
    
    Call mobjSquare.zlSetBrushCardObject(mobjCardObject.CardPreporty.�ӿ����, txt����, strExpend, _
                                        mobjCardObject.CardPreporty.���ѿ�)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
     '�����:58066
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If InStr(":��;��?��'��||", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
     
     If (Len(txt����.Text) = mobjCardObject.CardPreporty.���ų��� - 1 And KeyAscii <> 8) Or (KeyAscii = 13 And Trim(txt����.Text) <> "") Then
            If KeyAscii <> 13 Then
                txt����.Text = txt����.Text & Chr(KeyAscii)
                txt����.SelStart = Len(txt����.Text)
            End If
            KeyAscii = 0
            If Not GetPatient(txt����.Text) Then
                Call ClearFace
                txt����.SetFocus: Exit Sub
            End If
            txtPass.SetFocus
        End If
End Sub

Private Sub txt����_LostFocus()
    '�����:56597
   If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled False
   If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
   If Not mobjSquare Is Nothing Then mobjSquare.SetEnabled False
   
End Sub
