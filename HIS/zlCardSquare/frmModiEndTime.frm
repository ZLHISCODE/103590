VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmModiEndTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ʹ��ʱ�����"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8730
   Icon            =   "frmModiEndTime.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picEndDate 
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   210
      ScaleHeight     =   3705
      ScaleWidth      =   6525
      TabIndex        =   10
      Top             =   450
      Width           =   6525
      Begin VB.CommandButton cmdDefualtSet 
         Caption         =   "����XX��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3690
         TabIndex        =   17
         Top             =   3240
         Visible         =   0   'False
         Width           =   2265
      End
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
         Left            =   1275
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   975
         Width           =   4665
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   -90
         TabIndex        =   11
         Top             =   810
         Width           =   7245
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
         Left            =   1275
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1545
         Width           =   4665
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
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1425
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
         Left            =   4305
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   435
         Left            =   4260
         TabIndex        =   6
         Top             =   2730
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   248446978
         UpDown          =   -1  'True
         CurrentDate     =   .999988425925926
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   435
         Left            =   2235
         TabIndex        =   5
         Top             =   2730
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   767
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483646
         CustomFormat    =   "yyyy-MM-dd hh:mm"
         Format          =   248446979
         CurrentDate     =   401769
      End
      Begin VB.CheckBox chkEndDate 
         Caption         =   "��ֹʱ��"
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
         Left            =   555
         TabIndex        =   4
         Top             =   2790
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   150
         Picture         =   "frmModiEndTime.frx":6852
         Top             =   30
         Width           =   720
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
         Left            =   510
         TabIndex        =   16
         Top             =   1020
         Width           =   660
      End
      Begin VB.Label lblNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "�뽫[XX]��ˢ���������Ữ����  Ȼ��ѡ����Ҫ���ĵ����ڣ�"
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
         Left            =   510
         TabIndex        =   14
         Top             =   1605
         Width           =   630
      End
      Begin VB.Label label3 
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
         Index           =   0
         Left            =   510
         TabIndex        =   13
         Top             =   2190
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
         Left            =   3555
         TabIndex        =   12
         Top             =   2190
         Width           =   690
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
      Height          =   360
      Left            =   7380
      TabIndex        =   8
      Top             =   900
      Width           =   1230
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
      Height          =   360
      Left            =   7380
      TabIndex        =   7
      Top             =   360
      Width           =   1230
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   4755
      Left            =   120
      TabIndex        =   9
      Top             =   90
      Width           =   7095
      _Version        =   589884
      _ExtentX        =   12515
      _ExtentY        =   8387
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmModiEndTime"
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
Private mobjCard As Card

Private mblnFirst As Boolean
Private mblnCheckOldPass As Boolean
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private WithEvents mobjIDCard As clsIDCard '�����:54278
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents '�����:56597
Attribute mobjCommEvents.VB_VarHelpID = -1
 
Private mstrPrivs As String
Public Function zlModifyEndTime(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional lng����ID As Long, Optional strCardNo As String, _
    Optional strPrivs As String) As Boolean
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
    mstrPrivs = strPrivs
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlModifyEndTime = mblnOk
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
    Set tkpGroup = wndTaskPanel.Groups.Add(1, "��ˢ��ѡ���޸�����")
    Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
    Set Item.Control = picEndDate
    tkpGroup.CaptionVisible = False
   ' Call Item.SetMargins(0, -19, 0, -4)
   
    picEndDate.BackColor = Item.BackColor
    Me.BackColor = Item.BackColor
    cmdOK.BackColor = Item.BackColor
    chkEndDate.BackColor = Item.BackColor
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
    Dim rsTemp As ADODB.Recordset
    On Error Resume Next
    
    If gobjOneCardComLib.zlGetCard(mlngCardTypeID, False, mobjCard) = False Then Exit Function
    If mobjCard Is Nothing Then Exit Function
    
    If mobjCard.���� = "���￨" And mobjCard.ϵͳ Then
             lbl����.BorderStyle = 1: lbl����.Tag = "1"
    Else
         If mobjCard.�Ƿ�Ӵ�ʽ���� Then
             lbl����.BorderStyle = 1: lbl����.Tag = "1"
         Else
             lbl����.BorderStyle = 0: lbl����.Tag = "0"
         End If
     End If
     
    lblNotes.Caption = Replace(lblNotes.Caption, "[XX]", "[" & mobjCard.���� & "]")
    If InStr(mobjCard.ȱʡ��Чʱ��, "��") Then
        cmdDefualtSet.Caption = "����" & Val(mobjCard.ȱʡ��Чʱ��) & "��(&A)"
        cmdDefualtSet.Tag = Val(nvl(rsTemp!ȱʡ��Чʱ��)) & "��"
    ElseIf InStr(mobjCard.ȱʡ��Чʱ��, "��") Then
        cmdDefualtSet.Caption = "����" & Val(mobjCard.ȱʡ��Чʱ��) & "��(&A)"
        cmdDefualtSet.Tag = Val(mobjCard.ȱʡ��Чʱ��) & "��"
    End If
    If mobjCard.ȱʡ��Чʱ�� <> "" And Val(mobjCard.ȱʡ��Чʱ��) > 0 Then cmdDefualtSet.Visible = True
       
     InitCardInfor = True
End Function

Private Sub chkEndDate_Click()
    dtpDate.Enabled = chkEndDate.value
    dtpTime.Enabled = chkEndDate.value
End Sub

Private Sub chkEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
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
    Dim str�ƺ� As String, Curdate As Date, CardEndDate As Date
    str�ƺ� = IIf(glngSys Like "8??", "�ͻ�", "����")

    On Error GoTo errHandle
    
    If Not zlstr.IsHavePrivs(mstrPrivs, "ʹ��ʱ�����") Then
        MsgBox "��û��Ȩ�޸���ʹ��ʱ�䣬������ģ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mrsInfo Is Nothing Then
        MsgBox "���ܶ�ȡ" & str�ƺ� & "��Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
        Call ClearFace: txt����.SetFocus: Exit Function
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "���ܶ�ȡ" & str�ƺ� & "��Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
        Call ClearFace: txt����.SetFocus: Exit Function
    End If
    Curdate = zlDatabase.Currentdate
    CardEndDate = Format(CStr(dtpDate.value) & " " & CStr(dtpTime.value), "YYYY-MM-DD HH:MM:SS")
    If CardEndDate < Curdate And chkEndDate.value = vbChecked Then
        MsgBox "��ѡ����ڵ�ǰʱ�����ֹ���ڽ��и��ģ�", vbInformation, gstrSysName
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ModifCardEndTime() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ŀ���Ч��ֹʱ��
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, Curdate As Date, cllPro As Collection
    Dim strSQL As String, strPassWord As String, strEndDate As String

    On Error GoTo errHandle
    lng����ID = Val(nvl(mrsInfo!����ID))
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
    strEndDate = CStr(dtpDate.value) & " " & CStr(dtpTime.value)
    
    'Zl_ҽ�ƿ��䶯_Insert_S
     strSQL = "Zl_ҽ�ƿ��䶯_Insert_S("
    '  �䶯����_In     Number,
    strSQL = strSQL & "7,"
    '  ����id_In       ����ҽ�ƿ���Ϣ.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �����id_In     ����ҽ�ƿ���Ϣ.�����id%Type,
    strSQL = strSQL & "" & mlngCardTypeID & ","
    '  ԭ����_In       ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & mstrCardNo & "',"
    '  ҽ�ƿ���_In     ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & mstrCardNo & "',"
    '  �䶯ԭ��_In     ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
    strSQL = strSQL & "'" & "��ֹʱ�����" & "',"
    '  ����_In         ����ҽ�ƿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & strPassWord & "',"
    '  ����Ա����_In   ����ҽ�ƿ��䶯.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �䶯ʱ��_In     ����ҽ�ƿ��䶯.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "to_date('" & Format(Curdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ��ʧ��ʽ_In     ����ҽ�ƿ��䶯.��ʧ��ʽ%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ��ֹʹ��ʱ��_In ����ҽ�ƿ���Ϣ.��ֹʹ��ʱ��%Type := Null,
    strSQL = strSQL & IIf(chkEndDate.value = vbUnchecked, "Null", "to_date('" & Format(strEndDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')") & ")"
    '  ����_In         ����ҽ�ƿ��䶯.����%Type := Null,
    '  ������_In       ����ҽ�ƿ��䶯.������%Type := Null,
    '  ���õ���_In     ����ҽ�ƿ��䶯.���õ���%Type := Null,
    '  Ԥ������_In     ���˽����쳣��¼.Ԥ������%Type := Null,
    '  �䶯id_In       ����ҽ�ƿ��䶯.Id%Type := Null,
    '  �쳣��־_In     Number := 0,
    '  �쳣id_In       ���˽����쳣��¼.Id%Type := Null,
    '  Ԥ�����_In     ���˽����쳣��¼.Ԥ�����%Type := Null
    Call zlAddArray(cllPro, strSQL)
    On Error GoTo ErrSaveRollTo:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ModifCardEndTime = True
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

Private Sub cmdDefualtSet_Click()
    If Format(dtpDate, "yyyy-MM-dd") < Format("3000-01-01", "yyyy-MM-dd") Then
        If InStr(cmdDefualtSet.Tag, "��") Then
            dtpDate = DateAdd("D", Val(cmdDefualtSet.Tag), dtpDate)
        ElseIf InStr(cmdDefualtSet.Tag, "��") Then
            dtpDate = DateAdd("M", Val(cmdDefualtSet.Tag), dtpDate)
        End If
        cmdOK.SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If ModifCardEndTime = False Then Exit Sub
    MsgBox "��ֹʹ��ʱ���޸ĳɹ�!", vbOKOnly + vbInformation, gstrSysName
    mblnOk = True
    mstrCardNo = ""
    Unload Me
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim Curdate As Date
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If InitCardInfor = False Then Unload Me: Exit Sub
    Call ClearFace
    If mstrCardNo <> "" Then
        If GetPatient(mstrCardNo) = False Then
            Call ClearFace: If txt����.Enabled Then txt����.SetFocus
            Exit Sub
        End If
    Else
        If txt����.Enabled Then txt����.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    If glngSys Like "8??" Then lbl����.Caption = "�ͻ�"
    
    Call InitTaskPancel
    '�����:56597
    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    Set mobjCommEvents = Nothing
    Set mrsInfo = Nothing
End Sub

Private Sub lbl����_Click()
    Dim strCardNo As String, strOutXml As String, strExpand As String
  
    If mlngCardTypeID = 0 Then Exit Sub
    If mobjCard.CardObject Is Nothing Then Exit Sub
    If Not mobjCard.�Ƿ�Ӵ�ʽ���� Then Exit Sub
    
    If mobjCard.���� Like "IC��*" And mobjCard.ϵͳ = True Then
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
    If mobjCard.���� Like "*���֤*" And mobjCard.�ӿڳ����� = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled True
        Exit Sub
    End If
    If gobjOneCardComLib.zlReadCard(Me, mlngModule, mobjCard.�ӿ����, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    
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
    If GetPatiID(mlngCardTypeID, strInput, False, lng����ID, strPassWord, strErrMsg, , , , , , , , , , True) = False Then GoTo NotFoundPati:
    If lng����ID = 0 Then GoTo NotFoundPati:
    mstrCardNo = strInput
    
    If lng����ID <= 0 Then GoTo NotFoundPati:
    strSQL = "" & _
        "Select a.����id, a.�����, a.סԺ��, a.���￨��, a.����, a.�Ա�, a.����, b.��ֹʹ��ʱ��" & vbNewLine & _
        "From ������Ϣ a, ����ҽ�ƿ���Ϣ b" & vbNewLine & _
        "Where a.����id = b.����id And b.���� = [1] And b.�����id = [2] And a.����id = [3]"

    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, mlngCardTypeID, lng����ID)
    If mrsInfo.EOF Then Exit Function
    txtPati.Text = nvl(mrsInfo!����)
    txtPati.Tag = Val(mrsInfo!����ID)
    txtSex.Text = nvl(mrsInfo!�Ա�)
    txtAge.Text = nvl(mrsInfo!����)
    If nvl(mrsInfo!��ֹʹ��ʱ��) <> "" Then
        dtpDate = Format(nvl(mrsInfo!��ֹʹ��ʱ��), "yyyy-MM-dd")
        dtpTime = Format(nvl(mrsInfo!��ֹʹ��ʱ��), "HH:mm:ss")
        chkEndDate.value = 1
    Else
        dtpDate = Format("3000-01-01", "yyyy-MM-dd")
        dtpTime = Format("23:59:59", "HH:mm:ss")
        chkEndDate.value = 0
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
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2018-11-23 14:21:13
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    txt����.PasswordChar = IIf(mobjCard.�������Ĺ��� <> "", "*", "")
    txt����.Text = ""
    txtSex.Text = "": txtAge.Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt����_Change()
    Dim strExpend As String
    On Error GoTo Errhand
   
    txt����.PasswordChar = IIf(mobjCard.�������Ĺ��� <> "", "*", "")
    '�����:56597
    '��ʼ��IC��
    If mobjCard.���� Like "IC��*" And mobjCard.ϵͳ = True Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        mobjICCard.SetEnabled Trim(txt����.Text) = ""
        Exit Sub
    End If
    '��ʼ���������֤
    If mobjCard.���� Like "*���֤*" And mobjCard.�ӿڳ����� = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled Trim(txt����.Text) = ""
        Exit Sub
    End If
    
    gobjOneCardComLib.SetEnabled Trim(txt����.Text) = ""
    
    If mobjCard.�ӿ���� = 0 Or mobjCard.�ӿڳ����� = "" Then Exit Sub
    If Not (mobjCard.�Ƿ�ˢ�� Or mobjCard.�Ƿ�ɨ��) Then Exit Sub
    
    Call gobjOneCardComLib.zlSetBrushCardObject(mobjCard.�ӿ����, txt����, strExpend, mobjCard.���ѿ�)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txt����_GotFocus()
    Dim strExpend As String
    
    On Error GoTo Errhand
    zlControl.TxtSelAll txt����
    txt����.PasswordChar = IIf(mobjCard.�������Ĺ��� <> "", "*", "")
    '�����:56597
    '��ʼ��IC��
    If mobjCard.���� Like "IC��*" And mobjCard.ϵͳ = True Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        mobjIDCard.SetEnabled Trim(txt����.Text) = ""
        Exit Sub
    End If
    '��ʼ���������֤
    If mobjCard.���� Like "*���֤*" And mobjCard.�ӿڳ����� = "" Then
        If mobjIDCard Is Nothing Then
            Set mobjIDCard = CreateObject("zlIDCard.clsIDCard")
        End If
        mobjIDCard.SetEnabled Trim(txt����.Text) = ""
        Exit Sub
    End If
    
    gobjOneCardComLib.SetEnabled Trim(txt����.Text) = ""
    If mobjCard.�ӿ���� = 0 Or mobjCard.�ӿڳ����� = "" Then Exit Sub
    If Not (mobjCard.�Ƿ�ˢ�� Or mobjCard.�Ƿ�ɨ��) Then Exit Sub
    
    Call gobjOneCardComLib.zlSetBrushCardObject(mobjCard.�ӿ����, txt����, strExpend, mobjCard.���ѿ�)
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
     
    If (Len(txt����.Text) = mobjCard.���ų��� - 1 And KeyAscii <> 8) Or (KeyAscii = 13 And Trim(txt����.Text) <> "") Then
        If KeyAscii <> 13 Then
            txt����.Text = txt����.Text & Chr(KeyAscii)
            txt����.SelStart = Len(txt����.Text)
        End If
        KeyAscii = 0
        If Not GetPatient(txt����.Text) Then
            Call ClearFace
            txt����.SetFocus: Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt����_LostFocus()
    '�����:56597
   If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled False
   If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
   If Not gobjOneCardComLib Is Nothing Then gobjOneCardComLib.SetEnabled False
   
End Sub
