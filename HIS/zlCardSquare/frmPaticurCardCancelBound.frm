VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPaticurCardCancelBound 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ȡ�����Ű�"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8700
   Icon            =   "frmPaticurCardCancelBound.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7155
      TabIndex        =   15
      Top             =   405
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7155
      TabIndex        =   11
      Top             =   1005
      Width           =   1395
   End
   Begin VB.PictureBox picPass 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   195
      ScaleHeight     =   3135
      ScaleWidth      =   6420
      TabIndex        =   0
      Top             =   495
      Width           =   6420
      Begin VB.CommandButton cmdALL 
         Caption         =   "ȡ�����а󶨵�ҽ�ƿ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1080
         TabIndex        =   14
         Top             =   3420
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.CommandButton cmdAllType 
         Caption         =   "ȡ�����а󶨵�[���￨]"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1080
         TabIndex        =   13
         Top             =   3255
         Visible         =   0   'False
         Width           =   4815
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1815
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
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1815
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
         Left            =   1110
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1215
         Width           =   4845
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   0
         TabIndex        =   2
         Top             =   900
         Width           =   6555
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
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2475
         Width           =   4845
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   135
         Picture         =   "frmPaticurCardCancelBound.frx":0ECA
         Top             =   90
         Width           =   720
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
         TabIndex        =   10
         Top             =   1920
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
         Left            =   375
         TabIndex        =   9
         Top             =   1920
         Width           =   630
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
         Left            =   375
         TabIndex        =   8
         Top             =   1275
         Width           =   630
      End
      Begin VB.Label lblNotes 
         BackStyle       =   0  'Transparent
         Caption         =   "ȡ���󶨲���,�����Ҫȡ��ָ�����İ�,����ȡ����������ˢ��������ָ���Ŀ��Ž���ȡ���󶨡�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   975
         TabIndex        =   7
         Top             =   270
         Width           =   5340
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
         Left            =   345
         TabIndex        =   6
         Top             =   2520
         Width           =   660
      End
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   3540
      Left            =   75
      TabIndex        =   12
      Top             =   345
      Width           =   6945
      _Version        =   589884
      _ExtentX        =   12250
      _ExtentY        =   6244
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmPaticurCardCancelBound"
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
Private mblnOK As Boolean
Private mrsInfo As ADODB.Recordset
Private mobjCardObject As clsCardObject
Private mblnFirst As Boolean
Private mblnCheckOldPass As Boolean
Public mstrPrepayPrivs As String  'Ԥ�������Ȩ��
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mobjSquare As Object

Public Function zlCancelBand(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional lng����ID As Long, Optional strCardNo As String, _
    Optional blnCheckOldPass As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���(ȡ���󶨲���)
    '���:frmMain-���õ�������
    '       lngModule -ģ���
    '       lngCardTypeId-�����ID
    '       lng����ID-����ID
    '       strCardNo-����
    '����:ȡ���ɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-07-29 11:08:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngCardTypeID = lngCardTypeID: mlngModule = lngModule: mlng����ID = lng����ID
    mstrCardNo = strCardNo: mblnOK = False
    mblnCheckOldPass = blnCheckOldPass
    On Error Resume Next
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlCancelBand = mblnOK
End Function
Private Sub InitTaskPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��InitTaskPancel
    '����:���˺�
    '����:2011-06-30 18:20:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    wndTaskPanel.Top = wndTaskPanel.Top + 50
    wndTaskPanel.Height = wndTaskPanel.Height - 150
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    
    Call wndTaskPanel.SetGroupInnerMargins(2, 0, 2, 0)
    Call wndTaskPanel.SetGroupOuterMargins(2, -10, 2, -10)
    Call wndTaskPanel.SetMargins(2, 16, 2, 10, 30)
    
    Set tkpGroup = wndTaskPanel.Groups.Add(1, "��ˢ��")
    Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
   Set Item.Control = picPass
    tkpGroup.CaptionVisible = False
   Call Item.SetMargins(0, -19, 0, -4)
    picPass.BackColor = Item.BackColor
    Me.BackColor = Item.BackColor
    cmdAllType.BackColor = Item.BackColor
    cmdCancel.BackColor = Item.BackColor
    cmdALL.BackColor = Item.BackColor
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
    Dim strExpend As String
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
    cmdAllType.Caption = Replace(cmdAllType.Caption, "[���￨]", "[" & mobjCardObject.CardPreporty.���� & "]")
    
    InitCardInfor = True
    '85565:���ϴ�,2015/7/21,����ˢ���ӿ�
    If mobjSquare Is Nothing Then Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
    If Err <> 0 Then Exit Function
    mobjSquare.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    Err = 0: On Error Resume Next
    If mobjCardObject.CardPreporty.�ӿ���� = 0 Or mobjCardObject.CardPreporty.�ӿڳ����� = "" Then Exit Function
    If Not (mobjCardObject.CardPreporty.�Ƿ�ˢ�� Or mobjCardObject.CardPreporty.�Ƿ�ɨ��) Then Exit Function
    If mobjSquare.zlSetBrushCardObject(mobjCardObject.CardPreporty.�ӿ����, txt����, strExpend, _
                mobjCardObject.CardPreporty.���ѿ�) Then
        Call mobjSquare.zlInitEvents(Me.hWnd, mobjCommEvents)
    End If
End Function
   
Private Sub cmdALL_Click()
    Call SaveData(2)
End Sub

Private Sub cmdAllType_Click()
    Call SaveData(1)
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub
Private Function CheckBindCard(ByVal lng����ID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�󶨿�
    '����:���˺�
    '����:2011-07-31 05:37:48
    '����׼:
    '  1.��סԺ���ü�¼���޼�¼,�ͱ�ʾ�󶨿�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strȱʡ�����ID As String
    strȱʡ�����ID = IIf(mobjCardObject.CardPreporty.���� = "���￨" And mobjCardObject.CardPreporty.ϵͳ, mobjCardObject.�ӿ����, "")
    strSQL = "" & _
    "   Select  1 " & _
    "   From סԺ���ü�¼  " & _
    "   Where ����id = [1] And ��¼���� = 5 and nvl(����,[3])=[2] and ʵ��Ʊ��=[4] And RowNum=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, strȱʡ�����ID, Trim(CStr(mlngCardTypeID)), strCardNo)
    CheckBindCard = rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function isValied(ByVal intType As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ƿ���Ч
    '���:intType:0-��ǰ����;1-��ǰ���;2-��ǰ��������
    '����:������Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-29 11:15:42
    '---------------------------------------------------------------------------------------------------------------------------------------------\
    Dim str�ƺ� As String
    str�ƺ� = IIf(glngSys Like "8??", "�ͻ�", "����")
    
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then
        MsgBox "���ܶ�ȡ" & str�ƺ� & "��Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
         txt����.SetFocus: Exit Function
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "���ܶ�ȡ" & str�ƺ� & "��Ϣ����ȷ���Ƿ���ȷˢ����", vbInformation, gstrSysName
         txt����.SetFocus: Exit Function
    End If
    If intType = 0 Then
       If Trim(txt����.Text) = "" Then
            MsgBox "δ������Ҫȡ���Ŀ���,����ȡ���󶨲���!", vbOKOnly + vbInformation, gstrSysName
            txt����.SetFocus: Exit Function
       End If
       If mlng����ID <> mrsInfo!����ID Then
            MsgBox "��ǰ���ŵĳ�����,������ѡ��Ĳ���,����ȡ���󶨲���!", vbOKOnly + vbInformation, gstrSysName
            txt����.SetFocus: Exit Function
       End If
       '��鵱ǰ�����Ƿ�󶨲���
       If CheckBindCard(mlng����ID, Trim(txt����)) = False Then
            MsgBox "��ǰ���Ų��ǰ󶨵Ŀ���,��ʹ���˿�����!", vbOKOnly + vbInformation, gstrSysName
            txt����.SetFocus: Exit Function
       End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function BlandCancel(ByVal intType As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ���󶨿�
    '���:intType:0-��ǰ����;1-��ǰ���;2-��ǰ��������
    '����:ȡ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-29 11:18:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim lng����ID As Long, Curdate As Date, cllPro As Collection
   Dim strSQL As String, strPassWord As String
   
    On Error GoTo errHandle
    lng����ID = Val(Nvl(mrsInfo!����ID))
    If intType = 0 Then
        '�����ǰ����ȡ����,Ԥ������
        If Not IsCheckCancel��Ԥ��(lng����ID) Then
            Exit Function
        End If
    End If
    Set cllPro = New Collection
    Curdate = zlDatabase.Currentdate
    '105590:���ϴ�,2017/3/10��ȡ����ʱ��д����Ա����
      'Zl_ҽ�ƿ��䶯_Insert
       strSQL = "Zl_ҽ�ƿ��䶯_Insert("
      '      �䶯����_In   Number,
      '��������=1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
      strSQL = strSQL & "" & 14 & ","
      '      ����id_In     סԺ���ü�¼.����id%Type,
      strSQL = strSQL & "" & lng����ID & ","
      '      �����id_In   ����ҽ�ƿ���Ϣ.�����id%Type,
      strSQL = strSQL & "" & IIf(intType = 2, "NULL", mlngCardTypeID) & ","
      '      ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
      strSQL = strSQL & "NULL,"
      '      ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
      strSQL = strSQL & IIf(intType <> 0, "NULL", "'" & txt����.Text & "'") & ","
      '      �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
      strSQL = strSQL & "'ȡ�����Ű�',"
      '      ����_In       ������Ϣ.����֤��%Type,
      strSQL = strSQL & "NULL,"
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
    BlandCancel = True
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

Private Function IsCheckCancel��Ԥ��(ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ������ʱʱ��鲡���Ƿ���Ԥ����δ��
     '����:��Ч,����true,���򷵻�False
    '����:����
    '����:2012-07-16 18:50:36
    '�����:51537
    '�����:50891
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim msgBoxResult As String
    Dim strSQL As String
    Dim rsBill As Recordset, rsCard As Recordset
    '69483,������,2014-01-15,����ҽ�ƿ��˿��˿��
    strSQL = "Select Count(1) As ҽ�ƿ��� From ����ҽ�ƿ���Ϣ Where ״̬=0 And ����ID=[1]"
    Set rsCard = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    strSQL = _
            "Select Ԥ�����,������� From ������� Where ����=1 And ����=1 And ����ID=[1]"
    Set rsBill = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    '����:48249
    If InStr(1, mstrPrepayPrivs, ";Ԥ���˿�;") > 0 Then
        If rsBill.RecordCount > 0 Then
            If Format(Nvl(rsBill!Ԥ�����, 0) - Nvl(rsBill!�������, 0), "0.00") > 0 Then
                
                '�����:51537
                '�����:50891
                msgBoxResult = zlCommFun.ShowMsgbox(gstrSysName, "�ò�������Ԥ�����δ��!" & " " & "�Ƿ��Ƚ�������˿����?", "��Ԥ�������,����,ȡ��", Me, vbQuestion)
                If msgBoxResult = "��Ԥ�������" Then '��Ԥ��������
                    '��������˿�
                     IsCheckCancel��Ԥ�� = zlPrepayFunc(2, lng����ID)
                     Exit Function
                ElseIf msgBoxResult = "����" Then
                    If rsCard!ҽ�ƿ��� = 1 Then
                        MsgBox "�ò�������Ԥ�������ܶԲ���Ψһ��ҽ�ƿ�����ȡ���󶨲���!", vbInformation, gstrSysName
                        IsCheckCancel��Ԥ�� = False
                        Exit Function
                    End If
                    IsCheckCancel��Ԥ�� = True
                ElseIf msgBoxResult = "ȡ��" Or msgBoxResult = "" Then
                     IsCheckCancel��Ԥ�� = False
                     Exit Function
                End If
            End If
'        Else
'        '�����:51537
'        '�����:50891
'           If ZL9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "�Ƿ��������ȡ�����󶨲���?", "�˿�,ȡ��", Me, vbQuestion) = "ȡ��" Then
'                IsCheckCancel��Ԥ�� = False
'                Exit Function
'           End If
        End If
    Else
        If rsBill.RecordCount > 0 Then
            If Format(Nvl(rsBill!Ԥ�����, 0) - Nvl(rsBill!�������, 0), "0.00") > 0 Then
                If rsCard!ҽ�ƿ��� = 1 Then
                    MsgBox "��û��Ԥ���˿�Ȩ�ޣ����ܶԲ���Ψһ��ҽ�ƿ�����ȡ���󶨲���!", vbInformation, gstrSysName
                    IsCheckCancel��Ԥ�� = False
                    Exit Function
                End If
            End If
        End If
        If MsgBox("��û��Ԥ���˿�Ȩ��,�Ƿ��������ȡ�����󶨲���?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then IsCheckCancel��Ԥ�� = False: Exit Function
    End If
        IsCheckCancel��Ԥ�� = True
End Function
Private Sub SaveData(ByVal intType As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ���󶨿�
    '���:intType:0-��ǰ����;1-��ǰ���;2-��ǰ��������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If isValied(intType) = False Then Exit Sub
    If BlandCancel(intType) = False Then Exit Sub
    MsgBox "ȡ���ɹ�!", vbOKOnly + vbInformation, gstrSysName
    mblnOK = True
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Call SaveData(0)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call ClearFace
    If GetPatient("-" & mlng����ID) = False Then Unload Me: Exit Sub
    If mstrCardNo <> "" Then
        If GetPatient(mstrCardNo) = False Then
              If txt����.Enabled Then txt����.SetFocus
            Exit Sub
        End If
    End If
    If txt����.Enabled Then txt����.SetFocus
End Sub

Private Sub Form_Load()
    mblnFirst = True
    If glngSys Like "8??" Then lbl����.Caption = "�ͻ�"
    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    If InitCardInfor = False Then
        '74539,Ƚ����,2014-6-27,���շѴ���Ժ�⿨�󣬵�ҽ�ƿ����Ź���ȡ����ʱ��ȡ���󶨴���һ���������޷�ȡ��
        MsgBox "�ÿ��豸δ���ã������ܽ���ȡ���󶨲������뵽����������>�豸���á������ã�", vbInformation, gstrSysName
        mblnFirst = False: Unload Me: Exit Sub
    End If
    Call InitTaskPancel
End Sub
 Private Function GetPatient(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '����:���˺�
    '����:2011-07-29 11:34:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, strWhere As String, blnReadPatiInfor As Boolean
    blnReadPatiInfor = Left(strInput, 1) = "-"
    
    Set mrsInfo = Nothing
    On Error GoTo errH
    If Not blnReadPatiInfor Then
        '�������ĺ���
        If GetPatiID(mlngCardTypeID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID = 0 Then GoTo NotFoundPati:
        mstrCardNo = strInput
        If lng����ID <= 0 Then GoTo NotFoundPati:
        txt����.Text = strInput
    Else
        lng����ID = Val(Mid(strInput, 2))
    End If
    
    strSQL = "" & _
    "   Select ����ID,�����,סԺ��,���￨��,����,�Ա�,����" & _
    "   From ������Ϣ " & _
    "   Where ����ID=[1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If Not blnReadPatiInfor Then
        GetPatient = True: Exit Function
    End If
    If mrsInfo.EOF Then Exit Function

    txtPati.Text = Nvl(mrsInfo!����)
    txtPati.Tag = Val(mrsInfo!����ID)
    txtSex.Text = Nvl(mrsInfo!�Ա�)
    txtAge.Text = Nvl(mrsInfo!����)
    Set mrsInfo = Nothing
'    If mblnCheckOldPass Then
'        If zlCommFun.VerifyPassWord(Me, strPassWord, txtPati.Text, txtSex.Text, txtAge.Text, True) = False Then
'            Call ClearFace
'            Exit Function
'        End If
'    End If
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    txtPati.Text = ""
    txtSex.Text = "": txtAge.Text = ""
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSquare = Nothing
    Set mobjCommEvents = Nothing
End Sub

Private Sub lbl����_Click()
    Dim strExpand As String, strCardNo As String, strOutXml As String
    If Not mobjCardObject.CardPreporty.�Ƿ�Ӵ�ʽ���� Then Exit Sub
'    If mobjICCard Is Nothing Then
'        Set mobjICCard = CreateObject("zlICCard.clsICCard")
'        Set mobjICCard.gcnOracle = gcnOracle
'    End If
    
'    If Not mobjICCard Is Nothing Then
'        txt����.Text = mobjICCard.Read_Card()
'        If txt����.Text <> "" Then
'            mblnICCard = True
'            Call CheckFreeCard(txt����.Text)
'        End If
'    End If
  
    If mobjCardObject.CardObject Is Nothing Then Exit Sub
    If mobjCardObject.CardObject.zlReadCard(Me, mlngModule, False, strExpand, strCardNo, strOutXml) = False Then Exit Sub
    txt����.Text = Trim(strCardNo)
    If txt����.Text <> "" Then
        If Not GetPatient(txt����.Text) Then
            Call txt����_GotFocus
            If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Exit Sub
        End If
        cmdOK.SetFocus
    End If
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
    txt����.PasswordChar = IIf(mobjCardObject.CardPreporty.�������Ĺ��� <> "", "*", "")
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
     If (Len(txt����.Text) = mobjCardObject.CardPreporty.���ų��� - 1 And KeyAscii <> 8) Or (KeyAscii = 13 And Trim(txt����.Text) <> "") Then
            If KeyAscii <> 13 Then
                txt����.Text = txt����.Text & Chr(KeyAscii)
                txt����.SelStart = Len(txt����.Text)
            End If
            KeyAscii = 0
            If Not GetPatient(txt����.Text) Then
                Call txt����_GotFocus
                If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                Exit Sub
            End If
            cmdOK.SetFocus
        End If
End Sub


Private Function zlPrepayFunc(ByVal intFunc As Integer, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ԥ���
    '���:intFunc-1-��Ԥ��;2-��Ԥ��;3-����,4-����תסԺ;5-סԺת����;
    '����:���˺�
    '����:2011-07-24 18:25:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun As Object, intԤ������ As Integer
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Function
    'bytԤ������: 0-��Ԥ����(ȱʡ,���л�����),1-�������(1),2-����״̬(1); 3-����˿�(37770), 4-����תסԺ;5-סԺת����
    Select Case intFunc
    Case 1  '1.��Ԥ��
        intԤ������ = 0
    Case 2 '�˿�
        intԤ������ = 3
    Case 3: intԤ������ = 2
    Case 4: intԤ������ = 4
    Case 5: intԤ������ = 5
    End Select
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ� ����Ԥ�����տ��
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ���ݿ��¼�û���
    '  bytCallObject:���˺����(0-Ԥ�������(ȱʡ��);1-���˷��ò�ѯ����,2-ҽ�ƿ�����)
    '  lng����ID-ȱʡ�Ĳ���ID
    '  lng��ҳID-ȱʡ����ҳID
    '  dblDefPrePayMoney-ȱʡ��Ԥ�����
    Set gfrmCardMgr = Me
    '����:48249
    If objFun.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 2, lng����ID, 0, 0, intԤ������) = False Then
        zlPrepayFunc = False
        Set gfrmCardMgr = Nothing
        Exit Function
    End If
    Set gfrmCardMgr = Nothing
    zlPrepayFunc = True
End Function

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    txt����.Text = Trim(strCardNo)
    If txt����.Text <> "" Then
        If Not GetPatient(txt����.Text) Then
            Call txt����_GotFocus
            If txt����.Enabled And txt����.Visible Then txt����.SetFocus
            Exit Sub
        End If
        cmdOK.SetFocus
    End If
End Sub

Private Sub txt����_LostFocus()
    If Not mobjSquare Is Nothing Then mobjSquare.SetEnabled False
End Sub
