VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatientEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����Ա��Ϣ"
   ClientHeight    =   4830
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8820
   Icon            =   "frmPatientEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   8820
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6420
      TabIndex        =   36
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7620
      TabIndex        =   37
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   105
      TabIndex        =   39
      Top             =   4275
      Width           =   1100
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4290
      Index           =   1
      Left            =   30
      ScaleHeight     =   4290
      ScaleWidth      =   8715
      TabIndex        =   38
      Top             =   -45
      Width           =   8715
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   4050
         TabIndex        =   3
         Top             =   225
         Width           =   1605
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1350
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1155
         Width           =   1605
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   0
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   29
         Top             =   1905
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   10
         Left            =   4320
         MaxLength       =   18
         TabIndex        =   23
         Top             =   795
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   600
         Index           =   24
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   3510
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   21
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   27
         Top             =   1500
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   1140
         Index           =   22
         Left            =   4320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2310
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   330
         Index           =   19
         Left            =   4320
         TabIndex        =   25
         Top             =   1140
         Width           =   3945
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2340
         Width           =   1605
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2715
         Width           =   1605
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   4
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1965
         Width           =   1605
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   5
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3105
         Width           =   1605
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   6
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3465
         Width           =   1605
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   7
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3825
         Width           =   1605
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   4
         Left            =   8295
         Picture         =   "frmPatientEdit.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3510
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         Height          =   330
         Index           =   3
         Left            =   8310
         Picture         =   "frmPatientEdit.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2310
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   780
         Width           =   1605
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   6
         Left            =   1350
         TabIndex        =   1
         Top             =   195
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1350
         TabIndex        =   9
         Top             =   1560
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         DateIsNull      =   -1  'True
         Format          =   105447427
         CurrentDate     =   38329
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&A)"
         Height          =   180
         Index           =   2
         Left            =   3150
         TabIndex        =   2
         Top             =   300
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&Y)"
         Height          =   180
         Index           =   1
         Left            =   675
         TabIndex        =   6
         Top             =   1245
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʼ�(&L)"
         Height          =   180
         Index           =   0
         Left            =   3240
         TabIndex        =   28
         Top             =   1965
         Width           =   990
      End
      Begin VB.Line ln 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   -150
         X2              =   10020
         Y1              =   4185
         Y2              =   4185
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ(&U)"
         Height          =   180
         Index           =   35
         Left            =   3240
         TabIndex        =   33
         Top             =   3510
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵绰(&T)"
         Height          =   180
         Index           =   23
         Left            =   3060
         TabIndex        =   26
         Top             =   1560
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵�ַ(&K)"
         Height          =   180
         Index           =   24
         Left            =   3060
         TabIndex        =   30
         Top             =   2355
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ������(&W)"
         Height          =   180
         Index           =   21
         Left            =   3060
         TabIndex        =   24
         Top             =   1185
         Width           =   1170
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��(&M)"
         Height          =   180
         Index           =   15
         Left            =   330
         TabIndex        =   10
         Top             =   2025
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��(&E)"
         Height          =   180
         Index           =   16
         Left            =   690
         TabIndex        =   16
         Top             =   3150
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&D)"
         Height          =   180
         Index           =   14
         Left            =   690
         TabIndex        =   14
         Top             =   2775
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   13
         Left            =   690
         TabIndex        =   12
         Top             =   2415
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ(&J)"
         Height          =   180
         Index           =   17
         Left            =   690
         TabIndex        =   18
         Top             =   3510
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���(&F)"
         Height          =   180
         Index           =   18
         Left            =   690
         TabIndex        =   20
         Top             =   3885
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��(&S)"
         Height          =   180
         Index           =   12
         Left            =   3240
         TabIndex        =   22
         Top             =   855
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&B)"
         Height          =   180
         Index           =   11
         Left            =   330
         TabIndex        =   8
         Top             =   1620
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   8
         Left            =   690
         TabIndex        =   0
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�(&I)"
         Height          =   180
         Index           =   9
         Left            =   690
         TabIndex        =   4
         Top             =   840
         Width           =   630
      End
      Begin VB.Line ln 
         BorderColor     =   &H80000003&
         Index           =   2
         X1              =   -195
         X2              =   9975
         Y1              =   660
         Y2              =   660
      End
   End
End
Attribute VB_Name = "frmPatientEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mblnDataChange As Boolean
Private mvarParam As Variant
Private mblnModify As Boolean


'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    mblnDataChange = vData
        
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    On Error Resume Next
    
    For lngLoop = 0 To txt.UBound
        txt(lngLoop).Text = ""
        txt(lngLoop).Tag = ""
    Next
    
    On Error GoTo 0
    
    Call InitData
    
    EditChanged = True
    
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef strParam As String, Optional ByVal blnModify As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
        
    mblnStartUp = True
    mblnOK = False
    
    '����id,����,���֤,����,����״��
    mvarParam = Split(strParam, "'")
    
    Set mfrmMain = frmMain
    mblnModify = blnModify
    
    If InitData = False Then Exit Function
    If ReadPatient(Val(mvarParam(0))) = False Then Exit Function
    
    EditChanged = False
                
    Me.Show 1, frmMain
    
    strParam = Join(mvarParam, "'")
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadPatient(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHand
    
    txt(6).Text = mvarParam(1)
    txt(10).Text = mvarParam(2)
            
    zlControl.CboLocate cbo(1), mvarParam(3)
    
    If Trim(mvarParam(4)) = "" Then
        dtp(1).Value = Null
    Else
        dtp(1).Value = Format(Trim(mvarParam(4)), dtp(1).CustomFormat)
    End If
    
    zlControl.CboLocate cbo(4), mvarParam(5)
    
    zlControl.CboLocate cbo(2), mvarParam(6)
    zlControl.CboLocate cbo(3), mvarParam(7)
    zlControl.CboLocate cbo(5), mvarParam(8)
    zlControl.CboLocate cbo(6), mvarParam(9)
    zlControl.CboLocate cbo(7), mvarParam(10)
    txt(19).Text = mvarParam(11)
    txt(21).Text = mvarParam(12)
    txt(0).Text = mvarParam(13)
    txt(22).Text = mvarParam(14)
    txt(24).Text = mvarParam(15)
    txt(1).Text = mvarParam(16)
    
    If UBound(mvarParam) > 16 Then txt(2).Text = mvarParam(17)
        
    ReadPatient = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID,ȱʡ��־ FROM �Ա� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(1), rs)
        
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID,ȱʡ��־ FROM ���� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(2), rs)
        
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID,ȱʡ��־ FROM ���� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(3), rs)
    
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID,ȱʡ��־ FROM ����״�� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(4), rs)
    
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID,ȱʡ��־ FROM ѧ�� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(5), rs)
    
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID,ȱʡ��־ FROM ְҵ ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(6), rs)
    
    gstrSQL = "SELECT ����||'-'||���� AS ����,0 AS ID FROM ��� ORDER BY ����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(7), rs)
    
    '����������볤��
    
    txt(6).MaxLength = GetMaxLength("������Ϣ", "����")
    txt(10).MaxLength = GetMaxLength("������Ϣ", "���֤��")
    txt(19).MaxLength = GetMaxLength("������Ϣ", "��ϵ������")
    
    txt(21).MaxLength = GetMaxLength("������Ϣ", "��ϵ�˵绰")
    txt(22).MaxLength = GetMaxLength("������Ϣ", "��ϵ�˵�ַ")
    txt(24).MaxLength = GetMaxLength("������Ϣ", "������λ")
    txt(2).MaxLength = GetMaxLength("������Ϣ", "������")
        
    dtp(1).Value = Null
    
    
    If mblnModify = False Then
        txt(6).Locked = True
        txt(10).Locked = True
        txt(19).Locked = True
        txt(21).Locked = True
'        txt(3).Locked = True
        txt(22).Locked = True
        txt(24).Locked = True
        txt(0).Locked = True
        txt(1).Locked = True
        
        dtp(1).Enabled = False
        
        cbo(1).Locked = True
        cbo(2).Locked = True
        cbo(3).Locked = True
        cbo(4).Locked = True
        cbo(5).Locked = True
        cbo(6).Locked = True
        cbo(7).Locked = True
        
    End If
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    
    If Trim(txt(6).Text) = "" Then
        ShowSimpleMsg "��������Ϊ�գ��������룡"
        LocationObj txt(6)
        Exit Function
    End If

    If Trim(cbo(1).Text) = "" Then
        ShowSimpleMsg "�Ա���Ϊ�գ��������룡"
        LocationObj cbo(1)
        Exit Function
    End If
    
'    '����ڱ��������Ƿ��Ѿ����ڴ���
'    If mblnNew = False Then
'        If Val(txt(5).Text) <> mlngԭ����id Then
'            gstrSQL = "SELECT 1 FROM �����Ա���� WHERE �Ǽ�id=" & mlngKey & " AND ����id=" & Val(txt(5).Text)
'        Else
'            gstrSQL = "SELECT 1 FROM dual WHERE 1=2"
'        End If
'    Else
'        If Val(txt(5).Text) <> mlngԭ����id Then
'            gstrSQL = "SELECT 1 FROM �����Ա���� WHERE �Ǽ�id=" & mlngKey & " AND ����id=" & Val(txt(5).Text)
'        Else
'            gstrSQL = "SELECT 1 FROM dual WHERE 1=2"
'        End If
'    End If
'
'    Call OpenRecord(rs, gstrSQL, Me.Caption)
'    If rs.BOF = False Then
'        ShowSimpleMsg "����Ա���ڵ�ǰ��������У������ٴ���ӣ�"
'        LocationObj txt(5)
'        Exit Function
'    End If
                                                                
    ValidEdit = True
    
End Function

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub cbo_Click(Index As Integer)
    EditChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub


Private Sub cmd_Click(Index As Integer)
    Dim lngKey As Long
    
    On Error GoTo errHand
    
    Select Case Index
    Case 3
        
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
    Dim lngKey As Long
    
    If mblnModify Then
        
        If ValidEdit = False Then Exit Sub

        mblnOK = True
        
        mvarParam(1) = txt(6).Text
        mvarParam(2) = txt(10).Text
        
        mvarParam(3) = zlCommFun.GetNeedName(cbo(1).Text)
        
        If IsNull(dtp(1).Value) Then
            mvarParam(4) = ""
        Else
            mvarParam(4) = Format(dtp(1).Value, dtp(1).CustomFormat)
        End If
        mvarParam(5) = zlCommFun.GetNeedName(cbo(4).Text)
        
        mvarParam(6) = zlCommFun.GetNeedName(cbo(2).Text)
        mvarParam(7) = zlCommFun.GetNeedName(cbo(3).Text)
        mvarParam(8) = zlCommFun.GetNeedName(cbo(5).Text)
        mvarParam(9) = zlCommFun.GetNeedName(cbo(6).Text)
        mvarParam(10) = zlCommFun.GetNeedName(cbo(7).Text)
        
        mvarParam(11) = txt(19).Text
        mvarParam(12) = txt(21).Text
        mvarParam(13) = txt(0).Text
        mvarParam(14) = txt(22).Text
        mvarParam(15) = txt(24).Text
        mvarParam(16) = txt(1).Text
        
        If UBound(mvarParam) > 16 Then mvarParam(17) = txt(2).Text
        
    End If
    
    EditChanged = False
    Unload Me
End Sub

Private Sub dtp_Change(Index As Integer)
    EditChanged = True
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChange Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    EditChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 0, 2
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 2
        zlCommFun.OpenIme False
    End Select
    
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
