VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUserCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ѫȷ��"
   ClientHeight    =   4785
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmUserCheck.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6915
   StartUpPosition =   1  '����������
   Begin VB.PictureBox PicTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4875
      ScaleHeight     =   270
      ScaleWidth      =   1890
      TabIndex        =   26
      Top             =   2010
      Width           =   1920
      Begin MSComCtl2.DTPicker DTPTime 
         Height          =   330
         Left            =   -30
         TabIndex        =   5
         Top             =   -30
         Width           =   1980
         _ExtentX        =   3493
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   265420803
         CurrentDate     =   43019
      End
   End
   Begin VB.PictureBox picBlood 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   885
      ScaleHeight     =   270
      ScaleWidth      =   2955
      TabIndex        =   23
      Top             =   450
      Width           =   2985
      Begin VB.ComboBox cboBlood 
         Height          =   300
         Left            =   -30
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   -30
         Width           =   3030
      End
   End
   Begin VB.PictureBox picOper 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   6540
      Picture         =   "frmUserCheck.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox TXT���� 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   4875
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1590
      Width           =   1920
   End
   Begin VB.PictureBox picOper 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   6540
      Picture         =   "frmUserCheck.frx":06C3
      ScaleHeight     =   240
      ScaleWidth      =   225
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3285
      Width           =   255
   End
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   4875
      TabIndex        =   7
      Top             =   3270
      Width           =   1920
   End
   Begin VB.TextBox txt���� 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   4875
      TabIndex        =   2
      Top             =   1185
      Width           =   1920
   End
   Begin VB.TextBox txt�û� 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   4875
      TabIndex        =   6
      Top             =   2850
      Width           =   1920
   End
   Begin VB.TextBox TXT���� 
      Appearance      =   0  'Flat
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4875
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3705
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   15
      TabIndex        =   13
      Top             =   4080
      Width           =   7350
   End
   Begin VB.CommandButton CMD���� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5715
      TabIndex        =   11
      Top             =   4335
      Width           =   1100
   End
   Begin VB.CommandButton CMDȷ�� 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4530
      TabIndex        =   10
      Top             =   4335
      Width           =   1100
   End
   Begin VB.TextBox txt�û� 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   4875
      TabIndex        =   1
      Top             =   765
      Width           =   1920
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
      Height          =   3210
      Left            =   120
      TabIndex        =   21
      Top             =   795
      Width           =   3750
      _cx             =   6615
      _cy             =   5662
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   12
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmUserCheck.frx":0A7C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblBlood 
      AutoSize        =   -1  'True
      Caption         =   "ѪҺ��Ϣ"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   495
      Width           =   720
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��      ��"
      Height          =   180
      Index           =   0
      Left            =   3930
      TabIndex        =   22
      Top             =   1650
      Width           =   900
   End
   Begin VB.Label lblInfo 
      Caption         =   "��ʾ��"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   150
      TabIndex        =   20
      Top             =   4305
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����֤"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4395
      TabIndex        =   19
      Top             =   180
      Width           =   720
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "��Ϣ�˶�(3��8��)"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   180
      Width           =   1440
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "ȡѪ������"
      Height          =   180
      Index           =   1
      Left            =   3930
      TabIndex        =   17
      Top             =   3330
      Width           =   900
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��Ѫ������"
      Height          =   180
      Index           =   0
      Left            =   3930
      TabIndex        =   16
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label Lbl�û��� 
      AutoSize        =   -1  'True
      Caption         =   "ȡѪ���ʺ�"
      Height          =   180
      Index           =   1
      Left            =   3930
      TabIndex        =   15
      Top             =   2910
      Width           =   900
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��      ��"
      Height          =   180
      Index           =   1
      Left            =   3930
      TabIndex        =   14
      Top             =   3765
      Width           =   900
   End
   Begin VB.Image imgFlag 
      Height          =   345
      Left            =   3945
      Picture         =   "frmUserCheck.frx":0AE6
      Stretch         =   -1  'True
      Top             =   105
      Width           =   405
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��Ѫ����"
      Height          =   180
      Index           =   0
      Left            =   4110
      TabIndex        =   12
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Lbl�û��� 
      AutoSize        =   -1  'True
      Caption         =   "��Ѫ���ʺ�"
      Height          =   180
      Index           =   0
      Left            =   3930
      TabIndex        =   0
      Top             =   825
      Width           =   900
   End
End
Attribute VB_Name = "frmUserCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean  'ΪTrue��ʾ�Ѿ�������ʾ��
Private mblnOk As Boolean
Private mstrOper As String '�˶���
Private mstrSendTime As String '�˶�ʱ��
Private mstrCheckResult As String
Private mblnTakeVerification As Boolean 'ȡѪ�������֤
Private mlngDeptID As Long '���˿���ID
Private mlngSendDeptID As Long '��Ѫ����ID
Private mstr����ʱ�� As String
Private mstr���ʱ�� As String
Private mblnSelectSendUser As Boolean
Private mintMode As Integer
Private marrHTitle(0 To 2) As String  '�˶���
Private marrFTitle(0 To 2) As String  '������
Private mstrIDs As String   'ѪҺ�շ�ID��Ϣ
Private mBloodResult As Collection  'ѪҺ�˶Խ����Ϣ
Private mlngPreBoodID As Long  '��һ��ѡ���ѪҺID
Private mlngModul As Long '����ģ��

Private Enum Vsf_COL
    COL_��� = 0
    COL_���� = 1
    COL_��� = 2
End Enum

Public Property Get SendAndTakeOper() As String '��Ѫ��'ȡѪ��/������'�˶���/�˶���'������
    SendAndTakeOper = mstrOper
End Property

Public Property Get BloodResult() As Collection
    Set BloodResult = mBloodResult
End Property

Public Property Get CheckResult() As String '���12��Ŀ�������
    CheckResult = mstrCheckResult
End Property

Public Property Get SendTime() As String '��Ѫʱ��
    SendTime = Format(mstrSendTime, "YYYY-MM-DD HH:mm")
End Property

Public Function ShowMe(ByVal frmParent As Object, ByVal lngModul As Enum_Inside_Program, ByVal lngDeptID As Long, ByVal lngSendDeptId As Long, ByVal str����ʱ�� As String, ByVal str���ʱ�� As String, _
    Optional ByVal blnSelectSendUser As Boolean = True, Optional ByVal intMode As Enum_CheckType = ��Ѫ�˶�, Optional ByVal strIDs As String = "") As Boolean
'���ܣ���Ѫ�����ա�ִ�й��̵�˫�˶�
'����: strIDs ѪҺ�շ�ID����ʽ�Զ��ŷָ�(������ɶ�û��ѪҺ�ĺ˶Խ����������,���򷵻�ͳһ���),����ID��ȡ���ͨ�����ԡ�BloodResult��������ͨ������"CheckResult"��ȡ
'1�����ڷ�Ѫ���ܶ��ԣ���Ѫʱ�䲻��С������ʱ�䣬���ܳ������˾���ʱ��(���ʱ��δ���򲻼��)
'2�����ڽ��չ��ܶ��ԣ�����ʱ�䲻��С�ڷ�Ѫʱ�䣬���ܳ������˾���ʱ��(���ʱ��δ���򲻼��)
'3������ִ�к˶Զ��ԣ��˶�ʱ�䲻��С�ڽ���ʱ����ϴ�ִ��ʱ�䣬���ܴ����´�ִ��ʱ��(���ʱ��Ϊ���򲻼��)
    mlngModul = lngModul
    mlngDeptID = lngDeptID
    mlngSendDeptID = lngSendDeptId
    mstr����ʱ�� = str����ʱ��
    mstr���ʱ�� = str���ʱ��
    mblnSelectSendUser = blnSelectSendUser
    mintMode = intMode
    mstrIDs = strIDs
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    CMD����.Enabled = BlnState
    CMDȷ��.Enabled = BlnState
End Sub

Private Sub cboBlood_Click()
    Dim strTmp As String, strCheck As String
    Dim i As Integer, j As Integer
    With cboBlood
        If mlngPreBoodID = cboBlood.ListIndex Then Exit Sub
        If mlngPreBoodID = -1 Then '�״ν���(Ĭ��ȫѡ)
            For i = vsfCheck.FixedRows To vsfCheck.Rows - 1
                strTmp = "11111111111"
            Next
        Else
            '������һ�ε�ѡ��
            For i = vsfCheck.FixedRows To vsfCheck.Rows - 1
                strCheck = strCheck & IIf(Abs(Val(vsfCheck.TextMatrix(i, COL_���))) = 1, "1", "0")
            Next
            mBloodResult.Remove ("B_" & IIf(.ItemData(mlngPreBoodID) = -1, 0, .ItemData(mlngPreBoodID)))
            mBloodResult.Add strCheck, "B_" & IIf(.ItemData(mlngPreBoodID) = -1, 0, .ItemData(mlngPreBoodID))
            'ˢ�±��ε�ѡ��
            strTmp = CStr(mBloodResult("B_" & IIf(.ItemData(.ListIndex) = -1, 0, .ItemData(.ListIndex))))
        End If
        j = 1
        For i = vsfCheck.FixedRows To vsfCheck.Rows - 1
            vsfCheck.TextMatrix(i, COL_���) = Mid(strTmp, j, 1)
            j = j + 1
        Next
        mlngPreBoodID = cboBlood.ListIndex
    End With
End Sub

Private Sub CMD����_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub CMDȷ��_Click()
    Dim strNote As String
    Dim strUserName As String, strPassword As String
    Dim arrUserName(0 To 1) As String, arrPassword(0 To 1) As String
    Dim strServerName As String
    Dim intCheck As Integer, blnSendUserCheck As Boolean
    Dim strCheck As String
    Dim i As Integer
    On Error GoTo InputError
    
    Call Me.ValidateControls
    
    SetConState False
    'ȡѪ�˺ͷ�Ѫ����֤���
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    blnSendUserCheck = Val(Lbl�û���(intCheck).Tag) <> UserInfo.id
    For intCheck = 0 To 1
        If blnSendUserCheck = True And intCheck = 0 Or intCheck = 1 Then
            strUserName = Trim(txt�û�(intCheck).Text)
            strPassword = Trim(TXT����(intCheck).Text)
            
            '��Ч�ַ���Ч��
            If Len(Trim(txt�û�(intCheck))) = 0 Then
                strNote = "������" & IIf(intCheck = 0, marrHTitle(mintMode) & "��", marrFTitle(mintMode) & "��") & "�ʺ�"
                Call gobjControl.ControlSetFocus(txt�û�(1))
                GoTo InputError
            End If
            
            If Len(strUserName) <> 1 Then
                If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
                    strNote = IIf(intCheck = 0, marrHTitle(mintMode) & "��", marrFTitle(mintMode) & "��") & "�ʺŴ���"
                    Call gobjControl.ControlSetFocus(txt�û�(intCheck))
                    SetConState
                    Exit Sub
                End If
            End If
            If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
                If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
                    strNote = IIf(intCheck = 0, marrHTitle(mintMode) & "��", marrFTitle(mintMode) & "��") & "�ʺ��������"
                    Call gobjControl.ControlSetFocus(TXT����(intCheck))
                    GoTo InputError
                End If
            End If
            
            If Len(Trim(strPassword)) = 0 Then
                strNote = "������" & IIf(intCheck = 0, marrHTitle(mintMode) & "��", marrFTitle(mintMode) & "��") & "�ʺ�����"
                Call gobjControl.ControlSetFocus(TXT����(intCheck))
                GoTo InputError
            End If
        End If
        arrUserName(intCheck) = strUserName
        arrPassword(intCheck) = strPassword
     Next
    
'    If IsDate(TXT����(0).Text) = False Then
'        strNote = marrHTitle(mintMode) & "���ڲ�����Ч�����ڸ�ʽ�����飡"
'        Call gobjControl.ControlSetFocus(TXT����(0))
'        GoTo InputError
'    End If
    '��Ѫ�˺�ȡѪ�˲�����ͬһ��
    If txt����(1).Text = "" Then
        strNote = "������" & marrHTitle(mintMode) & "��"
        Call gobjControl.ControlSetFocus(txt����(1))
        GoTo InputError
    End If
    If txt����(0).Text = txt����(1).Text Then
        strNote = marrHTitle(mintMode) & "�˲��ܺ�" & marrFTitle(mintMode) & "����ͬһ���ˣ�������ȷ��" & marrFTitle(mintMode) & "�ˣ�"
        Call gobjControl.ControlSetFocus(txt����(1))
        GoTo InputError
    End If
    '�û���¼��֤
    If blnSendUserCheck = True Then
        '�û���¼��֤
        If GetObjectRegister = False Then Exit Sub
        strServerName = gobjRegister.GetServerName
        If gobjRegister.LoginValidate(strServerName, arrUserName(0), arrPassword(0), strNote) = False Then
            TXT����(0).Text = ""
            Call gobjControl.ControlSetFocus(TXT����(0))
            SetConState
            GoTo InputError
        End If
    End If
    
    '�û���¼��֤
    If GetObjectRegister = False Then Exit Sub
    strServerName = gobjRegister.GetServerName
    If gobjRegister.LoginValidate(strServerName, arrUserName(1), arrPassword(1), strNote) = False Then
        TXT����(1).Text = ""
        Call gobjControl.ControlSetFocus(TXT����(1))
        SetConState
        GoTo InputError
    End If
    
    
    mstrOper = txt����(0).Text & "'" & txt����(1).Text
    mstrSendTime = DTPTime.Value
    strCheck = ""
    For i = vsfCheck.FixedRows To vsfCheck.Rows - 1
        strCheck = strCheck & IIf(Abs(Val(vsfCheck.TextMatrix(i, COL_���))) = 1, "1", "0")
    Next
    mstrCheckResult = strCheck
    If Not mBloodResult Is Nothing Then
        If mBloodResult.Count > 0 Then
            If cboBlood.ItemData(cboBlood.ListIndex) = -1 Then '��ʾ��������
                Call mBloodResult.Remove("B_0")
                For i = 1 To cboBlood.ListCount - 1
                    Call mBloodResult.Remove("B_" & cboBlood.ItemData(i))
                    mBloodResult.Add strCheck, "B_" & cboBlood.ItemData(i)
                Next
            Else
                Call mBloodResult.Remove("B_" & cboBlood.ItemData(cboBlood.ListIndex))
                mBloodResult.Add strCheck, "B_" & cboBlood.ItemData(cboBlood.ListIndex)
            End If
        End If
    End If
    
    mblnOk = True
    Unload Me
    Exit Sub
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbInformation, gstrSysName
    End If
    SetConState
    Exit Sub
End Sub

Private Sub DTPTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim lngScrH As Long
    If mblnFirst = False Then
        '��ʾ����
        lngScrH = GetSystemMetrics(17) * 15 '��Ļ���ø߶�
        If Me.Top + Me.Height > lngScrH Then
            Me.Top = lngScrH - Me.Height
        End If
    
        If Trim(txt�û�(1).Text) = "" Then
            CMDȷ��.Default = False
            txt�û�(1).SetFocus
        Else
            If TXT����(1).Enabled Then
                TXT����(1).SetFocus
            Else
                CMDȷ��.SetFocus
            End If
        End If
        mblnFirst = True
        If Trim(txt�û�(1).Text) <> "" And Trim(TXT����(1).Text) <> "" Then Call CMDȷ��_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strName As String, arrName
    Dim i As Integer
    
    marrHTitle(0) = "��Ѫ"
    marrHTitle(1) = "����"
    marrHTitle(2) = "�˲�"
    marrFTitle(0) = "ȡѪ"
    marrFTitle(1) = "�˲�"
    marrFTitle(2) = "����"
    
    If mstrIDs = "" Then
        lblBlood.Visible = False
        picBlood.Visible = False
        vsfCheck.Height = vsfCheck.Height + vsfCheck.Top - picBlood.Top
        vsfCheck.Top = picBlood.Top
    End If
    mlngPreBoodID = -1
    Call LoadBoold
    With vsfCheck
        .Clear
        .Rows = 12
        .Cols = 3
        .ColWidth(COL_���) = 500
        .ColWidth(COL_����) = 2000
        .ColWidth(COL_���) = 500
        .TextMatrix(0, COL_���) = "���"
        .TextMatrix(0, COL_����) = "�˲���Ŀ"
        .TextMatrix(0, COL_���) = "���"
        strName = "ѪҺ��Ʒ��Ч��'ѪҺ��Ʒ����'��Ѫװ���Ƿ����'��������'����סԺ��'���߲���'���ߴ���'����Ѫ��'Ѫ����'ѪҺ��Ʒ����'����"
        arrName = Split(strName, "'")
        For i = 0 To UBound(arrName)
            .TextMatrix(i + 1, COL_���) = i + 1
            .TextMatrix(i + 1, COL_����) = arrName(i)
        Next i
        .ColDataType(COL_���) = flexDTBoolean
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, COL_���) = 1
        Next
        .Editable = flexEDKbdMouse
    End With
    
    Lbl�û���(0).Tag = UserInfo.id
    txt�û�(0).Text = UserInfo.�û���
    txt�û�(0).Tag = txt�û�(0).Text
    txt�û�(0).locked = Not mblnSelectSendUser: txt�û�(0).ForeColor = IIf(mblnSelectSendUser = False, COLOR.���ɫ, COLOR.��ɫ)
    txt����(0).Text = UserInfo.����
    txt����(0).Tag = txt����(0).Text
    txt����(0).locked = Not mblnSelectSendUser: txt����(0).ForeColor = IIf(mblnSelectSendUser = False, COLOR.���ɫ, COLOR.��ɫ)
    TXT����(0).Text = "123"
    TXT����(0).locked = True: TXT����(0).ForeColor = COLOR.���ɫ: TXT����(0).Enabled = mblnSelectSendUser
    picOper(0).Visible = mblnSelectSendUser: picOper(0).Enabled = mblnSelectSendUser
    If IsDate(mstr���ʱ��) Then
        DTPTime.Value = Format(mstr���ʱ��, "YYYY-MM-DD HH:mm")
    Else
        DTPTime.Value = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    End If
    DTPTime.Tag = DTPTime.Value
    
    Lbl�û���(0).Caption = marrHTitle(mintMode) & "���ʺ�"
    Lbl����(0).Caption = marrHTitle(mintMode) & "������"
    lbl����(0).Caption = marrHTitle(mintMode) & "����"
    Lbl�û���(1).Caption = marrFTitle(mintMode) & "���ʺ�"
    Lbl����(1).Caption = marrFTitle(mintMode) & "������"
    
    If mintMode = ��Ѫ�˶� Then
        Me.Caption = "��Ѫ�˶�"
    ElseIf mintMode = ���պ˶� Then
        Me.Caption = "���պ˶�"
    Else
        Me.Caption = "ִ�к˶�"
    End If
    mblnFirst = False
    mblnOk = False
End Sub

Private Sub picOper_Click(Index As Integer)
    If GetUserName(txt����(Index), Index) = True Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Function GetUserName(ByVal objControl As TextBox, ByVal intIndex As Integer, Optional ByVal StrInput As String = "") As Boolean
    Dim rsUser As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim vPoint As RECT, blnCancel As Boolean
    Dim str�������� As String, str��Ա���� As String
    
    On Error GoTo ErrHand
    If objControl.locked = True Then GetUserName = True: Exit Function
    If StrInput <> "" Then
         If IsNumeric(StrInput) Then
            strWhere = " And b.��� Like [2]"
         ElseIf gobjCommFun.IsCharAlpha(StrInput) Then
            strWhere = " And b.���� Like [2]"
            StrInput = UCase(StrInput)
         Else
            strWhere = " And b.���� Like [2]"
         End If
    End If
    
    '�Լ�վ�����ֲ������ʺ���Ա
    If Not mlngModul = pҽ������վ Then
        '��Ѫ��ΪѪ����Ա���ɣ�ȡѪ��Ϊ�ٴ���ʿ
         strWhere = strWhere & _
            "   And Exists  (Select 1 From ��������˵�� Where ����id = d.����id And Instr([3], ',' || �������� || ',', 1) <> 0 And ������� In (0, 1, 2, 3))"
        If Not (mintMode = 0 And intIndex = 0) Then  '��Ѫ
            strWhere = strWhere & _
                "   And Exists  (Select 1 From ��Ա����˵�� Where ��Աid = b.Id And Instr([4], ',' || ��Ա���� || ',', 1) <> 0) "
        End If
        
        If mintMode = 0 Then
            str�������� = IIf(intIndex = 0, ",Ѫ��,", ",�ٴ�,����,")
            str��Ա���� = IIf(intIndex = 0, "", ",��ʿ,")
        ElseIf mintMode = 1 Then
            str�������� = ",�ٴ�,����,"
            str��Ա���� = ",ҽ��,��ʿ,"
        Else
            str�������� = ",�ٴ�,����,"
            str��Ա���� = ",ҽ��,��ʿ,"
        End If
    End If

    vPoint = GetControlRect(objControl.hWnd)
    strSQL = _
        " Select  Rownum || '-' || b.Id as id, c.�û���,b.���, b.����,b.����,a.���� as ����" & vbNewLine & _
        " From ���ű� a, ��Ա�� b, �ϻ���Ա�� c, ������Ա d" & vbNewLine & _
        " Where a.Id = d.����id  And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) " & vbNewLine & _
        " " & strWhere & " And b.Id = c.��Աid  And c.��Աid = d.��Աid And d.����id = [1]"
    Set rsUser = gobjDatabase.ShowSQLSelect(Me, strSQL, 0, "", False, txt�û�(intIndex).Text, "��ѡ��һ��" & IIf(intIndex = 0, marrHTitle(mintMode), marrFTitle(mintMode)) & "��Ա", False, False, True, vPoint.Left, vPoint.Top, objControl.Height, blnCancel, False, False, _
                    IIf(intIndex = 0, mlngSendDeptID, mlngDeptID), StrInput & "%", str��������, str��Ա����)
    If Not rsUser Is Nothing Then
        If blnCancel = False Then
            If rsUser.EOF Then Exit Function
            Lbl�û���(intIndex).Tag = Split(rsUser!id, "-")(1)
            txt�û�(intIndex).Text = Nvl(rsUser!�û���)
            txt�û�(intIndex).Tag = txt�û�(intIndex).Text
            objControl.Text = Nvl(rsUser!����)
            objControl.Tag = objControl.Text
            objControl.SetFocus
            If intIndex = 0 Then '��Ѫ��
                If Lbl�û���(intIndex).Tag = UserInfo.id Then
                    TXT����(intIndex).Text = "123"
                    TXT����(intIndex).ForeColor = COLOR.���ɫ
                    TXT����(intIndex).locked = True
                Else
                    TXT����(intIndex).Text = ""
                    TXT����(intIndex).ForeColor = COLOR.��ɫ
                    TXT����(intIndex).locked = False
                End If
            End If
            GetUserName = True
        End If
    Else
        If StrInput = "" And blnCancel = False Then
            If mlngModul = pҽ������վ Then
                MsgBox "û�ж�Ӧ��ҽ����Ա��Ϣ��������Ա���������ã�", vbInformation, gstrSysName
            Else
                If mintMode = 0 Then
                    If intIndex = 0 Then
                        MsgBox "û�ж�Ӧ��Ѫ����Ա��Ϣ��������Ա���������ã�", vbInformation, gstrSysName
                    Else
                        MsgBox "û�ж�Ӧ���ٴ���ʿ��Ϣ��������Ա���������ã�", vbInformation, gstrSysName
                    End If
                Else
                    MsgBox "û�ж�Ӧ���ٴ���ʿ��ҽ����Ϣ��������Ա���������ã�", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    Exit Function
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub TXT����_GotFocus(Index As Integer)
    GetFocus TXT����(Index)
End Sub

Private Sub TXT����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub DTPTime_Validate(Cancel As Boolean)
    Dim blnOk As Boolean
    Dim strMsg As String, strCurDate As String

    '���ںϷ��Լ��
    blnOk = True: strMsg = ""
    If IsDate(mstr����ʱ��) Then
        If Format(DTPTime.Value, "YYYY-MM-DD HH:mm") < Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") Then
            blnOk = False
            If mintMode = ��Ѫ�˶� Then
                strMsg = "��Ѫ����ֻ������������[" & Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") & "]֮��"
            ElseIf mintMode = ���պ˶� Then
                strMsg = "��������ֻ���ڷ�Ѫ����[" & Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") & "]֮��"
            Else
                strMsg = "�˶�����ֻ���ڽ������ڻ��ϴ�ִ������[" & Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") & "]֮��"
            End If
            GoTo ShowMsg
        End If
    End If
    If IsDate(mstr���ʱ��) = False Then
        strCurDate = gobjDatabase.Currentdate
        If Format(DTPTime.Value, "YYYY-MM-DD HH:mm") > Format(strCurDate, "YYYY-MM-DD HH:mm") Then
            blnOk = False
            strMsg = marrHTitle(mintMode) & "���ڲ��ܴ��ڵ�ǰ����[" & Format(strCurDate, "YYYY-MM-DD HH:mm") & "]"
            GoTo ShowMsg
        End If
    Else
        If Format(DTPTime.Value, "YYYY-MM-DD HH:mm") > Format(mstr���ʱ��, "YYYY-MM-DD HH:mm") Then
            blnOk = False
            If mintMode = ��Ѫ�˶� Then
                strMsg = "��Ѫ���ڲ��ܴ��ڲ�����ɾ�������[" & Format(mstr���ʱ��, "YYYY-MM-DD HH:mm") & "]"
            ElseIf mintMode = ���պ˶� Then
                strMsg = "�������ڲ��ܴ��ڲ�����ɾ�������[" & Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") & "]"
            Else
                strMsg = "�˶����ڲ��ܴ��ڽ������ڻ��´�ִ������[" & Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") & "]"
            End If
            GoTo ShowMsg
        End If
    End If
ShowMsg:
    If blnOk = False Then
        MsgBox strMsg, vbInformation, gstrSysName
        Cancel = True
        DTPTime.Value = DTPTime.Tag
        DTPTime.SetFocus
        Exit Sub
    End If
    DTPTime.Tag = DTPTime.Value
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    GetFocus txt����(Index)
End Sub

Private Sub txt����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If Index = 1 Then
            DTPTime.SetFocus
        End If
    End If
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim StrInput As String
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        StrInput = txt����(Index).Text
        If StrInput <> "" And txt����(Index).Text <> txt����(Index).Tag Then
            If GetUserName(txt����(Index), Index, StrInput) = False Then Exit Sub
        End If
        gobjCommFun.PressKey vbKeyTab
    Else
        If KeyAscii = 39 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Index As Integer, Cancel As Boolean)
    If Index = 1 Then
        If txt����(Index).Tag <> "" And txt����(Index).Tag <> txt����(Index).Text Then txt����(Index).Text = txt����(Index).Tag
    End If
End Sub

Private Sub txt�û�_Change(Index As Integer)
    If Not mblnFirst Then Exit Sub
    CMDȷ��.Default = False
End Sub

Private Sub txt�û�_GotFocus(Index As Integer)
    GetFocus txt�û�(Index)
End Sub

Private Sub txt�û�_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub txt�û�_Validate(Index As Integer, Cancel As Boolean)
    Dim strText As String
    Dim lngUserID As Long, strUser As String, strName As String
    Dim rsOper As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim str�������� As String, str��Ա���� As String
    Dim str�������� As String
    
    strText = txt�û�(Index).Text
    
    On Error GoTo ErrHand
    If Index = 0 Or Index = 1 Then
        '��Ա��ȡ
        If txt�û�(Index).locked = True Then Exit Sub
        If strText = "" Then txt�û�(Index).Tag = txt�û�(Index).Text: Exit Sub
        strSQL = "Select a.Id, a.����,B.�û��� From ��Ա�� a, �ϻ���Ա�� b Where a.Id = b.��Աid And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And b.�û��� = [1]"
        Set rsOper = gobjDatabase.OpenSQLRecord(strSQL, "", UCase(strText))
        If rsOper.EOF Then
            txt�û�(Index).Text = txt�û�(Index).Tag
            lblInfo.Caption = "��ʾ��" & IIf(Index = 0, marrHTitle(mintMode), marrFTitle(mintMode)) & "���ʺŲ���ȷ"
            Call txt�û�_GotFocus(Index)
            Cancel = True
            Exit Sub
        End If
        lngUserID = Val(rsOper!id)
        strUser = "" & rsOper!�û���
        strName = "" & rsOper!����
        
        If Not mlngModul = pҽ������վ Then
            strWhere = strWhere & _
              "   And Exists  (Select 1 From ��������˵�� Where ����id = a.id And Instr([3], ',' || �������� || ',', 1) <> 0 And ������� In (0, 1, 2, 3))"
            If Not (mintMode = 0 And Index = 0) Then '��Ѫ�����ķ�Ѫ�˲���ָ����Ա����
                strWhere = strWhere & _
                    "   And Exists  (Select 1 From ��Ա����˵�� Where ��Աid = b.��Աid And Instr([4], ',' || ��Ա���� || ',', 1) <> 0) "
            End If
            
            If mintMode = 0 Then
                str�������� = IIf(Index = 0, ",Ѫ��,", ",�ٴ�,����,")
                str��Ա���� = IIf(Index = 0, "", ",��ʿ,")
            ElseIf mintMode = 1 Then
                str�������� = ",�ٴ�,����,"
                str��Ա���� = ",ҽ��,��ʿ,"
            Else
                str�������� = ",�ٴ�,����,"
                str��Ա���� = ",ҽ��,��ʿ,"
            End If
        End If
        
        'ִ�к˶ԣ�ͨ�������ʺ���ȡ��Ա���������Ƿ��ǵ�ǰ���ҵ�(ӦΪ��Ѫ���ˣ����ܴ���������ͬ���ҵ���֤)
        If Not mintMode = ִ�к˶� Then
            strSQL = "Select a.����, b.��Աid From ���ű� a, ������Ա b Where a.Id = b.����id And a.Id = [1]   And b.��Աid = [2] " & strWhere
            Set rsOper = gobjDatabase.OpenSQLRecord(strSQL, "", IIf(Index = 0, mlngSendDeptID, mlngDeptID), lngUserID, str��������, str��Ա����)
            If rsOper.EOF Then
                strSQL = " Select ���� from ���ű� where ID=[1]"
                Set rsOper = gobjDatabase.OpenSQLRecord(strSQL, "��ȡ��������", IIf(Index = 0, mlngSendDeptID, mlngDeptID))
                If rsOper.EOF Then
                    str�������� = ""
                Else
                    str�������� = "[" & rsOper!���� & "]"
                End If
                txt�û�(Index).Text = txt�û�(Index).Tag
                If mlngModul = pҽ������վ Then
                    lblInfo.Caption = "��ʾ����" & marrFTitle(mintMode) & "�˲������ڿ���" & str�������� & "��ҽ����Ա"
                Else
                    If mintMode = 0 Then
                        If Index = 0 Then
                            lblInfo.Caption = "��ʾ����" & marrHTitle(mintMode) & "�˲������ڵ�ǰ��Ѫ����" & str��������
                        Else
                            lblInfo.Caption = "��ʾ����" & marrFTitle(mintMode) & "�˲������ڲ��˵�ǰ����" & str�������� & "�Ļ�ʿ"
                        End If
                    Else
                        lblInfo.Caption = "��ʾ����" & marrFTitle(mintMode) & "�˲������ڿ���" & str�������� & "��ҽ����ʿ"
                    End If
                End If
                Call txt�û�_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
        End If
        Lbl�û���(Index).Tag = lngUserID
        txt�û�(Index).Text = strUser
        txt�û�(Index).Tag = txt�û�(Index).Text
        txt����(Index).Text = strName
        txt����(Index).Tag = txt����(Index).Text
        
        If Index = 0 Then  '��Ѫ��
            If Lbl�û���(Index).Tag = UserInfo.id Then
                TXT����(Index).Text = "123"
                TXT����(Index).ForeColor = COLOR.���ɫ
                TXT����(Index).locked = True
            Else
                TXT����(Index).Text = ""
                TXT����(Index).ForeColor = COLOR.��ɫ
                TXT����(Index).locked = False
            End If
        End If
            
        lblInfo.Caption = ""
    End If
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfCheck_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> COL_��� Then Cancel = True
End Sub

Private Sub LoadBoold()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    If mstrIDs = "" Then Exit Sub
    Set mBloodResult = New Collection
    picBlood.Enabled = True
    On Error GoTo ErrHand
    strSQL = _
        " Select /*+CARDINALITY(b 10)*/" & vbNewLine & _
        " a.Id, a.Ѫ�����, c.����, c.���" & vbNewLine & _
        " From �շ���ĿĿ¼ c, ѪҺ�շ���¼ a, Table(f_Num2list([1])) b" & vbNewLine & _
        " Where  c.Id = a.ѪҺid And a.Id = b.Column_Value"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, "��ȡѪҺ��Ϣ", mstrIDs)
    With cboBlood
        .Clear
        If rsTemp.RecordCount > 1 Then
            .AddItem "��������ѪҺͳһ����"
            .ItemData(.NewIndex) = -1
            mBloodResult.Add "11111111111", "B_0"
        End If
        Do While Not rsTemp.EOF
            'Ĭ�Ϻ˶Խ��������
            mBloodResult.Add "11111111111", "B_" & rsTemp!id
            .AddItem "���:" & rsTemp!Ѫ����� & "   ����:" & rsTemp!���� & "   ���" & rsTemp!���
            .ItemData(.NewIndex) = rsTemp!id
        rsTemp.MoveNext
        Loop
        gobjComlib.cbo.SetListHeight cboBlood, 360
        gobjComlib.cbo.SetListWidthAuto cboBlood
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0
    End With
    
    Exit Sub
ErrHand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub
