VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistoryDataMgr 
   BackColor       =   &H80000005&
   Caption         =   "����ת�ƿռ����"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmHistoryDataMgr.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFunction 
      Caption         =   "ת��(&T)"
      Height          =   350
      Index           =   6
      Left            =   6960
      TabIndex        =   14
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "�ϲ�(&L)"
      Height          =   350
      Index           =   5
      Left            =   5850
      TabIndex        =   12
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CheckBox chkAllֻ�� 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȫ��ֻ��(&A)"
      Height          =   285
      Left            =   4335
      TabIndex        =   11
      ToolTipText     =   "ֻ�Ա�������ʷ���ݿռ���Ч"
      Top             =   1260
      Width           =   1485
   End
   Begin VB.CheckBox chkֻ�� 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��ǰ��ֻ��(&Z)"
      Height          =   240
      Left            =   2430
      TabIndex        =   10
      ToolTipText     =   "ֻ�Ա�������ʷ���ݿռ���Ч"
      Top             =   1275
      Width           =   1695
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "����(&Z)"
      Height          =   350
      Index           =   3
      Left            =   3600
      TabIndex        =   9
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "�л�(&Q)"
      Height          =   350
      Index           =   4
      Left            =   4730
      TabIndex        =   8
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "��ֲ(&R)"
      Height          =   350
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "��ж(&M)"
      Height          =   350
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   4200
      Width           =   1100
   End
   Begin VB.CommandButton cmdFunction 
      Caption         =   "����(&C)"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   1100
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   615
      Width           =   3570
   End
   Begin MSComctlLib.ImageList imgSys 
      Left            =   5280
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoryDataMgr.frx":04F9
            Key             =   "Other"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoryDataMgr.frx":158B
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoryDataMgr.frx":497D
            Key             =   "Lock"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHistoryDataMgr.frx":7D6F
            Key             =   "LockAndRun"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwHistory 
      Height          =   2460
      Left            =   150
      TabIndex        =   3
      Top             =   1560
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   4339
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgSys"
      SmallIcons      =   "imgSys"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "��ǰ"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "���"
         Text            =   "ֻ��"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "������"
         Text            =   "������"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�汾��"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "���ת������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "���������"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "DB����"
         Text            =   "DB����"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblMain 
      BackColor       =   &H8000000E&
      Height          =   15
      Left            =   120
      TabIndex        =   13
      Top             =   5250
      Width           =   6360
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�Ѵ�������ʷ���ݿռ�"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ"
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ת�ƿռ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   1920
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmHistoryDataMgr.frx":B161
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmHistoryDataMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim intCount As Integer
Dim mblnHavingSpace As Boolean '������ʷ���ݱ�ռ�
Dim mblnSel As Boolean
Dim mintColumn As Integer
Dim mblnChk As Boolean

Private Enum ENUFT
    F0���� = 0
    F1��ж = 1
    F2��ֲ = 2
    F3���� = 3
    F4�л� = 4
    F5�ϲ� = 5
    F6ת�� = 6
End Enum
'0-�����������ݿռ�,1-��ж��ʷ���ݿռ�,2-��ֲ��ʷ���ݿռ�,3-���Ʒ�ת������,4���л��ڵ�ǰ����ʷ���ݿռ�,5-�ϲ���ʷ��ռ�,6-�������ʷ��ռ�ת�Ƶ���ǰ��

Private Enum ENUCOL
    C0��� = 0
    C1���� = 1
    C2��ǰ = 2
    C3ֻ�� = 3
    C4������ = 4
    C5�汾�� = 5
    C6���ת������ = 6
    C7��������� = 7
    C8DBLink = 8
End Enum

Private Sub chkAllֻ��_Click()
    Dim lngϵͳ As Long
    If mblnChk = True Then Exit Sub
    If cmbSystem.ListIndex < 0 Then Exit Sub
    lngϵͳ = cmbSystem.ItemData(cmbSystem.ListIndex)
    
    If lvwHistory.ListItems.Count = 0 Then Exit Sub
    
    Call SetHistoreReadPro(chkAllֻ��.value = 1, 0, lngϵͳ, True)
    Call cmbSystem_Click
End Sub

Private Sub chkֻ��_Click()
    Dim lngϵͳ As Long
    Dim str��� As String
    Dim lng�ռ��� As Long
    Dim blnֻ�� As Boolean
    Dim strImgKey As String
    If mblnSel = True Then Exit Sub
    If cmbSystem.ListIndex < 0 Then Exit Sub
    lngϵͳ = cmbSystem.ItemData(cmbSystem.ListIndex)
    
    err = 0: On Error Resume Next
    
    If lvwHistory.SelectedItem Is Nothing Then Exit Sub
    
    '���ܶ�Զ�̵���ʷ�ռ��������ֻ��
    If Mid(lvwHistory.SelectedItem.Tag, 3, 1) <> "1" Then Exit Sub
    
    lng�ռ��� = Val(Mid(lvwHistory.SelectedItem.Key, 2))
    If lng�ռ��� < 0 Then Exit Sub
    
    If SetHistoreReadPro(chkֻ��.value = 1, lng�ռ���, lngϵͳ, False) = False Then
        Exit Sub
    End If
    
    
    lvwHistory.SelectedItem.SubItems(C3ֻ��) = IIf(chkֻ��.value = 1, "��", "")
    lvwHistory.SelectedItem.Tag = Mid(lvwHistory.SelectedItem.Tag, 1, 1) & IIf(chkֻ��.value = 1, "1", "0") & Mid(lvwHistory.SelectedItem.Tag, 3, 1)
            
    If Val(Mid(lvwHistory.SelectedItem.Tag, 1, 1)) = 1 Then
        If chkֻ��.value = 1 Then
            strImgKey = "LockAndRun"
        Else
            strImgKey = "Run"
        End If
    Else
        If chkֻ��.value = 1 Then
            strImgKey = "Lock"
        Else
            strImgKey = "Other"
        End If
    End If
    lvwHistory.SelectedItem.SmallIcon = strImgKey
    lvwHistory.SelectedItem.Icon = strImgKey
    mblnChk = True
    chkAllֻ��.value = 2
    mblnChk = False
End Sub
Private Function SetHistoreReadPro(ByVal blnֻ�� As Boolean, ByVal lng��� As String, lngSys As Long, ByVal blnAll As Boolean) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '����:������ص���ʷ���ݿռ��ֻ������
    '����:blnֻ��-����Ϊֻ��
    '     lng���-�ռ���
    '     lngSys-ϵͳ��
    '     blnAll-������Ŀ
    '����:�ɹ�����true,����False
    '--------------------------------------------------------------------------------------------------------------
        
    err = 0: On Error GoTo errHand:
    gstrSQL = "Update zltools.zlbakspaces set ֻ��=" & IIf(blnֻ��, 1, 0) & " where  DB���� is null and  ϵͳ=" & lngSys & IIf(blnAll, "", " and ���=" & lng���)
    gcnOracle.Execute gstrSQL
    SetHistoreReadPro = True
    Exit Function
errHand:
    MsgBox "����ʧ��,��ϸ������Ϣ����:" & vbCrLf & "(" & err.Number & ")" & err.Description
End Function

Private Sub cmbSystem_Click()
    Dim lngϵͳ As Long
    Dim rsTemp As New ADODB.Recordset
    
    lngϵͳ = Val(cmbSystem.ItemData(cmbSystem.ListIndex))
    cmbSystem.Tag = GetOwnerName(lngϵͳ, gcnOracle)
    
    gstrSQL = "Select  1 From zlBakTables where rownum<=1 and  ϵͳ=" & lngϵͳ
    rsTemp.Open gstrSQL, gcnOracle
    If rsTemp.EOF Then
        mblnHavingSpace = False
    Else
        mblnHavingSpace = True
    End If
    Call LoadHistorySpace(lngϵͳ)
    '���ÿؼ�����
    Call SetCtlEnable
End Sub
Private Function LoadHistorySpace(ByVal lngϵͳ As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:������ʷ���ݿռ�
    '����:lngϵͳ-ϵͳ���
    '����:���سɹ�,����true,���򷵻�false
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsbakspaces As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strImgKey As String
    Dim objItem As ListItem
    Dim lngMaxLen As Long
    
    LoadHistorySpace = False
    
    gstrSQL = "Select max(length(���)) as MaxLen From zltools.zlbakspaces where ϵͳ=" & lngϵͳ
    OpenRecordset rsbakspaces, gstrSQL, "��ȡ��ʷ���ݿռ�"
    lngMaxLen = Val(Nvl(rsbakspaces!MaxLen))
    
    gstrSQL = "Select ϵͳ, ���, ����, ������, db����, ��ǰ, ֻ�� From zltools.zlbakspaces where ϵͳ=" & lngϵͳ & " Order by ���"
    OpenRecordset rsbakspaces, gstrSQL, "��ȡ��ʷ���ݿռ�"
    
    mblnSel = True
    err = 0: On Error Resume Next
    
    strImgKey = ""
    
    With lvwHistory
        .ListItems.Clear
        Do While Not rsbakspaces.EOF
            Set objItem = .ListItems.Add(, "K" & Nvl(rsbakspaces!���), Lpad(Nvl(rsbakspaces!���), lngMaxLen), 0, 0)
            objItem.SubItems(C1����) = Nvl(rsbakspaces!����)
            objItem.SubItems(C2��ǰ) = IIf(Val(Nvl(rsbakspaces!��ǰ)) = 1, "��", "")
            objItem.SubItems(C3ֻ��) = IIf(Val(Nvl(rsbakspaces!ֻ��)) = 1, "��", "")
            
            objItem.Tag = Val(Nvl(rsbakspaces!��ǰ)) & Val(Nvl(rsbakspaces!ֻ��)) & IIf(Nvl(rsbakspaces!db����) <> "", "0", "1")
            objItem.SubItems(C4������) = Nvl(rsbakspaces!������)
            objItem.SubItems(C8DBLink) = Nvl(rsbakspaces!db����)
                        
            If Val(Nvl(rsbakspaces!��ǰ)) = 1 Then
                If Val(Nvl(rsbakspaces!ֻ��)) = 1 Then
                    strImgKey = "LockAndRun"
                Else
                    strImgKey = "Run"
                End If
            Else
                If Val(Nvl(rsbakspaces!ֻ��)) = 1 Then
                    strImgKey = "Lock"
                Else
                    strImgKey = "Other"
                End If
            End If
            objItem.SmallIcon = strImgKey
            objItem.Icon = strImgKey
            
            On Error Resume Next
            gstrSQL = "select ϵͳ,�汾��,��������,���ת������,��������� from " & rsbakspaces!������ & ".ZLBAKINFO" & IIf(IsNull(rsbakspaces!db����), "", "@" & rsbakspaces!db����) & " where ϵͳ=" & lngϵͳ
            If rsTmp.State = 1 Then rsTmp.Close
            Set rsTmp = New ADODB.Recordset
            Call OpenRecordset(rsTmp, gstrSQL, gstrSysName, , , gcnOldOra) '������ʷ�ռ������Ȩ�޶��ڸ���Ӧ��ϵͳ�������ߵģ�����Ӧ���ܷ���
            If err <> 0 Or gcnOldOra.Errors.Count > 0 Then
                MsgBox "����:" & vbCrLf & "  ��ʷ���ݿռ�" & rsbakspaces!���� & "������������,����" & _
                    IIf(IsNull(rsbakspaces!db����), "Ȩ��", "DB����""" & rsbakspaces!db���� & """") & "�Ƿ�����? ", vbInformation + vbDefaultButton1
            Else
                If Not rsTmp.EOF Then
                    objItem.SubItems(C5�汾��) = Nvl(rsTmp!�汾��)
                    objItem.SubItems(C6���ת������) = Format(rsTmp!���ת������, "yyyy-mm-dd")
                    objItem.SubItems(C7���������) = Format(rsTmp!���������, "yyyy-mm-dd")
                End If
            End If
            err.Clear: err = 0

            If Nvl(rsbakspaces!��ǰ) = 1 Then
                objItem.Selected = True
                objItem.EnsureVisible
            End If
            rsbakspaces.MoveNext
        Loop
    End With
    mblnSel = False
    
    LoadHistorySpace = True
End Function

Private Sub SetCtlEnable(Optional blnCtlEnable As Boolean = True)
    '------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enable����
    '����:
    '------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim blnSys As Boolean
    Dim blnOwner As Boolean
    Dim blnCurSys As Boolean
    Dim blnSel As Boolean
    Dim blnֻ�� As Boolean
    Dim bln���� As Boolean
    
    blnSys = cmbSystem.ListIndex >= 0
    blnOwner = cmbSystem.Tag = gstrUserName
    
    If Me.lvwHistory.SelectedItem Is Nothing Then
        blnCurSys = False
        blnSel = False
        blnֻ�� = True
        bln���� = False
    Else
        blnSel = True
        blnCurSys = Mid(Me.lvwHistory.SelectedItem.Tag, 1, 1) = "1"
        blnֻ�� = Mid(Me.lvwHistory.SelectedItem.Tag, 2, 1) = "1"
        bln���� = Mid(Me.lvwHistory.SelectedItem.Tag, 3, 1) = "1"
    End If
    
    cmdFunction(F0����).Enabled = blnCtlEnable And blnSys And blnOwner And mblnHavingSpace
    cmdFunction(F2��ֲ).Enabled = cmdFunction(F0����).Enabled
    
    cmdFunction(F1��ж).Enabled = blnSel And blnCtlEnable And blnSys And blnOwner And mblnHavingSpace
    cmdFunction(F4�л�).Enabled = cmdFunction(F1��ж).Enabled
    cmdFunction(F3����).Enabled = blnSel And blnCtlEnable And blnSys And blnOwner And mblnHavingSpace And bln����
    
    cmdFunction(5).Enabled = blnSel And blnCtlEnable And blnSys And blnOwner And mblnHavingSpace And bln����
    
    Me.chkֻ��.Enabled = blnSel
    mblnSel = True
    If blnֻ�� Then
        Me.chkֻ��.value = 1
    Else
        Me.chkֻ��.value = 0
    End If
    chkֻ��.Enabled = bln����
        
    mblnSel = False
   ' cmbSystem.Enabled = blnCtlEnable
End Sub
 
Private Sub cmdFunction_Click(Index As Integer)
    '--------------------------------------------------------------------------------------------------------------------------
    '������ʷ���ݿռ�
    '--------------------------------------------------------------------------------------------------------------------------
    Dim blnSucced As Boolean
    Dim lngϵͳ As Long
    Dim str�ռ����� As String, str�ϲ��ռ��� As String
    Dim lng�ռ��� As Long
    Dim bln��ǰ As Boolean
    Dim blnֻ�� As Boolean
    Dim strNote As String

    Dim lngSelNum As Long, i As Long
    
    If cmbSystem.ListIndex < 0 Then Exit Sub
    lngϵͳ = cmbSystem.ItemData(cmbSystem.ListIndex)
    
    Call SetCtlEnable(False)
    If Index <> 0 Then
        If lvwHistory.SelectedItem Is Nothing Then Exit Sub
        lng�ռ��� = Val(Mid(lvwHistory.SelectedItem.Key, 2))
        bln��ǰ = Val(Mid(lvwHistory.SelectedItem.Tag, 1, 1)) = 1
        blnֻ�� = Val(Mid(lvwHistory.SelectedItem.Tag, 2, 1)) = 1
        
    Else
        lng�ռ��� = 0
    End If
    
    '//todo
    Select Case Index
    Case F0����
        '��odbc���Ӵ��룬��Ϊoledb�����ڴ���dblinkʱ����ʹû�д���Ҳ�����cn.Errors.Count > 0,����vb��err���󲶻񲻵�����
        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lngϵͳ, 0, 0)
    Case F1��ж
        If bln��ǰ Then
            MsgBox "����ʷ���ݿռ�Ϊ��ǰ��ʷ���ݿռ䣬���ܲ�ж�����ȶ�������ʷ�ռ�ʹ���л�����!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        End If
        
        lngSelNum = 0
        For i = 1 To lvwHistory.ListItems.Count
            If lvwHistory.ListItems(i).Checked Then
                If Val(Mid(lvwHistory.ListItems(i).Tag, 1, 1)) = 1 Then
                    MsgBox "ѡ�����ʷ���ݿռ�" & lvwHistory.ListItems(i).SubItems(C1����) & "Ϊ��ǰ��ʷ���ݿռ䣬���ܲ�ж!", vbInformation + vbDefaultButton1, gstrSysName
                    Exit Sub
                End If
                lngSelNum = lngSelNum + 1
            End If
        Next
        
        If lngSelNum > 0 Then
            If MsgBox("��ѡ����" & lngSelNum & "��Ҫ��ж����ʷ���ݿռ䣬��ȷ��Ҫ������", vbOKCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbCancel Then
                Exit Sub
            End If
            For i = 1 To lvwHistory.ListItems.Count
                If lvwHistory.ListItems(i).Checked Then
                    lng�ռ��� = Val(Mid(lvwHistory.ListItems(i).Key, 2))
                    blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lngϵͳ, 1, lng�ռ���)
                End If
            Next
        Else
            blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lngϵͳ, 1, lng�ռ���)
        End If
    Case F2��ֲ
        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lngϵͳ, 2, lng�ռ���)
    Case F3����
        
        If gstrServer = "" Then
            MsgBox "��ǰ��¼�ķ�����Ϊ�գ����Ʋ���Ҫ�����ָ���������������µ�¼��", vbInformation, gstrSysName
            Exit Sub
        End If
        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lngϵͳ, 3, lng�ռ���)
    Case F4�л�
        If bln��ǰ Then
            MsgBox "����ʷ���ݿռ�Ϊ��ǰ��ʷ���ݿռ䣬�����л�!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        Else
            If MsgBox("�������´���H����ͼָ��ǰѡ�����ʷ���ݿռ䣬�Ա��ѯ��ʷ���ݡ�" & vbCrLf & "��ȷ��Ҫ����ǰ��ʷ���ݿռ��л�" & lvwHistory.SelectedItem.SubItems(C1����) & "Ϊ��", vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lngϵͳ, 4, lng�ռ���)
        If blnSucced Then
            '������Ҫ������־
            Call SaveAuditLog(2, "�л�", "����ʷ���ݿռ��л�Ϊ" & lvwHistory.SelectedItem.SubItems(C1����))
        End If
    Case F5�ϲ�
        lngSelNum = 0
        lng�ռ��� = 0
        For i = 1 To lvwHistory.ListItems.Count
            If lvwHistory.ListItems(i).Checked Then
                lngSelNum = lngSelNum + 1
                'ȡ��С�ռ���Ϊ�����ռ���
                If Val(Mid(lvwHistory.ListItems(i).Key, 2)) < lng�ռ��� Or lng�ռ��� = 0 Then
                    lng�ռ��� = Val(Mid(lvwHistory.ListItems(i).Key, 2))
                    str�ռ����� = lvwHistory.ListItems(i).SubItems(C1����)
                    strNote = str�ռ�����
                End If
            End If
        Next
        
        If lngSelNum < 2 Then
            MsgBox "�����ٹ�ѡ2��Ҫ�ϲ�����ʷ���ݿռ䡣", vbInformation + vbDefaultButton1, gstrSysName
        Else
            If MsgBox("��ѡ��" & lngSelNum & "���ռ�����ݽ��ᱻ�ϲ��������С�Ŀռ䡾" & str�ռ����� & "���С�" & vbCrLf & _
                    "��ɺ󣬱��ϲ��ռ估�����ļ����ᱻɾ��,��ȷ���ѽ�����Ч���ݡ�" & vbCrLf & "��ȷ��Ҫ������", vbOKCancel + vbDefaultButton1 + vbQuestion, gstrSysName) = vbOK Then
                
                For i = 1 To lvwHistory.ListItems.Count
                    If lvwHistory.ListItems(i).Checked Then
                        If lng�ռ��� <> Mid(lvwHistory.ListItems(i).Key, 2) Then
                            str�ϲ��ռ��� = str�ϲ��ռ��� & "," & Mid(lvwHistory.ListItems(i).Key, 2)
                            strNote = strNote & "," & lvwHistory.ListItems(i).SubItems(C1����)
                        End If
                    End If
                Next
                blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lngϵͳ, 5, lng�ռ���, Mid(str�ϲ��ռ���, 2))
                If blnSucced Then
                    '������Ҫ������־
                    Call SaveAuditLog(2, "�ϲ�", "����ʷ���ݿռ�" & strNote & "�ϲ�Ϊ" & str�ռ�����)
                End If
            End If
        End If
    Case F6ת��

        blnSucced = frmHistorySpaceSet.ShowInstall(Me, gcnOldOra, gstrUserName, gstrPassword, lngϵͳ, 6, 0)
    End Select
    
    Call SetCtlEnable(True)
    If blnSucced = True Then
        Call cmbSystem_Click
    End If
    
End Sub
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    If gblnDBA Then
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", gstrUserName)
    End If
    
    With rsTemp
        .Filter = "���=100 or ���=2100 or ���=2400"
        Do While Not .EOF
            cmbSystem.addItem !���� & " v" & !�汾�� & "��" & !��� & "��"
            cmbSystem.ItemData(cmbSystem.NewIndex) = !���
            .MoveNext
        Loop
        If cmbSystem.ListCount = 0 Then
            Call SetCtlEnable
        End If
        If cmbSystem.ListCount > 0 Then cmbSystem.ListIndex = 0
        If cmbSystem.ListCount = 1 Then cmbSystem.Locked = True
    End With
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub


Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    With cmbSystem
        .Width = ScaleWidth - .Left - 50
    End With
    
    With lvwHistory
        .Width = ScaleWidth - .Left - 50
    End With
    With chkAllֻ��
        .Left = lvwHistory.Left + lvwHistory.Width - .Width - 50
        chkֻ��.Left = .Left - chkֻ��.Width - 50
    End With
    
    With cmdFunction(F0����)
        .Left = lvwHistory.Left
        .Top = ScaleHeight - .Height - 100
        
        cmdFunction(F1��ж).Top = .Top
        cmdFunction(F2��ֲ).Top = .Top
        cmdFunction(F3����).Top = .Top
        cmdFunction(F4�л�).Top = .Top
        cmdFunction(F5�ϲ�).Top = .Top
        cmdFunction(F6ת��).Top = .Top
        
        cmdFunction(F1��ж).Left = .Left + .Width + 15
        cmdFunction(F2��ֲ).Left = cmdFunction(F1��ж).Left + cmdFunction(F1��ж).Width + 15
        
        cmdFunction(F6ת��).Left = ScaleWidth - cmdFunction(F6ת��).Width - 60
        cmdFunction(F5�ϲ�).Left = cmdFunction(F6ת��).Left - cmdFunction(F5�ϲ�).Width - 15
        cmdFunction(F4�л�).Left = cmdFunction(F5�ϲ�).Left - cmdFunction(F4�л�).Width - 15
        cmdFunction(F3����).Left = cmdFunction(F4�л�).Left - cmdFunction(F3����).Width - 15
        
    End With
    lvwHistory.Height = cmdFunction(F0����).Top - lvwHistory.Top - 10
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As zlPrintLvw
    
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "�Ѱ�װ���������ݿռ�"
    Set objPrint.Body.objData = lvwHistory
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

 

Private Sub lvwHistory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
        lvwHistory.SortOrder = IIf(lvwHistory.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        mintColumn = ColumnHeader.Index - 1
        lvwHistory.SortKey = mintColumn
        lvwHistory.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwHistory_DblClick()

    
    If lvwHistory.SelectedItem Is Nothing Then Exit Sub
    mblnSel = True
    If lvwHistory.SelectedItem.SubItems(C3ֻ��) = "��" Then
        chkֻ��.value = 0
    Else
        chkֻ��.value = 1
    End If
    mblnSel = False
   Call chkֻ��_Click
   
    '����ʱclick��ʽ���ò�ִ��
   If Me.Visible Then Call SetControlAllֻ��
End Sub

Private Sub SetControlAllֻ��()
    Dim i As Long, lngSel As Long

    lngSel = 0
    For i = 1 To lvwHistory.ListItems.Count
        If lvwHistory.ListItems(i).SubItems(C3ֻ��) = "��" Then lngSel = lngSel + 1
    Next
    
    mblnChk = True
    chkAllֻ��.value = IIf(lngSel = 0, 0, IIf(lngSel = lvwHistory.ListItems.Count, 1, 2))
    mblnChk = False
End Sub

Private Sub lvwHistory_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SetCtlEnable
End Sub


