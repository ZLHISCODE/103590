VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����벡��"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmCheckIn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraInfo 
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   5940
      Begin VB.CheckBox chk���� 
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   3255
         TabIndex        =   1
         Top             =   1155
         Width           =   660
      End
      Begin VB.CheckBox chk��� 
         Caption         =   "�Ƿ����"
         Height          =   195
         Left            =   4710
         TabIndex        =   7
         Top             =   1155
         Width           =   1020
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1102
         Width           =   1530
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   671
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   780
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4620
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
      End
      Begin VB.ComboBox cbo���λ�ʿ 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1965
         Width           =   1830
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   671
         Width           =   1635
      End
      Begin VB.ComboBox cbo����ȼ� 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1533
         Width           =   4770
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3930
         TabIndex        =   5
         Top             =   1965
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   570
         TabIndex        =   26
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���λ�ʿ"
         Height          =   180
         Left            =   210
         TabIndex        =   25
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lbl��λ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ"
         Height          =   180
         Left            =   570
         TabIndex        =   24
         Top             =   1162
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   570
         TabIndex        =   23
         Top             =   731
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ȼ�"
         Height          =   180
         Left            =   210
         TabIndex        =   22
         Top             =   1593
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4200
         TabIndex        =   21
         Top             =   731
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2910
         TabIndex        =   20
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4200
         TabIndex        =   19
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�벡��ʱ��"
         Height          =   180
         Left            =   3000
         TabIndex        =   15
         Top             =   2025
         Width           =   900
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6165
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4350
      Width           =   6165
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4935
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3750
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   105
         TabIndex        =   10
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fraLvw 
      Caption         =   "��������"
      Height          =   1830
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   5940
      Begin MSComctlLib.ListView lvw 
         Height          =   1425
         Left            =   150
         TabIndex        =   6
         Top             =   255
         Width           =   5610
         _ExtentX        =   9895
         _ExtentY        =   2514
         View            =   2
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��λ"
            Object.Width           =   5292
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCheckIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mlng����ID As Long
Public mlng��ҳID As Long
Public mlngUnit As Long

Public mstr���� As String '��:ȱʡ��λ�Ĵ���,��ʾ��ͥ����,��:��ס�Ĵ���,���ܶ��Ŵ�,��,�ŷָ�
Public mlng��λ����ID As Long
Public mstrPrivs As String

Private mstrIDs As String
Private mstrText As String
Private mrsPatiInfo As ADODB.Recordset
Private mfrmParent As Object

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
    ElseIf cbo����.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo����ȼ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo����ȼ�.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo����ȼ�.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo����ȼ�.ListIndex = lngIdx
    ElseIf cbo����ȼ�.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo���λ�ʿ_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    If cbo���λ�ʿ.Text = "����..." Then
        Set rsTmp = GetSelectPersonal("��ʿ", "", Me)
        If Not rsTmp Is Nothing Then
            For i = 0 To cbo���λ�ʿ.ListCount - 1
                If cbo���λ�ʿ.List(i) = rsTmp!���� & "-" & rsTmp!���� Then
                    cbo���λ�ʿ.ListIndex = i: Exit Sub
                End If
            Next
            cbo���λ�ʿ.AddItem rsTmp!���� & "-" & rsTmp!����, cbo���λ�ʿ.ListCount - 1
            cbo���λ�ʿ.ListIndex = cbo���λ�ʿ.NewIndex
            cbo���λ�ʿ.ItemData(cbo���λ�ʿ.NewIndex) = rsTmp!�ϼ�ID
        Else
            cbo���λ�ʿ.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo���λ�ʿ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cbo���λ�ʿ.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cbo���λ�ʿ.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cbo���λ�ʿ.ListIndex = lngIdx
    ElseIf cbo���λ�ʿ.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub chk����_Click()
    If chk����.Value = 1 Then
        lbl��λ.Caption = "��Ҫ��λ"
        Call LoadMainBed
        lvw.Visible = True
        Me.Height = Me.Height + fraLvw.Height ' + 100
        If Visible Then lvw.SetFocus
    Else
        lbl��λ.Caption = "��λ"
        Call ShowBeds
        lvw.Visible = False
        Me.Height = Me.Height - fraLvw.Height '- 100
        If Visible Then cmdOK.SetFocus
    End If
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Integer
    Dim strSQLҽ��С�� As String
    Dim strIDs As String, strID As String, strCode As String
    Dim strTmp As String
    
    On Error GoTo errH
    gblnOK = False

    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID, 2)
    '��ʼ������
    With mrsPatiInfo
        txt����.Enabled = False
        
        '��ѡ�����Ŀ����벡�˿��Ҳ�ͬʱ,��������.
        If mlng��λ����ID <> 0 Then
            If mlng��λ����ID <> !��ס����id Then
                MsgBox "���˵�ǰ���ҡ�" & !��ǰ���� & "����ѡ��Ĵ�λ�������ҡ�" & GetDeptName(mlng��λ����ID) & "����ͬ,������ס�ô�λ,��ѡ��������λ!", vbInformation, gstrSysName
                Unload Me: Exit Sub
            End If
        End If

        txtDate.Text = Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")

        '������Ϣ
        txt����.Text = !����
        txt�Ա�.Text = "" & !�Ա�
        txt����.Text = "" & !����
        txt����.Text = "" & !��ǰ����
        'txtסԺ��.Text = "" & !סԺ��
        txt����.Tag = "" & !��ס����id
        
        
        'ȷ�������ķ������
        strSql = "Select ������� From ��������˵�� Where ��������='����' And ����ID=[1]" '
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit)
            
        '�д�λ���ٴ�����
        If rsTmp!������� = 1 Then
            strTmp = "1,3"
        ElseIf rsTmp!������� = 2 Then
            strTmp = "2,3"
        ElseIf rsTmp!������� = 3 Then
            If Val("" & !��������) = 1 Then
                strTmp = "1,3"
            Else
                strTmp = "2,3"
            End If
        End If
        Set rsTmp = GetDeptOrUnit(0, mlngUnit, strTmp)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                strIDs = strIDs & "," & rsTmp!ID
                rsTmp.MoveNext
            Next
        Else
            'û�ж�Ӧ�Ĵ�λ����
            MsgBox "�ڵ�ǰ����û�����ö�Ӧ����,���˲�����ס��" & vbCrLf, vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        
        '����
        strSql = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From ���� Order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
                If rsTmp!ȱʡ = 1 And cbo����.ListIndex = -1 Then cbo����.ListIndex = cbo����.NewIndex
                If rsTmp!���� = "" & !��ǰ���� Then cbo����.ListIndex = cbo����.NewIndex
                rsTmp.MoveNext
            Next
        End If
    
        '����ȼ�
        cbo����ȼ�.Enabled = InStr(mstrPrivs, ";" & "��������ȼ�" & ";") > 0
        Set rsTmp = GetNurseGrade
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo����ȼ�.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo����ȼ�.ItemData(cbo����ȼ�.NewIndex) = rsTmp!ID
                If rsTmp!ID = !����ȼ�ID Then cbo����ȼ�.ListIndex = cbo����ȼ�.NewIndex
                rsTmp.MoveNext
            Next
        End If
        
        'סԺ��ʿ
        Set rsTmp = GetDoctorOrNurse(1, strIDs & "," & mlngUnit & ",")
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo���λ�ʿ.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo���λ�ʿ.ItemData(i - 1) = rsTmp!ID
                rsTmp.MoveNext
            Next
            Call SeekDoctor(cbo���λ�ʿ, "" & mrsPatiInfo!���λ�ʿ)
        End If
        cbo���λ�ʿ.AddItem "����..."
        
        '��ʾ�ÿ��ҵĴ�λ
        Call ShowBeds
        If Not Visible Then chk����_Click
        
    End With
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngUnit = 0
    mstrText = ""
    'ж����Ϣ����
    If Not (mclsMipModule Is Nothing) Then
        Call mclsMipModule.CloseMessage
        Call DelMipModule(mclsMipModule)
        Set mclsMipModule = Nothing
    End If
    If Not (mclsXML Is Nothing) Then
        Set mclsXML = Nothing
    End If
End Sub

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadMainBed
End Sub

Private Sub LoadMainBed()
    Dim i As Integer, strBed As String
    
    If cbo����.ListIndex <> -1 Then strBed = cbo����.Text
    cbo����.Clear
    For i = 1 To lvw.ListItems.Count
        If lvw.ListItems(i).Checked Then
            cbo����.AddItem lvw.ListItems(i).Text
            If lvw.ListItems(i).Text = strBed Then cbo����.ListIndex = cbo����.NewIndex
            If cbo����.ListIndex = -1 And mstr���� <> "" Then
                If lvw.ListItems(i).Text = mstr���� Then cbo����.ListIndex = cbo����.NewIndex
            End If
        End If
    Next
    If cbo����.ListIndex = -1 And cbo����.ListCount = 1 Then cbo����.ListIndex = 0
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If IsDate(txtDate.Text) And KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub ShowBeds()
'���ܣ���ʾ��ǰ������ǰ���ҿ��õĲ���
    Dim i As Integer, objItem As ListItem
    Dim lng����ID As Long
    Dim rsBeds As ADODB.Recordset
    
    lvw.ListItems.Clear
    cbo����.Clear
    If InStr(1, mstrPrivs, "��ͥ����") > 0 Then
        cbo����.AddItem "��ͥ����"
        If mstr���� = "��ͥ����" Then cbo����.ListIndex = 0
    End If
    lng����ID = txt����.Tag
    Set rsBeds = GetFreeBeds(mlngUnit, lng����ID, mrsPatiInfo!�Ա�, mlng����ID)
    
    With rsBeds
        For i = 1 To rsBeds.RecordCount
            Set objItem = lvw.ListItems.Add(, "_" & !����, !���� & IIf(IsNull(!�����), "", " ����:" & !�����))
            objItem.Tag = !�ȼ�ID
            cbo����.AddItem objItem.Text
            
            If !���� = mstr���� Then
                objItem.Checked = True: objItem.Selected = True: objItem.EnsureVisible
                cbo����.ListIndex = cbo����.NewIndex
            End If
            
            .MoveNext
        Next
    End With
    
    If cbo����.ListIndex = -1 And cbo����.ListCount > 0 Then cbo����.ListIndex = 0
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim dMax As Date, i As Integer, Curdate As Date
    Dim strPreRoom As String, intRoom As Integer, intCheck As Integer, lngNurseGrade As Long
    Dim strSql As String, strBed As String, strTmp As String
    Dim str���� As String, str����� As String, blnTrans As Boolean, strMainBed As String
    Dim rsTmp As New ADODB.Recordset
        Dim colSQL As New Collection, strSQLtmp As String, rsPati As Recordset

    If cbo����.ListIndex = -1 Then
        MsgBox "��ָ�����˵ĵ�ǰ������", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    
    If cbo����ȼ�.ListIndex = -1 And gbln���ȷ������ȼ� Then
        MsgBox "��ָ�����˵ĵ�ǰ����ȼ���", vbInformation, gstrSysName
        cbo����ȼ�.SetFocus: Exit Sub
    End If
    
    If cbo����ȼ�.ListIndex <> -1 Then
        lngNurseGrade = cbo����ȼ�.ItemData(cbo����ȼ�.ListIndex)
    End If
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ����)
    Curdate = zlDatabase.Currentdate
    If InStr(Trim(cbo����.Text), " ����") <> 0 Then
        str���� = Mid(Trim(cbo����.Text), 1, InStr(Trim(cbo����.Text), " ����") - 1)
        str����� = Mid(Trim(cbo����.Text), InStr(Trim(cbo����.Text), "����:") + 3)
    ElseIf InStr(cbo����.Text, "��ͥ����") > 0 Then
        str���� = ""
    Else
        str���� = Trim(cbo����.Text)
    End If
    If CDate(txtDate.Text) > Curdate Then
        MsgBox "�벡��ʱ������˵�ǰϵͳʱ��,���飡", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    '��������Ժʱ����ͬ
    dMax = GetMaxDate(mlng����ID, mlng��ҳID)

    If Format(txtDate.Text, "yyyyMMddhhmmss") <= Format(dMax, "yyyyMMddHHmmss") Then
        MsgBox "ʱ�������ڸò��˵��ϴα䶯ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If

    
    If chk����.Value = 1 Then
        strPreRoom = "һ����ͬ"
        For i = 1 To lvw.ListItems.Count
            If lvw.ListItems(i).Checked Then
                intCheck = intCheck + 1
                strTmp = lvw.ListItems(i).Text
                If InStr(1, strTmp, ":") > 0 Then   'ð�ź��Ƿ����
                    strTmp = Mid(strTmp, InStr(1, strTmp, ":") + 1)
                    If strTmp <> strPreRoom Then
                        intRoom = intRoom + 1
                        strPreRoom = strTmp
                    End If
                End If
            End If
        Next
        If intCheck < 2 Then
            MsgBox "�������˱�������������ϵĴ�λ��", vbInformation, gstrSysName
            lvw.SetFocus: Exit Sub
        End If
        If intRoom > 1 Then
            MsgBox "��������������Ĵ�λ������һ�������ڣ�", vbInformation, gstrSysName
            lvw.SetFocus: Exit Sub
        End If
    End If
    
    If chk����.Value = 0 Then
        strBed = str����
        strMainBed = str����
    Else
        strMainBed = str����
        For i = 1 To lvw.ListItems.Count
            If lvw.ListItems(i).Checked Then
                strBed = strBed & "," & Mid(lvw.ListItems(i).Key, 2)
            End If
        Next
        strBed = Mid(strBed, 2)
    End If

    strSql = "zl_���˱䶯��¼_InUnit(" & mlng����ID & "," & mlng��ҳID & ",'" & strBed & "'," & _
            mlngUnit & "," & lngNurseGrade & ",'" & zlCommFun.GetNeedName(cbo����.Text) & "'," & chk���.Value & ",'" & zlCommFun.GetNeedName(cbo���λ�ʿ.Text) & "'," & _
            "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & UserInfo.��� & "','" & UserInfo.���� & "','" & strMainBed & "')"

    'ת�������ü��
    If CreatePublicExpenseBillOperation() And gblnת����ת���� Then
        strSQLtmp = "Select ID, ����id" & vbNewLine & _
                    "From ���˱䶯��¼" & vbNewLine & _
                    "Where ����id = [1] And ��ҳid = [2] And ��ʼʱ�� Is Not Null And ��ֹʱ�� Is Null And NVL(���Ӵ�λ,0) = 0"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQLtmp, Me.Caption, mlng����ID, mlng��ҳID)
        If rsPati.RecordCount > 0 Then
            If gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(Me, 0, mlng����ID, mlng��ҳID, Val(rsPati!ID & ""), Val(rsPati!����ID & ""), mlngUnit, colSQL) = False Then Exit Sub
        End If
    End If
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSql, Me.Caption
        For i = 1 To colSQL.Count
            zlDatabase.ExecuteProcedure colSQL(i), Me.Caption
        Next

    If Val("" & mrsPatiInfo!����) <> 0 Then
        If Not gclsInsure.ModiPatiSwap(mlng����ID, mlng��ҳID, Val("" & mrsPatiInfo!����), "1") Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    '����96847��118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng����ID, mlng��ҳID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    mstr���� = strBed
    gblnOK = True
    
    On Error Resume Next
    '�벡���ɹ��󴥷���Ϣ
    If mclsMipModule.IsConnect = True Then
        mclsXML.ClearXmlText '��������е�XML
        '--������Ϣ��װ
        '������Ϣ
        mclsXML.AppendNode "in_patient"
        'patient_id      ����id  1   N
        mclsXML.appendData "patient_id", mlng����ID, xsNumber  '����ID
        'page_id     ��ҳid  1   N
        mclsXML.appendData "page_id", mlng��ҳID, xsNumber '��ҳID
        'patient_name        ����    1   S
        mclsXML.appendData "patient_name", txt����.Text, xsString '����
        'patient_sex     �Ա�    0..1    S
        mclsXML.appendData "patient_sex", txt�Ա�.Text, xsString '�Ա�
        'in_number       סԺ��  1   S
        mclsXML.appendData "in_number", Nvl(mrsPatiInfo!סԺ��), xsString 'סԺ��
        mclsXML.AppendNode "in_patient", True
        
        strSql = " Select A.ID,B.���� ��λ�ȼ�,C.���� ��������  From  ���˱䶯��¼ A,�շ���ĿĿ¼ B,���ű� C" & _
            " Where NVl(A.���Ӵ�λ,0)=0 And A.��λ�ȼ�id=B.id(+) And A.����Id=C.id(+) And A.����ID=[1] And A.��ҳID=[2] And A.��ʼԭ��=[3] And ��ʼʱ��+0=[4]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���˱䶯��¼", mlng����ID, mlng��ҳID, 15, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        
        'סԺ��Ϣ
        mclsXML.AppendNode "in_hospital"
        'in_date     ��Ժʱ��    1   s
        mclsXML.appendData "in_date", Format(mrsPatiInfo!��Ժʱ��, "yyyy-MM-dd HH:mm:ss"), xsString
        'out_area_id     ת������id  0..1    N
        mclsXML.appendData "out_area_id", Val(Nvl(mrsPatiInfo!��ǰ����ID)), xsNumber
        'out_area_title      ת������    0..1    S
        mclsXML.appendData "out_area_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'out_dept_id     ת������id  1   N
        mclsXML.appendData "out_dept_id", Val(Nvl(mrsPatiInfo!��Ժ����id, 0)), xsNumber
        'out_dept_title      ת������    1   S
        mclsXML.appendData "out_dept_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'in_area_id      ת�벡��id  0..1    N
        mclsXML.appendData "in_area_id", mlngUnit, xsNumber
        'in_area_title       ת�벡��    0..1    S
        mclsXML.appendData "in_area_title", Nvl(rsTmp!��������), xsString
        'in_dept_id      ת�����id  1   N
        mclsXML.appendData "in_dept_id", Val(txt����.Tag), xsNumber
        'in_dept_title       ת�����    1   S
        mclsXML.appendData "in_dept_title", txt����.Text, xsString
        mclsXML.AppendNode "in_hospital", True
        'ת�����
        mclsXML.AppendNode "change_dept_arrange"
        'change_id       �䶯id  1   N
        mclsXML.appendData "change_id", rsTmp!ID, xsNumber '�䶯ID
        'in_room     ��ס����    0..1    S
        mclsXML.appendData "in_room", str�����, xsString
        'in_bed      ��ס����    1   S
        mclsXML.appendData "in_bed", strMainBed, xsString
        'in_tendgrade        ����ȼ�    0..1    S
        If cbo����ȼ�.ListIndex <> -1 Then
            mclsXML.appendData "in_tendgrade", zlCommFun.GetNeedName(cbo����ȼ�.Text), xsString
        Else
            mclsXML.appendData "in_tendgrade", "", xsString
        End If
        'in_bedgrade     ��λ�ȼ�    0..1    S
        mclsXML.appendData "in_bedgrade", Nvl(rsTmp!��λ�ȼ�), xsString
        'in_doctor       סԺҽʦ    0..1    S
        mclsXML.appendData "in_doctor", Nvl(mrsPatiInfo!סԺҽʦ), xsString
        'duty_nurse      ���λ�ʿ    0..1    S
        mclsXML.appendData "duty_nurse", zlCommFun.GetNeedName(cbo���λ�ʿ.Text), xsString
        'change_operator         ����Ա      1   S
        mclsXML.appendData "change_operator", UserInfo.����, xsString
        mclsXML.AppendNode "change_dept_arrange", True
        mclsMipModule.CommitMessage "ZLHIS_PATIENT_012", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
'����28432 by lesfeng 2010-03-10
Private Function GetDeptName(ByVal lngID As Long) As String

    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select ���� From ���ű� Where ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngID)
    
    If Not rsTmp.EOF Then
        GetDeptName = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SeekDoctor(cbo As ComboBox, Optional strPre As String)
    Dim strIDs As String, i As Integer
    
    If strPre <> "" Then
        For i = 0 To cbo.ListCount - 1
            If zlCommFun.GetNeedName(cbo.List(i)) = strPre Then cbo.ListIndex = i: Exit Sub
        Next
    End If
    
    strIDs = GetDeptDoctors(txt����.Tag)
    For i = 0 To cbo.ListCount - 1
        If InStr("," & strIDs & ",", "," & cbo.ItemData(i) & ",") > 0 Then cbo.ListIndex = i: Exit Sub
    Next
    
    strIDs = GetDeptDoctors(mlngUnit)
    For i = 0 To cbo.ListCount - 1
        If InStr("," & strIDs & ",", "," & cbo.ItemData(i) & ",") > 0 Then cbo.ListIndex = i: Exit Sub
    Next
End Sub


Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
        ByRef str���� As String, ByVal lng��λ����ID As Long, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstr���� = str����
    mlng��λ����ID = lng��λ����ID
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    str���� = mstr����
    ShowMe = gblnOK
End Function


