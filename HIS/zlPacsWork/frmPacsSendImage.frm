VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.2#0"; "DicomObjects.ocx"
Begin VB.Form frmPacsSendImage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ͼ��"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9645
   Icon            =   "frmPacsSendImage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame3 
      Caption         =   "ͼ��ԤԤ��"
      Height          =   3915
      Left            =   5160
      TabIndex        =   31
      Top             =   60
      Width           =   4425
      Begin DicomObjects.DicomViewer Viewer 
         Height          =   3135
         Left            =   120
         TabIndex        =   35
         Top             =   570
         Width           =   4155
         _Version        =   262146
         _ExtentX        =   7329
         _ExtentY        =   5530
         _StockProps     =   35
      End
      Begin VB.CheckBox ChkShowImage 
         Caption         =   "Ԥ��"
         Height          =   285
         Left            =   150
         TabIndex        =   34
         Top             =   210
         Value           =   1  'Checked
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ѯ"
      Height          =   1545
      Left            =   30
      TabIndex        =   25
      Top             =   4020
      Width           =   9555
      Begin VB.TextBox txtPatientID 
         Height          =   300
         Left            =   7140
         MaxLength       =   18
         TabIndex        =   4
         Top             =   420
         Width           =   2130
      End
      Begin VB.CommandButton CmdRefresh 
         Cancel          =   -1  'True
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   8190
         TabIndex        =   9
         Top             =   1020
         Width           =   1100
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   3690
         MaxLength       =   12
         TabIndex        =   6
         Top             =   1050
         Width           =   1590
      End
      Begin VB.CheckBox chk��Դ 
         Caption         =   "���ﲡ��"
         Height          =   195
         Index           =   0
         Left            =   5460
         TabIndex        =   7
         Top             =   1110
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.CheckBox chk��Դ 
         Caption         =   "סԺ����"
         Height          =   195
         Index           =   1
         Left            =   6750
         TabIndex        =   8
         Top             =   1110
         Value           =   1  'Checked
         Width           =   1020
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1050
         Width           =   3270
      End
      Begin VB.TextBox txtChkNoEnd 
         Height          =   300
         Left            =   5460
         MaxLength       =   18
         TabIndex        =   3
         Top             =   420
         Width           =   1260
      End
      Begin VB.TextBox txtChkNoBegin 
         Height          =   300
         Left            =   3690
         MaxLength       =   18
         TabIndex        =   2
         Top             =   420
         Width           =   1260
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   25165825
         CurrentDate     =   38082
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   150
         TabIndex        =   0
         Top             =   420
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   25165825
         CurrentDate     =   38082
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ��"
         Height          =   180
         Left            =   7140
         TabIndex        =   36
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   5130
         TabIndex        =   30
         Top             =   480
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   1680
         TabIndex        =   29
         Top             =   480
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         Height          =   180
         Left            =   3690
         TabIndex        =   23
         Top             =   810
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Դ"
         Height          =   180
         Left            =   5460
         TabIndex        =   24
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���˿���"
         Height          =   180
         Left            =   150
         TabIndex        =   22
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3690
         TabIndex        =   21
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��"
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   210
         Width           =   720
      End
   End
   Begin MSComctlLib.TreeView tvwImageDate 
      Height          =   3915
      Left            =   30
      TabIndex        =   28
      Top             =   60
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   6906
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8340
      TabIndex        =   18
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   7005
      TabIndex        =   17
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton CmdSelectAll 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   5685
      TabIndex        =   16
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton CmdSelectClear 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   4350
      TabIndex        =   15
      Top             =   6750
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   19
      Top             =   6750
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "��������"
      Height          =   945
      Left            =   60
      TabIndex        =   26
      Top             =   5640
      Width           =   9525
      Begin VB.OptionButton ChkImageFormat 
         Caption         =   "JPG"
         Height          =   225
         Index           =   2
         Left            =   5730
         TabIndex        =   13
         Top             =   510
         Width           =   1005
      End
      Begin VB.OptionButton ChkImageFormat 
         Caption         =   "BMP"
         Height          =   225
         Index           =   1
         Left            =   4770
         TabIndex        =   12
         Top             =   510
         Width           =   1005
      End
      Begin VB.OptionButton ChkImageFormat 
         Caption         =   "DICOM"
         Height          =   225
         Index           =   0
         Left            =   3630
         TabIndex        =   11
         Top             =   510
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.CheckBox ChkSendTool 
         Caption         =   "���Ͷ�����Ƭվ"
         Height          =   315
         Left            =   7290
         TabIndex        =   14
         Top             =   450
         Width           =   2025
      End
      Begin VB.ComboBox CboPath 
         Height          =   300
         ItemData        =   "frmPacsSendImage.frx":000C
         Left            =   150
         List            =   "frmPacsSendImage.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ͼ���ʽ"
         Height          =   180
         Left            =   3660
         TabIndex        =   33
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����Ŀ¼"
         Height          =   180
         Left            =   150
         TabIndex        =   32
         Top             =   240
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   27
      Top             =   7245
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   1764
            Picture         =   "frmPacsSendImage.frx":0010
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12409
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPacsSendImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************API����*****************************************
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)


Private Sub Check1_Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub ShowMe(frmobj As Object)
    
    Me.Show vbModal, frmobj
End Sub

Private Sub CmdRefresh_Click()
    Me.CmdRefresh.Enabled = False
    If zlCommFun.StrIsValid(Me.txtChkNoBegin) = False Then
        With Me.txtChkNoBegin
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    If zlCommFun.StrIsValid(Me.txtChkNoEnd) = False Then
        With Me.txtChkNoEnd
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    If zlCommFun.StrIsValid(Me.txtChkNoEnd) = False Then
        With Me.txtChkNoEnd
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    zl9comlib.zlCommFun.ShowFlash "��ȴ����ڶ�ȡ����.....", Me
    zl9comlib.zlCommFun.ShowFlash
    RefreshImageDate
    zl9comlib.zlCommFun.StopFlash
    AllSelectOrAllClear True
    Me.CmdRefresh.Enabled = True
End Sub

Private Sub CmdSelectAll_Click()
    AllSelectOrAllClear True
End Sub

Private Sub CmdSelectClear_Click()
    AllSelectOrAllClear False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdSend_Click()
    '����ͼ���ٷ���
    Dim i As Long, j As Long, m As Long, n As Long
    Dim strPath As String                   '����·��
    Dim strDate As String                   '�ڵ�ʱ��
    Dim strName As String                   '�ڵ�����
    Dim strSq As String                     '�ڵ����к�
    Dim xNodes As Node                      '�ڵ����
    Dim blWriteSucceed  As Boolean          'д���Ƿ�ɹ�
    Dim strSql As String                    '���SQL������
    Dim strTmp As String                    '��ʱ�ִ��ֽ����
    Dim strPas As String                    'Զ��Ŀ¼����
    Dim strUse As String                    'Զ���û���
    Dim rsTmp As New ADODB.Recordset        '��ʱ��¼��
    Dim duTime As Double                    '��¼ʱ���������������糬ʱ
    Dim strRemotePath As String             'Զ��Ŀ¼·��
    Dim DicomPath As New DicomDataSet       '����DIR�ļ�
    Dim objFile As New Scripting.FileSystemObject           '�����ļ�ʹ��
    On Error GoTo SendErr
    
    If Me.CboPath.Text = "" Then
        MsgBox "��ѡ��Ҫ���͵�Ŀ¼!", vbInformation, Me.Caption
        Exit Sub
    End If
    '���Ŀ¼�Ƿ��д
    
    If Me.CboPath.List(Me.CboPath.ListIndex) <> "" Then
        strTmp = Mid(Me.CboPath.Text, 1, InStr(1, Me.CboPath.Text, "_") - 1)
        strSql = "select �豸��,����Ŀ¼,�û���,���� from Ӱ���豸Ŀ¼ where ���� = 5 and �豸�� = [1]"
        Set rsTmp = OpenSQLRecord(strSql, Me.Caption, strTmp)
        If rsTmp.EOF = True Then
            MsgBox "û���ҵ�Ҫ����Ŀ¼����Ϣ��", vbInformation, Me.Caption
            Exit Sub
        End If
        zl9comlib.zlCommFun.ShowFlash "�������Ӳ�����ȴ�.....", Me
        zl9comlib.zlCommFun.ShowFlash
        strRemotePath = rsTmp("����Ŀ¼")
        strPas = Nvl(rsTmp("����"))
        strUse = Nvl(rsTmp("�û���"))
        Shell "net use " & strRemotePath & " " & strPas & " /user:" & strUse, vbHide
            
        duTime = Timer
        Do Until CLng(Timer - duTime) >= 20
            Shell "net use " & strRemotePath & " " & strPas & " /user:" & strUse, vbHide
            If WriteTest(False, strRemotePath) = True Then
                Exit Do
            End If
            DoEvents
        Loop
        zl9comlib.zlCommFun.StopFlash
    End If
    
    If WriteTest(False, strRemotePath) = False Then
        MsgBox "д�����ʧ�����鹲��Ŀ¼!", vbQuestion, App.EXEName
        Exit Sub
    End If
    
    With Me.tvwImageDate
        If .Nodes.Count = 0 Then
            MsgBox "û�п��Է��͵��ļ�!��ѡ���ѯ��������ˢ�¸����б�!", vbInformation, App.EXEName
            Exit Sub
        End If
        '�����ǰ�Ĵ�����־
        If Dir(App.Path & "\WriteErrLog.txt") <> vbNullString Then
            Kill App.Path & "\WriteErrLog.txt"
        End If
        
        Me.CmdOK.Enabled = False
        Me.cmdCancel.Enabled = False
        Me.CmdSelectAll.Enabled = False
        Me.CmdSelectClear.Enabled = False
        Me.CmdRefresh.Enabled = False
        Me.MousePointer = 11
        zlCommFun.ShowFlash "���ڶ����ļ���ȴ�.....", Me
        zlCommFun.ShowFlash
        For i = 1 To .Nodes.Count
            If .Nodes(i).Checked = True And .Nodes(i).Child Is Nothing Then
                strDate = Mid(.Nodes(i).Parent.Parent.Text, InStr(1, .Nodes(i).Parent.Parent.Text, "[") + 1)
                strDate = Format(Mid(strDate, 1, InStr(1, strDate, "]") - 1), "yyyymmdd")
                strName = Mid(.Nodes(i).Parent.Parent.Text, 1, InStr(1, .Nodes(i).Parent.Parent.Text, "[") - 1)
                strSq = .Nodes(i).Key
                strPath = strDate & "\" & strName & "\" & Mid(strSq, 1, InStr(strSq, "_") - 1)
                '����
                If SendFilesToDir(strSq, DicomPath, strRemotePath & "\DICOM\IMAGE\", strPath) Then
                    'ʧ��
                    .Nodes(i).Checked = False
                    blWriteSucceed = False
                    m = m + 1
                Else
                    '�ɹ�
                    blWriteSucceed = True
                    n = n + 1
                End If
                j = j + 1
            End If
            '�²���ʱ��������
            If .Nodes(i).Parent Is Nothing Then
                j = 0
            End If
            DoEvents
            Me.stbThis.Panels(2).Text = "���ڷ��Ͳ���[" & strName & "]�ĵ�" & j & "��ͼ��. �����" & CInt((i) / .Nodes.Count * 100) & "%."
            .Nodes(i).Checked = blWriteSucceed
            blWriteSucceed = False
        Next
        
        If Me.ChkSendTool.Value = 1 Then
            If Dir(App.Path & "\PacsLite\") <> "" Then
                objFile.CopyFile App.Path & "\PacsLite\*.*", strRemotePath & "\"
            End If
        End If
        If m > 0 Then
            DicomPath.WriteDirectory IIf(Len(strRemotePath) > 3, strRemotePath & "\DICOM\DICOMDIR", strRemotePath & "\DICOM\DICOMDIR")
        End If
        zl9comlib.zlCommFun.StopFlash
        
        Me.stbThis.Panels(2).Text = "�������!���ͳɹ�" & m & "��ͼ��,ʧ��" & n & "��ͼ��."
        If n > 0 Then
            If MsgBox("�������!���ͳɹ�" & m & "��ͼ��,ʧ��" & n & "��ͼ��." & _
            vbCrLf & "�鿴��־��ѡ��[��]", vbYesNo + vbDefaultButton2 + vbInformation, App.EXEName) = vbYes Then
                Shell "Notepad " & App.Path & "\WriteErrLog.txt", vbNormalFocus
            End If
        Else
            MsgBox "�������!���ͳɹ�" & m & "��ͼ��,ʧ��" & n & "��ͼ��.", vbInformation, App.EXEName
        End If
        Me.MousePointer = 0
    End With
    Shell "net use " & strRemotePath & " /delete "
    Me.CmdOK.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.CmdSelectAll.Enabled = True
    Me.CmdSelectClear.Enabled = True
    Me.CmdRefresh.Enabled = True
    Exit Sub
SendErr:
    Me.MousePointer = 0
    zl9comlib.zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Me.CmdOK.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.CmdSelectAll.Enabled = True
    Me.CmdSelectClear.Enabled = True
    Me.CmdRefresh.Enabled = True
End Sub



Private Sub dtpBegin_Change()
    Me.dtpBegin.MaxDate = Me.dtpEnd.Value
End Sub

Private Sub dtpEnd_Change()
    Me.dtpEnd.MinDate = Me.dtpBegin.Value
End Sub


Private Function LoadData() As Boolean
'���ܣ����ݲ�����Դ��ȡ���˿���
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim lngPre As Long
    
    
    If cboDept.ListIndex <> -1 Then
        lngPre = cboDept.ItemData(cboDept.ListIndex)
    End If
    strSql = "Select Distinct A.ID,A.����,A.����,B.�������" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.�������� IN('���')" & _
        " And B.������� IN(3," & IIf(chk��Դ(0).Value, 1, -1) & "," & IIf(chk��Դ(1).Value, 2, -1) & ")" & _
        " And (A.����ʱ�� is NULL Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.����"
    On Error GoTo errH
    Call OpenRecord(rsTmp, strSql, Me.Caption)
    On Error GoTo 0
    cboDept.Clear
    cboDept.AddItem "���п���"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        If rsTmp!ID = lngPre Then cboDept.ListIndex = cboDept.NewIndex
        rsTmp.MoveNext
    Next
    strSql = "select �豸��, �豸��,����Ŀ¼,�û���,����  from Ӱ���豸Ŀ¼ where ���� = 5"
    
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption)
    Me.CboPath.Clear
    Do Until rsTmp.EOF
        Me.CboPath.AddItem rsTmp("�豸��") & "_" & rsTmp("�豸��")
        rsTmp.MoveNext
    Loop
    If CboPath.ListCount > 0 Then
        CboPath.ListIndex = 0
    End If
    
    LoadData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chk��Դ_Click(Index As Integer)
    If chk��Դ(0).Value = 0 And chk��Դ(1).Value = 0 Then
        chk��Դ((Index + 1) Mod 2).Value = 1
    End If
    Call LoadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "�������", CInt(Me.dtpEnd - Me.dtpBegin)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "�������ʱ��", Format(Me.dtpEnd.Value, "yyyy-mm-dd")
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "����", Me.cboDept.ListIndex
'    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "����·��", Me.TxtPath
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "���ﲡ��", Me.chk��Դ(0).Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "סԺ����", Me.chk��Դ(1).Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "����Ŀ¼", Me.CboPath.ListIndex
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "Ԥ��", Me.ChkShowImage.Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "��ʽ0", Me.ChkImageFormat(0).Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "��ʽ1", Me.ChkImageFormat(1).Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "��ʽ2", Me.ChkImageFormat(2).Value
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "��Ƭվ", Me.ChkSendTool.Value
    
    Unload Me
End Sub
Private Sub Form_Load()
    Dim intDept As Integer
    Dim intDiffDay As Integer
    Dim intPath As String
    intDiffDay = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "�������", 3)
    Me.dtpBegin = Format(Now - intDiffDay, "yyyy-mm-dd")
    Me.dtpEnd = Format(Now, "yyyy-mm-dd")
    intDept = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "����", 0)
'    Me.TxtPath = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "����·��", "")
    Me.chk��Դ(0).Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "���ﲡ��", 1)
    Me.chk��Դ(1).Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "סԺ����", 1)
    Me.ChkShowImage.Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "Ԥ��", 1)
    Me.ChkImageFormat(0).Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "��ʽ0", True)
    Me.ChkImageFormat(1).Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "��ʽ1", False)
    Me.ChkImageFormat(2).Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "��ʽ2", False)
    Me.ChkSendTool.Value = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "��Ƭվ", 1)
    intPath = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\frmPacsSendImage", "����Ŀ¼", 0)
    LoadData
    SendMessage Me.cboDept.Hwnd, CB_SETCURSEL, intDept, 0
    SendMessage Me.CboPath.Hwnd, CB_SETCURSEL, intPath, 0
End Sub
Private Sub RefreshImageDate()
    'ˢ�²���ͼ��
    Dim rsMain As New ADODB.Recordset
    Dim rsNode As New ADODB.Recordset
    Dim strSql As String
    Dim strDeptNo As String
    Dim intMZ As Integer
    Dim intZY As Integer
    Dim blnMoved As Boolean
    Dim strSQLBak As String
    Dim strSerialUID As String
    Dim strPatientUID As String
    Dim i As Integer
    On Error GoTo RefreshError
    
    blnMoved = MovedByDate(Me.dtpBegin.Value)
    
    strDeptNo = Me.cboDept.Text
    
    If strDeptNo <> "���п���" Then
        strDeptNo = Mid(strDeptNo, 1, InStr(1, strDeptNo, "-") - 1)
    End If
    strSql = "select a.ҽ��ID,a.Ӱ�����,a.����,a.����,a.���UID,a.��������,b.������Դ,c.����||'-'||c.���� as ��������, " & _
             " d.�״�ʱ��,e.����UID,f.ͼ��UID,e.�������� from Ӱ�����¼ a , " & _
             " ����ҽ����¼ b , ���ű� c , ����ҽ������ d , Ӱ�������� e , Ӱ����ͼ�� f , ������Ϣ g " & _
             " where a.ҽ��id = b.id and b.ִ�п���id = c.id and b.id = d.ҽ��id and  " & _
             " a.���UID = e.���UID and e.����UID = f.����UID and b.����ID = g.����ID and " & _
             " d.�״�ʱ�� >= [1] and d.�״�ʱ�� <= [2] "
    If strDeptNo = "���п���" Then
        strSql = strSql & " and [3] = [3] "
    Else
        strSql = strSql & " and c.���� = [3] "
    End If
    If Trim(txtChkNoBegin) = "" Or Trim(txtChkNoEnd) = "" Then
        strSql = strSql & " and [4] = [4] and [5] = [5] "
    Else
        strSql = strSql & " and a.���� >= [4] and a.���� <= [5] "
    End If
    If Trim(txt����) = "" Then
        strSql = strSql & " and [6]= [6] "
    Else
        strSql = strSql & " and a.���� = [6] "
    End If
    If Me.chk��Դ(0).Value = 1 Then
        intMZ = 1
    Else
        intMZ = 3
    End If
    If Me.chk��Դ(1).Value = 1 Then
        intZY = 2
    Else
        intZY = 3
    End If
    strSql = strSql & " and  b.������Դ in (3,4,[7],[8]) "
    If Len(Trim(Me.txtPatientID)) <= 0 Then
        strSql = strSql & " and [9] = [9] "
    Else
        strSql = strSql & " and Decode(B.������Դ,1,G.�����,2,G.סԺ��,NULL)= " & Me.txtPatientID
    End If
    If blnMoved Then
        strSQLBak = strSql
        strSQLBak = Replace(strSQLBak, "Ӱ�����¼", "HӰ�����¼")
        strSQLBak = Replace(strSQLBak, "����ҽ����¼", "H����ҽ����¼")
        strSql = strSql & " Union ALL " & strSQLBak & " order by ���UID,����UID,ͼ��UID"
    Else
        strSql = strSql & " order by ���UID,����UID,ͼ��UId"
    End If
    Set rsMain = OpenSQLRecord(strSql, Me.Caption, CDate(Format(Me.dtpBegin.Value, "yyyy-mm-dd")), _
    CDate(Format(Me.dtpEnd.Value, "yyyy-mm-dd 23:59:59")), strDeptNo, _
    IIf(Trim(txtChkNoBegin) = "", 1, txtChkNoBegin), IIf(Trim(txtChkNoEnd) = "", 1, txtChkNoEnd), _
    IIf(Trim(txt����) = "", 1, txt����), intMZ, intZY, IIf(Trim(txtPatientID) = "", 1, txtPatientID))
    Me.tvwImageDate.Nodes.Clear
        
    Do Until rsMain.EOF
        With Me.tvwImageDate.Nodes
            '���˼�
            If strPatientUID <> rsMain("���UID") Then
                .Add , , "A" & rsMain("���UID"), rsMain("����") & "[" & rsMain("��������") & "]"
            End If
            
            '������м�
            If strSerialUID <> rsMain("����UID") Then
                .Add "A" & rsMain("���UID"), tvwChild, rsMain("���UID") & "_" & rsMain("����UID"), "[" & rsMain("��������") & "]" & rsMain("����UID")
            End If
            
            'ͼ��
            .Add rsMain("���UID") & "_" & rsMain("����UID"), tvwChild, rsMain("����UID") & "_" & rsMain("ͼ��UID"), rsMain("ͼ��UID")
            
            DoEvents
            
            strPatientUID = rsMain("���UID")
            strSerialUID = rsMain("����UID")
            rsMain.MoveNext
        End With
    Loop
    
    
'    Do Until rsMain.EOF
'        With Me.tvwImageDate.Nodes
'            .Add , , "A" & rsMain("ҽ��ID"), rsMain("����") & "[" & rsMain("��������") & "]"
'            strSQL = "select ����UID,�������� from Ӱ�������� where ���UID = [1]"
'            If blnMoved Then
'                strSQLBak = strSQL
'                strSQLBak = Replace(strSQLBak, "Ӱ��������", "HӰ��������")
'                strSQL = strSQL & " Union ALL " & strSQLBak
'            End If
'            Set rsNode = OpenSQLRecord(strSQL, Me.Caption, Nvl(rsMain("���UID"), 0))
'            Do Until rsNode.EOF
'                .Add "A" & rsMain("ҽ��ID"), tvwChild, "A" & rsNode("����UID"), "[" & rsNode("��������") & "]" & rsNode("����UID")
'                rsNode.MoveNext
'            Loop
'            rsNode.Close
'            Set rsNode = Nothing
'        End With
'        rsMain.MoveNext
'        DoEvents
'    Loop
    rsMain.Close
    Set rsMain = Nothing
    Exit Sub
RefreshError:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tvwImageDate_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim blSelOrCls As Boolean
    Dim blLoopEOF As Boolean
    Dim objNode As Node

    '���˼�
    If Node.Parent Is Nothing Then
        SelectAllChild Node, Node.Checked
        Set objNode = Node.Child.FirstSibling
        objNode.Checked = Node.Checked
        
        For i = 1 To Node.Children
            objNode.Checked = Node.Checked
            SelectAllChild objNode, Node.Checked
            Set objNode = objNode.Next
        Next
    End If
    
    '���м�
    If Not Node.Parent Is Nothing And Not Node.Child Is Nothing Then
        SelectAllChild Node, Node.Checked
    End If
    
    If Node.Checked = True Then
        'ѡ��
        If Not Node.Parent Is Nothing Then
            Node.Parent.Checked = True
            If Not Node.Parent.Parent Is Nothing Then
                Node.Parent.Parent.Checked = True
            End If
        End If
    Else
        'ȡ��
        If Not Node.Parent Is Nothing Then
            Set objNode = Node.Parent.Child.FirstSibling
            '������һ��
            For i = 1 To Node.Parent.Children
                If objNode.Checked = True Then
                    blLoopEOF = True
                    Exit For
                End If
                Set objNode = objNode.Next
            Next
            '�������ϼ�
            If blLoopEOF = False Then
                Node.Parent.Checked = False
                If Not Node.Parent.Parent Is Nothing Then
                    Set objNode = Node.Parent.Parent.Child.FirstSibling
                    For i = 1 To Node.Parent.Parent.Children
                        If objNode.Checked = True Then
                            blLoopEOF = True
                            Exit For
                        End If
                        Set objNode = objNode.Next
                    Next
                    If blLoopEOF = False Then
                        Node.Parent.Parent.Checked = False
                    End If
                End If
            End If
        End If
    End If
End Sub
Private Sub AllSelectOrAllClear(TrueOrFalse As Boolean)
    With Me.tvwImageDate
        For i = 1 To .Nodes.Count
            .Nodes(i).Checked = TrueOrFalse
        Next
    End With
End Sub
'��ʾ����Ŀ¼
Private Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    On Error GoTo OpenFileError
    With udtBI
        '�����������
        .hWndOwner = lWindowHwnd
        '����ѡ�е�Ŀ¼
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "��ѡ����ʼ�������ļ��У�"
        Else
            .lpszTitle = sTitle
        End If
    End With
    '�����������
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '��ȡ·��
        SHGetPathFromIDList lpIDList, sPath
        '�ͷ��ڴ�
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
    Exit Function
OpenFileError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function WriteTest(ShowErrMsg As Boolean, strPath As String) As Boolean
    Dim strTmpPath As String
    On Error GoTo CopyError
    strTmpPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "temp.txt"
    Open strTmpPath For Output As #1
    Close #1
    FileCopy strTmpPath, IIf(Len(strPath) > 3, strPath & "\", strPath) & "temp.txt"
    Kill IIf(Len(strPath) > 3, strPath & "\", strPath) & "temp.txt"
    Kill strTmpPath
    WriteTest = True
    Exit Function
CopyError:
    If ShowErrMsg = False Then Exit Function
    If Err.Number = 75 Then
        MsgBox "д�����ʧ��!��鿴[" & strPath & "]�Ƿ���д��Ȩ��!", vbInformation, App.EXEName
    Else
        MsgBox "������������", vbQuestion, App.EXEName
    End If
End Function

Private Function SendFilesToDir(lngSeqUID As String, DicomDirPath As DicomDataSet, DestinationBoot As String, DestinationDir As String) As Boolean
    '����:��FTP�����ļ�
    '����:����UID
    '�������سɹ�����ļ�·��
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strCachePath As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strTmpFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim DicomImg As New DicomImages
    Dim DicomImgFormat As String
    Dim ImageUID As String
    Dim SerialUID As String
    
    On Error GoTo WriteFileErr
    SendFilesToDir = True
    strSql = "Select A.ͼ���,D.�û��� As User1,D.���� As Pwd1,a.ͼ��UID, " & _
        "D.IP��ַ As Host1," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL1,d.�豸�� as �豸��1, " & _
        "E.�û��� As User2,E.���� As Pwd2," & _
        "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL2 , e.�豸�� as �豸��2 " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And A.����UID= [1] And A.ͼ��UID = [2] Order By A.ͼ���"
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
    SerialUID = Mid(lngSeqUID, 1, InStr(lngSeqUID, "_") - 1)
    ImageUID = Mid(lngSeqUID, InStr(lngSeqUID, "_") + 1)
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, SerialUID, ImageUID)
    strCachePath = App.Path & "\TmpImage\"
    ClearCacheFolder strCachePath
    If rsTmp.EOF <> True Then
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
    End If
    If Me.ChkImageFormat(1).Value = True Then
        DicomImgFormat = ".BMP"
    ElseIf Me.ChkImageFormat(2).Value = True Then
        DicomImgFormat = ".JPG"
    End If
    
    Do While Not rsTmp.EOF
        
'        If rsTmp("URL1") Is Nothing Then
'            strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
'        Else
'            strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
'        End If
            
'        Inet.strIPAddress = Nvl(rsTmp("Host1")): Inet.strUser = Nvl(rsTmp("User1")): Inet.strPsw = Nvl(rsTmp("Pwd1"))
        strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
    
        If strDeviceNO1 <> rsTmp("�豸��1") Then
            strDeviceNO1 = rsTmp("�豸��1")
            Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
        End If
        
        If strDeviceNO2 <> rsTmp("�豸��2") Then
            strDeviceNO2 = rsTmp("�豸��2")
            Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
        End If
        
        If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFile.GetFileName(rsTmp("URL1"))) <> 0 Then
'            Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
            Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFile.GetFileName(rsTmp("URL2")))
        End If

        On Error Resume Next
        MkLocalDir DestinationBoot & DestinationDir
        DicomImg.ReadFile Replace(strTmpFile, "/", "\")
        '�����ʽ
        If DicomImgFormat <> "" Then
            DicomImg(1).FileExport strTmpFile & DicomImgFormat, Mid(DicomImgFormat, 2)
        End If
        
        DicomDirPath.Name = "ZLPACS"
        DicomDirPath.AddToDirectory DicomImg(1), "IMAGE\" & DestinationDir & "\" & rsTmp("ͼ��UId") & _
                                    DicomImgFormat, "1.2.840.10008.1.2.1", 0
        DicomImg.Clear
        Err.Clear
        
        If Dir(DestinationBoot & DestinationDir & "\" & rsTmp("ͼ��UId")) = vbNullString Then
            FileCopy strTmpFile & DicomImgFormat, DestinationBoot & DestinationDir & "\" & rsTmp("ͼ��UId") & DicomImgFormat
        End If
         
        If Err.Number <> 0 Then
            Open App.Path & "\WriteErrLog.txt" For Append As #1
                Print #1, "����[" & strCachePath & Nvl(rsTmp("URL1")) & "]��[" & DestinationBoot & DestinationDir & "]" & vbCrLf & _
                "����" & Err.Description & "�����:" & Err.Number
            Close #1
            SendFilesToDir = False
        End If
        
        DoEvents
        rsTmp.MoveNext
    Loop
    Exit Function
WriteFileErr:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tvwImageDate_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strImageUID As String
    Dim strSerialUID As String
    
    If Node.Child Is Nothing And Me.ChkShowImage.Value = 1 Then
        strImageUID = Mid(Node.Key, InStr(Node.Key, "_") + 1)
        strSerialUID = Mid(Node.Key, 1, InStr(Node.Key, "_") - 1)
        ShowImage strImageUID, strSerialUID
    End If
End Sub

Private Sub txtChkNoBegin_GotFocus()
    With Me.txtChkNoBegin
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtChkNoBegin_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Function IsStrValib() As Boolean
    '����ִ��ĺϷ���
    If zlCommFun.StrIsValid(Me.txtChkNoBegin) = False Then
        MsgBox "��ʼ�����а����˷Ƿ��ַ�����", vbInformation, App.EXEName
        With Me.txtChkNoBegin
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    If zlCommFun.StrIsValid(Me.txtChkNoEnd) = False Then
        MsgBox "���������а����˷Ƿ��ַ�����", vbInformation, App.EXEName
        With Me.txtChkNoEnd
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    If zlCommFun.StrIsValid(Me.txt����, 12) = False Then
        MsgBox "�����а����˷Ƿ��ַ�����", vbInformation, App.EXEName
        With Me.txt����
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Function

Private Sub txtChkNoEnd_GotFocus()
    With Me.txtChkNoEnd
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TxtPath_Change()

End Sub

Private Sub txt����_GotFocus()
    With Me.txt����
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Function ShowImage(lngImageUID As String, lngSerialUID As String) As Boolean
    '����:��FTP�����ļ�
    '����:����UID
    '�������سɹ�����ļ�·��
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim strCachePath As String
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim strDeviceNO1 As String
    Dim strDeviceNO2 As String
    Dim strTmpFile As String
    Dim objFile As New Scripting.FileSystemObject
    Dim DicomImg As New DicomImages
    
    On Error GoTo WriteFileErr
    ShowImage = True
    strSql = "Select A.ͼ���,D.�û��� As User1,D.���� As Pwd1,a.ͼ��UID, " & _
        "D.IP��ַ As Host1," & _
        "'/'||D.FtpĿ¼||'/' As Root1,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL1,d.�豸�� as �豸��1, " & _
        "E.�û��� As User2,E.���� As Pwd2," & _
        "E.IP��ַ As Host2," & _
        "'/'||E.FtpĿ¼||'/' As Root2,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')" & _
        "||C.���UID||'/'||A.ͼ��UID As URL2 , e.�豸�� as �豸��2 " & _
        "From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E " & _
        "Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) " & _
        "And A.ͼ��UID= [1]  and a.����UID = [2]  Order By A.ͼ���"
    If mblnMoved Then
        strSql = Replace(strSql, "Ӱ����ͼ��", "HӰ����ͼ��")
        strSql = Replace(strSql, "Ӱ��������", "HӰ��������")
        strSql = Replace(strSql, "Ӱ�����¼", "HӰ�����¼")
    End If
            
    Set rsTmp = OpenSQLRecord(strSql, Me.Caption, lngImageUID, lngSerialUID)
    strCachePath = App.Path & "\TmpImage\"
    ClearCacheFolder strCachePath
    If rsTmp.EOF <> True Then
        MkLocalDir strCachePath & objFile.GetParentFolderName(Nvl(rsTmp("URL1")))
    End If
    Do While Not rsTmp.EOF
        If strDeviceNO1 <> rsTmp("�豸��1") Then
            strDeviceNO1 = rsTmp("�豸��1")
            Inet1.FuncFtpConnect Nvl(rsTmp("Host1")), Nvl(rsTmp("User1")), Nvl(rsTmp("Pwd1"))
        End If
        
        If strDeviceNO2 <> rsTmp("�豸��2") Then
            strDeviceNO2 = rsTmp("�豸��2")
            Inet2.FuncFtpConnect Nvl(rsTmp("Host2")), Nvl(rsTmp("User2")), Nvl(rsTmp("Pwd2"))
        End If
        
        strTmpFile = strCachePath & Nvl(rsTmp("URL1"))
        If Dir(strTmpFile) = "" Then
'            Inet.strIPAddress = Nvl(rsTmp("Host1")): Inet.strUser = Nvl(rsTmp("User1")): Inet.strPsw = Nvl(rsTmp("Pwd1"))
            If Inet1.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root1")) & rsTmp("URL1")), strTmpFile, objFile.GetFileName(rsTmp("URL1"))) <> 0 Then
                strTmpFile = strCachePath & Nvl(rsTmp("URL2"))
'                Inet.strIPAddress = Nvl(rsTmp("Host2")): Inet.strUser = Nvl(rsTmp("User2")): Inet.strPsw = Nvl(rsTmp("Pwd2"))
                Call Inet2.FuncDownloadFile(objFile.GetParentFolderName(Nvl(rsTmp("Root2")) & rsTmp("URL2")), strTmpFile, objFile.GetFileName(rsTmp("URL2")))
            End If
        End If
        On Error Resume Next
        Viewer.Images.ReadFile strTmpFile
        Kill strTmpFile
        DoEvents
        rsTmp.MoveNext
    Loop
    Exit Function
WriteFileErr:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SelectAllChild(xNode As Node, blCheck As Boolean)
    '����           ȡ����һ���е������ּ�
    '����           xNode    Node����
    '                blCheck �Ƿ�ѡ��
    Dim nNode As Node
    
    If xNode.Children = 0 Then Exit Sub
    Set nNode = xNode.Child
    For i = 1 To xNode.Children
        nNode.Checked = blCheck
        Set nNode = nNode.Next
    Next
End Sub
