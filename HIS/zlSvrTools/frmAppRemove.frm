VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppRemove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ӧ��ϵͳ��ж"
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmAppRemove.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6600
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   1605
      TabIndex        =   17
      Top             =   3510
      Width           =   1100
   End
   Begin VB.Frame fraSys 
      Height          =   1365
      Left            =   2085
      TabIndex        =   9
      Top             =   1125
      Width           =   3945
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   210
         TabIndex        =   16
         Top             =   990
         Width           =   540
      End
      Begin VB.Label lblOwner 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   795
         TabIndex        =   15
         Top             =   930
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ϵͳ��"
         Height          =   180
         Left            =   210
         TabIndex        =   13
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblSysName 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   795
         TabIndex        =   12
         Top             =   225
         Width           =   2895
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   795
         TabIndex        =   11
         Top             =   570
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�汾��"
         Height          =   180
         Left            =   210
         TabIndex        =   10
         Top             =   630
         Width           =   540
      End
   End
   Begin VB.PictureBox PicSetup 
      Align           =   3  'Align Left
      Height          =   4005
      Left            =   0
      ScaleHeight     =   3945
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   0
      Width           =   1335
      Begin VB.Image imgRemove 
         Height          =   2550
         Left            =   120
         Picture         =   "frmAppRemove.frx":058A
         Stretch         =   -1  'True
         Top             =   135
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdGetIni 
      Caption         =   "ѡ��(&S)��"
      Height          =   350
      Left            =   4935
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4095
      TabIndex        =   3
      Top             =   3510
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   60
      TabIndex        =   5
      Top             =   3405
      Width           =   7140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5185
      TabIndex        =   2
      Top             =   3510
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   4005
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAppRemove.frx":5B70
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8070
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "16:13"
            Key             =   "STANUM"
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
   Begin VB.Label lblNote 
      Caption         =   "    ��������������������������ȷָ��Ӧ�ð�װ�����ļ���ִ�в�ж�����Զ������ϵͳ���������ݡ������������ߺͶ����Ĵ洢�ռ䡣"
      Height          =   525
      Index           =   1
      Left            =   1605
      TabIndex        =   8
      Top             =   540
      Width           =   4680
   End
   Begin VB.Label lbliniFile 
      AutoSize        =   -1  'True
      Caption         =   "Ӧ�ð�װ�����ļ�"
      Height          =   180
      Left            =   2085
      TabIndex        =   6
      Top             =   2760
      Width           =   1440
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2085
      TabIndex        =   1
      Top             =   3000
      Width           =   3945
   End
   Begin VB.Label lblNote 
      Caption         =   "    ��ж�����Ƕ�ָ��ϵͳ�ĳ�������������ڲ�жǰ������������ɿ������ݱ��ݣ�"
      Height          =   375
      Index           =   0
      Left            =   1605
      TabIndex        =   4
      Top             =   105
      Width           =   4680
   End
End
Attribute VB_Name = "frmAppRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intDefSysCode As Integer                'ϵͳ���
Dim strDefSysName As String                 'ϵͳ����
Dim strDefVersion As String                 '�汾��
Dim strDefSpace   As String                 '��ռ�

Dim mbln���� As Boolean    '���ΰ�װ�Ƿ����������װ�װ

Dim objFile As New FileSystemObject
Dim objText As TextStream

Dim rsTemp As New ADODB.Recordset
Dim strSQL As String, strTemp As String
Dim intCount As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetIni_Click()
    With frmMDIMain.DlgMain
        .FileName = lblFileName.Caption
        .DialogTitle = "ѡ��Ӧ�ð�װ�����ļ�"
        .Filter = "(Ӧ�ð�װ�����ļ�)|zlSetup.ini"
        .ShowOpen
        If .FileName = "" Then
            Exit Sub
        Else
            lblFileName.Caption = .FileName
        End If
    End With
    
    If CheckIniFile(lblFileName.Caption, True) = False Then
        cmdOk.Enabled = False
        lblFileName.Caption = ""
        cmdGetIni.SetFocus
    Else
        cmdOk.Enabled = True
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub cmdOK_Click()
    Me.MousePointer = 11
    If DeleteSystem = False Then
        Me.MousePointer = 0
        Exit Sub
    End If
    Me.MousePointer = 0
    Unload Me
End Sub

Private Function DeleteSystem() As Boolean
    Dim msgSystem As VbMsgBoxResult
    Dim strMsg As String
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 1 From zltools.zlbakspaces where ��ǰ<>1 and ϵͳ=" & Val(lblSysName.Tag)
    rsTemp.Open gstrSQL, gcnOracle
    If Not rsTemp.EOF Then
        strMsg = "����жϵͳ������ʷ���ݿռ�,�Ƿ������ж��" & vbCrLf & _
             "ѡ���ǡ�����������ʷҵ������(���Զ�����)����Ҫʱ��ͨ������ֲ���ָ���" & vbCrLf & _
             "ѡ�񡾷񡿣����˳���ж����������ڹ�����->���ݹ���->����ת����ɾ" & vbCrLf & Space(12) & "����ʷ���ݿռ���ٽ��в�ж��"
        msgSystem = MsgBox(strMsg, vbQuestion Or vbYesNo Or vbDefaultButton3, gstrSysName)
        If msgSystem = vbNo Then Exit Function
    End If
    
   strMsg = "��ж���������㱣���û�ҵ������(��ɾ����ϵͳ���й�������)��" & vbCrLf _
            & "�Ƿ����û�ҵ�����ݣ�" & vbCrLf & vbCrLf _
            & "ѡ���ǡ����������û�ҵ�����ݣ���Ҫʱ��ͨ������ֲ���ָ���" & vbCrLf _
            & "                        ����ע�Ᵽ���Ѿ��������޸Ĺ����Զ��屨��" & vbCrLf & vbCrLf _
            & "ѡ�񡾷񡿣�������ɾ����ϵͳ��������(�������б�ͱ�ռ�)���޷��ָ���" & vbCrLf _
            & "                        ���ǿ�ҽ��������˲���ǰ�����ݿ����һ�α��ݡ�"
    
    msgSystem = MsgBox(strMsg, vbQuestion Or vbYesNoCancel Or vbDefaultButton3, gstrSysName)
        
    If msgSystem = vbCancel Then Exit Function
    
    If msgSystem = vbNo Then
        '��ȫɾ��
        If MsgBox("ϵͳ��ж��������ɾ����ϵͳ���ݣ��޷��ָ���" & vbCrLf & vbCrLf & "������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
        
    cmdGetIni.Enabled = False
    cmdCancel.Enabled = False
    cmdOk.Enabled = False
    
    '-------------��ж�㷨-----------------
    '   �����װ��¼
    '   If ��ϵͳ�����߲�������ϵͳ�������� Then
    '       ����ɾ����ϵͳ������ (���ݶ���ȫ��ɾ��)
    '   End if;
    '   ɾ���ع���
    '   ɾ����ռ�
    '--------------------------------------
    Dim strSpaces As String, strErrInfo As String
    Dim aryTbs() As String, aryChr() As String
    Dim blnEnjoy As Boolean
    
    On Error GoTo 0
    
    DoEvents
    If msgSystem = vbNo Then
        'Ҫɾ���û��ͱ�ռ�
        With rsTemp
            If .State = adStateOpen Then .Close
            .Open "select 1 from gv$session where USERNAME='" & UCase(lblOwner.Caption) & "'", gcnOracle
            If .EOF = False Then
                MsgBox "ϵͳ�����������ӵ����ݿ��ϣ��޷����ж�ز�����", vbExclamation, gstrSysName
                cmdGetIni.Enabled = True
                cmdCancel.Enabled = True
                cmdOk.Enabled = True
                Exit Function
            End If
        End With
        
        If mbln���� = False Then
            '������ռ估�����ļ�
            aryTbs = Split(strDefSpace, "||")
            strSpaces = ""
            For intCount = 0 To UBound(aryTbs)
                aryChr = Split(aryTbs(intCount), "|")
                strSpaces = strSpaces & ",'" & UCase(Trim(aryChr(1))) & "'"
            Next
        End If
        
        '�ж��Ƿ�������ϵͳ
        With rsTemp
            strSQL = "select 1 from zlSystems where upper(������)='" & UCase(lblOwner.Caption) & "'"
            If .State = adStateOpen Then .Close
            .Open strSQL, gcnOracle, adOpenKeyset
            blnEnjoy = (.RecordCount > 1)
        End With
        
        
        'ɾ����ϵͳ������
        err = 0
        On Error Resume Next
        If blnEnjoy = False Then
            stbThis.Panels(2).Text = "ɾ��ϵͳ�����ߡ�"
            intCount = 0
            Do
                gcnOracle.Execute "drop user " & lblOwner.Caption & " cascade"
                With rsTemp
                    If .State = adStateOpen Then .Close
                    .Open "select * from all_users where username='" & UCase(lblOwner.Caption) & "'", gcnOracle
                    If .EOF Then Exit Do
                End With
                intCount = intCount + 1
                DoEvents
                If intCount > 1000 Then
                    strErrInfo = strErrInfo & vbCr & "�û�:" & lblOwner.Caption
                    Exit Do
                End If
            Loop
        End If
        
        If mbln���� = False Then
            '8i֮���޻ع��θ������Undo��ռ�
            
            'ɾ����ϵͳ���ݿռ�
            
            stbThis.Panels(2).Text = "ɾ�����ݱ�ռ䡭"
            Refresh
            aryChr = Split(Mid(strSpaces, 2), ",")
            
            For intCount = 0 To UBound(aryChr)
                DoEvents
                strTemp = Mid(Mid(aryChr(intCount), 2), 1, Len(Mid(aryChr(intCount), 2)) - 1)
                
                '����װʱ��һ��ϵͳ����Щ����Ŀǰû�м�¼���޷��жϵ�ǰɾ���ı�ռ��Ƿ�������ϵͳ�Ķ���
                
                If CheckSpaceIsUse("��ռ�", strTemp, lblOwner.Caption) = False Then
                    'û�������û�ʹ�ã�����ɾ��
                    gcnOracle.Execute "alter tablespace " & strTemp & " offline"
                    gcnOracle.Execute "drop tablespace " & strTemp & " including contents and datafiles cascade constraints"
                End If
            Next
            
            
            '�����ļ�һ�����ڷ������ϵģ����ҿ�����ASM�ϵģ�ɾ����ռ�ʱ�Ѽ���ɾ��
        End If
    End If
    
    'ɾ����װ��¼
    err = 0: On Error GoTo 0
    stbThis.Panels(2).Text = "ɾ����װ��¼��"
    
    If msgSystem = vbNo Then
        '���˺�:��Ҫɾ�����߲��ֵ���ʷ���ݿռ�
        Call frmHistorySpaceSet.ShowInstall(Me, gcnOracle, gstrUserName, gstrPassword, Val(lblSysName.Tag), 1, 0)
    End If
    DoEvents
    strSQL = "delete from zlSystems where ���=" & lblSysName.Tag
    gcnOracle.Execute strSQL
    
    '������Ч�˵�
    With rsTemp
        Do
            If .State = adStateOpen Then .Close
            strSQL = "select 1 from zlMenus A where ģ�� is null and not exists(select 1 from zlMenus B where B.�ϼ�ID=A.ID)"
            .Open strSQL, gcnOracle
            If .EOF Then Exit Do
            strSQL = "delete from zlMenus A where ģ�� is null and not exists(select 1 from zlMenus B where B.�ϼ�ID=A.ID)"
            gcnOracle.Execute strSQL
        Loop
    End With
    
    If strErrInfo <> "" Then
        MsgBox strDefSysName & "��ж��ɣ����ֹ�ɾ���������ݣ�" & strErrInfo, vbExclamation, gstrSysName
    Else
        MsgBox strDefSysName & "��ж��ɡ�", vbExclamation, gstrSysName
    End If
    DeleteSystem = True
End Function

Private Sub Form_Load()
    Call ApplyOEM(stbThis)
    With imgRemove
        .Top = PicSetup.ScaleTop
        .Left = PicSetup.ScaleLeft
        .Height = PicSetup.ScaleHeight
        .Width = PicSetup.ScaleWidth
    End With
    With frmAppStart.lvwSys.SelectedItem
        lblSysName.Tag = Mid(.Key, 2)
        lblSysName.Caption = .Text
        lblVersion.Caption = .SubItems(1)
        lblOwner.Caption = .SubItems(3)
    End With
    
    Call Judge����
    
    If mbln���� = False Then
        '��ȫɾ��
        With rsTemp
            strSQL = "select �ļ��� from zlSysFiles where ϵͳ=" & lblSysName.Tag & " and ����=1"
            If .State = adStateOpen Then .Close
            .Open strSQL, gcnOracle, adOpenKeyset
            If Not .EOF And Not .BOF Then
                lblFileName.Caption = .Fields(0).value
            End If
        End With
        If Not gobjFile.FileExists(lblFileName.Caption) Then
            If gobjFile.FileExists(App.Path & "\zlSetup.ini") Then
                lblFileName.Caption = App.Path & "\zlSetup.ini"
            End If
        End If
        
        If Trim(lblFileName.Caption) <> "" Then
            If CheckIniFile(lblFileName.Caption) = False Then
                lblFileName.Caption = ""
            Else
                cmdOk.Enabled = True
            End If
        End If
    Else
        '����ɾ��
        cmdOk.Enabled = True
        cmdGetIni.Enabled = False
        lbliniFile.Enabled = False
        lblFileName.Enabled = False
    End If
End Sub

Private Sub Judge����()
    '�ж��Ƿ�Ӧ�ðѱ��ΰ�װ��Ϊ���װ�װ
    Dim lngϵͳ�� As Long, lngTemp As Long
    Dim lstTemp As ListItem

    
    mbln���� = False
    lngϵͳ�� = lblSysName.Tag \ 100
    For Each lstTemp In frmAppStart.lvwSys.ListItems
        lngTemp = Mid(lstTemp.Key, 2)
        If lngTemp \ 100 = lngϵͳ�� Then
            'ϵͳ��ͬ
            
            If lngTemp <> lblSysName.Tag Then
                '����һ�����״���
                mbln���� = True
                Exit For
            End If
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If cmdCancel.Enabled = False Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Function CheckIniFile(strFile As String, Optional blnMsg As Boolean) As Boolean
    err = 0
    On Error Resume Next
        
    '�����ļ���ȷ�Լ��
    Set objText = objFile.OpenTextFile(strFile)
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        intDefSysCode = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[ϵͳ��]" Then
        strDefSysName = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[�汾��]" Then
        strDefVersion = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    strTemp = Trim(objText.ReadLine)
    If Left(strTemp, 5) = "[��ռ�]" Then
        strDefSpace = Mid(strTemp, 6)
    Else
        err.Raise 10
    End If
    objText.Close
    
    If err <> 0 Then
        CheckIniFile = False
        If blnMsg Then MsgBox "��װ�����ļ�����ȷ", vbExclamation, gstrSysName
        Exit Function
    End If
    
    '�����ļ������Լ��
    If intDefSysCode <> Int(lblSysName.Tag / 100) Then
        err.Raise 10
        If blnMsg Then MsgBox "ѡ���ļ����� " & lblSysName.Caption & " �İ�װ�����ļ�", vbExclamation, gstrSysName
    ElseIf Trim(strDefVersion) <> lblVersion.Caption Then
        err.Raise 10
        If blnMsg Then MsgBox "ѡ���ļ��� " & lblSysName.Caption & " �汾���� ", vbExclamation, gstrSysName
    End If
    If err = 0 Then
        CheckIniFile = True
    Else
        CheckIniFile = False
    End If
End Function
