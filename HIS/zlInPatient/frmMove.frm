VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMove 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˻���"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   Icon            =   "frmMove.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   5805
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3750
      Width           =   5805
      Begin VB.CheckBox chk���� 
         Caption         =   "����(&M)"
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   1455
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4440
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3240
         TabIndex        =   7
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.Frame fraBed 
      Height          =   1950
      Left            =   120
      TabIndex        =   12
      Top             =   75
      Width           =   5565
      Begin VB.TextBox txtUnit 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1065
         Width           =   1800
      End
      Begin VB.ComboBox cboNew 
         Height          =   300
         Left            =   3510
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1065
         Width           =   1845
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   660
         Width           =   1800
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1800
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   300
         Left            =   3510
         TabIndex        =   5
         Top             =   1500
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   19
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   105
         TabIndex        =   20
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   465
         TabIndex        =   18
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   465
         TabIndex        =   17
         Top             =   330
         Width           =   360
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2910
         TabIndex        =   16
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   2730
         TabIndex        =   15
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label lblNew 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�²���"
         Height          =   180
         Left            =   2925
         TabIndex        =   14
         Top             =   1125
         Width           =   525
      End
      Begin VB.Label lblPre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ����"
         Height          =   180
         Left            =   2910
         TabIndex        =   13
         Top             =   720
         Width           =   540
      End
   End
   Begin VB.Frame fraLvw 
      Caption         =   "ѡ�񲡴�"
      Height          =   1545
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   5565
      Begin MSComctlLib.ListView lvw 
         Height          =   1140
         Left            =   165
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   255
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   2011
         View            =   2
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public mbytInFun As Byte '0-����,1-����������,2-������Ժʱ���´�
Public mstrĿ�괲�� As String '�룺��mbytInFun=0ʱ��ʾ�϶���ָ���Ĵ�λ�ϣ�=2ʱ��ʾ��λ��Ӧ�ĵȼ�ID������������´���(���ܶ���)
Public mstr���� As String   '��mbytInFun=2ʱ����ֵ,��ʾ��Ժǰ�Ĵ�λ

Public mstrPrivs As String
Public mlngUnit As Long
Public mlng����ID As Long, mlng��ҳID As Long
Private mfrmParent As Object

Private mrsPatiInfo As ADODB.Recordset
Private mrsBeds As ADODB.Recordset '��ѡ��λ��

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub chk����_Click()
        
    If chk����.Value = 1 Then
        lvw.Visible = True
        lvw.TabStop = True
        lblNew.Caption = "����λ"
        If Visible Then Call LoadMainBed
        
        fraLvw.Top = fraBed.Top + fraBed.Height + 200
        Me.Height = fraLvw.Top + fraLvw.Height + Picture1.Height + 400
        If lvw.Visible Then lvw.SetFocus
        
    Else
        lvw.Visible = False
        lvw.TabStop = False
        lblNew.Caption = "�²���"
        If Visible Then Call InitBed(mlngUnit)
        
        Me.Height = fraBed.Top + fraBed.Height + Picture1.Height + 400
        If cboNew.Visible Then cboNew.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_Activate()
    If lvw.Visible And lvw.Enabled Then
        lvw.SetFocus
    ElseIf cboNew.Visible And cboNew.Enabled Then
        cboNew.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    gblnOK = False
    
    '�������
    Me.Height = fraBed.Top + fraBed.Height + Picture1.Height + 400
    
    Call InitData
    
    Select Case mbytInFun
        Case 0  '0-����
            If cboNew.ListCount = 0 Then
                MsgBox "�������ڿ��ҵĲ�����û�к��ʵĴ�λ�ɹ�������", vbInformation, gstrSysName
                Unload Me: Exit Sub
            End If
        Case 1 '1-����������
            If lvw.ListItems.Count = 0 Then
                MsgBox "���˵�ǰ����û�������մ���", vbInformation, gstrSysName
                Unload Me: Exit Sub
            End If
        Case 2 '2-������Ժʱ���´�
            If UBound(Split(txtPre.Text, ";")) > 0 Then '��Ժ֮ǰ���Ŵ�λ
                If lvw.ListItems.Count = 0 Then
                    MsgBox "���˵�ǰ����û�������մ���", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            Else
                If cboNew.ListCount = 0 Then
                    MsgBox "�������ڿ��ҵĲ�����û�к��ʵĴ�λ�ɹ�������", vbInformation, gstrSysName
                    Unload Me: Exit Sub
                End If
            End If
    End Select
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Sub InitData()
    Dim i As Integer, rsTmp As ADODB.Recordset, str���� As String, str����� As String
    
    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With mrsPatiInfo
        txt����.Text = !����
        txt����.Tag = "" & !�Ա�
        txtסԺ��.Text = "" & !סԺ��
        txt����.Text = !��ǰ����
    End With
    
    str����� = ""
    If mbytInFun = 2 Then
        txtPre.Text = mstr����
    Else
        Set rsTmp = GetPatiBeds(mlng����ID)
        If rsTmp.RecordCount = 0 Then
            str���� = "��ͥ����"
        Else
            Do While Not rsTmp.EOF
                str���� = str���� & "," & rsTmp!����
                If Nvl(rsTmp!����) = Nvl(mrsPatiInfo!��Ҫ����) And Nvl(rsTmp!����ID) = Nvl(mrsPatiInfo!��ס����id) Then
                    str����� = Nvl(rsTmp!�����)
                End If
                rsTmp.MoveNext
            Loop
            str���� = Mid(str����, 2)
        End If
        txtPre.Text = str����
        txtPre.Tag = str�����
    End If
                                
    If UBound(Split(txtPre.Text, ",")) > 0 And mstrĿ�괲�� <> "��ͥ����" Or mbytInFun = 1 Then
        chk����.Value = 1   '����click�¼�
    Else
        Call chk����_Click  'ȱʡֵΪ0
    End If
    If mbytInFun = 1 Or mbytInFun = 2 Then chk����.Visible = False
    If mbytInFun <> 0 Then txtUnit.Visible = False: lblUnit.Visible = False
    If mbytInFun = 2 Then lblDate.Visible = False: txtDate.Visible = False
    
    Select Case mbytInFun
        Case 0 '����
            Me.Caption = "���˻���"
            'Ŀǰ��������۲���
            Set rsTmp = GetDeptOrUnit(1, mrsPatiInfo!��Ժ����id, "1,2,3")
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    If rsTmp!ID = mlngUnit Then txtUnit.Text = rsTmp!���� '��ǰ��������
                    rsTmp.MoveNext
                Next
            End If
            Call InitBed(mlngUnit)
            
        Case 1 '1-����������
            Me.Caption = "���˰����Ӽ���λ"
            lblDate.Caption = "�䶯ʱ��"
            
            lblDate.Top = lblNew.Top: txtDate.Top = cboNew.Top
            lblNew.Left = lblUnit.Left: cboNew.Left = txtUnit.Left
            
            fraBed.Height = fraBed.Height - cboNew.Height - 100
            fraLvw.Top = fraLvw.Top - cboNew.Height - 100
            Me.Height = Me.Height - cboNew.Height - 100
            
            Call InitBed(0)
            If lvw.ListItems.Count = 0 Then Exit Sub
            
        Case 2 '2-������Ժʱ���´�
            Me.Caption = "������Ժ�����´�λ"
            fraBed.Height = fraBed.Height - cboNew.Height - 100
            fraLvw.Top = fraLvw.Top - cboNew.Height - 100
            Me.Height = Me.Height - cboNew.Height - 100
            
            Call InitBed(mlngUnit)
     End Select
    
End Sub

Private Sub InitBed(ByVal lng����ID As Long)
'���ܣ���ʼ����λ,��ʱȡ�ò��������Ҷ�Ӧ�����пմ�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Integer
    Dim bytLen As Byte
    Dim strTmp As String
    
    On Error GoTo errH
        
    If InStr(mrsPatiInfo!�Ա�, "��") > 0 Then
        strTmp = "�д�,���޴�"
    ElseIf InStr(mrsPatiInfo!�Ա�, "Ů") > 0 Then
        strTmp = "Ů��,���޴�"
    Else
        strTmp = "���޴�"
    End If
        
    lvw.ListItems.Clear
    cboNew.Clear
    
    Select Case mbytInFun
        Case 0, 2 '0-���� '2-������Ժʱ���´�
            If mbytInFun = 0 Then
                If InStr(1, mstrPrivs, "��ͥ����") > 0 And txtPre.Text <> "��ͥ����" And chk����.Value = 0 Then
                    cboNew.AddItem "��ͥ����", 0
                    If mstrĿ�괲�� = "��ͥ����" Then cboNew.ListIndex = 0
                End If
            End If
            
            bytLen = GetMaxBedLen(lng����ID)
            '��ǰ�����Ĺ��ÿմ�+��ǰ������ǰ���ҵĿմ�
            strSql = "Select ����,�Ա����,�����,�ȼ�ID From ��λ״����¼" & vbNewLine & _
                    " Where ״̬='�մ�'" & vbNewLine & _
                    " And instr([1],�Ա����)>0 And (����ID is Null Or ����ID=[2]) And ����ID=[3] " & vbNewLine & _
                    " Order by  LPad(NVL(�����,0), 10, ' '),LPad(����, 10, ' ')"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp, Val(mrsPatiInfo!��Ժ����id), lng����ID)
            Set mrsBeds = rsTmp.Clone
            
            For i = 1 To rsTmp.RecordCount
                cboNew.AddItem Space(bytLen - Len(rsTmp!����)) & rsTmp!���� & IIf(IsNull(rsTmp!�����), "", " ����:" & rsTmp!�����)
                If mbytInFun = 0 Then
                    If rsTmp!���� = mstrĿ�괲�� Then cboNew.ListIndex = cboNew.NewIndex
                ElseIf mbytInFun = 2 Then
                    If rsTmp!���� = "" & mrsPatiInfo!��Ҫ���� Then cboNew.ListIndex = cboNew.NewIndex
                End If
                                
                If mbytInFun = 0 Or UBound(Split(txtPre.Text, ",")) > 0 Then
                    lvw.ListItems.Add , "_" & rsTmp!����, rsTmp!���� & IIf(IsNull(rsTmp!�����), "", " ����:" & rsTmp!�����)
                    lvw.ListItems(lvw.ListItems.Count).Tag = "" & rsTmp!�����
                    If mbytInFun = 2 Then
                        If InStr(1, "," & txtPre.Text & ",", "," & rsTmp!���� & ",") > 0 Then
                            lvw.ListItems(lvw.ListItems.Count).Checked = True
                        End If
                    End If
                End If
                
                rsTmp.MoveNext
            Next
            
            If chk����.Value = 1 Then
                If Not lvw.ListItems.Count > 0 Then Exit Sub
                lvw.ListItems(1).Selected = True
                lvw.SelectedItem.EnsureVisible
            Else
                If cboNew.ListIndex = -1 And cboNew.ListCount > 0 Then cboNew.ListIndex = 0
            End If
                             
        Case 1 '1-����������
            
            strSql = "Select A.����,A.״̬ From ��λ״����¼ A" & vbNewLine & _
                    " Where (A.����ID,A.����ID,A.�����) In (Select Distinct B.����ID,B.����ID,B.����� From ��λ״����¼ B Where ����ID = [1]) " & vbNewLine & _
                    " And (A.״̬ = 'ռ��' And ����ID = [1] Or A.״̬ = '�մ�') And instr([2],�Ա����)>0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng����ID, strTmp)
            If rsTmp.RecordCount < 2 Then Exit Sub
            
            For i = 1 To rsTmp.RecordCount
                lvw.ListItems.Add , "_" & rsTmp!����, rsTmp!����
                
                If rsTmp!״̬ = "ռ��" Then
                    lvw.ListItems(lvw.ListItems.Count).Checked = True
                    lvw.ListItems(lvw.ListItems.Count).Selected = True
                    
                    cboNew.AddItem rsTmp!���� '������Ҫ����
                    If rsTmp!���� = "" & mrsPatiInfo!��Ҫ���� Then cboNew.ListIndex = cboNew.NewIndex
                End If
                
                rsTmp.MoveNext
            Next
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
    mlngUnit = 0
    
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

Private Function StringDelItem(ByVal strAll As String, ByVal strItem As String) As String
'���ܣ���ָ�����ַ����б���ɾ��һ��(����ж��ƥ���,ֻ�Ƴ���һ��)
    Dim i As Long, arrTmp As Variant
    
    arrTmp = Split(strAll, ",")
    For i = 0 To UBound(arrTmp)
        If arrTmp(i) = strItem Then
            strItem = ""
        Else
            StringDelItem = StringDelItem & "," & arrTmp(i)
        End If
    Next
    StringDelItem = Mid(StringDelItem, 2)
End Function

Private Sub cmdOK_Click()
    Dim strBeds As String, strBed As String, strSql As String, strUnitID As String, strMainBed As String
    Dim dMax As Date, i As Integer, j As Integer, blnTrans As Boolean
    Dim strRoom As String, Curdate As Date, strBedGrids As String, strBedGridsNew As String
    Dim rsTmp As New ADODB.Recordset
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ����)
    If mbytInFun <> 2 Then
        Curdate = zlDatabase.Currentdate
        If CDate(txtDate.Text) > Curdate Then
            If CDate(txtDate.Text) - Curdate > 30 Then
                MsgBox "����ʱ��ȵ�ǰʱ���ù���,���飡", vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
            If MsgBox("����ʱ������˵�ǰϵͳʱ��,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        dMax = GetMaxDate(mlng����ID, mlng��ҳID)
        If CDate(txtDate.Text) <= dMax Then
            MsgBox "���˻���ʱ���������ϴα䶯��ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    End If
    
    If chk����.Value = 1 Then
        If Trim(cboNew.Text) = "" Then
            MsgBox "��Ϊ����ָ������λ��", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(Trim(cboNew.Text), "��ͥ����") > 0 Then
            strMainBed = ""
        ElseIf InStr(Trim(cboNew.Text), " ����") > 0 Then
            strMainBed = Mid(Trim(cboNew.Text), 1, InStr(Trim(cboNew.Text), " ����") - 1)
        Else
            strMainBed = Trim(cboNew.Text)
        End If
    Else
        If cboNew.ListIndex = -1 Then
            MsgBox "��ѡ��Ҫ����Ĵ�λ��", vbInformation, gstrSysName
            cboNew.SetFocus: Exit Sub
        End If
        strMainBed = Trim(Split(cboNew.Text, "����:")(0))
    End If
    
    Select Case mbytInFun
        Case 0
            'ȡ��λ
            If chk����.Value = 0 Then
                If InStr(Trim(cboNew.Text), "��ͥ����") > 0 Then
                    strBeds = ""
                Else
                    If InStr(Trim(cboNew.Text), " ����") > 0 Then
                        strBeds = Mid(Trim(cboNew.Text), 1, InStr(Trim(cboNew.Text), " ����") - 1)
        
                        strRoom = Mid(Trim(cboNew.Text), InStr(Trim(cboNew.Text), "����:") + 3)
                        
                        strSql = "Select �Ա� From ������Ϣ A,��λ״����¼ B  Where A.����ID = b.����id And b.����ID Is Not Null And ����ID = [1] And ����� =[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit, strRoom)
                        
                        Do While Not rsTmp.EOF
                         
                            If Trim(txt����.Tag) <> rsTmp!�Ա� Then
                                If (MsgBox("ָ����λ���ڷ��������Ů��ס������Ƿ������ס��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                                    Exit Do
                                Else
                                    Exit Sub
                                    cboNew.SetFocus
                                End If
                            End If
                            rsTmp.MoveNext
                        Loop
                    Else
                        strBeds = Trim(cboNew.Text)
                    End If
                End If
                
            Else
                For i = 1 To lvw.ListItems.Count
                    If lvw.ListItems(i).Checked Then
                        j = j + 1
                        If j = 1 Then
                            strRoom = lvw.ListItems(i).Tag
                        ElseIf lvw.ListItems(i).Tag <> strRoom Then
                            MsgBox "���˰���ʱ����ѡ��ͬһ�������ڵĴ�λ��", vbInformation, gstrSysName
                            lvw.SetFocus: Exit Sub
                        End If
                        strBeds = strBeds & "," & Mid(lvw.ListItems(i).Key, 2)
                    End If
                Next
                If strBeds = "" Then
                    MsgBox "��ѡ��Ҫ����Ĵ�λ��", vbInformation, gstrSysName
                    lvw.SetFocus: Exit Sub
                End If
                strBeds = Mid(strBeds, 2)
                If UBound(Split(strBeds, ",")) = 0 Then
                    MsgBox "���˰���ʱ����Ӧѡ���������ϵĴ�λ��", vbInformation, gstrSysName
                    lvw.SetFocus: Exit Sub
                End If
                
                If strBeds = txtPre.Text Then
                    MsgBox "�����»���İ�����λ��ԭ��λ��ͬ,������ѡ��", vbInformation, gstrSysName
                    lvw.SetFocus: Exit Sub
                End If
            End If
            
            strUnitID = mlngUnit
        Case 1
            For i = 1 To lvw.ListItems.Count
                If lvw.ListItems(i).Checked Then Exit For
            Next
            If i = lvw.ListItems.Count + 1 Then
                MsgBox "������ѡ��һ�Ų�����", vbInformation, gstrSysName
                lvw.SetFocus: Exit Sub
            End If
            For i = 1 To lvw.ListItems.Count
                If lvw.ListItems(i).Checked Then
                    strBeds = strBeds & "," & Mid(lvw.ListItems(i).Key, 2)
                End If
            Next
            strBeds = Mid(strBeds, 2)
            strUnitID = mlngUnit
        Case 2
            If UBound(Split(txtPre.Text, ",")) > 0 Then '���Ŵ�
                j = 0
                For i = 1 To lvw.ListItems.Count
                    If lvw.ListItems(i).Checked Then j = j + 1
                Next
                If j <> UBound(Split(txtPre.Text, ",")) + 1 Then
                    MsgBox "�°��ŵĴ�λ������ԭ��ס��λ������һ�������飡", vbExclamation, gstrSysName
                    Exit Sub
                End If
                
                strBeds = strMainBed  '�����ŷŵ���һ�������ڴ���
                strBedGrids = mstrĿ�괲��
                For i = 1 To lvw.ListItems.Count
                    If lvw.ListItems(i).Checked Then
                        strBed = Trim(lvw.ListItems(i).Text)
                        If UBound(Split(strBed, "����:")) = 0 Then '�޷����
                            mrsBeds.Filter = "����='" & Trim(strBed) & "'"
                        Else
                            mrsBeds.Filter = "����='" & Trim(Split(strBed, "����:")(0)) & "' and �����='" & Split(strBed, "����:")(1) & "'"
                        End If
                        
                        strBedGridsNew = StringDelItem(strBedGrids, mrsBeds!�ȼ�ID)
                        If strBedGridsNew = strBedGrids Then
                            MsgBox "�°��ŵĴ�λ�ȼ���ԭ��ס��λ�ȼ���һ�������飡", vbExclamation, gstrSysName
                            Exit Sub
                        Else
                            strBedGrids = strBedGridsNew
                        End If
                        If InStr(1, strBeds, Trim(Split(strBed, "����:")(0))) = 0 Then    '�������ѷŵ���һ��
                            strBeds = strBeds & "," & Trim(Split(strBed, "����:")(0))
                        End If
                    End If
                Next
                If strBedGrids <> "" Then
                    MsgBox "�°��ŵĴ�λ�ȼ���ԭ��ס��λ�ȼ���һ�������飡", vbExclamation, gstrSysName
                    Exit Sub
                End If
            Else '���Ŵ�
                '�����м�ͥ����,��Ϊ������Ժʱ��������ָ����λ
                strBed = Trim(cboNew.Text)
                If UBound(Split(strBed, "����:")) = 0 Then
                    mrsBeds.Filter = "����='" & strBed & "'"
                Else
                    mrsBeds.Filter = "����='" & Trim(Split(strBed, "����:")(0)) & "' and �����='" & Split(strBed, "����:")(1) & "'"
                End If
                '�ȼ��Ƚ�
                If mstrĿ�괲�� <> CStr(mrsBeds!�ȼ�ID) Then
                    MsgBox "�°��ŵĴ�λ�ȼ���ԭ��ס��λ�ȼ���һ�������飡", vbExclamation, gstrSysName
                    Exit Sub
                End If
                strBeds = Trim(Split(strBed, "����:")(0))
            End If
    End Select
    
    If mbytInFun <> 2 Then
        strSql = "zl_���˱䶯��¼_MOVE(" & mlng����ID & "," & mlng��ҳID & "," & _
            "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),'" & strBeds & "'," & _
            "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & strUnitID & ",'" & strMainBed & "')"
            
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            zlDatabase.ExecuteProcedure strSql, Me.Caption
        
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
    End If
    
    mstrĿ�괲�� = strBeds
    gblnOK = True
    
    On Error Resume Next
    '�����ɹ��󴥷���Ϣ
    If mclsMipModule.IsConnect = True And mbytInFun <> 2 Then
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
        mclsXML.appendData "patient_sex", txt����.Tag, xsString '�Ա�
        'in_number       סԺ��  1   S
        mclsXML.appendData "in_number", txtסԺ��.Text, xsString  'סԺ��
        mclsXML.AppendNode "in_patient", True
        
        '��ǰ���
        'current_state       ��ǰ���    1
        mclsXML.AppendNode "current_state"
        'current_area_id     ��ǰ����id  0..1    N
        mclsXML.appendData "current_area_id", Val(Nvl(mrsPatiInfo!��ǰ����ID)), xsNumber
        'current_area_title      ��ǰ����    0..1    S
        mclsXML.appendData "current_area_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'current_dept_id     ��ǰ����id  1   N
        mclsXML.appendData "current_dept_id", Val(Nvl(mrsPatiInfo!��Ժ����id, 0)), xsNumber
        'current_dept_title      ��ǰ����    1   S
        mclsXML.appendData "current_dept_title", Nvl(mrsPatiInfo!��ǰ����), xsString
        'current_room        ��ǰ����    0..1    S
        mclsXML.appendData "current_room", txtPre.Tag, xsString
        'current_bed     ��ǰ����    1   S
        mclsXML.appendData "current_bed", Nvl(mrsPatiInfo!��Ҫ����), xsString
        mclsXML.AppendNode "current_state", True
        
        strSql = " Select ID �䶯id,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ Where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] And ��ʼʱ��+0=[4] And NVL(���Ӵ�λ,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���˱䶯��¼", mlng����ID, mlng��ҳID, 4, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        'ת����Ϣ
        'change_state        ת����Ϣ    1
        mclsXML.AppendNode "change_state"
        'change_id       ת�Ʊ��id  1   N
        mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
        'change_date     ���ʱ��    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
        'change_room     ��ס����    0..1    S
        mclsXML.appendData "change_room", strRoom, xsString
        'change_bed      ��ס����    1   S
        mclsXML.appendData "change_bed", strMainBed, xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_004", mclsXML.XmlText
    End If
    If Err <> 0 Then Err.Clear
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadMainBed
End Sub

Private Sub LoadMainBed()
    Dim i As Integer, strBed As String
    
    If cboNew.ListIndex <> -1 Then strBed = cboNew.Text
    cboNew.Clear
    For i = 1 To lvw.ListItems.Count
        If lvw.ListItems(i).Checked Then
            cboNew.AddItem lvw.ListItems(i).Text
            If lvw.ListItems(i).Text = strBed Then cboNew.ListIndex = cboNew.NewIndex
            If cboNew.ListIndex = -1 Then
                If lvw.ListItems(i).Text = mrsPatiInfo!��Ҫ���� Then cboNew.ListIndex = cboNew.NewIndex
            End If
        End If
    Next
    If cboNew.ListIndex = -1 And cboNew.ListCount = 1 Then cboNew.ListIndex = 0
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    lvw.ToolTipText = lvw.ListItems(Item.Index).Text
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
        ByVal bytInFun As Byte, ByRef strĿ�괲�� As String, ByVal str���� As String, ByVal strPrivs As String) As Boolean
'#########################################################################################################
'### ������bytInFun :'0-����,1-����������,2-������Ժʱ���´�
'###       strĿ�괲�� :��mbytInFun=0ʱ��ʾ�϶���ָ���Ĵ�λ�ϣ�=2ʱ��ʾ���˵�ǰ��λ��Ӧ�ĵȼ�ID
'###       str���� :��mbytInFun=2ʱ����ֵ,��ʾ��Ժǰ�Ĵ�λ
'### ���أ�Ŀ�괲��
'#########################################################################################################
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mbytInFun = bytInFun
    mstrĿ�괲�� = strĿ�괲��
    mstr���� = str����
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    strĿ�괲�� = mstrĿ�괲��
    
    ShowMe = gblnOK
End Function
