VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Begin VB.Form frmBedSwap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���˴�λ�Ի�"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBedSwap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboNew 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   1845
   End
   Begin VB.Frame fraBedSwap 
      Height          =   1150
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   5565
      Begin VB.TextBox txtSwap���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   660
         Width           =   1800
      End
      Begin VB.TextBox txtSwapסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txtSwap���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   270
         Width           =   1800
      End
      Begin VB.TextBox txtSwapPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblSwapDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   345
         TabIndex        =   24
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblSwapPName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   345
         TabIndex        =   23
         Top             =   330
         Width           =   360
      End
      Begin VB.Label lblSwapInHosNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2910
         TabIndex        =   22
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ����"
         Height          =   180
         Left            =   2910
         TabIndex        =   21
         Top             =   720
         Width           =   540
      End
   End
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3165
      Width           =   5805
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   345
         Left            =   3240
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   345
         Left            =   4440
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame fraBed 
      Height          =   1150
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   5565
      Begin VB.TextBox txtPre 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   660
         Width           =   1845
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   1800
      End
      Begin VB.TextBox txtסԺ�� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   1845
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   660
         Width           =   1800
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
      Begin VB.Label lblInHosNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2910
         TabIndex        =   12
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lblPName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   345
         TabIndex        =   11
         Top             =   330
         Width           =   360
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   345
         TabIndex        =   10
         Top             =   720
         Width           =   360
      End
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   300
      Left            =   3630
      TabIndex        =   2
      Top             =   1440
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd hh:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.Label lblNew 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ŀ�겡��"
      Height          =   195
      Left            =   105
      TabIndex        =   25
      Top             =   1500
      Width           =   720
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      Height          =   180
      Left            =   2850
      TabIndex        =   15
      Top             =   1500
      Width           =   720
   End
End
Attribute VB_Name = "frmBedSwap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng����ID As Long              '��ǰ����ID
Private mlng��ҳID As Long              '��ǰ������ҳID
Private mlngBeSwap����ID As Long        '��������ID
Private mlngBeSwap��ҳID As Long        '����������ҳID
Private mstr���� As String              '��ǰ���˴���
Private mstrĿ�괲�� As String          '��λ�Ի���Ŀ�괲�ţ����������˴��ţ�
Private mfrmParent As Object

Public mstrPrivs As String              'Ȩ��
Public mlngUnit As Long                 '���˲���ID

Private mrsPatiInfo As ADODB.Recordset
Private mrsSwapPatiInfo As ADODB.Recordset
Private mrsBeds As ADODB.Recordset '��ѡ��λ��

Private WithEvents mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mclsXML As zl9ComLib.clsXML

Private Sub cboNew_Click()
    Dim strBed As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errHandle:
    
    If cboNew.ListIndex <> -1 And cboNew.ListCount > 0 Then
        'ȥ��λ
        If InStr(Trim(cboNew.Text), " ����") > 0 Then
            strBed = Mid(Trim(cboNew.Text), 1, InStr(Trim(cboNew.Text), " ����") - 1)
        Else
            strBed = Trim(cboNew.Text)
        End If
        '���ݴ��ż� �������Ҳ��Ҳ���ID
        gstrSQL = "Select ����ID From ��λ״����¼ Where (����ID is Null Or ����ID=[1] Or ����=1) And ����ID=[2] And ����=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mrsPatiInfo!��Ժ����id), mlngUnit, strBed)
        
        mlngBeSwap����ID = rsTmp!����ID
        mlngBeSwap��ҳID = GetMax��ҳID(rsTmp!����ID) - 1
        
        '���ݲ���ID �� ��ҳID ��ȡ������Ϣ
        Set mrsSwapPatiInfo = GetPatiInfo(mlngBeSwap����ID, mlngBeSwap��ҳID)
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        With mrsSwapPatiInfo
            txtSwap����.Text = !����
            txtSwapסԺ��.Text = "" & !סԺ��
            txtSwap����.Text = !��ǰ����
        End With
        
        txtSwapPre.Text = cboNew.Text
        'Ŀ�괲���޸�Ϊ��ѡ����
        mstrĿ�괲�� = Trim(cboNew.Text)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboNew_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii <> 13 Then
        If SendMessage(cboNew.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
        lngIdx = cbo.MatchIndex(cboNew.hWnd, KeyAscii, 0.5)
        If lngIdx <> -2 Then cboNew.ListIndex = lngIdx
    ElseIf cboNew.ListIndex <> -1 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strBeds As String, strSql As String, strMainBed As String
    Dim dMax As Date, dMaxSwap As Date, i As Integer, j As Integer, blnTrans As Boolean
    Dim strRoom As String, strOldRoom As String, Curdate As Date, strBedGrids As String, strBedGridsNew As String
    Dim rsTmp As ADODB.Recordset, intMainBed As Integer
    Dim arrSQL() As String, intLoop As Integer
    Dim strErr As String
    
    'ʱ�䲻�ܳ�����ǰʱ��̫��(һ����)
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
    dMaxSwap = GetMaxDate(mlngBeSwap����ID, mlngBeSwap��ҳID)
    If CDate(txtDate.Text) <= dMax Then
        MsgBox "���˻���ʱ���������ϴα䶯��ʱ�� " & Format(dMax, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    If CDate(txtDate.Text) <= dMaxSwap Then
        MsgBox "���˻���ʱ���������ϴα䶯��ʱ�� " & Format(dMaxSwap, "yyyy-MM-dd HH:mm:ss") & " ��", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Sub
    End If
    
    
    If cboNew.ListIndex = -1 Then
        MsgBox "��ѡ��Ҫ����Ĵ�λ��", vbInformation, gstrSysName
        cboNew.SetFocus: Exit Sub
    End If
        
    '��ѡ�в��˴�λ����ָ����λ
    strMainBed = Trim(Split(cboNew.Text, "����:")(0))

    'ȡ��λ
    '�ж�Ŀ�괲λ���ڷ����Ƿ������Ů��ס���
    If InStr(Trim(cboNew.Text), " ����") > 0 Then
        strBeds = Mid(Trim(cboNew.Text), 1, InStr(Trim(cboNew.Text), " ����") - 1)
        
        strRoom = Mid(Trim(cboNew.Text), InStr(Trim(cboNew.Text), "����:") + 3)
        
        strSql = "Select �Ա� From ������Ϣ A,��λ״����¼ B  Where A.����ID = b.����id And b.����ID Is Not Null And ����ID = [1] And ����� =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit, strRoom)
        
        Do While Not rsTmp.EOF
         
            If Trim(mrsPatiInfo!�Ա�) <> rsTmp!�Ա� Then
                If (MsgBox("Ŀ�괲λ���ڷ��������Ů��ס������Ƿ������ס��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
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
    '�жϵ�ǰ��λ���ڷ����Ƿ������Ů��ס���
    strSql = "select ����� from ��λ״����¼ where ����id=[1] and ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit, mstr����)
    If Nvl(rsTmp!�����) <> "" Then
        strOldRoom = Nvl(rsTmp!�����)
        strSql = "Select �Ա� From ������Ϣ A,��λ״����¼ B  Where A.����ID = b.����id And b.����ID Is Not Null And ����ID = [1] And ����� =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngUnit, Val("" + rsTmp!�����))
        
        Do While Not rsTmp.EOF
            If Trim(mrsSwapPatiInfo!�Ա�) <> rsTmp!�Ա� Then
                If (MsgBox("Ŀ�괲λ���ڷ��������Ů��ס������Ƿ������ס��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)) = vbYes Then
                    Exit Do
                Else
                    Exit Sub
                    cboNew.SetFocus
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '�������˲�������д�λ�Ի�
    Set rsTmp = GetPatiBeds(mlng����ID)
    If rsTmp.RecordCount > 1 Then
        MsgBox mstr���� & "������Ϊ�������ˣ���������д�λ�Ի���", vbInformation, gstrSysName
        Exit Sub
    Else
        Set rsTmp = GetPatiBeds(mlngBeSwap����ID)
        If rsTmp.RecordCount > 1 Then
            MsgBox strBeds & "������Ϊ�������ˣ���������д�λ�Ի���", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '82383:LPF,������鴲λ�Ƿ�Ϊ��,��Ϊ�ղ�����������λ�Ի����Ŀ�괲λ��Ϊ�գ������Ƿ��Ǵ�λ�Ի��Ĳ��ˣ����˲�ͬ��������л���
    ReDim Preserve arrSQL(0)
    arrSQL(UBound(arrSQL)) = "zl_���˱䶯��¼_Move(" & mlng����ID & "," & mlng��ҳID & "," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),'" & strBeds & "'," & _
        "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngUnit & ",'" & strBeds & "'," & mlngBeSwap����ID & ")"
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "zl_���˱䶯��¼_Move(" & mlngBeSwap����ID & "," & mlngBeSwap��ҳID & "," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),'" & mstr���� & "'," & _
        "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & mlngUnit & ",'" & mstr���� & "'," & mlng����ID & ")"
            
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For intLoop = LBound(arrSQL) To UBound(arrSQL)
        zlDatabase.ExecuteProcedure arrSQL(intLoop), Me.Caption
    
        If Val("" & mrsPatiInfo!����) <> 0 Then
            If Not gclsInsure.ModiPatiSwap(mlng����ID, mlng��ҳID, Val("" & mrsPatiInfo!����), "1") Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        ElseIf Val("" & mrsSwapPatiInfo!����) <> 0 Then
            If Not gclsInsure.ModiPatiSwap(mlngBeSwap����ID, mlngBeSwap��ҳID, Val("" & mrsSwapPatiInfo!����), "1") Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
    Next
    gcnOracle.CommitTrans: blnTrans = False
    '����96847��118004
    If CreateXWHIS() Then
        If gobjXWHIS.HISModPati(2, mlng����ID, mlng��ҳID) <> 1 Or gobjXWHIS.HISModPati(2, mlngBeSwap����ID, mlngBeSwap��ҳID) <> 1 Then
            MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISModPati)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        End If
    ElseIf gblnXW = True Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������RIS�ӿڴ���ʧ��δ����(HISModPati)�ӿڣ�����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
    End If
    mstrĿ�괲�� = strBeds
    gblnOK = True
           
    On Error Resume Next
    '�����ɹ��󴥷���Ϣ
    '--����һ
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
        mclsXML.appendData "patient_sex", Nvl(mrsPatiInfo!�Ա�), xsString '�Ա�
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
        mclsXML.appendData "current_room", strOldRoom, xsString
        'current_bed     ��ǰ����    1   S
        mclsXML.appendData "current_bed", mstr����, xsString
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
        mclsXML.appendData "change_bed", mstrĿ�괲��, xsString
        mclsXML.AppendNode "change_state", True

        mclsMipModule.CommitMessage "ZLHIS_PATIENT_004", mclsXML.XmlText
    End If
    '--���˶�
    If mclsMipModule.IsConnect = True Then
         mclsXML.ClearXmlText '��������е�XML
        '--������Ϣ��װ
        '������Ϣ
        mclsXML.AppendNode "in_patient"
        'patient_id      ����id  1   N
        mclsXML.appendData "patient_id", mlngBeSwap����ID, xsNumber  '����ID
        'page_id     ��ҳid  1   N
        mclsXML.appendData "page_id", mlngBeSwap��ҳID, xsNumber  '��ҳID
        'patient_name        ����    1   S
        mclsXML.appendData "patient_name", txtSwap����.Text, xsString '����
        'patient_sex     �Ա�    0..1    S
        mclsXML.appendData "patient_sex", Nvl(mrsSwapPatiInfo!�Ա�), xsString '�Ա�
        'in_number       סԺ��  1   S
        mclsXML.appendData "in_number", txtSwapסԺ��.Text, xsString  'סԺ��
        mclsXML.AppendNode "in_patient", True
        
        '��ǰ���
        'current_state       ��ǰ���    1
        mclsXML.AppendNode "current_state"
        'current_area_id     ��ǰ����id  0..1    N
        mclsXML.appendData "current_area_id", Val(Nvl(mrsSwapPatiInfo!��ǰ����ID)), xsNumber
        'current_area_title      ��ǰ����    0..1    S
        mclsXML.appendData "current_area_title", Nvl(mrsSwapPatiInfo!��ǰ����), xsString
        'current_dept_id     ��ǰ����id  1   N
        mclsXML.appendData "current_dept_id", Val(Nvl(mrsSwapPatiInfo!��Ժ����id, 0)), xsNumber
        'current_dept_title      ��ǰ����    1   S
        mclsXML.appendData "current_dept_title", Nvl(mrsSwapPatiInfo!��ǰ����), xsString
        'current_room        ��ǰ����    0..1    S
        mclsXML.appendData "current_room", strRoom, xsString
        'current_bed     ��ǰ����    1   S
        mclsXML.appendData "current_bed", txtSwapPre.Text, xsString
        mclsXML.AppendNode "current_state", True
        
        strSql = " Select ID �䶯id,��ʼʱ�� �䶯ʱ�� From ���˱䶯��¼ Where ����ID=[1] And ��ҳId=[2] And ��ʼԭ��=[3] And ��ʼʱ��+0=[4] And NVL(���Ӵ�λ,0)=0"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���˱䶯��¼", mlngBeSwap����ID, mlngBeSwap��ҳID, 4, CDate(Format(txtDate.Text, "YYYY-MM-DD HH:mm:ss")))
        'ת����Ϣ
        'change_state        ת����Ϣ    1
        mclsXML.AppendNode "change_state"
        'change_id       ת�Ʊ��id  1   N
        mclsXML.appendData "change_id", rsTmp!�䶯ID, xsNumber
        'change_date     ���ʱ��    1   S
        mclsXML.appendData "change_date", Format(Nvl(rsTmp!�䶯ʱ��), "YYYY-MM-DD HH:mm:ss"), xsString
        'change_room     ��ס����    0..1    S
        mclsXML.appendData "change_room", strOldRoom, xsString
        'change_bed      ��ס����    1   S
        mclsXML.appendData "change_bed", mstr����, xsString
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    gblnOK = False
    
    If Not InitData Then Unload Me: Exit Sub
    
    fraBed.Caption = mstr���� & "������"
    fraBedSwap.Caption = mstrĿ�괲�� & "������"
    If cboNew.ListCount = 0 Then
        MsgBox "�������ڿ��ҵĲ�����û�к��ʵĴ�λ�ɹ��Ի���", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    '������Ϣ����
    Set mclsMipModule = New zl9ComLib.clsMipModule
    Call mclsMipModule.InitMessage(glngSys, 1132, mstrPrivs, mfrmParent.hWnd)
    Call AddMipModule(mclsMipModule)
    Set mclsXML = New zl9ComLib.clsXML
End Sub

Private Function InitData() As Boolean
    Dim i As Integer, rsTmp As ADODB.Recordset, str���� As String
    
    Set mrsPatiInfo = GetPatiInfo(mlng����ID, mlng��ҳID)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    
    With mrsPatiInfo
        txt����.Text = !����
        txtסԺ��.Text = "" & !סԺ��
        txt����.Text = !��ǰ����
        If Trim(mstr����) = "" Then mstr���� = !��ǰ����
    End With
    
    txtPre.Text = mstr����
    '��ʼ����λ
    If InitBed(mlngUnit) = False Then Exit Function
    
    InitData = True
End Function

Private Function InitBed(ByVal lng����ID As Long) As Boolean
'���ܣ���ʼ����λ,��ʱȡ�ò��������Ҷ�Ӧ�����пմ�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strSQLtmp As String, i As Integer
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
        
    cboNew.Clear

    bytLen = GetMaxBedLen(lng����ID)
    
    Set rsTmp = GetPatiBeds(mlng����ID)
    
    If rsTmp.RecordCount > 1 Then
        MsgBox "�ò���Ϊ�������ˣ���������д�λ�Ի���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�Ա�����
    If rsTmp!�Ա���� = "���޴�" Then
        strSQLtmp = " And (A.�Ա���� ='���޴�' Or (A.�Ա���� = '�д�' And '" & rsTmp!�Ա� & "'='��') Or (A.�Ա���� = 'Ů��' And '" & rsTmp!�Ա� & "'='Ů')) "
    ElseIf rsTmp!�Ա���� = "�д�" Then
        strSQLtmp = " And ((A.�Ա���� = '���޴�' And B.�Ա� = '��') Or (A.�Ա���� = '�д�' And '" & rsTmp!�Ա� & "'='��'))"
    ElseIf rsTmp!�Ա���� = "Ů��" Then
        strSQLtmp = " And ((A.�Ա���� = '���޴�' And B.�Ա� = 'Ů') Or (A.�Ա���� = 'Ů��' And '" & rsTmp!�Ա� & "'='Ů'))"
    End If
    
    strSql = "Select Distinct A.����,A.�Ա����,A.�����,A.�ȼ�ID,B.�Ա�,C.״̬ From ��λ״����¼ A, ������Ϣ B, ������ҳ C, ���˱䶯��¼ D " & vbNewLine & _
                " Where A.����ID=c.����ID And B.����ID=C.����ID And C.����ID=D.����ID And C.��ҳID=D.��ҳID And (" & _
                IIf(rsTmp!���� = 1, " A.����ID is Null Or A.����ID=[1] Or A.����=1 ", "A.����ID is Null Or A.����ID=[1] Or (A.����=1 And A.����id=[1])") & _
                ") And A.����ID=[2] And A.״̬='ռ��' And C.״̬ Not In(2,3)" & vbNewLine & _
                strSQLtmp & " Order by  LPad(NVL(A.�����,0), 10, ' '),LPad(A.����, 10, ' ')"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mrsPatiInfo!��Ժ����id), lng����ID)
    Set mrsBeds = rsTmp.Clone
    
    For i = 1 To rsTmp.RecordCount
        If Not rsTmp!���� = mstr���� Then cboNew.AddItem Space(bytLen - Len(rsTmp!����)) & rsTmp!���� & IIf(IsNull(rsTmp!�����), "", " ����:" & rsTmp!�����)
        If rsTmp!���� = mstrĿ�괲�� Then cboNew.ListIndex = cboNew.NewIndex
        rsTmp.MoveNext
    Next
    InitBed = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowMe(frmParent As Object, ByVal lngUnit As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str���� As String, _
            ByRef strĿ�괲�� As String, ByVal strPrivs As String) As Boolean
    On Error Resume Next
    Set mfrmParent = frmParent
    mlngUnit = lngUnit
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstr���� = str����
    mstrĿ�괲�� = strĿ�괲��
    mstrPrivs = strPrivs
    
    Me.Show 1, frmParent
    strĿ�괲�� = mstrĿ�괲��
    ShowMe = gblnOK
End Function

Private Sub Form_Unload(Cancel As Integer)
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

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    If Not IsDate(txtDate.Text) Then txtDate.SetFocus
End Sub
