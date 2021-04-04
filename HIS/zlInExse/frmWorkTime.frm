VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWorkTime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ӡ�ɿ���"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmWorkTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboTimes 
      Height          =   300
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "�ɿ���(&P)"
      Height          =   350
      Left            =   405
      TabIndex        =   5
      Top             =   2880
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   2610
      Width           =   6555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   3960
      TabIndex        =   6
      Top             =   2895
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   2865
      TabIndex        =   4
      Top             =   2895
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Index           =   0
      Left            =   -270
      TabIndex        =   11
      Top             =   1440
      Width           =   6555
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   300
      Left            =   1155
      TabIndex        =   3
      Top             =   2175
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
      Format          =   134086659
      CurrentDate     =   38175
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   300
      Left            =   1155
      TabIndex        =   2
      Top             =   1710
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
      Format          =   134086659
      CurrentDate     =   38175
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Left            =   1155
      TabIndex        =   0
      Top             =   990
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   134086659
      CurrentDate     =   2
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   3720
      TabIndex        =   13
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ֹʱ��"
      Height          =   180
      Left            =   375
      TabIndex        =   10
      Top             =   2235
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼʱ��"
      Height          =   180
      Left            =   375
      TabIndex        =   9
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϰ�����"
      Height          =   180
      Left            =   375
      TabIndex        =   8
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "�ڴ�ӡ�ɿ���֮ǰ������ָ���ϰ�����,����ʵ�ʹ���ʱ����д�����������ϰ࿪ʼʱ��ͽ���ʱ��(ҹ����ܿ���)��"
      Height          =   540
      Left            =   975
      TabIndex        =   7
      Top             =   195
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   255
      Picture         =   "frmWorkTime.frx":000C
      Top             =   390
      Width           =   480
   End
End
Attribute VB_Name = "frmWorkTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte '1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-���￨,6-���ѿ��ɿ�
Private mrsTimes As ADODB.Recordset '��ǰ�ɿ����ڵĽɿ����

Public Sub ShowMe(frmParent As Object, bytType As Byte)
    mbytType = bytType
    
    On Error Resume Next
    Me.Show 1, frmParent
End Sub


Private Sub cboTimes_Click()
    
    If cboTimes.Visible Then
        Call SetTimeRange
        cboTimes.Tag = "Click"
    End If
End Sub

Private Sub cboTimes_Validate(Cancel As Boolean)
    If cboTimes.Tag = "Click" Then
        cboTimes.Tag = ""
    Else
        Call SetTimeRange
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim strReport As String
    
    If cmdSave.Enabled Then
        MsgBox "�ڴ�ӡ�ɿ���֮ǰ�����ȱ����ϰ࿪ʼ��ֹʱ�䡣", vbInformation, gstrSysName
        cmdSave.SetFocus: Exit Sub
    End If
    
    Select Case mbytType
        Case 1 'Ԥ����ɿ���
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1103_1"
        Case 2 '���ʽɿ���
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1137_1"
        Case 3 '�շѽɿ���
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1121_1"
        Case 4 '�ҺŽɿ���
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1111_1"
        Case 5 '���￨�ɿ���
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1102_1"
        Case 6 '���ѿ��ɿ���
            strReport = "ZL" & glngSys \ 100 & "_INSIDE_1503_2"
    End Select
    
    Call ReportOpen(gcnOracle, glngSys, strReport, Me, _
        "��ʼʱ��=" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss"), _
        "����ʱ��=" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss"), _
        "����Ա=" & UserInfo.����)
End Sub

Private Sub cmdSave_Click()
    Dim strSQL As String, strOldDate As String
    
    If cboTimes.ListIndex < 0 Then MsgBox "��ѡ��ɿ����!", vbInformation, App.ProductName
    If dtpBegin.Value >= dtpEnd.Value Then
        MsgBox "��ʼʱ��Ӧ��С����ֹʱ�䡣", vbInformation, gstrSysName
        If dtpBegin.Enabled Then dtpBegin.SetFocus
        Exit Sub
    End If
    
    If InStr(";" & gstrPrivs & ";", ";�޸��ϰ�ʱ��;") = 0 Then
        If MsgBox("���浱ǰ���õ��ϰ�ʱ��󽫲������޸�,��ȷ��Ҫ������?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    If cboTimes.ItemData(cboTimes.ListIndex) <> 0 Then '�޸Ľɿ���
        mrsTimes.Filter = "����=" & cboTimes.ItemData(cboTimes.ListIndex)
        strOldDate = "To_Date('" & Format(mrsTimes!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strOldDate = "Null"
    End If
    
    
    strSQL = "ZL_�շ�����¼_Insert(To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','YYYY-MM-DD')," & mbytType & "," & _
        "To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
        "To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & strOldDate & ")"
    On Error GoTo errH
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    
    'ˢ�����ݣ��Ա�û��Ȩ���޸ĵģ��ض��̶����ϰ�ʱ��
    Call dtpDate_Change
    
    cmdSave.Enabled = False
    cmdPrint.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpBegin_Change()
    cmdSave.Enabled = True
    lblMessage.Caption = ""
    If dtpEnd.Value < dtpBegin.Value Then
        lblMessage.Top = Label3.Top
        lblMessage.Caption = "���ܱȽ���ʱ���!"
        cmdSave.Enabled = False
    End If
End Sub

Private Sub SetTimeRange()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, curDate As Date, DatBegin As Date, strBDate As String, strEDate As String
    Dim blnHave As Boolean, dat�ϴ���ֹ As Date, dat�´ο�ʼ As Date
    
    On Error GoTo errH
    '���������˽���,��ѡ����ʱҪ�ָ�
    dtpBegin.Enabled = True
    dtpEnd.Enabled = True
    
    'һ���ж�νɿ�
    If cboTimes.ListCount > 1 Then
        If cboTimes.ItemData(cboTimes.ListIndex) = 0 Then  '�½ɿ�
            mrsTimes.Filter = "����=" & cboTimes.ItemData(cboTimes.ListCount - 1)
            DatBegin = mrsTimes!��ʼʱ��
            
            strBDate = " And ����=[1] And ��ʼʱ��=[4]"
            strEDate = " And ����>[1]"
        Else
            mrsTimes.Filter = "����=" & cboTimes.ItemData(cboTimes.ListIndex)
            DatBegin = mrsTimes!��ʼʱ��
            
            If cboTimes.ItemData(cboTimes.ListIndex) = 1 Then  '�����1�νɿ�
                strBDate = " And ����<[1]"
            Else
                strBDate = " And ����=[1] And ��ʼʱ��<[4]"
            End If
            
            If cboTimes.ListIndex = cboTimes.ListCount - 1 Then '�������һ�νɿ�
                strEDate = " And ����>[1]"
            Else
                strEDate = " And ����=[1] And ��ʼʱ��>[4]"
            End If
        End If
    Else
        strBDate = " And ����<[1]"
        strEDate = " And ����>[1]"
    End If
    
    
    '���ø�������޸ĵ�ʱ�䷶Χ
    '------------------------------------------------------------------------------------------
    curDate = zldatabase.Currentdate
    dtpBegin.MinDate = "1601-01-01": dtpBegin.MaxDate = "9999-12-31"
    dtpEnd.MinDate = "1601-01-01": dtpEnd.MaxDate = "9999-12-31"
    
    '��ʼʱ�䣺Ӧ�����ϴ���ֹʱ��,��ǰһ����
    strSQL = "Select Max(��ֹʱ��) as �ϴ���ֹ From �շ�����¼ Where �տ�Ա=[2] And ����=[3]" & strBDate
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.����, mbytType, DatBegin)
    
    blnHave = False
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!�ϴ���ֹ) Then blnHave = True
    End If
    If blnHave Then
        dat�ϴ���ֹ = rsTmp!�ϴ���ֹ
        dtpBegin.MinDate = DateAdd("s", 1, rsTmp!�ϴ���ֹ)
    Else
        dtpBegin.MinDate = Int(dtpDate.Value - 1)
    End If
    
    '��ֹʱ�䣺ӦС���´ο�ʼʱ��,���һ����(��������ǰʱ��)
    strSQL = "Select Min(��ʼʱ��) as �´ο�ʼ From �շ�����¼ Where �տ�Ա=[2] And ����=[3]" & strEDate
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.����, mbytType, DatBegin)
    
    blnHave = False
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!�´ο�ʼ) Then blnHave = True
    End If
    If blnHave Then
        dat�´ο�ʼ = rsTmp!�´ο�ʼ
        dtpEnd.MaxDate = DateAdd("s", -1, rsTmp!�´ο�ʼ)
    Else
        dtpEnd.MaxDate = curDate
    End If
    dtpBegin.MaxDate = dtpEnd.MaxDate
    dtpEnd.MinDate = dtpBegin.MinDate
    
    
        
    '����ȱʡ�ϰ�ʱ�䷶Χ
    '------------------------------------------------------------------------------------------
    '�ɿ��ش���޸�
    If cboTimes.ItemData(cboTimes.ListIndex) > 0 Then
        If cboTimes.ListCount = 1 Then
            strSQL = "Select ��ʼʱ��,��ֹʱ�� From �շ�����¼ Where �տ�Ա=[2] And ����=[3] And ����=[1] Order by ��ֹʱ�� Desc"
        Else
            strSQL = "Select ��ʼʱ��,��ֹʱ�� From �շ�����¼ Where �տ�Ա=[2] And ����=[3] And ����=[1] And ��ʼʱ��=[4] Order by ��ֹʱ�� Desc"
        End If
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.����, mbytType, DatBegin)
    
        cmdSave.Enabled = False
        
        If rsTmp.RecordCount = 1 Then
            dtpBegin.Value = rsTmp!��ʼʱ��
            dtpEnd.Value = rsTmp!��ֹʱ��
        Else
            '���һ��ɿ���,��ȱʡ���������ɿ�
            dtpBegin.Value = DateAdd("s", 1, rsTmp!��ֹʱ��)
            '��ֹʱ��ȱʡΪ�´ο�ʼʱ��-1s,��Ϊ�������ʱ��
            If dat�´ο�ʼ <> CDate(0) Then
                dtpEnd.Value = DateAdd("s", -1, dat�´ο�ʼ)
            Else
                If Format(dtpDate.Value, "yyyy-MM-dd 23:59:59") <= curDate And dtpEnd.MinDate <= dtpDate.Value Then
                    dtpEnd.Value = Format(dtpDate.Value, "yyyy-MM-dd 23:59:59")
                Else
                    dtpEnd.Value = curDate
                End If
            End If
        End If
        
        '�޸��ϰ�ʱ��Ȩ��,ָ�Ƿ����޸��ϰ�������һ��Ŀ�ʼʱ��ͽ���ʱ���Ȩ��,�������ϰ�ʱ�䱾����Ϊ����һ�δ�ӡ����Ľɿ���
        If InStr(";" & gstrPrivs & ";", ";�޸��ϰ�ʱ��;") = 0 Then
            dtpBegin.Enabled = False
            dtpEnd.Enabled = False
        End If
    Else
        '�µĽɿ�
        cmdSave.Enabled = True
        
        '��ʼʱ��ȱʡ�ϴ���ֹʱ��+1s
        If dat�ϴ���ֹ <> CDate(0) Then
            dtpBegin.Value = DateAdd("s", 1, dat�ϴ���ֹ)
            If InStr(";" & gstrPrivs & ";", ";�޸��ϰ�ʱ��;") = 0 Then dtpBegin.Enabled = False
        Else
            dtpBegin.Value = Int(dtpDate.Value) '�����޸���ǰһ��,ȱʡΪ����
        End If
        
        '��ֹʱ��ȱʡΪ�´ο�ʼʱ��-1s,��Ϊ�������ʱ��
        If dat�´ο�ʼ <> CDate(0) Then
            dtpEnd.Value = DateAdd("s", -1, dat�´ο�ʼ)
        Else
            If Format(dtpDate.Value, "yyyy-MM-dd 23:59:59") <= curDate And dtpEnd.MinDate <= dtpDate.Value Then
                dtpEnd.Value = Format(dtpDate.Value, "yyyy-MM-dd 23:59:59")
            Else
                dtpEnd.Value = curDate
            End If
        End If
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTimes(datThis As Date)
'���ؽɿ����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long
    
    'ע��:˳��Ӱ�����ȡ������
    strSQL = "Select Rownum ����, ��ʼʱ�� From (Select ��ʼʱ�� From �շ�����¼ Where ���� = [1] And �տ�Ա=[2] And ����=[3] Order By ��ʼʱ��)"
    On Error GoTo errH
    Set mrsTimes = zldatabase.OpenSQLRecord(strSQL, Me.Caption, datThis, UserInfo.����, mbytType)
        
    '����������ֹʱ���뵱��֮�����һ�νɿ����ڵ���С��ʼʱ��֮���޼��ʱ���������������ɿ�
    strSQL = "Select 1" & vbNewLine & _
            "From (Select Min(��ʼʱ��) ��ʼʱ��" & vbNewLine & _
            "       From �շ�����¼" & vbNewLine & _
            "       Where ���� > [1] And �տ�Ա = [2] And ���� = [3]) A," & vbNewLine & _
            "     (Select Max(��ֹʱ��) ��ֹʱ��" & vbNewLine & _
            "       From �շ�����¼" & vbNewLine & _
            "       Where ���� = [1] And �տ�Ա = [2] And ���� = [3]) B" & vbNewLine & _
            "Where A.��ʼʱ�� = B.��ֹʱ�� + 1 / 60 / 60 / 24"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, datThis, UserInfo.����, mbytType)
        
    With cboTimes
        .Clear
        If rsTmp.RecordCount = 0 Then .AddItem "�����ɿ�": .ItemData(.NewIndex) = 0
        For i = 1 To mrsTimes.RecordCount
            .AddItem "��" & i & "�νɿ�": .ItemData(.NewIndex) = i
        Next
    End With
    Call zlControl.CboSetIndex(cboTimes.hWnd, 0)
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpDate_Change()
    If dtpDate.Tag = "SelfChange" Then Exit Sub
    
    Call ValidateDate
    
    Call LoadTimes(dtpDate.Value)
    Call SetTimeRange
End Sub

Private Sub ValidateDate()
    Dim rsTmp As ADODB.Recordset, blnDo As Boolean
    Dim strSQL As String
        
    On Error GoTo errH
    
    '���������ϰ�ʱ���Ƿ��Ѱ������Ѵ��ڵĽɿ�ʱ�����
    '�����һ���νɿ�,����ʱ�ټ��
    '����:�����Ѵ�������[�շ�����¼],��ʱ����2006-12-11��2006-12-13֮������ڶ��ǲ������,�Զ���Ϊ2006-12-10��2006-12-14
    '    ����    �տ�Ա  ����    ��ʼʱ��    ��ֹʱ��
    '1   2006-12-10  ������  3   2006-12-10  2006-12-13 11:59:59
    '2   2006-12-14  ������  3   2006-12-13 12:00:01 2006-12-14 08:00:00
    strSQL = "Select ���� From �շ�����¼ Where ����<>[1] And [1] Between ��ʼʱ�� And ��ֹʱ�� And �տ�Ա=[2] And ����=[3]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.����, mbytType)
    blnDo = rsTmp.RecordCount > 0
    If Not blnDo Then
        '����֮ǰ�������ֹʱ���뵱��֮�����һ�νɿ����ڵ���С��ʼʱ��֮���޼��ʱ��������ѡ��
        strSQL = "Select 1" & vbNewLine & _
                "From (Select Min(��ʼʱ��) ��ʼʱ��" & vbNewLine & _
                "       From �շ�����¼" & vbNewLine & _
                "       Where ���� > [1] And �տ�Ա = [2] And ���� = [3]) A," & vbNewLine & _
                "     (Select Max(��ֹʱ��) ��ֹʱ��" & vbNewLine & _
                "       From �շ�����¼" & vbNewLine & _
                "       Where ���� < [1] And �տ�Ա = [2] And ���� = [3]) B" & vbNewLine & _
                "Where A.��ʼʱ�� = B.��ֹʱ�� + 1 / 60 / 60 / 24"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, dtpDate.Value, UserInfo.����, mbytType)
        blnDo = rsTmp.RecordCount > 0
    End If
    
    If blnDo Then
        MsgBox "�ϰ�ʱ��:" & Format(dtpDate.Value, "YYYY-MM-DD") & "�Ѱ�����" & Format(rsTmp!����, "YYYY-MM-DD") & "�Ľɿ�ʱ�䷶Χ��!", vbInformation, gstrSysName
        dtpDate.Tag = "SelfChange"
        dtpDate.Value = rsTmp!����
        dtpDate.Tag = ""
        If dtpDate.Enabled And dtpDate.Visible Then dtpDate.SetFocus
        '����:38829
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpEnd_Change()
    cmdSave.Enabled = True
    lblMessage.Caption = ""
    If dtpEnd.Value < dtpBegin.Value Then
        lblMessage.Top = Label4.Top
        lblMessage.Caption = "���ܱȿ�ʼʱ��С!"
        cmdSave.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    dtpDate.Value = Int(zldatabase.Currentdate)
    dtpDate.MaxDate = Int(dtpDate.Value)
    Call dtpDate_Change
End Sub

