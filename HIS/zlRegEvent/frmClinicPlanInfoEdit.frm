VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanInfoEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�༭"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   Icon            =   "frmClinicPlanInfoEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picTemplet 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   90
      ScaleHeight     =   1755
      ScaleWidth      =   5295
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   5295
      Begin VB.Frame fraTempletType 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   750
         TabIndex        =   21
         Top             =   60
         Width           =   1875
         Begin VB.OptionButton optTempletType 
            Caption         =   "���Ű�"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optTempletType 
            Caption         =   "���Ű�"
            Height          =   180
            Index           =   1
            Left            =   930
            TabIndex        =   22
            Top             =   0
            Width           =   885
         End
      End
      Begin VB.CheckBox chkTempletByDay 
         Caption         =   "���찲�ų���"
         Height          =   180
         Left            =   2640
         TabIndex        =   4
         Top             =   60
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txt��ע 
         Height          =   1050
         Left            =   750
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   4485
      End
      Begin VB.ComboBox cbo���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3660
         TabIndex        =   9
         Text            =   "cbo����"
         Top             =   330
         Width           =   1665
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Caption         =   "��������"
         Height          =   180
         Index           =   2
         Left            =   2610
         TabIndex        =   8
         Top             =   390
         Width           =   1065
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Caption         =   "������"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   390
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optӦ�÷�Χ 
         Caption         =   "ȫԺ"
         Height          =   180
         Index           =   0
         Left            =   750
         TabIndex        =   6
         Top             =   390
         Width           =   705
      End
      Begin VB.Label lblTempletType 
         AutoSize        =   -1  'True
         Caption         =   "ģ������"
         Height          =   180
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lbl��ע 
         AutoSize        =   -1  'True
         Caption         =   "��ע"
         Height          =   180
         Left            =   360
         TabIndex        =   10
         Top             =   690
         Width           =   360
      End
      Begin VB.Label lblӦ�÷�Χ 
         AutoSize        =   -1  'True
         Caption         =   "Ӧ�÷�Χ"
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.Frame fraSplitY 
      Height          =   25
      Left            =   -30
      TabIndex        =   20
      Top             =   2340
      Width           =   5730
   End
   Begin VB.Frame fraSplitX 
      Height          =   1875
      Left            =   3090
      TabIndex        =   19
      Top             =   -120
      Width           =   25
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   870
      MaxLength       =   50
      TabIndex        =   1
      Top             =   180
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   330
      Left            =   4200
      TabIndex        =   18
      Top             =   2580
      Width           =   915
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   330
      Left            =   2970
      TabIndex        =   17
      Top             =   2580
      Width           =   915
   End
   Begin VB.PictureBox picFixedRule 
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   90
      ScaleHeight     =   1065
      ScaleWidth      =   2985
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   540
      Width           =   2985
      Begin MSComCtl2.DTPicker dtpEndTime 
         Height          =   300
         Left            =   780
         TabIndex        =   16
         Top             =   630
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy/MM/dd HH:mm:ss"
         Format          =   171180035
         CurrentDate     =   42340
      End
      Begin MSComCtl2.DTPicker dtpStartTime 
         Height          =   300
         Left            =   780
         TabIndex        =   14
         Top             =   150
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy/MM/dd HH:mm:ss"
         Format          =   171180035
         CurrentDate     =   42340
      End
      Begin VB.Label lblEndTime 
         AutoSize        =   -1  'True
         Caption         =   "��ֹʱ��"
         Height          =   180
         Left            =   30
         TabIndex        =   15
         Top             =   690
         Width           =   720
      End
      Begin VB.Label lblStartTime 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   30
         TabIndex        =   13
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "ģ������"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmClinicPlanInfoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytFun As Byte '1-ģ�壬2-�̶�����
Private mlngModule As Long
Private mblnOK As Boolean
Private mobj���ﰲ�� As ���ﰲ��
Private mrsDepts As ADODB.Recordset '��Ա��������
Private mblnSaveAsTemplet As Boolean '�Ƿ�������Ϊģ�壬��ΪTrue���������޸�ģ������
Private mblnUpdate As Boolean

Public Function ShowMe(frmParent As Form, ByVal lngModule As Long, ByVal bytFun As Byte, _
    Optional ByRef obj���ﰲ�� As ���ﰲ��, Optional ByVal blnUpdate As Boolean, _
    Optional ByVal blnSaveAsTemplet As Boolean)
    '�������
    mbytFun = bytFun: Set mobj���ﰲ�� = obj���ﰲ��
    mlngModule = lngModule
    If mobj���ﰲ�� Is Nothing Then Set mobj���ﰲ�� = New ���ﰲ��
    mblnSaveAsTemplet = blnSaveAsTemplet: mblnUpdate = blnUpdate
    
    On Error Resume Next
    mblnOK = False
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Err = 0: On Error GoTo errHandle
    If KeyAscii = 13 Then
        If cbo����.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If cbo����.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If Select����(Me, mlngModule, mrsDepts, cbo����, cbo����.Text) = True Then
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If cbo����.Enabled Then cbo����.SetFocus
        zlControl.TxtSelAll cbo����
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, blnDoCheck As Boolean
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Sub
    If zlControl.TxtCheckInput(txtName, lblName.Caption, 50) = False Then Exit Sub
    If mbytFun = 1 Then
        If zlControl.TxtCheckInput(txt��ע, lbl��ע.Caption, 100, True) = False Then Exit Sub
        If optӦ�÷�Χ(2).Value And cbo����.ListIndex = -1 Then
            MsgBox "�������Ҳ���Ϊ�գ�", vbInformation, gstrSysName
            If cbo����.Visible And cbo����.Enabled Then cbo����.SetFocus
            Exit Sub
        End If
    ElseIf dtpStartTime.Enabled Then
        If dtpEndTime.Value <= dtpStartTime.Value Then
            MsgBox "��ֹʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            If dtpEndTime.Visible And dtpEndTime.Enabled Then dtpEndTime.SetFocus
            Exit Sub
        End If
        If dtpStartTime.Value < Now Then
            MsgBox "��ʼʱ�䲻��С�ڵ�ǰʱ�䣡", vbInformation, gstrSysName
            If dtpStartTime.Visible And dtpStartTime.Enabled Then dtpStartTime.SetFocus
            Exit Sub
        End If
        
        If mobj���ﰲ��.��ʼʱ�� = "" Then
            blnDoCheck = True
        Else
            If DateDiff("s", mobj���ﰲ��.��ʼʱ��, dtpStartTime.Value) <> 0 Then blnDoCheck = True
        End If
        If blnDoCheck Then
'            strSQL = "Select Max(a.��ʼʱ��) As ��ʼʱ��" & vbNewLine & _
'                    " From �ٴ����ﰲ�� A, �ٴ������ B" & vbNewLine & _
'                    " Where a.����id = b.Id And b.�Ű෽ʽ = 0 And a.��ʼʱ�� > [1] And b.����ʱ�� Is Not Null"
'            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������Ϣ", dtpStartTime.Value)
'            If Not rsTemp.EOF Then
'                If Nvl(rsTemp!��ʼʱ��) <> "" Then
'                    MsgBox "��ǰ��ʼʱ�䲻��С����һ���ѷ����Ĺ̶����ŵĿ�ʼʱ��(" & Nvl(rsTemp!��ʼʱ��, "yyyy-mm-dd hh:mm:ss") & ")��", vbInformation, gstrSysName
'                    Exit Sub
'                End If
'            End If
        End If
    End If
    
    If mobj���ﰲ��.������� <> Trim(txtName.Text) Then
        strSQL = "Select 1 From �ٴ������ Where ������� = [1] And �Ű෽ʽ = [2] And Nvl(վ��,'-') = Nvl([3],'-') And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�������Ϣ", Trim(txtName.Text), IIf(mbytFun = 1, 3, 0), gstrNodeNo)
        If Not rsTemp.EOF Then
            MsgBox "��ǰ�Ѵ�����Ϊ��" & Trim(txtName.Text) & "����" & IIf(mbytFun = 1, "ģ�壡", "�����"), vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    With mobj���ﰲ��
        .������� = Trim(txtName.Text)
        If mbytFun = 1 Then
            'ģ�����ͣ�0-���Ű�ģ�壬1-���ǰ����Ű�����Ű�ģ�壬2-�����Ű�����Ű�ģ��
            .ģ������ = IIf(optTempletType(1).Value, IIf(chkTempletByDay.Value = vbChecked, 2, 1), 0)
            'Ӧ�÷�Χ��0-����;1-��Ա��������(ָ������);2-ȫԺͨ��
            .Ӧ�÷�Χ = Choose(GetSelectedIndex(optӦ�÷�Χ) + 1, 2, 0, 1)
            If .Ӧ�÷�Χ = 1 Then  '��������
                .����ID = cbo����.ItemData(cbo����.ListIndex)
                .�������� = cbo����.Text
            End If
            .��ע = Trim(txt��ע.Text)
        Else
            mobj���ﰲ��.��ʼʱ�� = Format(dtpStartTime.Value, "yyyy-mm-dd hh:mm:ss")
            mobj���ﰲ��.��ֹʱ�� = Format(dtpEndTime.Value, "yyyy-mm-dd hh:mm:ss")
        End If
    End With
    mblnOK = True: Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub dtpEndTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStartTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStartTime_Validate(Cancel As Boolean)
    If dtpEndTime.Value < dtpStartTime.Value Then
        dtpEndTime.Value = dtpStartTime.Value
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim intӦ�÷�Χ As Integer, blnϵͳ���� As Boolean
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = 1 Then
        strSQL = "Select b.ID,b.����,b.����,b.����" & vbNewLine & _
                " From ������Ա A, ���ű� B" & vbNewLine & _
                " Where a.����ID=b.ID And a.��ԱID=[1]" & vbNewLine & _
                "       And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
                "       And (b.վ��='" & gstrNodeNo & "' Or b.վ�� is Null)"
        Set mrsDepts = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.id)
        cbo����.Clear
        If Not mrsDepts Is Nothing Then
            Do While Not mrsDepts.EOF
                cbo����.AddItem Nvl(mrsDepts!����)
                cbo����.ItemData(cbo����.NewIndex) = Nvl(mrsDepts!id)
                mrsDepts.MoveNext
            Loop
        End If
    End If
    
    With mobj���ﰲ��
        txtName.Text = .�������
        If mbytFun = 1 Then
            'ģ�����ͣ�0-���Ű�ģ�壬1-���ǰ����Ű�����Ű�ģ�壬2-�����Ű�����Ű�ģ��
            If .ģ������ = 0 Then '���Ű�
                optTempletType(0).Value = True
            Else '���Ű�
                optTempletType(1).Value = True
                chkTempletByDay.Value = IIf(.ģ������ = 2, vbChecked, vbUnchecked)
            End If
            intӦ�÷�Χ = .Ӧ�÷�Χ '0-����;1-��������;2-ȫԺͨ��
            If intӦ�÷�Χ = 1 Then
                optӦ�÷�Χ(2).Value = True
                zlControl.CboLocate cbo����, .����ID, True
            Else
                optӦ�÷�Χ(IIf(intӦ�÷�Χ = 0, 1, 0)).Value = True
            End If
            txt��ע.Text = .��ע
            
            If mblnSaveAsTemplet Or mblnUpdate Then
                optTempletType(0).Enabled = False
                optTempletType(1).Enabled = False
                chkTempletByDay.Enabled = False
            Else
                optTempletType(0).Enabled = True
                optTempletType(1).Enabled = True
                chkTempletByDay.Enabled = True
            End If
        Else
            blnϵͳ���� = (.�Ű෽ʽ = 0 And .��ע = "ϵͳ����")
            
            dtpStartTime.Value = Format(IIf(mobj���ﰲ��.��ʼʱ�� = "", Format(Now + 1, "yyyy-mm-dd 00:00:00"), mobj���ﰲ��.��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
            dtpEndTime.Value = Format(IIf(mobj���ﰲ��.��ֹʱ�� = "", "3000-01-01 00:00:00", mobj���ﰲ��.��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
        End If
    End With
    
    picTemplet.Visible = False
    picFixedRule.Visible = False
    fraSplitX.Visible = False
    fraSplitY.Visible = False
    If mbytFun = 1 Then 'ģ��
        picTemplet.Visible = True
        picTemplet.Left = 90
        picTemplet.Top = 540
        fraSplitY.Visible = True
        cmdOk.Left = 2970
        cmdOk.Top = 2450
        cmdCancel.Left = cmdOk.Left + cmdOk.Width + 300
        cmdCancel.Top = cmdOk.Top
        Me.Width = 5560
        Me.Height = 3370
        Me.Caption = IIf(mobj���ﰲ��.����ID = 0, "����ģ��", "����ģ��")
        lblName.Caption = "ģ������"
    Else '�̶�����
        picFixedRule.Visible = True
        picFixedRule.Left = 90
        picFixedRule.Top = 540
        fraSplitX.Visible = True
        cmdOk.Left = 3240
        cmdOk.Top = txtName.Top
        cmdCancel.Left = cmdOk.Left
        cmdCancel.Top = cmdOk.Top + cmdOk.Height + 100
        Me.Width = 4360
        Me.Height = 2110
        Me.Caption = IIf(mobj���ﰲ��.����ID = 0, "���ӹ̶�����", "��������")
        lblName.Caption = "��������"
        If blnϵͳ���� Then
            dtpStartTime.Enabled = False
            dtpEndTime.Enabled = False
        End If
    End If
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsDepts = Nothing
End Sub

Private Sub optTempletType_Click(index As Integer)
    chkTempletByDay.Visible = index = 1
End Sub

Private Sub optӦ�÷�Χ_Click(index As Integer)
    cbo����.Enabled = index = 2
    If index <> 2 Then
        cbo����.ListIndex = -1
    End If
End Sub

Private Sub optӦ�÷�Χ_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtName_GotFocus()
    zlControl.TxtSelAll txtName
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt��ע_GotFocus()
    zlControl.TxtSelAll txt��ע
End Sub

Private Sub txt��ע_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
