VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicPlanStopVisitFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdAudit 
      Height          =   240
      Left            =   2670
      Picture         =   "frmClinicPlanStopVisitFilter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����Ŀ(F4)"
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdApply 
      Height          =   240
      Left            =   2670
      Picture         =   "frmClinicPlanStopVisitFilter.frx":00F6
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����Ŀ(F4)"
      Top             =   330
      Width           =   255
   End
   Begin VB.TextBox txtAudit 
      Height          =   300
      Left            =   885
      MaxLength       =   8
      TabIndex        =   4
      Top             =   810
      Width           =   2070
   End
   Begin VB.CheckBox chkAudit 
      Caption         =   "������"
      Height          =   255
      Index           =   1
      Left            =   3300
      TabIndex        =   5
      Top             =   833
      Value           =   1  'Checked
      Width           =   885
   End
   Begin VB.CheckBox chkAudit 
      Caption         =   "δ����"
      Height          =   255
      Index           =   0
      Left            =   3300
      TabIndex        =   2
      Top             =   323
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.Frame frmSplit 
      Height          =   25
      Left            =   -270
      TabIndex        =   11
      Top             =   2280
      Width           =   6525
   End
   Begin VB.CheckBox chkShowInvalid 
      Caption         =   "��ʾ��ʧЧͣ�ﰲ��"
      Height          =   255
      Left            =   885
      TabIndex        =   6
      Top             =   1320
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4410
      TabIndex        =   13
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3150
      TabIndex        =   12
      Top             =   2520
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpStopBegin 
      Height          =   300
      Left            =   885
      TabIndex        =   8
      Top             =   1755
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   159252483
      CurrentDate     =   36588
   End
   Begin MSComCtl2.DTPicker dtpStopEnd 
      Height          =   300
      Left            =   3300
      TabIndex        =   10
      Top             =   1755
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      CalendarTitleBackColor=   -2147483647
      CalendarTitleForeColor=   -2147483634
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   159252483
      CurrentDate     =   36588
   End
   Begin VB.TextBox txtApply 
      Height          =   300
      Left            =   885
      MaxLength       =   8
      TabIndex        =   1
      Top             =   300
      Width           =   2070
   End
   Begin VB.Label lblStopTimeRange 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ͣ��ʱ��"
      Height          =   180
      Left            =   90
      TabIndex        =   7
      Top             =   1815
      Width           =   720
   End
   Begin VB.Label lblApply 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
   Begin VB.Label lblAudit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   870
      Width           =   540
   End
   Begin VB.Label lbl���ʱ�䷶Χ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   3000
      TabIndex        =   9
      Top             =   1815
      Width           =   180
   End
End
Attribute VB_Name = "frmClinicPlanStopVisitFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrFilter As String
Public mblnOK As Boolean
Private mrsApply As ADODB.Recordset
Private mrsAudit As ADODB.Recordset

Private Sub chkAudit_Click(index As Integer)
    If chkAudit(0).Value = vbUnchecked And chkAudit(1).Value = vbUnchecked Then
        chkAudit(IIf(index = 0, 1, 0)).Value = vbChecked
    End If
    txtAudit.Enabled = chkAudit(1).Value = vbChecked
    txtAudit.BackColor = IIf(txtAudit.Enabled, vbWindowBackground, vbButtonFace)
    cmdAudit.Enabled = chkAudit(1).Value = vbChecked
End Sub

Private Sub chkAudit_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkShowInvalid_Click()
    dtpStopBegin.Enabled = chkShowInvalid.Value = vbChecked
    dtpStopEnd.Enabled = chkShowInvalid.Value = vbChecked
End Sub

Private Sub chkShowInvalid_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    If txtApply.Visible And txtApply.Enabled Then txtApply.SetFocus
    Me.Hide '���ش���
End Sub

Private Sub cmdOK_Click()
    Err = 0: On Error GoTo ErrHandler
    If zlControl.FormCheckInput(Me) = False Then Exit Sub
    If chkShowInvalid.Value = vbChecked Then
        If DateDiff("s", dtpStopBegin.Value, dtpStopEnd.Value) <= 0 Then
            MsgBox "ͣ��ʱ��Ľ���ʱ�������ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            If dtpStopBegin.Visible And dtpStopBegin.Enabled Then dtpStopBegin.SetFocus
            Exit Sub
        End If
    End If
    
    Call MakeFilter
    mblnOK = True
    
    If txtApply.Visible And txtApply.Enabled Then txtApply.SetFocus
    Me.Hide '���ش���
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpStopBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpStopEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer, lngOldID As Long, strListFeeItem As String
    Dim Curdate As Date, index As Integer, arrItem As Variant
    
    Err = 0: On Error GoTo ErrHandler
    mblnOK = False
    Curdate = zlDatabase.Currentdate
    
    dtpStopBegin.Value = Format(DateAdd("m", -1, Curdate), "yyyy-MM-dd 00:00:00")
    dtpStopEnd.Value = Format(DateAdd("m", 1, Curdate), "yyyy-MM-dd 23:59:59")
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub MakeFilter()
    Err = 0: On Error GoTo ErrHandler
    
    mstrFilter = ""
    If Trim(txtApply.Text) <> "" Then mstrFilter = mstrFilter & " And ������=[1]"
    If Trim(txtAudit.Text) <> "" And chkAudit(1).Value = vbChecked Then mstrFilter = mstrFilter & " And ������=[2] "
    
    If chkAudit(0).Value = vbChecked And chkAudit(1).Value = vbChecked Then
    ElseIf chkAudit(0).Value = vbChecked Then '��δ����
        mstrFilter = mstrFilter & " And (������ Is Null Or ȡ���� Is Not Null)"
    Else '��������
        mstrFilter = mstrFilter & " And ������ Is Not Null And ȡ���� Is Null"
    End If
    
    If chkShowInvalid.Value = vbChecked Then
        mstrFilter = mstrFilter & " And ��ʼʱ��<=[4] And Nvl(ʧЧʱ��,��ֹʱ��)>=[3]"
    Else
        mstrFilter = mstrFilter & " And Nvl(ʧЧʱ��,��ֹʱ��)>sysdate"
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsApply Is Nothing Then Set mrsApply = Nothing
    If Not mrsAudit Is Nothing Then Set mrsAudit = Nothing
End Sub

Private Sub txtApply_GotFocus()
    zlControl.TxtSelAll txtApply
End Sub

Private Sub txtApply_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtApply_Validate(Cancel As Boolean)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    Err = 0: On Error GoTo ErrHandler
    If Trim(txtApply.Text) = "" Then Exit Sub
    
    strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & vbNewLine & _
            " From ��Ա�� A,��Ա����˵�� B" & vbNewLine & _
            " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & vbNewLine & _
            "       And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & vbNewLine & _
            "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "       And (A.��� Like [1] Or Upper(A.����) Like [2] Or A.���� Like [2] )" & vbNewLine & _
            " Order by A.���"
    
    vRect = zlControl.GetControlRect(txtApply.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtApply.Height, blnCancel, False, True, Trim(txtApply.Text) & "%", gstrLike & UCase(Trim(txtApply.Text)) & "%")
    If rsTemp Is Nothing Then Cancel = True: Exit Sub
    If rsTemp.EOF Then Cancel = True: Exit Sub
    If blnCancel Then Cancel = True: Exit Sub
    
    txtApply.Text = rsTemp!����
    If chkAudit(0).Visible And chkAudit(0).Enabled Then chkAudit(0).SetFocus
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdApply_Click()
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strSQL As String, lngDeptID As Long
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & vbNewLine & _
            " From ��Ա�� A,��Ա����˵�� B" & vbNewLine & _
            " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & vbNewLine & _
            "       And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & vbNewLine & _
            "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.���"
    
    vRect = zlControl.GetControlRect(txtApply.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtApply.Height, blnCancel, False, True)
    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.EOF Then Exit Sub
    If blnCancel Then Exit Sub
    
    txtApply.Text = rsTemp!����
    If chkAudit(0).Visible And chkAudit(0).Enabled Then chkAudit(0).SetFocus
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtAudit_GotFocus()
    zlControl.TxtSelAll txtAudit
End Sub

Private Sub txtAudit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtAudit_Validate(Cancel As Boolean)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    Err = 0: On Error GoTo ErrHandler
    If Trim(txtAudit.Text) = "" Then Exit Sub
    
    strSQL = "Select Distinct d.ID, d.���, d.����, d.����" & vbNewLine & _
            " From ��Ա�� D,�ϻ���Ա�� C, Zluserroles B, zlRoleGrant A" & vbNewLine & _
            " Where A.��ɫ = B.��ɫ And B.�û� = C.�û��� And C.��Աid = D.ID" & vbNewLine & _
            "       And A.ϵͳ = 100 And A.��� = 1114 And A.���� = 'ͣ������'" & vbNewLine & _
            "       And (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null) " & vbNewLine & _
            "       And (d.վ��='" & gstrNodeNo & "' Or d.վ�� is Null)" & vbNewLine & _
            "       And (d.��� Like [1] Or Upper(d.����) Like [2] Or d.���� Like [2] )" & vbNewLine & _
            " Order by d.���"

    vRect = zlControl.GetControlRect(txtAudit.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtAudit.Height, blnCancel, False, True, Trim(txtAudit.Text) & "%", gstrLike & UCase(Trim(txtAudit.Text)) & "%")
    If rsTemp Is Nothing Then Cancel = True: Exit Sub
    If rsTemp.EOF Then Cancel = True: Exit Sub
    If blnCancel Then Cancel = True: Exit Sub
    
    txtAudit.Text = rsTemp!����
    If chkAudit(1).Visible And chkAudit(1).Enabled Then chkAudit(1).SetFocus
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdAudit_Click()
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strSQL As String, lngDeptID As Long
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select Distinct d.ID, d.���, d.����, d.����" & vbNewLine & _
            " From ��Ա�� D,�ϻ���Ա�� C, Zluserroles B, zlRoleGrant A" & vbNewLine & _
            " Where A.��ɫ = B.��ɫ And B.�û� = C.�û��� And C.��Աid = D.ID" & vbNewLine & _
            "       And A.ϵͳ = 100 And A.��� = 1114 And A.���� = 'ͣ������'" & vbNewLine & _
            "       And (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null) " & vbNewLine & _
            "       And (d.վ��='" & gstrNodeNo & "' Or d.վ�� is Null)" & vbNewLine & _
            " Order by d.���"
    
    vRect = zlControl.GetControlRect(txtAudit.hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txtAudit.Height, blnCancel, False, True)
    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.EOF Then Exit Sub
    If blnCancel Then Exit Sub
    
    txtAudit.Text = rsTemp!����
    If chkAudit(1).Visible And chkAudit(1).Enabled Then chkAudit(1).SetFocus
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

