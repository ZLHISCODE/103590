VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmEPRSearchTerms 
   BorderStyle     =   0  'None
   Caption         =   "������������"
   ClientHeight    =   7890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picContent 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   1080
      ScaleHeight     =   600
      ScaleWidth      =   3000
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3135
      Width           =   3000
      Begin VB.TextBox txtContent 
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   300
         Width           =   2370
      End
      Begin VB.ComboBox cboCompend 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   0
         Width           =   2385
      End
   End
   Begin VB.PictureBox picDept 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1710
      Left            =   1080
      ScaleHeight     =   1710
      ScaleWidth      =   3000
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   4800
      Width           =   3000
      Begin VB.ListBox lstDept 
         Enabled         =   0   'False
         Height          =   1320
         Left            =   450
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   390
         Width           =   2055
      End
      Begin VB.OptionButton optDept 
         Caption         =   "ָ������(&P)"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   195
         Width           =   1485
      End
      Begin VB.OptionButton optDept 
         Caption         =   "���п���(&A)"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   0
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.PictureBox picSearch 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   1080
      ScaleHeight     =   1335
      ScaleWidth      =   3000
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6525
      Width           =   3000
      Begin VB.TextBox txtElement 
         Height          =   900
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   0
         Width           =   2355
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   1740
         TabIndex        =   24
         Top             =   945
         Width           =   1100
      End
   End
   Begin VB.PictureBox picKind 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   1080
      ScaleHeight     =   1035
      ScaleWidth      =   3000
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3750
      Width           =   3000
      Begin VB.CheckBox chkKind 
         Caption         =   "���ﲡ��(&1)"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Tag             =   "1"
         Top             =   0
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.CheckBox chkKind 
         Caption         =   "סԺ����(&2)"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Tag             =   "2"
         Top             =   210
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.CheckBox chkKind 
         Caption         =   "������(&3)"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Tag             =   "4"
         Top             =   420
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.CheckBox chkKind 
         Caption         =   "����֤������(&4)"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Tag             =   "5"
         Top             =   630
         Width           =   1830
      End
      Begin VB.CheckBox chkKind 
         Caption         =   "֪���ļ�(&5)"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Tag             =   "6"
         Top             =   840
         Width           =   1830
      End
   End
   Begin VB.PictureBox picPati 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2145
      Left            =   1080
      ScaleHeight     =   2145
      ScaleWidth      =   3000
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   765
      Width           =   3000
      Begin VB.TextBox txtPati 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1365
         TabIndex        =   35
         Top             =   375
         Width           =   1320
      End
      Begin VB.OptionButton optPati 
         Caption         =   "ҽ����(&Y)"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   420
         Width           =   1125
      End
      Begin VB.TextBox txtPati 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1365
         TabIndex        =   11
         ToolTipText     =   "��ָ��ʱ��ʱ������ģ����������"
         Top             =   1425
         Width           =   1320
      End
      Begin VB.OptionButton optPati 
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   1470
         Width           =   1125
      End
      Begin VB.CheckBox chkSex 
         Caption         =   "Ů(&W)"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Top             =   1755
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkSex 
         Caption         =   "��(&M)"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   1365
         TabIndex        =   12
         Top             =   1755
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.OptionButton optPati 
         Caption         =   "סԺ��(&I)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1125
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.TextBox txtPati 
         Height          =   300
         Index           =   3
         Left            =   1365
         TabIndex        =   9
         Top             =   1080
         Width           =   1320
      End
      Begin VB.OptionButton optPati 
         Caption         =   "�����(&O)"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   1125
      End
      Begin VB.TextBox txtPati 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1365
         TabIndex        =   7
         Top             =   735
         Width           =   1320
      End
      Begin VB.TextBox txtPati 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1365
         TabIndex        =   5
         Top             =   30
         Width           =   1320
      End
      Begin VB.OptionButton optPati 
         Caption         =   "���￨(&A)"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   75
         Width           =   1125
      End
   End
   Begin VB.PictureBox picDate 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   1080
      ScaleHeight     =   600
      ScaleWidth      =   3000
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   3000
      Begin VB.CheckBox chkDtp 
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   180
         Value           =   1  'Checked
         Width           =   195
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   300
         Left            =   705
         TabIndex        =   3
         Top             =   300
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   112590851
         CurrentDate     =   38683
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   300
         Left            =   705
         TabIndex        =   1
         Top             =   0
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   112590851
         CurrentDate     =   38683
      End
      Begin VB.Label lblDateTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   495
         TabIndex        =   2
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblDateFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   510
         TabIndex        =   0
         Top             =   60
         Width           =   180
      End
   End
   Begin XtremeSuiteControls.TaskPanel tplThis 
      Height          =   7185
      Left            =   15
      TabIndex        =   25
      Top             =   30
      Width           =   2085
      _Version        =   589884
      _ExtentX        =   3678
      _ExtentY        =   12674
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmEPRSearchTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum FT
    T����� = 0
    Tҽ���� = 1
    T����� = 2
    TסԺ�� = 3
    T���� = 4
End Enum
Public Event SearchClick(rsResult As ADODB.Recordset)   '��ť����¼�

Public mlngDeptId As Long           'ָ����Ĭ����д����id
Public mbytKind As Byte             'Ҫ����ҵ��ļ����ࣺ0-��ʾ�ٴ���д�Ĳ����������Ͳ����ļ�������ͬ
Public mlngFileID As Long         'ָ�����ҵĲ����ļ�id��0-δָ��; >0,ֻ�����ض��Ĳ����ļ���ͨ�����ڲ����༭�еĲ��Ҳ���;
Public mstrPrivs As String

'-----------------------------------------------------------------------------------------------------------------
Private Sub cboCompend_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboCompend_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkDtp_Click()
If chkDtp.Value = vbChecked Then
        dtpDateFrom.Enabled = True
        dtpDateTo.Enabled = True
    Else
        dtpDateFrom.Enabled = False
        dtpDateTo.Enabled = False
    End If
End Sub

Private Sub chkKind_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkKind_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkSex_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkSex_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdSearch_Click()
Dim strKinds As String, strDepts As String, strTemp As String, blnMoved As Boolean
Dim strFromDate As String, strToDate As String, blnIn As Boolean, blnSpecify As Boolean
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    If Trim(txtPati(Decode(True, optPati(T�����).Value, T�����, optPati(Tҽ����), Tҽ����, optPati(T�����), T�����, optPati(TסԺ��), TסԺ��, optPati(T����), T����, 5)).Text) = "" Then
'    If Trim(txtPati(IIf(optPati(0).Value, 0, IIf(optPati(1).Value, 1, IIf(optPati(2).Value, 2, IIf(optPati(3).Value, 3, IIf(optPati(4).Value, 4, 5)))))).Text) = "" Then
        chkDtp.Value = vbChecked
    End If
    
    If chkDtp.Value <> vbChecked Then
        blnSpecify = True
    End If
    '----------------------------------------
    '�����������
    strKinds = ""
    
    If mlngFileID = 0 Then
        If mbytKind = 0 Then
            For lngCount = 0 To Me.chkKind.Count - 1
                If Me.chkKind(lngCount).Value = 1 Then strKinds = strKinds & "," & Me.chkKind(lngCount).Tag
            Next
            If strKinds = "" Then MsgBox "û��ѡ���κ��ļ����࣬�޷�������", vbExclamation, gstrSysName: Exit Sub
            strKinds = Mid(strKinds, 2)
        Else
            strKinds = CStr(mbytKind)
        End If
    Else
        strKinds = "0"      'ָ���ļ�ʱ���ⰴ����ɸѡ��������Ч
    End If
    
    strDepts = ""
    If Me.optDept(1).Value Then
        If Me.lstDept.SelCount = 0 Then MsgBox "û��ѡ���κβ�����д���ң��޷�������", vbExclamation, gstrSysName: Exit Sub
        For lngCount = 0 To Me.lstDept.ListCount - 1
            If Me.lstDept.Selected(lngCount) Then strDepts = strDepts & "," & Me.lstDept.ItemData(lngCount)
        Next
        strDepts = Mid(strDepts, 2)
    End If
    blnIn = optPati(TסԺ��).Value
    '----------------------------------------
    '��ѯ�����֯
    If Trim(Me.txtContent.Text) = "" Then
        gstrSQL = "Select l.Id, l.����id, l.��ҳid, l.������Դ, i.����, i.�Ա�,i.����,decode(l.������Դ,2,l.��ҳid,'') סԺ����,l.��������, l.��������, d.���� As ����, l.������, l.���ʱ��,0 as ����ת��,�༭��ʽ,��ӡ��,��ӡʱ�� " & _
                " From ���Ӳ�����¼ l," & IIf(blnIn, "������ҳ A,", "") & " ������Ϣ i, ���ű� d" & _
                " Where l.���ʱ�� " & IIf(blnSpecify, "Is Not Null ", "Between To_Date([1],'yyyy-mm-dd') And To_Date([2],'yyyy-mm-dd')+1-1/24/60/60 ") & IIf(blnIn, "And l.����ID=A.����ID and l.��ҳID=A.��ҳID ", "") & "And l.����id = i.����id And" & _
                "      (l.�������� In (" & strKinds & ") Or l.�ļ�id = " & mlngFileID & ") And l.����id = d.Id"
        If strDepts <> "" Then gstrSQL = gstrSQL & " And l.����id + 0 In (" & strDepts & ")"
        If Me.optPati(T�����).Value Then
            If Val(Me.txtPati(T�����).Text) <> 0 Then gstrSQL = gstrSQL & " And i.���￨��='" & Me.txtPati(T�����).Text & "'"
        ElseIf Me.optPati(Tҽ����).Value Then
            If Trim(Me.txtPati(Tҽ����).Text) <> "" Then gstrSQL = gstrSQL & " And i.ҽ����='" & Me.txtPati(Tҽ����).Text & "'"
        ElseIf Me.optPati(T�����).Value Then
            If Val(Me.txtPati(T�����).Text) <> 0 Then gstrSQL = gstrSQL & " And i.�����=" & Me.txtPati(T�����).Text
        ElseIf Me.optPati(TסԺ��).Value Then
            If Val(Me.txtPati(TסԺ��).Text) <> 0 Then gstrSQL = gstrSQL & " And A.סԺ��=" & Me.txtPati(TסԺ��).Text
        ElseIf Me.optPati(T����).Value Then
            If Trim(Me.txtPati(T����).Text) <> "" Then gstrSQL = gstrSQL & IIf(blnSpecify, " And i.����='" & Trim(Me.txtPati(T����).Text) & "'", " And i.���� like '%" & Trim(Me.txtPati(T����).Text) & "%'")
            If Me.chkSex(0).Value = vbChecked And Me.chkSex(1).Value = vbUnchecked Then gstrSQL = gstrSQL & " And i.�Ա� = '��'"
            If Me.chkSex(0).Value = vbUnchecked And Me.chkSex(1).Value = vbChecked Then gstrSQL = gstrSQL & " And i.�Ա� = 'Ů'"
        End If
    Else
        If Me.cboCompend.ListIndex = 0 Then
            gstrSQL = "Select Distinct l.Id, l.����id, l.��ҳid, l.������Դ, i.����, i.�Ա�,i.����,decode(l.������Դ,2,l.��ҳid,'') סԺ����, l.��������, l.��������, d.���� As ����, l.������, l.���ʱ��,0 as ����ת��,�༭��ʽ,l.��ӡ��,l.��ӡʱ��" & _
                    " From ���Ӳ�����¼ l," & IIf(blnIn, "������ҳ A,", "") & " ������Ϣ i, ���ű� d, ���Ӳ������� c" & _
                    " Where l.���ʱ�� " & IIf(blnSpecify, "Is Not Null ", "Between To_Date([1],'yyyy-mm-dd') And To_Date([2],'yyyy-mm-dd')+1-1/24/60/60 ") & IIf(blnIn, "And l.����ID=A.����ID and l.��ҳID=A.��ҳID ", "") & "And l.����id = i.����id And" & _
                    "      (l.�������� In (" & strKinds & ") Or l.�ļ�id = " & mlngFileID & ") And l.����id = d.Id And" & _
                    "      l.Id = c.�ļ�id And Nvl(c.��ֹ��, 0) = 0 And c.�����ı� Like '%" & Trim(Me.txtContent.Text) & "%'"
        ElseIf Me.cboCompend.ListIndex = 1 Then
            gstrSQL = "Select Distinct l.Id, l.����id, l.��ҳid, l.������Դ, i.����, i.�Ա�,i.����,decode(l.������Դ,2,l.��ҳid,'') סԺ����, l.��������, l.��������, d.���� As ����, l.������, l.���ʱ��,0 as ����ת��,�༭��ʽ,l.��ӡ��,l.��ӡʱ��" & _
                    " From ���Ӳ�����¼ l," & IIf(blnIn, "������ҳ A,", "") & " ������Ϣ i, ���ű� d, ���Ӳ������� c" & _
                    " Where l.���ʱ�� " & IIf(blnSpecify, "Is Not Null ", "Between To_Date([1],'yyyy-mm-dd') And To_Date([2],'yyyy-mm-dd')+1-1/24/60/60 ") & IIf(blnIn, "And l.����ID=A.����ID and l.��ҳID=A.��ҳID ", "") & "And l.����id = i.����id And" & _
                    "      (l.�������� In (" & strKinds & ") Or l.�ļ�id = " & mlngFileID & ") And l.����id = d.Id And" & _
                    "      l.Id = c.�ļ�id And  Nvl(c.��ֹ��, 0) = 0 And c.�����ı� Like '%" & Trim(Me.txtContent.Text) & "%' And c.�������� = 7"
        Else
            gstrSQL = "Select Distinct l.Id, l.����id, l.��ҳid, l.������Դ, i.����, i.�Ա�,i.����,decode(l.������Դ,2,l.��ҳid,'') סԺ����, l.��������, l.��������, d.���� As ����, l.������, l.���ʱ��,0 as ����ת��,�༭��ʽ,l.��ӡ��,l.��ӡʱ��" & _
                    " From ���Ӳ�����¼ l," & IIf(blnIn, "������ҳ A,", "") & " ������Ϣ i, ���ű� d, ���Ӳ������� c, ���Ӳ������� p" & _
                    " Where l.���ʱ�� " & IIf(blnSpecify, "Is Not Null ", "Between To_Date([1],'yyyy-mm-dd') And To_Date([2],'yyyy-mm-dd')+1-1/24/60/60 ") & IIf(blnIn, "And l.����ID=A.����ID and l.��ҳID=A.��ҳID ", "") & "And l.����id = i.����id And" & _
                    "      (l.�������� In (" & strKinds & ") Or l.�ļ�id = " & mlngFileID & ") And l.����id = d.Id And" & _
                    "      l.Id = c.�ļ�id And Nvl(c.��ֹ��, 0) = 0 And c.�����ı� Like '%" & Trim(Me.txtContent.Text) & "%' And" & _
                    "      c.��id = p.Id And p.Ԥ�����id + 0 =" & Me.cboCompend.ItemData(Me.cboCompend.ListIndex)
        End If
        If strDepts <> "" Then gstrSQL = gstrSQL & " And l.����id + 0 In (" & strDepts & ")"
        If Me.optPati(T�����).Value Then
            If Trim(Me.txtPati(T�����).Text) <> "" Then gstrSQL = gstrSQL & " And i.���￨��='" & Me.txtPati(T�����).Text & "'"
        ElseIf Me.optPati(Tҽ����).Value Then
            If Trim(Me.txtPati(Tҽ����).Text) <> "" Then gstrSQL = gstrSQL & " And i.ҽ����='" & Me.txtPati(Tҽ����).Text & "'"
        ElseIf Me.optPati(T�����).Value Then
            If Val(Me.txtPati(T�����).Text) <> 0 Then gstrSQL = gstrSQL & " And i.�����=" & Me.txtPati(T�����).Text
        ElseIf Me.optPati(TסԺ��).Value Then
            If Val(Me.txtPati(TסԺ��).Text) <> 0 Then gstrSQL = gstrSQL & " And A.סԺ��=" & Me.txtPati(TסԺ��).Text
        ElseIf Me.optPati(T����).Value Then
            If Trim(Me.txtPati(T����).Text) <> "" Then gstrSQL = gstrSQL & IIf(blnSpecify, " And i.����='" & Trim(Me.txtPati(T����).Text) & "'", " And i.���� like '%" & Trim(Me.txtPati(T����).Text) & "%'")
            If Me.chkSex(0).Value = vbChecked And Me.chkSex(1).Value = vbUnchecked Then gstrSQL = gstrSQL & " And i.�Ա� = '��'"
            If Me.chkSex(0).Value = vbUnchecked And Me.chkSex(1).Value = vbChecked Then gstrSQL = gstrSQL & " And i.�Ա� = 'Ů'"
        End If
    End If
    
    If Trim(Me.txtElement.Tag) <> "" Then
    
        Dim blnAnd As Boolean
        Dim strWhere As String, strDecodeName As String, strDecodeText As String
        Dim aryTerm() As String, aryField() As String
        Dim aryValue() As String, strValues As String, lngPoint As Long
        
        If Val(Left(Trim(Me.txtElement.Tag), 1)) = 1 Then blnAnd = True
        aryTerm = Split(Mid(Trim(Me.txtElement.Tag), 3), "|")
        
        strWhere = "": strDecodeName = ""
        For lngCount = 0 To UBound(aryTerm)
            aryField = Split(aryTerm(lngCount), ";")
            strWhere = strWhere & " Or c.Ҫ������ = '" & aryField(1) & "' And c.Ҫ������ = " & Val(aryField(2))
            If Val(aryField(2)) = 0 Then
                Select Case aryField(3)
                Case "����":    strDecodeText = "Decode(Zl_To_Number(c.�����ı�)," & Val(aryField(4)) & ",1,0)"
                Case "������":  strDecodeText = "Decode(Zl_To_Number(c.�����ı�)," & Val(aryField(4)) & ",0,1)"
                Case "����":    strDecodeText = "Decode(Sign(Zl_To_Number(c.�����ı�)-" & Val(aryField(4)) & "),1,1,0)"
                Case "С��":    strDecodeText = "Decode(Sign(Zl_To_Number(c.�����ı�)-" & Val(aryField(4)) & "),-1,1,0)"
                Case "����":    strDecodeText = "Decode(Sign(Zl_To_Number(c.�����ı�)-" & Val(aryField(4)) & "),1,0,1)"
                Case "����":    strDecodeText = "Decode(Sign(Zl_To_Number(c.�����ı�)-" & Val(aryField(4)) & "),-1,0,1)"
                Case "����"
                    aryValue = Split(Trim(aryField(4)), ",")
                    strDecodeText = "Decode(Sign(Zl_To_Number(c.�����ı�)-" & Val(aryValue(0)) & "),-1,0,Decode(Sign(Zl_To_Number(c.�����ı�)-" & Val(aryValue(1)) & "),1,0,1))"
                Case "����", "������"
                    aryValue = Split(Trim(aryField(4)), ",")
                    strValues = ""
                    For lngPoint = 0 To UBound(aryValue)
                        strValues = strValues & "," & Val(aryValue(lngPoint)) & "," & IIf(aryField(3) = "����", "1", "0")
                    Next
                    strValues = Mid(strValues, 2)
                    strDecodeText = "Decode(Zl_To_Number(c.�����ı�)," & strValues & "," & IIf(aryField(3) = "����", "0", "1") & ")"
                End Select
            Else
                Select Case aryField(3)
                Case "����":    strDecodeText = "Decode(Trim(c.�����ı�),'" & Trim(aryField(4)) & "',1,0)"
                Case "������":  strDecodeText = "Decode(Trim(c.�����ı�),'" & Trim(aryField(4)) & "',0,1)"
                Case "����":    strDecodeText = "Decode(Sign(Instr(c.�����ı�,'" & Trim(aryField(4)) & "')),1,1,0)"
                Case "������":  strDecodeText = "Decode(Sign(Instr(c.�����ı�,'" & Trim(aryField(4)) & "')),1,0,1)"
                Case "����", "������"
                    aryValue = Split(Trim(aryField(4)), ",")
                    strValues = ""
                    For lngPoint = 0 To UBound(aryValue)
                        strValues = strValues & ",'" & Trim(aryValue(lngPoint)) & "'," & IIf(aryField(3) = "����", "1", "0")
                    Next
                    strValues = Mid(strValues, 2)
                    strDecodeText = "Decode(Trim(c.�����ı�)," & strValues & "," & IIf(aryField(3) = "����", "0", "1") & ")"
                End Select
            End If
            strDecodeName = strDecodeName & "+Decode(c.Ҫ������, '" & aryField(1) & "'," & strDecodeText & ",0)"
            
        Next
        gstrSQL = "Select b.*" & _
                " From (" & gstrSQL & ") b," & vbCrLf & _
                "      (Select Id" & _
                "       From (Select l.Id," & Mid(strDecodeName, 2) & " As ������" & _
                "             From ���Ӳ�����¼ l, ���Ӳ������� c" & _
                "             Where l.���ʱ�� " & IIf(blnSpecify, "Is Not Null ", "Between To_Date([1],'yyyy-mm-dd') And To_Date([2],'yyyy-mm-dd')+1-1/24/60/60") & " And" & _
                "                   l.Id = c.�ļ�id And l.�������� In (" & strKinds & ") And Nvl(c.��ֹ��, 0) = 0 And (" & Mid(strWhere, 5) & "))" & _
                "       Group By Id" & _
                "       Having Sum(������)" & IIf(blnAnd, " = " & UBound(aryTerm) + 1, " > 0") & ") e" & _
                " Where b.id = e.Id"
    End If
    
    strFromDate = Format(Me.dtpDateFrom.Value, "yyyy-MM-dd")
    strToDate = Format(Me.dtpDateTo.Value, "yyyy-MM-dd")
    blnMoved = MovedByDate(strFromDate)
    If blnMoved Then
        strTemp = Replace(gstrSQL, "0 as ����ת��", "1 as ����ת��")
        strTemp = Replace(strTemp, "���Ӳ�����¼", "H���Ӳ�����¼")
        strTemp = Replace(strTemp, "���Ӳ�������", "H���Ӳ�������")
        gstrSQL = gstrSQL & " Union All " & strTemp
    End If
    gstrSQL = gstrSQL & " order by ���ʱ�� desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strFromDate, strToDate)

    RaiseEvent SearchClick(rsTemp)
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpDateFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpDateTo_Validate(Cancel As Boolean)
    Me.dtpDateFrom.MaxDate = Me.dtpDateTo.Value
    If Me.dtpDateFrom.Value > Me.dtpDateFrom.MaxDate Then Me.dtpDateFrom.Value = Me.dtpDateFrom.MaxDate
End Sub

Private Sub Form_Load()
Dim tplGroup As TaskPanelGroup
Dim tplItem As TaskPanelGroupItem
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    
    Err = 0: On Error GoTo errHand
    '-----------------------------------------------------
    '��������װ��:
    gstrSQL = "Select Sysdate From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With Me.dtpDateTo
        .Value = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd")
        .MaxDate = .Value: .MinDate = Format("1990-01-01", "yyyy-MM-dd")
    End With
    With Me.dtpDateFrom
        .Value = Me.dtpDateTo.Value - 7
        .MaxDate = Me.dtpDateTo.MaxDate: .MinDate = Me.dtpDateTo.MinDate
    End With
    
    If mlngFileID > 0 Then
        gstrSQL = "Select Distinct p.Id, p.�������, p.�����ı�" & vbNewLine & _
                "From �����ļ��ṹ p, �����ļ��ṹ d" & vbNewLine & _
                "Where p.Id = d.Ԥ�����id And p.�ļ�id Is Null And d.�ļ�id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    Else
        If mbytKind = cpr���Ʊ��� Then
            gstrSQL = "Select Id, �������, �����ı� From �����ļ��ṹ Where �ļ�id Is Null And Substr(ʹ��ʱ��, " & cpr���Ʊ��� & ", 1) <> '0'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        Else
            gstrSQL = "Select Id, �������, �����ı� From �����ļ��ṹ Where �ļ�id Is Null And Substr(ʹ��ʱ��, 1, " & cpr���Ʊ��� - 1 & ") <> '000000'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        End If
    End If
    With rsTemp
        Me.cboCompend.Clear
        Me.cboCompend.AddItem "<�κ����>": Me.cboCompend.ListIndex = 0
        Me.cboCompend.AddItem "<�������>"
        Do While Not .EOF
            Me.cboCompend.AddItem !������� & "-" & !�����ı�
            Me.cboCompend.ItemData(Me.cboCompend.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    If mbytKind = cpr���Ʊ��� Then
        gstrSQL = "Select Distinct a.Id, a.����, a.����" & vbNewLine & _
                "From ���ű� a, ��������˵�� b" & vbNewLine & _
                "Where b.����id = a.Id And b.�������� In ('����', '���')" & vbNewLine & _
                "Order By a.����"
    Else
        gstrSQL = "Select Distinct a.Id, a.����, a.����" & vbNewLine & _
                "From ���ű� a, ��������˵�� b" & vbNewLine & _
                "Where b.����id = a.Id And b.�������� In ('�ٴ�', '����')" & vbNewLine & _
                "Order By a.����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lstDept.Clear
        Do While Not .EOF
            Me.lstDept.AddItem !���� & "-" & !����
            Me.lstDept.ItemData(Me.lstDept.NewIndex) = !ID
            If !ID = mlngDeptId Then Me.lstDept.Selected(Me.lstDept.NewIndex) = True
            .MoveNext
        Loop
    End With
    
    '-----------------------------------------------------
    Set tplGroup = Me.tplThis.Groups.Add(0, "��������:"): tplGroup.Expandable = False
    Set tplItem = tplGroup.Items.Add(0, "��д���ڷ�Χ:", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picDate
    Me.picDate.BackColor = tplItem.BackColor
    Set tplItem = tplGroup.Items.Add(0, "������Ϣ(����Ϊ��������):", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picPati
    Me.picPati.BackColor = tplItem.BackColor
    For lngCount = 0 To Me.optPati.Count - 1: Me.optPati(lngCount).BackColor = tplItem.BackColor: Next
    For lngCount = 0 To Me.chkSex.Count - 1: Me.chkSex(lngCount).BackColor = tplItem.BackColor: Next
    Set tplItem = tplGroup.Items.Add(0, "��ټ������ݰ����ı�:", xtpTaskItemTypeText)
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picContent
    Me.picContent.BackColor = tplItem.BackColor
    
    If mlngFileID = 0 And mbytKind <> cpr���Ʊ��� Then
        Set tplGroup = Me.tplThis.Groups.Add(0, "�ļ�����:"): tplGroup.Expanded = False
        Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picKind
        Me.picKind.BackColor = tplItem.BackColor
        For lngCount = 0 To Me.chkKind.Count - 1: Me.chkKind(lngCount).BackColor = tplItem.BackColor: Next
    Else
        Me.picKind.Visible = False
        For lngCount = 0 To Me.chkKind.Count - 1: Me.chkKind(lngCount).Value = vbUnchecked: Next
    End If
    
    Set tplGroup = Me.tplThis.Groups.Add(0, "��д����:"): tplGroup.Expanded = False
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picDept
    Me.picDept.BackColor = tplItem.BackColor
    Me.optDept(0).BackColor = tplItem.BackColor: Me.optDept(1).BackColor = tplItem.BackColor

    Set tplGroup = Me.tplThis.Groups.Add(0, "�߼�����:(˫������������)"): tplGroup.Expandable = False
    Set tplItem = tplGroup.Items.Add(0, "", xtpTaskItemTypeControl): Set tplItem.Control = Me.picSearch
    Me.picSearch.BackColor = tplItem.BackColor

    '-----------------------------------------------------
    Me.tplThis.Reposition
    If InStr(1, mstrPrivs, "���ﲡ��") < 1 And InStr(1, mstrPrivs, "סԺ����") < 1 Then
        Me.chkKind(0).Enabled = False
        Me.chkKind(0).Value = 0
        Me.chkKind(1).Enabled = False
        Me.chkKind(1).Value = 0
        Me.chkKind(2).Enabled = False
        Me.chkKind(2).Value = 0
        Me.chkKind(3).Enabled = False
        Me.chkKind(3).Value = 0
        Me.chkKind(4).Enabled = False
        Me.chkKind(4).Value = 0
    Else
        If InStr(1, mstrPrivs, "���ﲡ��") < 1 Then
            Me.chkKind(0).Enabled = False
            Me.chkKind(0).Value = 0
            Me.optPati(T�����).Enabled = False
        End If
        If InStr(1, mstrPrivs, "סԺ����") < 1 Then
            Me.chkKind(1).Enabled = False
            Me.chkKind(1).Value = 0
            Me.chkKind(2).Enabled = False
            Me.chkKind(2).Value = 0
            Me.optPati(TסԺ��).Enabled = False
        End If
        If InStr(1, mstrPrivs, "������") < 1 Then
            Me.chkKind(2).Enabled = False
            Me.chkKind(2).Value = 0
        End If
        If InStr(1, mstrPrivs, "����֤��") < 1 Then
            Me.chkKind(3).Enabled = False
            Me.chkKind(3).Value = 0
        End If
        If InStr(1, mstrPrivs, "֪���ļ�") < 1 Then
            Me.chkKind(4).Enabled = False
            Me.chkKind(4).Value = 0
        End If
    End If
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    With Me.tplThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.ScaleTop: .Height = Me.ScaleHeight
    End With
End Sub

Private Sub lstDept_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub lstDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optDept_Click(Index As Integer)
    If Me.optDept(1).Value Then
        Me.lstDept.Enabled = True: If Me.lstDept.Visible Then Me.lstDept.SetFocus
    Else
        Me.lstDept.Enabled = False
    End If
End Sub

Private Sub optDept_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub optDept_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optPati_Click(Index As Integer)
Dim lngCount As Long
    For lngCount = 0 To Me.txtPati.Count - 1
        Me.txtPati(lngCount).Enabled = False
    Next
    For lngCount = 0 To Me.chkSex.Count - 1
        Me.chkSex(lngCount).Enabled = False
    Next
    If Me.optPati(T�����).Value Then
        Me.txtPati(T�����).Enabled = True: Me.txtPati(T�����).SetFocus
    ElseIf Me.optPati(Tҽ����).Value Then
        Me.txtPati(Tҽ����).Enabled = True: Me.txtPati(Tҽ����).SetFocus
    ElseIf Me.optPati(T�����).Value Then
        Me.txtPati(T�����).Enabled = True: Me.txtPati(T�����).SetFocus
    ElseIf Me.optPati(TסԺ��).Value Then
        Me.txtPati(TסԺ��).Enabled = True: Me.txtPati(TסԺ��).SetFocus
    Else
        Me.txtPati(T����).Enabled = True: Me.txtPati(T����).SetFocus
        For lngCount = 0 To Me.chkSex.Count - 1
            Me.chkSex(lngCount).Enabled = True
        Next
    End If
End Sub

Private Sub picContent_Resize()
    Err = 0: On Error Resume Next
    Me.cboCompend.Width = Me.picContent.ScaleWidth - Me.cboCompend.Left
    Me.txtContent.Width = Me.picContent.ScaleWidth - Me.txtContent.Left
End Sub

Private Sub picDate_Resize()
    Err = 0: On Error Resume Next
    Me.dtpDateFrom.Width = Me.picDate.ScaleWidth - Me.dtpDateFrom.Left
    Me.dtpDateTo.Width = Me.picDate.ScaleWidth - Me.dtpDateTo.Left
End Sub

Private Sub PicDept_Resize()
    Err = 0: On Error Resume Next
    Me.lstDept.Width = Me.picDept.ScaleWidth - Me.lstDept.Left
End Sub

Private Sub picPati_Resize()
Dim lngCount As Long
    Err = 0: On Error Resume Next
    For lngCount = 0 To Me.txtPati.Count - 1
        Me.txtPati(lngCount).Width = Me.picPati.ScaleWidth - Me.txtPati(lngCount).Left
    Next
End Sub

Private Sub picSearch_Resize()
    Err = 0: On Error Resume Next
    Me.txtElement.Width = Me.picSearch.ScaleWidth - Me.txtElement.Left
    Me.cmdSearch.Left = Me.picSearch.ScaleWidth - Me.cmdSearch.Width + 15
    Me.cmdSearch.Top = Me.picSearch.ScaleHeight - Me.cmdSearch.Height + 15
End Sub

Private Sub txtContent_Change()
    ValidControlText txtContent
End Sub

Private Sub txtContent_GotFocus()
    Me.txtContent.SelStart = 0: Me.txtContent.SelLength = Len(Me.txtContent.Text)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtElement_Change()
    ValidControlText txtElement
End Sub

Private Sub txtElement_DblClick()
Dim strTerms As String, lngStyle As Long
Dim lngCount As Long
    strTerms = Me.txtElement.Tag
    lngStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
    If frmEPRSearchElement.ShowMe(Me, strTerms) Then
        Me.txtElement.Tag = strTerms
    
        Dim aryTerm() As String, aryField() As String
        aryTerm = Split(Mid(Trim(Me.txtElement.Tag), 3), "|")
        strTerms = ""
        For lngCount = 0 To UBound(aryTerm)
            aryField = Split(aryTerm(lngCount), ";")
            strTerms = strTerms & vbCrLf & Space(2) & aryField(1) & " " & aryField(3) & " " & aryField(4)
        Next
        If strTerms <> "" Then strTerms = Mid(strTerms, 3)
        Me.txtElement.Text = strTerms
    End If
    SetWindowLong Me.hWnd, GWL_STYLE, lngStyle And Not WS_DISABLED
End Sub

Private Sub txtPati_Change(Index As Integer)
    ValidControlText txtPati(Index)
    If Trim(txtPati(Index).Text) = "" Then
        chkDtp.Value = vbChecked
    Else
        chkDtp.Value = vbUnchecked
    End If
End Sub

Private Sub txtPati_GotFocus(Index As Integer)
    Me.txtPati(Index).SelStart = 0: Me.txtPati(Index).SelLength = Len(Me.txtPati(Index).Text)
    Select Case Index
    Case 0, 1
        Call zlCommFun.OpenIme(False)
    Case 2
        Call zlCommFun.OpenIme(True)
    End Select
End Sub

Private Sub txtPati_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 0, 2
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case vbKeyReturn
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        Case Else
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        End Select
        KeyAscii = 0
    Case 3
        If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End Select
End Sub


