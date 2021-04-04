VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTimeSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ù�������"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmTimeSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraFilter 
      Caption         =   "��������"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3765
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   1830
         TabIndex        =   4
         Top             =   870
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   170852355
         CurrentDate     =   36279
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1830
         TabIndex        =   2
         Top             =   390
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   170852355
         CurrentDate     =   36279
         MinDate         =   2
      End
      Begin VB.ComboBox cboOperator 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1830
         TabIndex        =   6
         Text            =   "cboOperator"
         Top             =   1320
         Width           =   1785
      End
      Begin VB.Label lblOperator 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&P)"
         Height          =   180
         Left            =   960
         TabIndex        =   5
         Top             =   1395
         Width           =   810
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   150
         Picture         =   "frmTimeSet.frx":000C
         Top             =   420
         Width           =   480
      End
      Begin VB.Label lblTimeStart 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��(&B)"
         Height          =   180
         Left            =   780
         TabIndex        =   1
         Top             =   450
         Width           =   990
      End
      Begin VB.Label lblTimeStop 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��(&E)"
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   930
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   7
      Top             =   240
      Width           =   1100
   End
End
Attribute VB_Name = "frmTimeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mbytInFun As Byte '0-�շѽɿ���ˣ�1-Ʊ��ʹ�ù���
Private mdatBegin As Date, mdatEnd As Date
Private mstrOperator As String, mstrPrivs As String
Private mrsPerson As ADODB.Recordset
Private mlngModule  As Long
Private mlngPreID As Long
Private mblnDateMoved As Boolean '�Ƿ���ת������֮ǰ
 
Private Sub cboOperator_Click()
    If cboOperator.ListIndex >= 0 Then mlngPreID = cboOperator.ItemData(cboOperator.ListIndex)
End Sub

Private Sub cboOperator_KeyPress(KeyAscii As Integer)
   Dim lngIdx As Long, lngҽ��ID As Long
     '���˺� ����:27378 ����:2010-01-27 16:20:02
    Dim strAllCaption As String
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cboOperator.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mbytInFun <> 0 And InStr(mstrPrivs, ";���в���Ա;") > 0 Then
        strAllCaption = "������Ա"
    Else
    End If

    If mrsPerson Is Nothing Then Exit Sub
    If zlPersonSelect(Me, mlngModule, cboOperator, mrsPerson, _
        cboOperator.Text, True, strAllCaption) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub
Private Sub cboOperator_Validate(Cancel As Boolean)
    If cboOperator.ListIndex < 0 Then zlControl.CboLocate cboOperator, mlngPreID, True
    If cboOperator.ListIndex < 0 And cboOperator.Text <> "" Then cboOperator.Text = ""
End Sub

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then dtpEnd.SetFocus
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If dtpBegin.Value > dtpEnd.Value Then
        MsgBox "��ʼʱ�䲻Ӧ���ڽ���ʱ�䡣", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    mblnDateMoved = zlDatabase.DateMoved(Format(dtpBegin.Value, "yyyy-MM-dd hh:mm:ss"), , , Me.Caption)
    mdatBegin = dtpBegin.Value
    mdatEnd = dtpEnd.Value
    
    If cboOperator.Text <> "������Ա" Then
        mstrOperator = zlCommFun.GetNeedName(cboOperator.Text)
    Else
        mstrOperator = ""
    End If
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function ShowMe(ByVal frmOwner As Form, ByVal bytInFun As Byte, _
    ByVal bytInvoiceKind As gBillType, ByVal lngModule As Long, ByVal strPrivs As String, _
    datBegin As Date, datEnd As Date, strOperator As String, blnDateMoved As Boolean, _
    Optional strPersonelKind As String, Optional blnOnlyHave As Boolean) As Boolean
'������
'    bytInFun:0-�շѽɿ���ˣ�1-Ʊ��ʹ�ù���
'    bytInvoiceKind:��bytInFun=1ʱ��1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨,6-���ѿ�,7-��Ա��
'    strPersonelKind:��Ա���ʣ�Ϊ��ʱ��ʾ��������
'    blnOnlyHave:ֻ������������Ա

    mbytInFun = bytInFun
    mstrPrivs = strPrivs: mlngModule = lngModule
                        
    dtpBegin.Value = TruncateDate(datBegin)
    dtpEnd.Value = TruncateDate(datEnd)
    dtpBegin.MaxDate = TruncateDate(zlDatabase.Currentdate)
    dtpEnd.MaxDate = dtpBegin.MaxDate
                    
    If mbytInFun = 0 Then
        lblOperator.Caption = "�ɿ���(&P)"
        Call Fill�ɿ���(strOperator, strPersonelKind, blnOnlyHave)
    Else
        lblOperator.Caption = "������(&P)"
        Call FillOperator(bytInvoiceKind)
    End If
    
    frmTimeSet.Show vbModal, frmOwner
    ShowMe = mblnOK
    If mblnOK = True Then
        datBegin = mdatBegin
        datEnd = mdatEnd
        strOperator = mstrOperator
        blnDateMoved = mblnDateMoved
    End If
End Function

Private Sub Fill�ɿ���(strOperator As String, strPersonelKind As String, blnOnlyHave As Boolean)
    Dim i As Long
    Dim strSQL As String
    
    On Error GoTo errH
    cboOperator.Clear
    
    If strPersonelKind = "" Then
        strSQL = " And C.��Ա���� in " & _
                "       ('����Һ�Ա','�����շ�Ա','Ԥ���տ�Ա','סԺ����Ա','��Ժ�Ǽ�Ա','�����Ǽ���')"
    Else
        strSQL = " And C.��Ա����=[1]"
    End If
                
    If blnOnlyHave Then
        '��ָ�㶨�ڼ������ݴ��Ĳ���Ա
        strSQL = _
            "Select Distinct B.ID,B.���, B.����,B.����" & vbNewLine & _
            "From ��Ա�ɿ���� A,��Ա�� B,��Ա����˵�� C" & vbNewLine & _
            "Where A.�տ�Ա=B.���� And B.id=C.��ԱID And a.���<>0" & vbNewLine & _
            "      And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
                   strSQL & vbNewLine & _
            "Order by ����"
    Else
        '�����ڼ��ڲ���Ա
        strSQL = _
            "Select Distinct A.ID,A.���, A.����,A.����" & vbNewLine & _
            "From ��Ա�� A,��Ա����˵�� C " & vbNewLine & _
            "Where A.ID=C.��ԱID And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "      And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & vbNewLine & _
                   strSQL & vbNewLine & _
            "Order by ����"
    End If
    Set mrsPerson = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPersonelKind)
    If mrsPerson.EOF Then Exit Sub
    
    For i = 1 To mrsPerson.RecordCount
        cboOperator.AddItem mrsPerson!��� & "-" & mrsPerson!����
        cboOperator.ItemData(cboOperator.NewIndex) = Val(Nvl(mrsPerson!ID))
        If strOperator = mrsPerson!���� Then cboOperator.ListIndex = cboOperator.NewIndex
        mrsPerson.MoveNext
    Next
    If cboOperator.ListIndex = -1 And cboOperator.ListCount > 0 Then cboOperator.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillOperator(ByVal bytInvoiceKind As gBillType)
    Dim strSQL As String
    Dim strValue As String, i As Long, strID As Long
    
    If InStr(mstrPrivs, "���в���Ա") = 0 Then
        cboOperator.Clear
        cboOperator.AddItem UserInfo.��� & "-" & UserInfo.����
        cboOperator.ItemData(cboOperator.NewIndex) = UserInfo.ID
        cboOperator.ListIndex = 0
    Else
        If bytInvoiceKind > 0 And bytInvoiceKind <= 7 Then
            '�������Ժ�Ǽ�Ա������Ҫͬʱ���ö�Ӧ�ķ�����Ԥ����Ա�����������ʾ��������Ϣ����ͬ��Ҳ�����������
            strValue = Choose(bytInvoiceKind, "�����շ�Ա", "Ԥ���տ�Ա", "סԺ����Ա", "����Һ�Ա", _
                "�����Ǽ���", "�����Ǽ���", "�����Ǽ���")
        End If
        strSQL = _
            "Select Distinct A.ID, A.���, A.����,A.����" & vbNewLine & _
            "From ��Ա�� A, ��Ա����˵�� B" & vbNewLine & _
            "Where A.ID = B.��Աid And B.��Ա���� = [1] " & vbNewLine & _
            "      And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            "      And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)"

        On Error GoTo errH
        Set mrsPerson = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue)
       
        cboOperator.Clear
        cboOperator.AddItem "������Ա"
        For i = 1 To mrsPerson.RecordCount
            cboOperator.AddItem mrsPerson!��� & "-" & mrsPerson!����
            cboOperator.ItemData(cboOperator.NewIndex) = Val(Nvl(mrsPerson!ID))
            mrsPerson.MoveNext
        Next
        cboOperator.ListIndex = 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
