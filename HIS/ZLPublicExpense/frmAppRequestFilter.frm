VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAppRequestFilter 
   BorderStyle     =   0  'None
   Caption         =   "��¼����"
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&O)"
      Height          =   350
      Left            =   2505
      TabIndex        =   9
      Top             =   3795
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3660
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   3825
      Begin VB.CheckBox chkShowSet 
         Caption         =   "��ʾ�Ѵ����¼"
         Height          =   375
         Left            =   420
         TabIndex        =   16
         Top             =   255
         Width           =   1665
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "������ʱ�����"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   420
         TabIndex        =   14
         Top             =   1365
         Width           =   1665
      End
      Begin VB.ComboBox cbo���﷽ʽ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3045
         Width           =   2085
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "���Ǽ�ʱ�����"
         Height          =   375
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.ComboBox cbo������ 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2640
         Width           =   2085
      End
      Begin VB.ComboBox cbo�Ǽ��� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2235
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Index           =   0
         Left            =   2355
         TabIndex        =   2
         Top             =   1005
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   42991619
         CurrentDate     =   42338
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   1005
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   42991619
         CurrentDate     =   42328
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Index           =   1
         Left            =   2355
         TabIndex        =   12
         Top             =   1770
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   42991619
         CurrentDate     =   42338
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Index           =   1
         Left            =   720
         TabIndex        =   13
         Top             =   1770
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   42991619
         CurrentDate     =   42328
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   2115
         TabIndex        =   15
         Top             =   1830
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���﷽ʽ"
         Height          =   180
         Left            =   285
         TabIndex        =   10
         Top             =   3105
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   465
         TabIndex        =   7
         Top             =   2700
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   2115
         TabIndex        =   5
         Top             =   1065
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ���"
         Height          =   180
         Left            =   465
         TabIndex        =   4
         Top             =   2295
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmAppRequestFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mfrmParent As Object

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Public Sub SetForm(frmParent As Object)
    Set mfrmParent = frmParent
End Sub

Private Sub chkShowSet_Click()
    If chkShowSet.Value = 1 Then
        cbo������.Enabled = True
        chkDate(1).Enabled = True
        dtpBegin(1).Enabled = True
        dtpEnd(1).Enabled = True
    Else
        cbo������.Enabled = False
        chkDate(1).Enabled = False
        chkDate(1).Value = False
        dtpBegin(1).Enabled = False
        dtpEnd(1).Enabled = False
    End If
End Sub

Private Sub cmdOK_Click()
    Call mfrmParent.RefreshRecord
End Sub

Private Sub Form_Load()
    Call LoadData
End Sub

Private Function zlGetFullFieldsTable(Optional strTableName As String = "������ü�¼", Optional bytHistory As Byte = 2, _
    Optional strWhere As String = "", Optional blnSubTable As Boolean = True, Optional strAliasName As String = "A", Optional blnReadDatabaseFields As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡһ�����ݱ��е��ֶ�.������Select Id,....
    '��Σ�bytHistory-0-��������ʷ����,1-��������ʷ����,2-����������( select * from tablename Union select * from Htablename)
    '      strWhere-����
    '      blnSubTable-�Ƿ��ӱ�
    '      strAliasName-����
    '���Σ�
    '���أ�select ID ... From tableName Union ALL
    '���ƣ����˺�
    '���ڣ�2010-03-10 11:19:11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, strSQL As String
    
    strFields = zlGetFeeFields(Trim(strTableName), blnReadDatabaseFields)
    Select Case bytHistory
    Case 0 '��
        strSQL = "  Select  " & strFields & " From " & strTableName & " " & strWhere
    Case 1 '����ʷ
        strSQL = " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    Case Else '���߶�����
        strSQL = " Select  " & strFields & " From " & Trim(strTableName) & " " & strWhere & " UNION ALL " & " Select  " & strFields & " From H" & Trim(strTableName) & " " & strWhere
    End Select
    If blnSubTable Then strSQL = " (" & strSQL & ") " & strAliasName
    zlGetFullFieldsTable = strSQL
    
End Function

Private Function GetPersonnel(str���� As String, Optional blnBaseInfo As Boolean) As ADODB.Recordset
'���ܣ���ȡָ�����ʵ���Ա�б�
    Dim strSQL As String
    On Error GoTo errH
    
    If str���� <> "" Then
        If blnBaseInfo Then
            strSQL = "Select a.id,a.���,a.����,a.���� From ��Ա�� a,��Ա����˵�� b" & _
            " Where a.ID = b.��ԱID And b.��Ա����=[1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by a.����"
        Else
            strSQL = "Select a.Id, a.���, a.����, a.����, a.���֤��, a.��������, a.�Ա�, a.����, a.��������, a.�칫�ҵ绰, a.�����ʼ�, a.ִҵ���, a.ִҵ��Χ, " & _
                    "a.����ְ��, a.רҵ����ְ��, a.Ƹ�μ���ְ��, a.ѧ��, a.��ѧרҵ, a.��ѧʱ��, a.��ѧ����, a.������ѵ, a.���п���, a.���˼��, a.����ʱ��, " & _
                    "a.����ʱ��, a.����ԭ��, a.����, a.վ�� From ��Ա�� a,��Ա����˵�� b" & _
            " Where a.ID = b.��ԱID And b.��Ա����=[1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by a.����"
        End If
    Else
        If blnBaseInfo Then
            strSQL = "Select id,���,����,���� From ��Ա�� A" & _
            " Where (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by ����"
        Else
            strSQL = zlGetFullFieldsTable("��Ա��", 0, "", False) & _
            " Where (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by ����"
        End If
    End If
    Set GetPersonnel = gobjDatabase.OpenSQLRecord(strSQL, App.ProductName, str����)
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function


Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ػ�������
    '���ƣ�������
    '���ڣ�2016-01-11
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset

    Set rsTemp = GetPersonnel("", True)

    cbo�Ǽ���.Clear
    cbo�Ǽ���.AddItem "���еǼ���-"
    cbo�Ǽ���.ListIndex = 0
    If rsTemp.RecordCount > 0 Then
        Call rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            cbo�Ǽ���.AddItem rsTemp!���� & "-" & rsTemp!����
            If Nvl(rsTemp!����) = UserInfo.���� Then cbo�Ǽ���.ListIndex = cbo�Ǽ���.NewIndex
            rsTemp.MoveNext
        Next
    End If
    
    cbo������.Clear
    cbo������.AddItem "���д�����-"
    cbo������.ListIndex = 0
    If rsTemp.RecordCount > 0 Then
        Call rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            cbo������.AddItem rsTemp!���� & "-" & rsTemp!����
            If Nvl(rsTemp!����) = UserInfo.���� Then cbo������.ListIndex = cbo������.NewIndex
            rsTemp.MoveNext
        Next
    End If
    
    cbo���﷽ʽ.Clear
    cbo���﷽ʽ.AddItem "���з�ʽ-"
    cbo���﷽ʽ.ListIndex = 0
    cbo���﷽ʽ.AddItem "1-���Ƴ̸���"
    cbo���﷽ʽ.AddItem "2-���¸���"
    cbo���﷽ʽ.AddItem "3-���ܸ���"
    cbo���﷽ʽ.AddItem "4-���츴��"
    
    dtpBegin(0).Value = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd 00:00:00")
    dtpEnd(0).Value = Format(gobjDatabase.CurrentDate + 1, "yyyy-mm-dd 23:59:59")
    dtpBegin(1).Value = Format(gobjDatabase.CurrentDate - 7, "yyyy-mm-dd 00:00:00")
    dtpEnd(1).Value = Format(gobjDatabase.CurrentDate, "yyyy-mm-dd 23:59:59")
    
    LoadData = True
End Function

