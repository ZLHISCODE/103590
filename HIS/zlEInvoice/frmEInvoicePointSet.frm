VERSION 5.00
Begin VB.Form frmEInvoicePointSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Ʊ�ݿ�Ʊ������"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmEInvoicePointSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5220
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Height          =   3400
      Left            =   3800
      TabIndex        =   25
      Top             =   -90
      Width           =   10
   End
   Begin VB.CommandButton cmd�շ�Ա 
      Caption         =   "��"
      Height          =   250
      Left            =   3360
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   7
      Left            =   840
      MaxLength       =   50
      TabIndex        =   23
      Top             =   3720
      Width           =   2475
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   250
      Left            =   3360
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2085
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   6
      Left            =   840
      TabIndex        =   6
      Top             =   2070
      Width           =   2475
   End
   Begin VB.CommandButton cmd�ͻ��� 
      Caption         =   "��"
      Height          =   250
      Left            =   3360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1725
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   970
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "����"
      Text            =   "111111"
      Top             =   615
      Width           =   2535
   End
   Begin VB.TextBox txtTemp 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   840
      MaxLength       =   10
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "����"
      Text            =   "1111111111"
      Top             =   570
      Width           =   2775
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   180
      Width           =   2475
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   1695
      Width           =   2475
   End
   Begin VB.ComboBox cmbStationNo 
      Height          =   300
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2835
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   840
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   840
      MaxLength       =   50
      TabIndex        =   2
      Top             =   945
      Width           =   2775
   End
   Begin VB.CommandButton cmd�ϼ� 
      Caption         =   "��"
      Height          =   250
      Left            =   3350
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   195
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   5
      Left            =   840
      MaxLength       =   100
      TabIndex        =   8
      Top             =   2445
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3960
      TabIndex        =   12
      Top             =   240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3960
      TabIndex        =   13
      Top             =   720
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�շ�Ա"
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   22
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   6
      Left            =   360
      TabIndex        =   21
      Top             =   2130
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�ͻ���"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   20
      Top             =   1740
      Width           =   540
   End
   Begin VB.Label lblStationNo 
      AutoSize        =   -1  'True
      Caption         =   "Ժ��"
      Height          =   180
      Left            =   360
      TabIndex        =   19
      Top             =   2895
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�"
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   1365
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   615
      Width           =   360
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "λ��"
      Height          =   180
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   2505
      Width           =   360
   End
   Begin VB.Menu mnuShort 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPatient 
         Caption         =   "���ﲡ��(&O)"
         Index           =   0
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "סԺ����(&I)"
         Index           =   1
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "�����סԺ����(&B)"
         Index           =   2
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuPatient 
         Caption         =   "�������ڲ���(&N)"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmEInvoicePointSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const mlng���볤�� As Long = 10
Private mstr�ϼ�ID As String     '��ǰ�༭���ϼ���Ʊ��ID
Private mstrID As String            '��ǰ�༭�ķ�Ʊ��ID
Private mblnĩ�� As Boolean     '��ǰ�༭�ķ�Ʊ���Ƿ�Ϊĩ��
Private mstr���� As String         'ԭʼ�ı��������ֵ
Private mstr�ϼ����� As String   'ԭʼ���ϼ������ֵ
Private mint���� As Integer       '�޸�ǰ�����¼����ڵı�����ĳ���
Private mblnChange As Boolean     '�Ƿ�ı���
Private Enum mEdit
    Edit_�ͻ��� = 0
    Edit_���� = 1
    Edit_���� = 2
    Edit_���� = 3
    Edit_�ϼ� = 4
    Edit_λ�� = 5
    Edit_���� = 6
    Edit_�շ�Ա = 7
End Enum
Private mbln��Ʊ����� As Boolean
Private mintMode As Integer  '���뷽ʽ0-���ͻ��˶�,1-���շ�Ա��;2-���շ�Ա+�ͻ��˶�
Private mlng����ID As Long
Private mblnOK  As Boolean

Private Sub IniStationNo()
    Dim rsRecord As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    lblStationNo.Visible = True
    cmbStationNo.Visible = True
    
    strSQL = "select ���,���� from zlnodelist"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "վ���ѯ")
    
    If rsRecord.RecordCount = 0 Then
        lblStationNo.Visible = False
        cmbStationNo.Visible = False
    Else
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!��� & "-" & rsRecord!����
                rsRecord.MoveNext
            Loop
        End With
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStationNo(ByVal strNO As String)
    Dim n As Integer
    If cmbStationNo.ListCount = 0 Then Exit Sub
    
    If strNO = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNO Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL  As String
    
    If mbln��Ʊ����� Then
        If Save��Ʊ�����() = False Then Exit Sub
        mblnChange = False: mbln��Ʊ����� = False
        mblnOK = True: Unload Me: Exit Sub
    End If
    
    If IsValid() = False Then Exit Sub
        
    '��鿪Ʊ���Ƿ������п�Ʊ��������ͬ
    If CheckSame(txtEdit(mEdit.Edit_����).Text, Val(mstrID)) Then
        If MsgBox("��ǰ¼��Ŀ�Ʊ�����������п�Ʊ��������ͬ!", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    
    If Save��Ʊ��() = False Then Exit Sub
    mblnOK = True
    '�ı������ڵ���ʾ
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    Else
    
    End If
    '��������
    mstrID = ""
    txtEdit(mEdit.Edit_����).Text = ""
    txtEdit(mEdit.Edit_����).Text = ""
    txtEdit(mEdit.Edit_λ��).Text = ""
    txtEdit(mEdit.Edit_����).Text = GetMaxLocalCode(mstr�ϼ�ID, "����Ʊ�ݿ�Ʊ��")
    
    txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ�ID, "����Ʊ�ݿ�Ʊ��")
    txtEdit(mEdit.Edit_����).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(txtTemp.Text)

    zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
    
    mblnChange = False
End Sub

Private Function IsValid() As Boolean
    Dim i As Long
    Dim blnTmp As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim str�������� As String
    Dim strMsg As String
    
    On Error GoTo errHandle
    
    For i = 1 To 5
        If i <> 4 Then
            If zlCommFun.StrIsValid(Trim(txtEdit(mEdit.Edit_����).Text), txtEdit(i).MaxLength) = False Then
                zlControl.ControlSetFocus txtEdit(i)
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        End If
    Next
    txtEdit(mEdit.Edit_����).Text = Trim(txtEdit(mEdit.Edit_����).Text)

    If Len(Trim(txtEdit(mEdit.Edit_�ϼ�).Text)) = 0 And Me.Tag = "�ָ�" Then
        MsgBox "�ϼ�����Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_�ϼ�)
        Exit Function
    End If
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(mEdit.Edit_����).Text) = 0 Then
            MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
            Exit Function
        End If
    Else
        If Len(txtEdit(mEdit.Edit_����).Text) < txtEdit(mEdit.Edit_����).MaxLength Then
            MsgBox "����ĳ��Ȳ�����", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(mEdit.Edit_����).Text) Or InStr(txtEdit(mEdit.Edit_����).Text, ",") > 0 Or InStr(txtEdit(mEdit.Edit_����).Text, ".") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
        Exit Function
    End If
    If Len(Trim(txtEdit(mEdit.Edit_����).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(mEdit.Edit_����).Text = ""
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
        Exit Function
    End If
    If LenB(StrConv(txtEdit(mEdit.Edit_����).Text, vbFromUnicode)) > 20 Then
        MsgBox "���Ƴ��Ȳ��ܳ���10�����ֻ���20���ַ���������¼�룡", vbInformation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
        Exit Function
    End If
    If LenB(StrConv(txtEdit(mEdit.Edit_����).Text, vbFromUnicode)) > 20 Then
        MsgBox "���볤�Ȳ��ܳ���20���ַ���������¼�룡", vbInformation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
        Exit Function
    End If

    IsValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function reMoveSpe(ByVal strChar As String) As String
'129884,ȥ��������Ļس������У�����β�ո�
    reMoveSpe = Trim(Replace(Replace(Replace(strChar, vbCrLf, ""), vbCr, ""), vbLf, ""))
End Function

Private Function Save��Ʊ��() As Boolean
    'blnDelete-�Ƿ�ɾ������Ʊ�ݿ�Ʊ��
    Dim i As Integer, strSQL As String
    Dim nod As Node
    Dim lst As ListItem
    Dim strվ�� As String
    Dim lngID As Long
    
    On Error GoTo errHandle
    
    txtEdit(mEdit.Edit_����).Text = reMoveSpe(txtEdit(mEdit.Edit_����).Text)
    If cmbStationNo.Text = "" Then
        strվ�� = "Null"
    Else
        strվ�� = "'" & Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1) & "'"
    End If
        
    If mstrID = "" Then       '����һ����¼
        If Check�ظ�����(mstr�ϼ�ID, Trim(txtEdit(mEdit.Edit_����).Text)) = True Then
            MsgBox "�ü��������иò��ţ����������ͬ���ţ�", vbInformation, gstrSysName
            Exit Function
        End If
        lngID = zlDatabase.GetNextId("����Ʊ�ݿ�Ʊ��")
        '  Zl_����Ʊ�ݿ�Ʊ��_Insert
        strSQL = "Zl_����Ʊ�ݿ�Ʊ��_Insert("
        '  Id_In       In ����Ʊ�ݿ�Ʊ��.Id%Type,
        strSQL = strSQL & lngID & ","
        '  �ϼ�id_In   In ����Ʊ�ݿ�Ʊ��.�ϼ�id%Type,
        strSQL = strSQL & ZVal(Val(txtEdit(mEdit.Edit_�ϼ�).Tag)) & ","
        '  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
        strSQL = strSQL & "'" & txtTemp.Text & txtEdit(mEdit.Edit_����).Text & "',"
        '  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_����).Text & "',"
        '  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_����).Text & "',"
        '  Ժ��_In     In ����Ʊ�ݿ�Ʊ��.Ժ��%Type,
        strSQL = strSQL & strվ�� & ","
        '  �ͻ���_In   In ����Ʊ�ݿ�Ʊ��.�ͻ���%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_�ͻ���).Text & "',"
        '  ����id_In   In ����Ʊ�ݿ�Ʊ��.����id%Type,
        strSQL = strSQL & IIf(Val(txtEdit(mEdit.Edit_����).Tag) = 0, "NULL", "'" & Val(txtEdit(mEdit.Edit_����).Tag) & "'") & ","
        '  λ��_In     In ����Ʊ�ݿ�Ʊ��.λ��%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_λ��).Text & "',"
        '  ĩ��_In     In ����Ʊ�ݿ�Ʊ��.ĩ��%Type := Null,
        strSQL = strSQL & "" & ZVal(IIf(mblnĩ��, 1, 0)) & ")"
    Else
        '�޸�
        lngID = Val(mstrID)
        '  Zl_����Ʊ�ݿ�Ʊ��_Update
        strSQL = "Zl_����Ʊ�ݿ�Ʊ��_Update("
        '  Id_In       In ����Ʊ�ݿ�Ʊ��.Id%Type,
        strSQL = strSQL & lngID & ","
        '  �ϼ�id_In   In ����Ʊ�ݿ�Ʊ��.�ϼ�id%Type,
        strSQL = strSQL & ZVal(Val(txtEdit(mEdit.Edit_�ϼ�).Tag)) & ","
        '  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
        strSQL = strSQL & "'" & txtTemp.Text & txtEdit(mEdit.Edit_����).Text & "',"
        '  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_����).Text & "',"
        '  ����_In     In ����Ʊ�ݿ�Ʊ��.����%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_����).Text & "',"
        '  Ժ��_In     In ����Ʊ�ݿ�Ʊ��.Ժ��%Type,
        strSQL = strSQL & strվ�� & ","
        '  �ͻ���_In   In ����Ʊ�ݿ�Ʊ��.�ͻ���%Type,
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_�ͻ���).Text & "',"
        '  ����id_In   In ����Ʊ�ݿ�Ʊ��.����id%Type,
        strSQL = strSQL & IIf(Val(txtEdit(mEdit.Edit_����).Tag) = 0, "NULL", "'" & Val(txtEdit(mEdit.Edit_����).Tag) & "'") & ","
        '  λ��_In     In ����Ʊ�ݿ�Ʊ��.λ��%Type
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_λ��).Text & "')"
    End If

    Call zlDatabase.ExecuteProcedure(strSQL, "����Ʊ�ݿ�Ʊ��")
    
    Save��Ʊ�� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save��Ʊ�����() As Boolean
    Dim strSQL As String
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    
    If mlng����ID = 0 Then       '����һ����¼
        lng����ID = zlDatabase.GetNextId("����Ʊ�ݿ�Ʊ��")
        '  Zl_Ʊ�ݿ�Ʊ�����_Update
        strSQL = "Zl_Ʊ�ݿ�Ʊ�����_Update("
        '  ����_In     In Number,
        strSQL = strSQL & 0 & ","
        '  Id_In       In Ʊ�ݿ�Ʊ�����.Id%Type := Null,
        strSQL = strSQL & lng����ID & ","
        '  ��Ʊ��id_In In ����Ʊ�ݿ�Ʊ��.Id%Type := Null,
        strSQL = strSQL & Val(mstrID) & ","
        '  ��Աid_In   In Ʊ�ݿ�Ʊ�����.��Աid%Type := Null,
        strSQL = strSQL & ZVal(txtEdit(Edit_�շ�Ա).Tag) & ","
        '  �ͻ���_In   In Ʊ�ݿ�Ʊ�����.�ͻ���%Type := Null
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_�ͻ���).Text & "')"
    Else
        '�޸�
        '  Zl_Ʊ�ݿ�Ʊ�����_Update
        strSQL = "Zl_Ʊ�ݿ�Ʊ�����_Update("
        '  ����_In     In Number,
        strSQL = strSQL & 1 & ","
        '  Id_In       In Ʊ�ݿ�Ʊ�����.Id%Type := Null,
        strSQL = strSQL & mlng����ID & ","
        '  ��Ʊ��id_In In ����Ʊ�ݿ�Ʊ��.Id%Type := Null,
        strSQL = strSQL & Val(mstrID) & ","
        '  ��Աid_In   In Ʊ�ݿ�Ʊ�����.��Աid%Type := Null,
        strSQL = strSQL & ZVal(txtEdit(Edit_�շ�Ա).Tag) & ","
        '  �ͻ���_In   In Ʊ�ݿ�Ʊ�����.�ͻ���%Type := Null
        strSQL = strSQL & "'" & txtEdit(mEdit.Edit_�ͻ���).Text & "')"
    End If

    Call zlDatabase.ExecuteProcedure(strSQL, "����Ʊ�ݿ�Ʊ��")
    
    Save��Ʊ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Init��Ʊ������(ByVal strID As String, Optional ByVal str�ϼ�ID As String, Optional ByVal blnĩ�� As Boolean, Optional blnRefresh As Boolean)
    On Error GoTo errHandle
    'strID-����Ʊ�ݿ�Ʊ��.id
    'str�ϼ�ID:�ϼ�id
    'blnĩ��:true-��ĩ��,false-����ĩ��
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    mstrID = strID
    mblnĩ�� = blnĩ��
    mblnOK = False
    Call IniStationNo
    If Not mblnĩ�� Then Call SetControlVisib
    If strID <> "" Then
        strSQL = "Select a.Id, a.�ϼ�id, a.����, b.���� as �ϼ�����,a.����, a.����, a.Ժ��, a.�ͻ���, a.λ��,a.����id, b.���� As �ϼ�����,c.���� As ���� " & _
        "   From ����Ʊ�ݿ�Ʊ�� A, ����Ʊ�ݿ�Ʊ�� B,���ű� C " & _
        "   Where a.�ϼ�id = b.Id(+) And a.����id=c.id(+) And a.ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
        mstr�ϼ�ID = IIf(IsNull(rsTemp("�ϼ�ID")), "", rsTemp("�ϼ�ID"))
        mstr�ϼ����� = IIf(IsNull(rsTemp("�ϼ�����")), "", rsTemp("�ϼ�����"))
        mstr���� = rsTemp("����")
        txtEdit(mEdit.Edit_�ϼ�).Text = IIf(IsNull(rsTemp("�ϼ�����")), "��", rsTemp("�ϼ�����"))
        txtEdit(mEdit.Edit_�ϼ�).Tag = IIf(IsNull(rsTemp("�ϼ�id")), "0", rsTemp("�ϼ�id"))
        txtTemp.Text = mstr�ϼ�����
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ�ID, "����Ʊ�ݿ�Ʊ��")
        'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
        txtEdit(mEdit.Edit_����).Text = Mid(rsTemp("����"), Len(txtTemp.Text) + 1)
        '��������ӽڵ����ڵ������
        mint���� = GetDownCodeLength(mstrID, "����Ʊ�ݿ�Ʊ��")
        '10 - (mint���� - Len(mstr����))�����ʽ����˼��ҪΪ���ĺ��ӵı����������
        txtEdit(mEdit.Edit_����).MaxLength = IIf(txtTemp.MaxLength = 0, 10 - (mint���� - Len(mstr����)), txtTemp.MaxLength) - Len(mstr�ϼ�����)
        txtEdit(mEdit.Edit_����).Text = rsTemp("����")
        txtEdit(mEdit.Edit_����).Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        txtEdit(mEdit.Edit_λ��).Text = IIf(IsNull(rsTemp("λ��")), "", rsTemp("λ��"))
        txtEdit(mEdit.Edit_�ͻ���).Text = IIf(IsNull(rsTemp("�ͻ���")), "", rsTemp("�ͻ���"))
        txtEdit(mEdit.Edit_����).Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        txtEdit(mEdit.Edit_����).Tag = IIf(IsNull(rsTemp("����id")), "", rsTemp("����id"))
        SetStationNo (IIf(IsNull(rsTemp("Ժ��")), "", rsTemp("Ժ��")))
    Else
        If str�ϼ�ID = "oot" Then
            mstr�ϼ�ID = ""
            mstr�ϼ����� = ""
            txtTemp.Text = ""
            txtEdit(mEdit.Edit_�ϼ�).Text = "��"
            'ȡ���ϼ����룬�������볤�ȵ�ֵ
            txtTemp.MaxLength = GetLocalCodeLength("", "����Ʊ�ݿ�Ʊ��")
        Else
            strSQL = "select ���� as �ϼ�����,���� as �ϼ�����,ID as �ϼ�ID from ����Ʊ�ݿ�Ʊ�� where ID=[1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str�ϼ�ID))
                        
            mstr�ϼ�ID = IIf(IsNull(rsTemp("�ϼ�ID")), "", rsTemp("�ϼ�ID"))
            mstr�ϼ����� = IIf(IsNull(rsTemp("�ϼ�����")), "", rsTemp("�ϼ�����"))
            txtEdit(mEdit.Edit_�ϼ�).Text = IIf(IsNull(rsTemp("�ϼ�����")), "��", rsTemp("�ϼ�����"))
            txtEdit(mEdit.Edit_�ϼ�).Tag = IIf(IsNull(rsTemp("�ϼ�id")), "0", rsTemp("�ϼ�id"))
            txtTemp.Text = mstr�ϼ�����
            '�жϱ����Ƿ�����
            If Len(mstr�ϼ�����) = mlng���볤�� Then
                MsgBox "�����������Ӳ����ˣ����볤���Ѿ��þ���", vbExclamation, gstrSysName
                Exit Sub
            End If
            'ȡ���ϼ����룬�������볤�ȵ�ֵ
            txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ�ID, "����Ʊ�ݿ�Ʊ��")
            'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
        End If
        txtEdit(mEdit.Edit_����).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(mstr�ϼ�����)
        txtEdit(mEdit.Edit_����).Text = GetMaxLocalCode(mstr�ϼ�ID, "����Ʊ�ݿ�Ʊ��")
        mstr���� = mstr�ϼ����� & txtEdit(1).Text
    End If

    mblnChange = False
    Me.Show vbModal
    blnRefresh = mblnOK
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub Init��Ʊ�����(ByVal intMode As Integer, ByVal lng��Ʊ��id As Long, Optional ByVal lng����ID As Long, Optional blnRefresh As Boolean)
    On Error GoTo errHandle
    'intMode:0-���ͻ��˶�,1-���շ�Ա��;2-���շ�Ա+�ͻ��˶�
    'lng��Ʊ��id:����Ʊ�ݿ�Ʊ��.id
    'lng����ID-�޸Ŀ�Ʊ�����ʱ����
    
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strFilter As String
    
    If lng��Ʊ��id = 0 Then Exit Sub
    mbln��Ʊ����� = True
    mblnOK = False
    mstrID = lng��Ʊ��id
    mintMode = intMode
    mlng����ID = lng����ID
    If lng����ID > 0 Then strFilter = " And b.id=[2] "
    strSQL = "Select a.Id, a.����, nvl(b.�ͻ���,a.�ͻ���)As �ͻ���,b.��Աid As �շ�Աid,c.���� As �շ�Ա " & _
    "   From ����Ʊ�ݿ�Ʊ�� A,Ʊ�ݿ�Ʊ����� B,��Ա�� C " & _
    "   Where a.id=b.��Ʊ��id(+) And b.��Աid=c.id(+)  And a.ID=[1]" & strFilter
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ʊ��id, lng����ID)
    If rsTemp.EOF Then Exit Sub
    txtEdit(mEdit.Edit_����).Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
    txtEdit(mEdit.Edit_����).Enabled = False
    If lng����ID <> 0 Then
        If mintMode = 0 Then
            txtEdit(mEdit.Edit_�ͻ���).Text = IIf(IsNull(rsTemp("�ͻ���")), "", rsTemp("�ͻ���"))
        ElseIf mintMode = 1 Then
            txtEdit(mEdit.Edit_�շ�Ա).Text = IIf(IsNull(rsTemp("�շ�Ա")), "", rsTemp("�շ�Ա"))
            txtEdit(mEdit.Edit_�շ�Ա).Tag = IIf(IsNull(rsTemp("�շ�Աid")), "0", rsTemp("�շ�Աid"))
        Else
            txtEdit(mEdit.Edit_�ͻ���).Text = IIf(IsNull(rsTemp("�ͻ���")), "", rsTemp("�ͻ���"))
            txtEdit(mEdit.Edit_�շ�Ա).Text = IIf(IsNull(rsTemp("�շ�Ա")), "", rsTemp("�շ�Ա"))
            txtEdit(mEdit.Edit_�շ�Ա).Tag = IIf(IsNull(rsTemp("�շ�Աid")), "0", rsTemp("�շ�Աid"))
        End If
    End If
    mblnChange = False
    Call SetControlStation
    Me.Show vbModal
    blnRefresh = mblnOK
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    strSQL = "Select Distinct a.id,a.����,a.����, a.λ�� From ���ű� A, ��������˵�� B Where a.Id = b.����id And Nvl(b.�������, 0) <> 0 Order  By a.���� "
    vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_����).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ����", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_����).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(mEdit.Edit_����).Text = rsTemp("����")
        txtEdit(mEdit.Edit_����).Tag = rsTemp("id")
        txtEdit(mEdit.Edit_λ��).Text = NVL(rsTemp("λ��"))
    End If
    zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
End Sub

Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '�������������ID������
    '����������ɹ����� �¼�������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ�ID is null " & strWhere & " connect by prior id=�ϼ�id"
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ�ID=" & strID & strWhere & " connect by prior id=�ϼ�id"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetDownCodeLength")
    
    If rsTemp.RecordCount = 0 Then
        GetDownCodeLength = 0
    Else
        GetDownCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetDownCodeLength = 0
End Function

Private Sub cmd�ͻ���_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    strSQL = "Select Rownum As id,Upper(����վ) as ����վ, Upper(��;) as ��;,Upper(����) as ���� From zlclients Order  By ����վ "
    vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_�ͻ���).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�ͻ���", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_�ͻ���).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(mEdit.Edit_�ͻ���).Text = rsTemp("����վ")
    End If
    zlControl.ControlSetFocus txtEdit(mEdit.Edit_�ͻ���)
End Sub

Private Sub cmd�ϼ�_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim str���� As String
    Dim int����  As Integer
    Dim vRect As RECT, rsTemp As ADODB.Recordset
    
    If mstrID <> "" Then
        strSQL = "select id,�ϼ�id,����,����,���� from ����Ʊ�ݿ�Ʊ�� where ����ʱ��=to_date('3000-01-01','YYYY-MM-DD') and Nvl(ĩ��, 0) = 0 and id<>" & mstrID & " start with �ϼ�id is null connect by prior id =�ϼ�id And �ϼ�id<>" & mstrID
    Else
        strSQL = "select id,�ϼ�id,����,����,���� from ����Ʊ�ݿ�Ʊ�� where ����ʱ��=to_date('3000-01-01','YYYY-MM-DD') and Nvl(ĩ��, 0) = 0 start with �ϼ�id is null connect by prior id =�ϼ�id "
    End If
    strID = mstr�ϼ�ID
    str���� = txtEdit(mEdit.Edit_�ϼ�).Text
    str���� = txtTemp.Text
    vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_�ϼ�).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�ϼ�", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_λ��).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(mEdit.Edit_�ϼ�).Text = rsTemp("����")
        txtEdit(mEdit.Edit_�ϼ�).Tag = rsTemp("id")
        int���� = GetLocalCodeLength(txtEdit(mEdit.Edit_�ϼ�).Tag, "����Ʊ�ݿ�Ʊ��")
        strID = rsTemp("id")
        str���� = rsTemp("����")
        str���� = rsTemp("����")
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_�ϼ�)
        'ֻ���޸Ĳ��б�Ҫ���
        If mstrID <> "" Then
            If mint���� - Len(mstr����) + IIf(int���� = 0, Len(str����) + 1, int����) > 10 Then
                MsgBox "����ϼ������ʣ���Ϊ���ı���̫���ˡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        mstr�ϼ�ID = strID
        txtEdit(mEdit.Edit_�ϼ�).Text = str����
        txtTemp.MaxLength = int����
        txtTemp.Text = str����
        If mstrID <> "" Then
            txtEdit(mEdit.Edit_����).MaxLength = IIf(txtTemp.MaxLength = 0, 10 - (mint���� - Len(mstr����)), txtTemp.MaxLength) - Len(str����)
        Else
            txtEdit(mEdit.Edit_����).MaxLength = IIf(txtTemp.MaxLength = 0, 10, txtTemp.MaxLength) - Len(str����)
        End If
        txtEdit(mEdit.Edit_����).Text = GetMaxLocalCode(mstr�ϼ�ID, "����Ʊ�ݿ�Ʊ��")
    End If

    mblnChange = True
End Sub

Private Sub cmd�շ�Ա_Click()
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    strSQL = "Select Distinct a.Id, a.����, a.�Ա�, a.��������" & vbNewLine & _
                    "From ��Ա�� A, ��Ա����˵�� B" & vbNewLine & _
                    "Where a.Id = b.��Աid And b.��Ա���� In ('����Һ�Ա', '�����շ�Ա', 'Ԥ���տ�Ա', 'סԺ����Ա')"
    vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_�շ�Ա).hWnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�շ�Ա", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_�շ�Ա).Height, True, False, False)
     If Not rsTemp Is Nothing Then
        txtEdit(mEdit.Edit_�շ�Ա).Text = rsTemp("����")
        txtEdit(mEdit.Edit_�շ�Ա).Tag = rsTemp("ID")
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_�շ�Ա)
    Else
        MsgBox "û���ҵ���Ч���շ�Ա��", vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Activate()
    If Not mbln��Ʊ����� Then
        zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
    
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Edit_���� Then
        txtEdit(mEdit.Edit_����).Text = zlStr.GetCodeByVB(txtEdit(mEdit.Edit_����).Text)
    ElseIf Index = Edit_�շ�Ա Then
        If txtEdit(Index) = "" Then txtEdit(Index).Tag = "0"
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = mEdit.Edit_���� Or Index = mEdit.Edit_λ�� Then
        OS.OpenIme True
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    
    If Index = mEdit.Edit_���� Then
        If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    ElseIf Index = mEdit.Edit_���� Or Index = mEdit.Edit_���� Then
        If LenB(StrConv(txtEdit(mEdit.Edit_����).Text & Chr(KeyAscii), vbFromUnicode)) > 100 And (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack) Then
            KeyAscii = 0
        End If
    ElseIf Index = mEdit.Edit_λ�� Then
'        If KeyAscii = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
    ElseIf Index = mEdit.Edit_�ͻ��� Then
        If KeyAscii <> vbKeyReturn Then Exit Sub
        strSQL = "Select Rownum As id,Upper(����վ) as ����վ, Upper(��;) as ��;,Upper(����) as ����  From zlClients " & _
                  "Where ����վ Like Upper([1]) Or ��; Like Upper([1]) Or ���� Like Upper([1]) " & _
                  "   Or Upper(zlPinYinCode(����վ)) Like Upper([1]) Or Upper(zlPinYinCode(��;)) Like Upper([1]) Or Upper(zlPinYinCode(����)) Like Upper([1]) Order By ����վ "
                  
        vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_�ͻ���).hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�ͻ���", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_�ͻ���).Height, True, False, False, "%" & txtEdit(mEdit.Edit_�ͻ���).Text & "%")
         If Not rsTemp Is Nothing Then
            txtEdit(mEdit.Edit_�ͻ���).Text = rsTemp("����վ")
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_�ͻ���)
        Else
            MsgBox "�����������Ϣδ�ҵ���Ч�Ŀͻ��ˣ������ԣ�", vbInformation, gstrSysName
            txtEdit(mEdit.Edit_�ͻ���).Text = ""
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_�ͻ���)
        End If
    ElseIf Index = mEdit.Edit_���� Then
        If KeyAscii <> vbKeyReturn Then Exit Sub
        strSQL = "Select Distinct a.ID,a.����,a.����,a.λ�� From ���ű� A, ��������˵�� B Where a.Id = b.����id And Nvl(b.�������, 0) <> 0 " & _
                  " And A.���� Like Upper([1]) Or A.���� Like Upper([1]) Or A.���� Like Upper([1]) " & _
                  "   Or Upper(zlPinYinCode(A.����)) Like Upper([1]) Or Upper(zlPinYinCode(A.����)) Like Upper([1]) Order  By a.���� "
                  
        vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_����).hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�ͻ���", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_����).Height, True, False, False, "%" & txtEdit(mEdit.Edit_����).Text & "%")
         If Not rsTemp Is Nothing Then
            txtEdit(mEdit.Edit_����).Text = rsTemp("����")
            txtEdit(mEdit.Edit_����).Tag = rsTemp("ID")
            txtEdit(mEdit.Edit_λ��).Tag = NVL(rsTemp("λ��"))
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
        Else
            MsgBox "�����������Ϣδ�ҵ���Ч�Ĳ��ţ������ԣ�", vbInformation, gstrSysName
            txtEdit(mEdit.Edit_����).Text = ""
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_����)
        End If
    ElseIf Index = mEdit.Edit_�շ�Ա Then
        If KeyAscii <> vbKeyReturn Then Exit Sub
        strSQL = "Select Distinct a.Id, a.����, a.�Ա�, a.��������" & vbNewLine & _
                        "From ��Ա�� A, ��Ա����˵�� B" & vbNewLine & _
                        "Where a.Id = b.��Աid And b.��Ա���� In ('����Һ�Ա', '�����շ�Ա', 'Ԥ���տ�Ա', 'סԺ����Ա')" & vbNewLine & _
                        " And a.���� Like Upper([1]) Or A.���� Like Upper([1]) Or A.��� Like Upper([1])"
        vRect = zlControl.GetControlRect(txtEdit(mEdit.Edit_�շ�Ա).hWnd)
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ȡ�շ�Ա", False, "", "", False, False, _
                True, vRect.Left, vRect.Top, txtEdit(mEdit.Edit_�շ�Ա).Height, True, False, False, "%" & txtEdit(mEdit.Edit_�շ�Ա).Text & "%")
         If Not rsTemp Is Nothing Then
            txtEdit(mEdit.Edit_�շ�Ա).Text = rsTemp("����")
            txtEdit(mEdit.Edit_�շ�Ա).Tag = rsTemp("ID")
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_�շ�Ա)
        Else
            MsgBox "û���ҵ���Ч���շ�Ա��", vbInformation, gstrSysName
            txtEdit(mEdit.Edit_�շ�Ա).Text = ""
            zlControl.ControlSetFocus txtEdit(mEdit.Edit_�շ�Ա)
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Or Index = 5 Then
        OS.OpenIme False
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(1).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(1).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

Private Function CheckSame(ByVal strName As String, Optional ByVal lngID As Long) As Boolean
'----------------------------------------------
'���ܣ���鿪Ʊ���Ƿ������п�Ʊ��������ͬ
'----------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    If lngID = 0 Then
        strSQL = "Select 1 From ����Ʊ�ݿ�Ʊ�� " & _
              "Where  ���� = [1] "
    Else
      strSQL = "Select 1 From ����Ʊ�ݿ�Ʊ�� " & _
              "Where  ���� = [1]  and id<> [2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��鿪Ʊ���Ƿ������п�Ʊ��������ͬ", strName, lngID)
    CheckSame = Not rsTemp.EOF

    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check�ظ�����(ByVal str�ϼ�ID As String, ByVal str���� As String) As Boolean
    '���ܣ���������Ƿ��Ѿ��иò���
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strSQL = "select ���� from ����Ʊ�ݿ�Ʊ�� where �ϼ�id=[1] and ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�Ƿ����ظ�����", str�ϼ�ID, str����)
    If rsTemp.EOF Then
        Check�ظ����� = False
    Else
        Check�ظ����� = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetMaxLocalCode(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '��������������ָ������ϼ�ID ��ȡ������������
    '����������ϼ�ID,����
    '����������ɹ����� ������; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strAllCode As String
    Dim intLength   As Integer
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ�ID is null" & strWhere
        
        '����ǲ��ű���Ҫ�ų�"��ɾ������"�����ID
        If strTableName = "���ű�" Then
            strSQL = strSQL & " And ���� <> '-'"
        End If
    Else
        strSQL = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
    End If
    intCode = GetLocalCodeLength(str�ϼ�ID, strTableName, strWhere)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetMaxLocalCode")
    
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    'strCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
    'GetMaxLocalCode = String(intCode - Len(strAllCode), "0") & strCode
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
    If GetMaxLocalCode = "" Then GetMaxLocalCode = "1"
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Function GetParentCode(ByVal str�ϼ�ID As String, ByVal strTableName As String) As String
    '������������ȡ�ϼ�����
    '����������ϼ�ID,����
    '����������ɹ����� �ϼ�����; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        GetParentCode = ""
        Exit Function
    Else
        strSQL = "select ���� from " & strTableName & " where ID=" & str�ϼ�ID
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetParentCode")
    
    If rsTemp.RecordCount = 0 Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("����").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetLocalCodeLength(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '����������ϼ�ID������
    '����������ɹ����� ������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID is null" & strWhere
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "GetLocalCodeLength")
    
    If rsTemp.RecordCount = 0 Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Private Sub SetControlVisib()
    '���ÿؼ��Ŀɼ���
    Me.Caption = "��Ʊ���������"
    Me.Height = 2250
    lblEdit(Edit_�ͻ���).Visible = False: txtEdit(Edit_�ͻ���).Visible = False
    lblEdit(Edit_λ��).Visible = False: txtEdit(Edit_λ��).Visible = False
    lblEdit(Edit_����).Visible = False: txtEdit(Edit_����).Visible = False
    lblStationNo.Visible = False: cmbStationNo.Visible = False
    cmd�ͻ���.Enabled = False: cmd�ͻ���.Visible = False
    cmd����.Enabled = False: cmd����.Visible = False
End Sub

Private Sub SetControlStation()
    '���ÿؼ���λ��
    Me.Caption = "��Ʊ�����"
    Me.Height = 1900
    lblEdit(Edit_����).Visible = False: txtEdit(Edit_����).Visible = False: txtTemp.Visible = False
    lblEdit(Edit_�ϼ�).Visible = False: txtEdit(Edit_�ϼ�).Visible = False: cmd�ϼ�.Visible = False
    lblEdit(Edit_����).Visible = False: txtEdit(Edit_����).Visible = False
    
    lblEdit(Edit_�ͻ���).Top = lblEdit(Edit_����).Top: txtEdit(Edit_�ͻ���).Top = txtTemp.Top: cmd�ͻ���.Top = txtTemp.Top + 15
    lblEdit(Edit_�շ�Ա).Top = lblEdit(Edit_����).Top: txtEdit(Edit_�շ�Ա).Top = txtEdit(Edit_����).Top: cmd�շ�Ա.Top = txtEdit(Edit_����).Top + 15

    lblEdit(Edit_����).Top = lblEdit(Edit_�ϼ�).Top: txtEdit(Edit_����).Top = txtEdit(Edit_�ϼ�).Top
    If mintMode = 2 Then Exit Sub
    cmdOK.Top = lblEdit(Edit_�ϼ�).Top: cmdCancel.Top = 600
    Me.Height = 1585
    If mintMode = 0 Then
        lblEdit(Edit_�ͻ���).Top = 700: txtEdit(Edit_�ͻ���).Top = 650: cmd�ͻ���.Top = 665
        lblEdit(Edit_�շ�Ա).Visible = False: txtEdit(Edit_�շ�Ա).Visible = False: cmd�շ�Ա.Visible = False
    Else
        lblEdit(Edit_�շ�Ա).Top = 700: txtEdit(Edit_�շ�Ա).Top = 650: cmd�շ�Ա.Top = 665
        lblEdit(Edit_�ͻ���).Visible = False: txtEdit(Edit_�ͻ���).Visible = False: cmd�ͻ���.Visible = False
    End If
End Sub

