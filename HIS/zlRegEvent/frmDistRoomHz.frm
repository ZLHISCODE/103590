VERSION 5.00
Begin VB.Form frmDistRoomHz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ﲡ��ǩ��"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5490
   Icon            =   "frmDistRoomHz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5490
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2595
      TabIndex        =   8
      Top             =   2865
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3810
      TabIndex        =   7
      Top             =   2865
      Width           =   1100
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      ItemData        =   "frmDistRoomHz.frx":058A
      Left            =   2055
      List            =   "frmDistRoomHz.frx":058C
      TabIndex        =   6
      Top             =   1125
      Width           =   2025
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   2055
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2010
      Width           =   2025
   End
   Begin VB.ComboBox cboҽ�� 
      Height          =   300
      Left            =   2055
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   2025
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   2700
      Width           =   6900
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   5490
      TabIndex        =   0
      Top             =   0
      Width           =   5490
      Begin VB.Image Image1 
         Height          =   720
         Left            =   4350
         Picture         =   "frmDistRoomHz.frx":058E
         Top             =   45
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   2
         Top             =   135
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ָ��������Ҫ���ﵽ��Ŀ����ҵ���Ϣ��"
         Height          =   180
         Left            =   600
         TabIndex        =   1
         Top             =   390
         Width           =   3420
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -45
         X2              =   5500
         Y1              =   765
         Y2              =   765
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   180
      Left            =   1275
      TabIndex        =   11
      Top             =   1185
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   1275
      TabIndex        =   10
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽ��"
      Height          =   180
      Left            =   1275
      TabIndex        =   9
      Top             =   1620
      Width           =   720
   End
End
Attribute VB_Name = "frmDistRoomHz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNo As String
Private mlng����ID As Long
Private mstr���� As String
Private mstrҽ�� As String
Private mlngҽ��ID As Long
Private mstrԭ���� As String
Private mlngԭ����ID As Long
Private mstrԭҽ�� As String
Private mlng�Һ�ID As Long
Private mrsDept As ADODB.Recordset
Private mlngPreDept As Long
Private mstrLike As String
Private mblnOk As Boolean
Private mlngModule As Long, mstrPrivs As String
Public Function ShowMe(frmParent As Object, ByVal lngModule As Long, strPrivs As String, ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ǩ��
    '���:strNO=Ҫ����ĹҺŵ�
    '����:
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-16 14:59:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mstrNo = strNO: mlngModule = lngModule: mstrPrivs = strPrivs: mblnOk = False
    Me.Show 1, frmParent
    ShowMe = mblnOk
End Function

Private Sub cbo����_Click()
    If cbo����.ListIndex <> -1 Then
        If mlngPreDept <> cbo����.ItemData(cbo����.ListIndex) Then
            mlngPreDept = cbo����.ItemData(cbo����.ListIndex)
            '��ȡ�ÿ���ҽ��������
            Call LoadDoctor
            Call LoadRoom
        End If
    Else
        mlngPreDept = 0
    End If
End Sub
Private Sub cbo����_GotFocus()
    Call zlControl.TxtSelAll(cbo����)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    If cbo����.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Exit Sub
    If zlSelectDept(Me, mlngModule, cbo����, mrsDept, cbo����.Text, True, "����ѡ��") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
    Dim lngID As Long
    If cbo����.ListIndex >= 0 Then Exit Sub
    lngID = mlngPreDept
   zlControl.CboLocate cbo����, lngID, True
   If cbo����.ListIndex < 0 And cbo����.ListCount <> 0 Then cbo����.ListIndex = 0
End Sub

Private Sub cboҽ��_Click()
    Call LoadRoom
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function Valied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵĺϷ���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-16 15:25:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnYes As Boolean
    On Error GoTo errHandle
    If cbo����.ListIndex = -1 Then
        MsgBox "��ȷ��Ҫ����Ŀ��ҡ�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Function
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "��ȷ��Ҫ��������ҡ�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Function
    End If
    If cbo����.ItemData(cbo����.ListIndex) <> mlngԭ����ID And blnYes = False Then
        If MsgBox("ע��:" & vbCrLf & "  ��ѡ��Ŀ��������Ŀ��Ҳ�һ��,���Ƿ�Ҫ�������˻������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            cbo����.SetFocus: Exit Function
        End If
        blnYes = True
                
    End If
    If cbo����.Text <> mstrԭ���� And blnYes = False Then
        If MsgBox("ע��:" & vbCrLf & "  ��ѡ����������������Ҳ�һ��,���Ƿ�Ҫ�������˵Ļ�������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            cbo����.SetFocus: Exit Function
        End If
        blnYes = True
    End If
    If NeedName(cboҽ��.Text) <> mstrԭҽ�� And blnYes = False Then
        If MsgBox("ע��:" & vbCrLf & "  ��ѡ���ҽ��������ҽ����һ��,���Ƿ�Ҫ�������˵Ļ���ҽ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        cboҽ��.SetFocus: Exit Function
        End If
        blnYes = True
    End If
    Valied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOk_Click()
    Dim strSQL As String
    If Valied = False Then Exit Sub
    '��������
    mlng����ID = cbo����.ItemData(cbo����.ListIndex)
    mstr���� = cbo����.Text
    mstrҽ�� = NeedName(cboҽ��.Text)
    If cboҽ��.ListIndex <> -1 Then
        mlngҽ��ID = cboҽ��.ItemData(cboҽ��.ListIndex)
    End If
    'Zl_���˹Һż�¼_����
    strSQL = "Zl_���˹Һż�¼_����("
    '  Id_In         ���˹Һż�¼.ID%Type,
    strSQL = strSQL & "" & mlng�Һ�ID & ","
    '  ��ִ�п���_In ���˹Һż�¼.ִ�в���id%Type,
    strSQL = strSQL & "" & mlng����ID & ","
    '  ������_In     ���˹Һż�¼.����%Type,
    strSQL = strSQL & "'" & mstr���� & "',"
    '  ��ҽ��_In     ���˹Һż�¼.ִ����%Type,
    strSQL = strSQL & "'" & mstrҽ�� & "',"
    '  �����_In Integer:=0
    strSQL = strSQL & "0,"
    'ԤԼ��ʽ
    strSQL = strSQL & "'" & zl_GetԤԼ��ʽByID(mlng�Һ�ID) & "')" '�����:48350
    zlDatabase.ExecuteProcedure strSQL, Me.Caption '�����:53508
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Me.ActiveControl Is cbo���� Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    mlng����ID = 0
    mstr���� = ""
    mstrҽ�� = ""
    mlngҽ��ID = 0
    mblnOk = False
    mlngPreDept = 0
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    
    On Error GoTo errH
    
    'ԭ�Һ������Ϣ
    strSQL = "Select ID, ִ�в���ID,����,ִ���� From ���˹Һż�¼ Where NO=[1] and ��¼����=1 and ��¼״̬=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNo)
    mstrԭ���� = Nvl(rsTmp!����)
    mlngԭ����ID = rsTmp!ִ�в���id
    mstrԭҽ�� = Nvl(rsTmp!ִ����)
    mlng�Һ�ID = Val(Nvl(rsTmp!id))
    '��ȡ�������:ȱʡΪ������
    strSQL = "" & _
    " Select Distinct B.ID,B.����,B.����,B.����,Decode(B.ID,[1],1,0) as ȱʡ" & _
    " From ���ű� B,��������˵�� C" & _
    " Where B.ID=C.����ID And C.������� In(1,3) And C.��������='�ٴ�'" & _
    "       And (B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or B.����ʱ�� is Null)" & _
    "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
    " Order by B.����"
    Set mrsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngԭ����ID)
    Do While Not mrsDept.EOF
        cbo����.AddItem mrsDept!���� & "-" & mrsDept!����
        cbo����.ItemData(cbo����.NewIndex) = mrsDept!id
        If Val(Nvl(mrsDept!ȱʡ)) = 1 Then
            cbo����.ListIndex = cbo����.NewIndex '��������Click
            mlngPreDept = mrsDept!id
        End If
        mrsDept.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDoctor()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
                
    cboҽ��.Clear
    If cbo����.ListIndex = -1 Then Exit Sub
    
    strSQL = "" & _
    " Select Distinct A.ID,A.���,A.����,A.����" & _
    " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
    " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
    "       And C.��Ա����='ҽ��' And B.����ID=[1]" & _
    "       And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
    "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
    " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo����.ItemData(cbo����.ListIndex))
    
    cboҽ��.AddItem ""
    Call zlControl.CboSetIndex(cboҽ��.Hwnd, 0)
    Do While Not rsTmp.EOF
        cboҽ��.AddItem Nvl(rsTmp!����) & "-" & Nvl(rsTmp!����)
        cboҽ��.ItemData(cboҽ��.NewIndex) = rsTmp!id
        If Nvl(rsTmp!����) = mstrԭҽ�� Then
            cboҽ��.ListIndex = cboҽ��.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRoom()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, bln�ٴ����� As Boolean
    
    On Error GoTo errH
    
    cbo����.Clear
    If cbo����.ListIndex = -1 Then Exit Sub
    
    bln�ٴ����� = False
    If gbytRegistMode = 1 Then
        If Sys.Currentdate >= gdatRegistTime Then bln�ٴ����� = True
    End If
    
    If bln�ٴ����� = False Then
        strSQL = _
            "Select Distinct �������� As ����" & vbNewLine & _
            "From �ҺŰ������� A, �ҺŰ��� B" & vbNewLine & _
            "Where a.�ű�id = b.Id And b.����id = [1] And Nvl(b.ҽ������,Nvl([2],'-')) = Nvl([2],'-')" & vbNewLine & _
            "Order By ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo����.ItemData(cbo����.ListIndex), NeedName(cboҽ��.Text))
    Else
        strSQL = _
            " Select Distinct c.����" & vbNewLine & _
            " From �ٴ��������Ҽ�¼ A, �ٴ������¼ B, �������� C" & vbNewLine & _
            " Where a.��¼id = b.Id And a.����id = c.Id And b.����id+0 = [1]" & vbNewLine & _
            "       And Nvl(b.ҽ������,Nvl([2],'-')) = Nvl([2],'-')" & vbNewLine & _
            "       And b.�������� Between Trunc(Sysdate) - 1 And Trunc(Sysdate)" & vbNewLine & _
            " Order By ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo����.ItemData(cbo����.ListIndex), NeedName(cboҽ��.Text))
        If rsTmp.RecordCount = 0 Then
            '���´ӿ������������ж�ȡ����,121589
            Set rsTmp = GetDoctorRooms(cbo����.ItemData(cbo����.ListIndex))
        End If
    End If
    
    cbo����.AddItem ""
    Call zlControl.CboSetIndex(cbo����.Hwnd, 0)
    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!����
        If cbo����.ItemData(cbo����.ListIndex) = mlngԭ����ID And rsTmp!���� = mstrԭ���� Then
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function zlSelectDept(ByVal frmMain As Form, ByVal lngModule As Long, ByVal cboDept As ComboBox, ByVal rsDept As ADODB.Recordset, _
    ByVal strSearch As String, Optional blnNot���ȼ� As Boolean = False, Optional str���в��� As String = "", Optional blnSendKeys As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���:cboDept-ָ���Ĳ��Ų���
    '     rsDept-ָ���Ĳ���
    '     strSearch-Ҫ�����Ĵ�
    '     blnNot���ȼ�-�Ƿ�������ȼ��ֶ�
    '     str���в���-���в�������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-26 10:20:11
    '����:27378
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, rsReturn As ADODB.Recordset
    Dim lngDeptID As Long, iCount As Integer
    Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
    Dim strCompents As String 'ƥ�䴮
    Dim strIDs As String, str���� As String
    
    '�ȸ��Ƽ�¼��
    Set rsTemp = zlDatabase.zlCopyDataStructure(rsDept)
    
    strSearch = UCase(strSearch)
    strCompents = Replace(gstrLike, "%", "*") & strSearch & "*"
    
    If IsNumeric(strSearch) Then
        intInputType = 0
    ElseIf zlCommFun.IsCharAlpha(strSearch) Then
        intInputType = 1
    Else
        intInputType = 2
    End If
    If str���в��� <> "" Then
        str���� = zlCommFun.SpellCode(str���в���)
        If intInputType = 1 Then
            If Trim(str����) Like strCompents Then
                rsTemp.AddNew
                rsTemp!id = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        Else
            If strSearch = "-" Or Trim(str����) Like strCompents Or UCase(str���в���) Like strCompents Then
                rsTemp.AddNew
                rsTemp!id = -1
                rsTemp!���� = "-"
                rsTemp!���� = str���в���
                rsTemp!���� = str����
                rsTemp.Update
            End If
        End If
    End If
    
    
    strIDs = ","
    With rsDept
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Select Case intInputType
            Case 0  '�������ȫ����
                '������������,��Ҫ���:
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                If Nvl(!����) = strSearch Then lngDeptID = Nvl(!id): iCount = 0: Exit Do
                
                '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                If Val(Nvl(!����)) = Val(strSearch) Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!id))
                    iCount = iCount + 1
                End If
                '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                 If Nvl(!����) Like strSearch & "*" Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                 End If
            Case 1  '�������ȫ��ĸ
                '����:
                ' 1.����ļ������,��ֱ�Ӷ�λ
                ' 2.���ݲ�����ƥ����ͬ����
                
                '1.����ļ������,��ֱ�Ӷ�λ
                If Trim(Nvl(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!id))   '���ܴ��ڶ����ͬ����
                    iCount = iCount + 1
                End If
                '2.���ݲ�����ƥ����ͬ����
                If Trim(Nvl(!����)) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                End If
            Case Else  ' 2-����
                '����:���ܴ��ں��ֵ����,����������N001���������LXH01�������
                '1.����\�������,ֱ�Ӷ�λ
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                
                '1.����\�������,ֱ�Ӷ�λ
                If Trim(!����) = strSearch Or Trim(!����) = strSearch Or UCase(Trim(!����)) = strSearch Then
                    If iCount = 0 Then lngDeptID = Val(Nvl(!id))   '���ܴ��ڶ����ͬ�Ķ��
                    iCount = iCount + 1
                End If
                '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                If UCase(Trim(!����)) Like strSearch & "*" Or Trim(Nvl(!����)) Like strCompents Or UCase(Trim(Nvl(!����))) Like strCompents Then
                    If InStr(1, strIDs, "," & Val(Nvl(!id)) & ",") = 0 Then Call zlDatabase.zlInsertCurrRowData(rsDept, rsTemp)
                    strIDs = strIDs & Val(Nvl(!id)) & ","
                End If
            End Select
            .MoveNext
        Loop
    End With
    strIDs = ""
    
    If iCount > 1 Then lngDeptID = 0
    If lngDeptID <> 0 And rsTemp.RecordCount = 1 Then lngDeptID = Nvl(rsTemp!id)
        
    '���˺�:ֱ�Ӷ�λ
    If lngDeptID <> 0 Then GoTo GoOver:
    If lngDeptID < 0 Then lngDeptID = 0
    
    '��Ҫ����Ƿ��ж������������ļ�¼
    If rsTemp.RecordCount = 0 Then GoTo GoNotSel:
    
    '�Ȱ�ĳ�ַ�ʽ��������
    Select Case intInputType
    Case 0 '����ȫ����
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case 1 '����ȫƴ��
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    Case Else
        rsTemp.Sort = IIf(blnNot���ȼ�, "", "���ȼ�,") & "����"
    End Select
    
    '����ѡ����
    If zlDatabase.zlShowListSelect(frmMain, glngSys, lngModule, cboDept, rsTemp, True, "", "ȱʡ," & IIf(blnNot���ȼ�, "", ",���ȼ�") & "", rsReturn) = False Then GoTo GoNotSel:
    
    If rsReturn Is Nothing Then GoTo GoNotSel:
    If rsReturn.State <> 1 Then GoTo GoNotSel:
    If rsReturn.RecordCount = 0 Then GoTo GoNotSel:
    lngDeptID = Val(Nvl(rsReturn!id))
    If lngDeptID < 0 Then lngDeptID = 0
GoOver:
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.CboLocate cboDept, lngDeptID, True
    If blnSendKeys Then zlCommFun.PressKey vbKeyTab
zlSelectDept = True
    Exit Function
GoNotSel:
    'δ�ҵ�
    rsTemp.Close: Set rsTemp = Nothing: Set rsReturn = Nothing
    zlControl.TxtSelAll cboDept
End Function


