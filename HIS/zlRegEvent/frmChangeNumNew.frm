VERSION 5.00
Begin VB.Form frmChangeNumNew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "���˻���"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra2 
      Caption         =   "�����ű�"
      Height          =   1065
      Left            =   165
      TabIndex        =   20
      Top             =   2070
      Width           =   6840
      Begin VB.ComboBox cmbDiagRoom2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2925
         TabIndex        =   25
         Text            =   "cmbDiagRoom2"
         Top             =   600
         Width           =   1620
      End
      Begin VB.ComboBox cmbSect2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   780
         TabIndex        =   23
         Top             =   615
         Width           =   1620
      End
      Begin VB.ComboBox cmbDocTor2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5040
         TabIndex        =   27
         Text            =   "cmbDocTor2"
         Top             =   600
         Width           =   1620
      End
      Begin VB.CommandButton cmdItemSel 
         Caption         =   "&P"
         Height          =   300
         Left            =   6285
         TabIndex        =   21
         ToolTipText     =   "ѡ���ºű�"
         Top             =   210
         Width           =   375
      End
      Begin VB.Label lblItem2 
         Alignment       =   1  'Right Justify
         Height          =   180
         Left            =   210
         TabIndex        =   31
         Top             =   255
         Width           =   5910
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   375
         TabIndex        =   22
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         Height          =   180
         Left            =   4635
         TabIndex        =   26
         Top             =   675
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   2520
         TabIndex        =   24
         Top             =   675
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7155
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7155
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   735
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7155
      TabIndex        =   28
      Top             =   255
      Width           =   1100
   End
   Begin VB.Frame fra1 
      Caption         =   "ԭ�Һŵ���Ϣ"
      Height          =   1815
      Left            =   165
      TabIndex        =   0
      Top             =   135
      Width           =   6825
      Begin VB.TextBox txtSect 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   780
         TabIndex        =   15
         Top             =   1290
         Width           =   1620
      End
      Begin VB.TextBox txtDiagRoom 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2925
         TabIndex        =   17
         Top             =   1290
         Width           =   1620
      End
      Begin VB.TextBox txtTime 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5025
         TabIndex        =   13
         Top             =   915
         Width           =   1620
      End
      Begin VB.TextBox txtExes 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2925
         TabIndex        =   11
         Top             =   900
         Width           =   1620
      End
      Begin VB.TextBox txtOutNum 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   780
         TabIndex        =   8
         Top             =   900
         Width           =   1620
      End
      Begin VB.TextBox txtOld 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5025
         TabIndex        =   7
         Top             =   525
         Width           =   1620
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   2925
         TabIndex        =   5
         Top             =   510
         Width           =   1620
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   300
         Left            =   780
         TabIndex        =   3
         Top             =   510
         Width           =   1620
      End
      Begin VB.ComboBox cmbDoctor 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   300
         Left            =   5025
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1305
         Width           =   1650
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   375
         TabIndex        =   14
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label lblOldItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һŵ�����"
         Height          =   180
         Left            =   315
         TabIndex        =   1
         Top             =   255
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ʱ��"
         Height          =   180
         Left            =   4635
         TabIndex        =   12
         Top             =   990
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   2535
         TabIndex        =   10
         Top             =   975
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��"
         Height          =   180
         Left            =   4620
         TabIndex        =   18
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   2520
         TabIndex        =   16
         Top             =   1350
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Left            =   210
         TabIndex        =   9
         Top             =   975
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   4635
         TabIndex        =   6
         Top             =   585
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2535
         TabIndex        =   4
         Top             =   585
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   390
         TabIndex        =   2
         Top             =   585
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmChangeNumNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const STR_COMP = "|',~" '�ָ��ַ���

Private strSQL As String
Private i As Long
Private mblnCancel As Boolean
Private mlng�Һ�ID As Long
Private mstrNo As String
Private mstr�ű� As String
Private mrsDoctor As New ADODB.Recordset
Private mlng�����¼ID As Long

Public Function ShowMe(ByVal lng�Һ�ID As String, frmParent As Form) As Boolean
'��ʾ�����岢����ѡ����Ƿ���ȷ
    On Error GoTo errHandle
    Dim rsTmp As ADODB.Recordset
    Dim strDoctor As String, lngִ�в���ID As Long
    
    mlng�Һ�ID = lng�Һ�ID
    mblnCancel = False
    
    '������ǰ�Ĳ��˹Һż�¼
    strSQL = _
        " Select A.NO,X.�ű�,X.����,X.�Ա�,X.����,X.�����," & _
        " A.�ѱ�,A.����ʱ��,X.����,X.ִ����,X.ִ�в���ID," & _
        " D.����,C.���� as �շ���Ŀ����,B.���� as ִ�в�������,D1.ID As �����¼ID" & _
        " From ������ü�¼ A,���ű� B,�շ���ĿĿ¼ C,�ٴ������Դ D,�ٴ������¼ D1,���˹Һż�¼ X" & _
        " Where A.��¼����=4 And A.��¼״̬=1 And A.���=1 And A.NO=X.NO" & _
        " AND X.��¼״̬=1 and x.��¼����=1 And A.�շ�ϸĿID=C.ID And X.�����¼ID=D1.ID And D.ID=D1.��ԴID And X.ִ�в���ID=B.ID" & _
        " And X.ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˻���", lng�Һ�ID)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        Me.lblOldItem.Caption = "�Һŵ���" & rsTmp!NO & "��   ����:" & Nvl(rsTmp!����) & _
            "    �ű�:" & Nvl(rsTmp!�ű�) & "    �Һ���Ŀ:" & rsTmp!�շ���Ŀ����
        mstrNo = rsTmp!NO
        mstr�ű� = zlCommFun.Nvl(rsTmp!�ű�)
        Me.txtName = zlCommFun.Nvl(rsTmp!����)
        Me.txtSex = zlCommFun.Nvl(rsTmp!�Ա�)
        Me.txtOld = zlCommFun.Nvl(rsTmp!����)
        Me.txtOutNum = zlCommFun.Nvl(rsTmp!�����)
        Me.txtExes = zlCommFun.Nvl(rsTmp!�ѱ�)
        Me.txtTime = Format(Nvl(rsTmp!����ʱ��), "YYYY-MM-DD HH:MM:SS")
        Me.txtSect = zlCommFun.Nvl(rsTmp!ִ�в�������)
        Me.txtDiagRoom = zlCommFun.Nvl(rsTmp!����)
        
        '����Ĭ��ҽ��
        strDoctor = "" & rsTmp!ִ����
        lngִ�в���ID = Val("" & rsTmp!ִ�в���id)
        strSQL = "SELECT b.ID,b.����,b.���� FROM ������Ա a,��Ա�� b,��Ա����˵�� c  " & _
            " WHERE b.id=a.��ԱID AND b.id=c.��Աid AND c.��Ա����='ҽ��' AND a.����ID=[1]" & _
            " And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
            " And (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˻���", lngִ�в���ID)
        Me.cmbDoctor.Clear
        Me.cmbDoctor.AddItem "W-��" & String(400, " ") & STR_COMP
        Me.cmbDoctor.ListIndex = 0
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                Me.cmbDoctor.AddItem rsTmp!���� & "-" & rsTmp!���� & String(400, " ") & STR_COMP & rsTmp!ID
                rsTmp.MoveNext
            Next
            If Trim(strDoctor) <> "" Then
                For i = 0 To Me.cmbDoctor.ListCount - 1
                    If Me.cmbDoctor.List(i) Like "*-" & strDoctor & " *" Then
                        Me.cmbDoctor.ListIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
    Else
        MsgBox "�޸ò��˹Һ���Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    Me.Show 1, frmParent
    If mblnCancel = False Then
        ShowMe = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetRoom(lng��¼ID As Long) As String
'���ܣ����ݺű�ķ��﷽ʽ��ȡ�ű������
    Dim strSQL As String, strRoomIDs As String
    Dim rsTmp As ADODB.Recordset, rsRoom As ADODB.Recordset
    On Error GoTo errH
    
    strSQL = "Select ID,Nvl(���﷽ʽ,0) as ���� From �ٴ������¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˻���", lng��¼ID)
    If rsTmp.EOF Then Exit Function
    If rsTmp!���� = 0 Then Exit Function '������
    
    '�������
    If rsTmp!���� = 1 Then
        'ָ������
        strSQL = "Select A.���� As �������� From �������� A,�ٴ��������Ҽ�¼ B Where B.��¼ID=[1] And A.ID=B.����ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˻���", lng��¼ID)
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 2 Then
        '��̬����ø��ű���Һ�δ�������ٵ�����   //todoδ����ԤԼ�Һ�
        strSQL = _
            " Select ��������,Sum(NUM) as NUM From (" & _
                " Select B.���� As ��������,0 as NUM From �ٴ��������Ҽ�¼ A,�������� B Where A.��¼ID=[1] And A.����ID=B.ID" & _
                " Union ALL" & _
                " Select ����,Count(����) as NUM From ���˹Һż�¼" & _
                " Where Nvl(ִ��״̬,0)=0 And �����¼ID=[1]" & _
                " And ��¼����=1 and ��¼״̬=1 and ����ʱ�� Between Trunc(Sysdate) And  Sysdate" & _
                " And ���� IN (Select B.���� From �ٴ��������Ҽ�¼ A,�������� B Where A.��¼ID=[1] And A.����ID=B.ID)" & _
                " Group by ����)" & _
            " Group by ��������" & _
            " Order by Num"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���˻���", lng��¼ID)
        If Not rsTmp.EOF Then GetRoom = rsTmp!��������
    ElseIf rsTmp!���� = 3 Then
        'ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����,(��������,����Ҫ��*,������»����)
        strSQL = "Select * From �ٴ��������Ҽ�¼ Where ��¼ID=" & rsTmp!ID
        '���ؿɸ��¼�¼��
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption, adOpenStatic, adLockOptimistic)
        
        If Not rsTmp.EOF Then
            Do While Not rsTmp.EOF
                If IIf(IsNull(rsTmp!��ǰ����), 0, rsTmp!��ǰ����) = 1 Then
                    strRoomIDs = rsTmp!����ID
                    rsTmp!��ǰ���� = 0

                    rsTmp.MoveNext
                    If rsTmp.EOF Then rsTmp.MoveFirst
                    rsTmp!��ǰ���� = 1
                    rsTmp.Update
                    Exit Do
                End If
                rsTmp.MoveNext
            Loop
            '�����һ��ƽ������
            If strRoomIDs = "" Then
                rsTmp.MoveFirst
                strRoomIDs = rsTmp!����ID
                rsTmp.MoveNext
                If rsTmp.EOF Then rsTmp.MoveFirst
                rsTmp!��ǰ���� = 1
                rsTmp.Update
            End If
        End If
        If strRoomIDs <> "" Then
            strSQL = "Select ���� From �������� Where ID = [1]"
            Set rsRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRoomIDs)
            If Not rsRoom.EOF Then
                GetRoom = rsRoom!����
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    mblnCancel = True
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdItemSel_Click()
On Error GoTo errHandle
'ѡ��ű�
    Dim strReturn As String
    '�ű�ID,��ĿID,ҽ��ID,ҽ��,����ID,����,����,�ű�,�����¼ID
    If frmNumSortSelNew.ShowMe(mlng�Һ�ID, strReturn, Me) Then
        '���༰�ű�
        lblItem2.Caption = "����:" & Trim(Split(strReturn, ",")(6)) & "   �ű�:" & Trim(Split(strReturn, ",")(7))
        lblItem2.Tag = Trim(Split(strReturn, ",")(7))
        mstr�ű� = Trim(Split(strReturn, ",")(7))
        mlng�����¼ID = Val(Split(strReturn, ",")(8))
        'ִ�в���
        Me.cmbSect2.Text = Trim(Split(strReturn, ",")(5))
        Me.cmbSect2.Tag = CLng(Trim(Split(strReturn, ",")(4)))
        
        '�ҵ�����
        Me.cmbDiagRoom2.Text = GetRoom(mlng�����¼ID)
        
        '����ҽ��
        If Trim(Split(strReturn, ",")(3)) = "" Then
            Me.cmbDocTor2.Text = "��"
        Else
            Me.cmbDocTor2.Text = Trim(Split(strReturn, ",")(3)) & String(400, " ") & STR_COMP & Trim(Split(strReturn, ",")(2))
        End If
    
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOk_Click()
Dim strDoctor  As String
On Error GoTo errHandle

    'NO_IN          ���˹Һż�¼.NO%TYPE:=NULL,
    '�ű�_IN        ���˹Һż�¼.�ű�%TYPE:=NULL,
    '����_IN        ���˹Һż�¼.����%TYPE:=NULL,
    'ִ�в���ID_IN  ���˹Һż�¼.ִ�в���ID%TYPE:=NULL,
    'ҽ��_IN        ���˹Һż�¼.ִ����%TYPE:=NULL,
    'ҽ��ID_IN      ���˹ҺŻ���.ҽ��ID%TYPE:=NULL,
    'ҽ��2_IN       ���˹Һż�¼.ִ����%TYPE:=NULL,
    'ҽ��ID2_IN     ���˹ҺŻ���.ҽ��ID%TYPE:=NULL
    If Trim(lblItem2.Tag) = "" Then MsgBox "��ѡ��һ���ű�", vbInformation, gstrSysName: Exit Sub
    If Trim(zlCommFun.GetNeedName(Trim(Split(cmbDoctor.Text, STR_COMP)(0)))) = "��" Then
        strDoctor = "'',null"
    Else
        strDoctor = "'" & zlCommFun.GetNeedName(Trim(Split(cmbDoctor.Text, STR_COMP)(0))) & "'," & Trim(Split(cmbDoctor.Text, STR_COMP)(1))
    End If
    If Trim(cmbDocTor2.Text) = "��" Then
        strSQL = "'',null"
    Else
        strSQL = "'" & Trim(Split(cmbDocTor2.Text, STR_COMP)(0)) & "'," & Trim(Split(cmbDocTor2.Text, STR_COMP)(1))
    End If
    If ExcPlugInFun(1, mlng�Һ�ID, Trim(Split(cmbDocTor2.Text, STR_COMP)(0)), Me.cmbDiagRoom2.Text, lblItem2.Tag, mlng�����¼ID) = False Then Exit Sub
    
    strSQL = "ZL_���˹Һż�¼_����('" & mstrNo & "','" & lblItem2.Tag & "','" & _
            Me.cmbDiagRoom2.Text & "'," & Me.cmbSect2.Tag & "," & strDoctor & "," & strSQL & "," & mlng�����¼ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    lblItem2.Tag = ""
    Me.cmbSect2.Tag = 0
    Me.cmbDiagRoom2.Text = ""
    Me.cmbDocTor2.Text = ""
End Sub
