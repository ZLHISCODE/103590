VERSION 5.00
Begin VB.Form frmClinicBill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ŀ���Ƶ���"
   ClientHeight    =   3192
   ClientLeft      =   2760
   ClientTop       =   3756
   ClientWidth     =   5448
   Icon            =   "frmClinicBill.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3192
   ScaleWidth      =   5448
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboTest 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.CheckBox chkTest 
      Caption         =   "������(&M)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   675
      Width           =   1290
   End
   Begin VB.CheckBox chkIn 
      Caption         =   "סԺ����(&I)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   660
      TabIndex        =   4
      Top             =   1380
      Width           =   1290
   End
   Begin VB.CheckBox chkOut 
      Caption         =   "�������(&T)"
      Enabled         =   0   'False
      Height          =   195
      Left            =   660
      TabIndex        =   2
      Top             =   1035
      Width           =   1290
   End
   Begin VB.ComboBox cboIn 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1335
      Width           =   3255
   End
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   -180
      TabIndex        =   13
      Top             =   525
      Width           =   6615
   End
   Begin VB.OptionButton optScope 
      Caption         =   "���ڱ�������Ŀ"
      Height          =   195
      Index           =   1
      Left            =   660
      TabIndex        =   7
      Top             =   2115
      Width           =   5610
   End
   Begin VB.Frame fraBottom 
      Height          =   30
      Left            =   -165
      TabIndex        =   12
      Top             =   2490
      Width           =   6585
   End
   Begin VB.OptionButton optScope 
      Caption         =   "���ڱ���Ŀ"
      Height          =   195
      Index           =   0
      Left            =   660
      TabIndex        =   6
      Top             =   1800
      Value           =   -1  'True
      Width           =   5610
   End
   Begin VB.ComboBox cboOut 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1965
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4170
      TabIndex        =   9
      Top             =   2655
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   150
      Picture         =   "frmClinicBill.frx":058A
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2655
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3060
      TabIndex        =   8
      Top             =   2655
      Width           =   1100
   End
   Begin VB.Image imgNote 
      Height          =   384
      Left            =   156
      Picture         =   "frmClinicBill.frx":06D4
      Top             =   12
      Width           =   384
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ����������Ŀ��Ӧ�����Ƶ��ݣ��Ա���ҽ������ִ�й����У����÷�����Ŀ���Եĵ��ݣ��������ƹ�����Ҫ��"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   645
      TabIndex        =   10
      Top             =   75
      Width           =   4680
   End
End
Attribute VB_Name = "frmClinicBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1����ǰ��Ŀ����me.optScope(0).tag���棬���ϼ�����ͨ��ShowMe��������
'---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim strTemp As String
Dim intCount As Integer

Public Sub ShowMe(ByVal frmParent As Object, Optional ByVal lng��Ŀid As Long)
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    Dim str������� As String
    Dim intControl As Integer       '�������Ƴ�ʼ��ʱ��ѡ��Ĺ�ѡ��0-����ѡ;1-��ѡ
    
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.���,I.����,I.����,I.����id,nvl(I.�������,0) as �������,K.���� as �����,K.���� as �����" & _
            " from ������ĿĿ¼ I,������Ŀ��� K" & _
            " where I.id=[1] and I.���=K.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng��Ŀid)
    
    With rsTemp
        If .BOF Or .EOF Then Unload Me: Exit Sub
        str������� = !���
        Me.optScope(0).Tag = !ID: Me.optScope(0).Caption = "&1��Ӧ���ڱ���Ŀ(" & !���� & "-" & !���� & ")"
        Me.optScope(1).Tag = !�����: Me.optScope(1).Caption = "&2��Ӧ�������С�" & !����� & "������Ŀ"
        
        If !������� = 1 Or !������� = 3 Then Me.chkOut.Enabled = True
        If !������� = 2 Or !������� = 3 Then Me.chkIn.Enabled = True
        If !������� = 4 Then Me.chkTest.Enabled = True
    End With
    
    gstrSql = "select ID,����,����" & _
            " from ���Ʒ���Ŀ¼" & _
            " start with id=[1] " & _
            " connect by prior �ϼ�id=id" & _
            " order by level"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rsTemp!����id))
        
    With rsTemp
        Do While Not .EOF
            Load Me.optScope(.AbsolutePosition + 1)
            Me.optScope(.AbsolutePosition + 1).Tag = !ID
            Me.optScope(.AbsolutePosition + 1).Caption = "&" & .AbsolutePosition + 2 & "��Ӧ���ڡ�[" & !���� & "]" & !���� & "������Ŀ"
            Me.optScope(.AbsolutePosition + 1).Left = Me.optScope(0).Left
            Me.optScope(.AbsolutePosition + 1).Top = Me.optScope(.AbsolutePosition).Top + Me.optScope(1).Top - Me.optScope(0).Top
            Me.optScope(.AbsolutePosition + 1).Visible = True
            .MoveNext
        Loop
    End With
    
    If Me.chkOut.Enabled Then
        gstrSql = "select �����ļ�id from ��������Ӧ�� where Ӧ�ó���=1 and ������Ŀid=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng��Ŀid)
        
        
        If Not rsTemp.EOF Then
            Me.cboOut.Tag = rsTemp!�����ļ�id
            intControl = 1
        Else
            intControl = 0
        End If
        
        gstrSql = "select ID,���,���� from �����ļ��б� where ����=7 "
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'            Call SQLTest
        With rsTemp
            If .EOF Or .BOF Then
                Me.chkOut.Value = 0: Me.chkOut.Enabled = False
            ElseIf intControl = 0 Then
                Me.chkOut.Value = 0: Me.cboOut.Enabled = False
            Else
                Me.chkOut.Value = 1: Me.cboOut.Enabled = True
            End If
            Me.cboOut.ListIndex = -1
            Do While Not .EOF
                Me.cboOut.AddItem !��� & "-" & !����
                Me.cboOut.ItemData(Me.cboOut.NewIndex) = !ID
                If !ID = Val(Me.cboOut.Tag) Then
                    Me.cboOut.ListIndex = Me.cboOut.NewIndex
                End If
                .MoveNext
            Loop
       
        End With
    End If
    If cboOut.ListIndex = -1 Then
        'ҩƷ����Ĭ�ϵ���Ŀ����ҩ��Ӧ��ҩ����ǩ���в�ҩ��Ӧ��ҩ����ǩ
        '��������ļ��б����漰ҩƷ�����ݻ�˳�����˸ı䣬����������ҲҪ����Ӧ����
        If str������� = "5" Or str������� = "6" Then
            cboOut.ListIndex = 0
        ElseIf str������� = "7" Then
            cboOut.ListIndex = 1
        Else
            cboOut.Enabled = False: chkOut.Value = 0
        End If
    End If
    chkOut.Tag = cboOut.ListIndex
    
    If Me.chkIn.Enabled Then
        gstrSql = "select �����ļ�id from ��������Ӧ�� where Ӧ�ó���=2 and ������Ŀid=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng��Ŀid)
        
        If Not rsTemp.EOF Then
            Me.cboIn.Tag = rsTemp!�����ļ�id
            intControl = 1
        Else
            intControl = 0
        End If
        
        gstrSql = "select ID,���,���� from �����ļ��б� where ����=7 "
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'            Call SQLTest
        With rsTemp
            If .EOF Or .BOF Then
                Me.chkIn.Value = 0: Me.chkIn.Enabled = False
            ElseIf intControl = 0 Then
                Me.chkIn.Value = 0: Me.cboIn.Enabled = False
            Else
                Me.chkIn.Value = 1: Me.cboIn.Enabled = True
            End If
            Me.cboIn.ListIndex = -1
            Do While Not .EOF
                Me.cboIn.AddItem !��� & "-" & !����
                Me.cboIn.ItemData(Me.cboIn.NewIndex) = !ID
                If !ID = Val(Me.cboIn.Tag) Then
                    Me.cboIn.ListIndex = Me.cboIn.NewIndex
                End If
                .MoveNext
            Loop
            
        End With
    End If
    If Me.cboIn.ListIndex = -1 Then
        'ҩƷ����Ĭ�ϵ���Ŀ����ҩ��Ӧ��ҩ����ǩ���в�ҩ��Ӧ��ҩ����ǩ
        '��������ļ��б����漰ҩƷ�����ݻ�˳�����˸ı䣬����������ҲҪ����Ӧ����
        If str������� = "5" Or str������� = "6" Then
            cboIn.ListIndex = 0
        ElseIf str������� = "7" Then
            cboIn.ListIndex = 1
        Else
            cboIn.Enabled = False: chkIn.Value = 0
        End If
    End If
    chkIn.Tag = cboIn.ListIndex
    
    If chkTest.Enabled Then
        gstrSql = "select �����ļ�id from ��������Ӧ�� where Ӧ�ó���=4 and ������Ŀid=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng��Ŀid)
        
        If Not rsTemp.EOF Then
            Me.cboTest.Tag = rsTemp!�����ļ�id
            intControl = 1
        Else
            intControl = 0
        End If
        
        gstrSql = "select ID,���,���� from �����ļ��б� where ����=7 "
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.Title, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'            Call SQLTest
        With rsTemp
            If .EOF Or .BOF Then
                chkTest.Value = 0: chkTest.Enabled = False
            ElseIf intControl = 0 Then
                chkTest.Value = 0: cboTest.Enabled = False
            Else
                chkTest.Value = 1: cboTest.Enabled = True
            End If
            cboTest.ListIndex = -1
            Do While Not .EOF
                cboTest.AddItem !��� & "-" & !����
                cboTest.ItemData(cboTest.NewIndex) = !ID
                If !ID = Val(cboTest.Tag) Then
                    cboTest.ListIndex = cboTest.NewIndex
                End If
                .MoveNext
            Loop
        End With
    End If
    chkTest.Tag = cboTest.ListIndex
    If cboTest.ListIndex = -1 Then
        cboTest.Enabled = False: chkTest.Value = 0
    End If
    
    Me.optScope(0).Value = True
    Me.fraBottom.Top = Me.optScope(Me.optScope.Count - 1).Top + 300
    Me.cmdHelp.Top = Me.fraBottom.Top + 150
    Me.cmdOK.Top = Me.cmdHelp.Top: Me.cmdCancel.Top = Me.cmdHelp.Top
    Me.Height = Me.cmdHelp.Top + Me.cmdHelp.Height + 500
    Me.Show 1, frmParent
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub cboIn_Click()
    chkIn.Tag = Me.cboIn.ListIndex
End Sub

Private Sub cboOut_Click()
    chkOut.Tag = cboOut.ListIndex
End Sub

Private Sub cboTest_Click()
    chkTest.Tag = cboTest.ListIndex
End Sub

Private Sub chkIn_Click()
    If Me.chkIn.Value = 1 Then
        Me.cboIn.Enabled = True
        If Me.cboIn.ListCount > 0 Then
            Me.cboIn.ListIndex = Val(chkIn.Tag)
        Else
            Me.cboIn.ListIndex = -1
        End If
    Else
        Me.cboIn.Enabled = False
    End If
End Sub

Private Sub chkOut_Click()
    If Me.chkOut.Value = 1 Then
        Me.cboOut.Enabled = True
        If Me.cboOut.ListCount > 0 Then
            Me.cboOut.ListIndex = Val(chkOut.Tag)
        Else
            Me.cboOut.ListIndex = -1
        End If
    Else
        Me.cboOut.Enabled = False
    End If
End Sub

Private Sub chkTest_Click()
    If Me.chkTest.Value = 1 Then
        Me.cboTest.Enabled = True
        If cboTest.ListCount > 0 Then
            Me.cboTest.ListIndex = Val(chkTest.Tag)
        Else
            Me.cboTest.ListIndex = -1
        End If
    Else
        Me.cboTest.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    If optScope(0).Value = False Then
        For i = 1 To optScope.UBound
            If optScope(i).Value = True Then
                If MsgBox("��ҩƷ���Ƶ���Ӧ�÷�ΧΪ��" & optScope(i).Caption & "���Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
    End If
        
    gstrSql = "zl_���Ƶ���Ӧ��_Update("
    
    If Me.cboOut.Enabled = False Or Me.cboOut.ListIndex = -1 Then
        gstrSql = gstrSql & "null"
    Else
        gstrSql = gstrSql & Me.cboOut.ItemData(Me.cboOut.ListIndex)
    End If
    
    If Me.cboIn.Enabled = False Or Me.cboIn.ListIndex = -1 Then
        gstrSql = gstrSql & ",null"
    Else
        gstrSql = gstrSql & "," & Me.cboIn.ItemData(Me.cboIn.ListIndex)
    End If
    
    If Me.optScope(0).Value = True Then
        gstrSql = gstrSql & ",0,'" & Me.optScope(0).Tag & "'"
    ElseIf Me.optScope(1).Value = True Then
        gstrSql = gstrSql & ",1,'" & Me.optScope(1).Tag & "'"
    Else
        For intCount = 2 To Me.optScope.Count - 1
            If Me.optScope(intCount).Value = True Then
                gstrSql = gstrSql & ",2,'" & Me.optScope(intCount).Tag & "'"
                Exit For
            End If
        Next
    End If
    
    If cboTest.Enabled = False Or cboTest.ListIndex = -1 Then
        gstrSql = gstrSql & ",null)"
    Else
        gstrSql = gstrSql & "," & cboTest.ItemData(cboTest.ListIndex) & ")"
    End If
    
    err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub optScope_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optScope.UBound
        If i = Index Then
            optScope(i).FontBold = True
        Else
            optScope(i).FontBold = False
        End If
    Next
End Sub
