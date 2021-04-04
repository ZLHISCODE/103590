VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStPathItemEdit 
   Caption         =   "��������"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   Icon            =   "frmStPathItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   5910
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5910
      TabIndex        =   7
      Top             =   5460
      Width           =   5910
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   4680
         TabIndex        =   9
         Top             =   160
         Width           =   1100
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   3240
         TabIndex        =   8
         Top             =   160
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.ComboBox cboNo 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Text            =   "cboNO"
      Top             =   82
      Width           =   2295
   End
   Begin VB.CheckBox chkContinual 
      Caption         =   "��������"
      Height          =   225
      Left            =   4675
      TabIndex        =   6
      Top             =   120
      Value           =   1  'Checked
      Width           =   1100
   End
   Begin RichTextLib.RichTextBox rtfContent 
      Height          =   4335
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7646
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmStPathItemEdit.frx":59D62
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   581
      Width           =   4935
   End
   Begin VB.Label lblContent 
      Caption         =   "����(&Q)"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblItemTile 
      Caption         =   "����(&T)"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   619
      Width           =   720
   End
   Begin VB.Label lblNO 
      Caption         =   "���(&N)"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmStPathItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintMode As Integer '0-����������Ŀ��1-�޸�·��������Ŀ��2-ɾ��·��������Ŀ
Private mlngStPathID As Long
Private mlng��� As Long
Private mrsCourseItems As New ADODB.Recordset
Private mblnOK As Boolean '�Ƿ���������ݲ���

Public Function ShowMe(ByRef FrmParent As Object, ByVal intMode As Integer, ByVal lngStPathID As Long, Optional ByVal lng��� As Long) As Boolean
'���ܣ���ʾ·��������Ŀ��ɾ�Ľ���
'������ intMode 0-����������Ŀ��1-�޸�·��������Ŀ��2-ɾ��·��������Ŀ
'       lngStPathID ��׼·��ID
'       lng��� ·��������Ŀ���

    mintMode = intMode
    mlngStPathID = lngStPathID
    mlng��� = lng���
    mblnOK = False
    Me.Show 1, FrmParent
    ShowMe = mblnOK
    
End Function


Private Sub cboNo_Click()
'���ܣ������������б����½�������
    Dim strSel As String
    
    If cboNo.ListIndex = -1 Then Exit Sub
    
    strSel = cboNo.List(cboNo.ListIndex)
    
    If InStr(strSel, "-") > 0 Then
        cboNo.Text = Mid(strSel, 1, InStr(strSel, "-") - 1)
    Else
        cboNo.Text = strSel
    End If
    mlng��� = Val(cboNo.Text)
    
    If Me.Visible Then
        mrsCourseItems.Filter = "���=" & Val(cboNo.Text)
        If mrsCourseItems.RecordCount <> 0 Then
            txtTitle.Text = IIf(mintMode <> 0, mrsCourseItems!���� & "", "")
            rtfContent.Text = IIf(mintMode <> 0, mrsCourseItems!���� & "", "")
        End If
        mrsCourseItems.Filter = ""
    Else '��ʼ���������б�
        txtTitle.Text = IIf(mintMode <> 0, mrsCourseItems!���� & "", "")
        rtfContent.Text = IIf(mintMode <> 0, mrsCourseItems!���� & "", "")
    End If
    
End Sub

Private Sub cboNo_KeyPress(KeyAscii As Integer)
'���ܣ�������

    'ֻ�������������Լ��س�
    If Not (InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then KeyAscii = 0: Exit Sub
    'if KeyAscii = vbKeyReturn then
    '�ڷǲ���·��������Ŀ������£��������ֲ��������������
    If Val(cboNo.Text & Chr(KeyAscii)) > mrsCourseItems.RecordCount And mintMode <> 0 Then KeyAscii = 0: Exit Sub
    '�ڲ���·��������Ŀ������£��������ֲ��������������+1
    If Val(cboNo.Text & Chr(KeyAscii)) > mrsCourseItems.RecordCount + 1 And mintMode = 0 Then KeyAscii = 0: Exit Sub

End Sub

Private Sub cboNo_LostFocus()
    Call cboNo_Click
End Sub

Private Sub cboNo_Validate(Cancel As Boolean)
    Dim strSel As String
    
    strSel = cboNo.Text
    
    If InStr(strSel, "-") > 0 Then
        cboNo.Text = Mid(strSel, 1, InStr(strSel, "-") - 1)
    Else
        cboNo.Text = strSel
    End If
    
    If Val(cboNo.Text) = 0 Then
        cboNo.Text = mlng���
    Else
        mlng��� = Val(cboNo.Text)
    End If
    
     '�ڷǲ���·��������Ŀ������£��������ֲ��������������
    If Val(cboNo.Text) > mrsCourseItems.RecordCount And mintMode <> 0 Then
        mlng��� = mrsCourseItems.RecordCount
        cboNo.Text = mlng���
        Exit Sub
    End If
    
    '�ڲ���·��������Ŀ������£��������ֲ��������������+1
    If Val(cboNo.Text) > mrsCourseItems.RecordCount + 1 And mintMode = 0 Then
        mlng��� = mrsCourseItems.RecordCount + 1
        cboNo.Text = mlng���
        Exit Sub
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
'���ܣ��������ݸ��£����Ҹ����Ƿ���������������ʼ���´ν�����˳�
    Dim strSql As String
    
    mblnOK = True
    
    On Error GoTo errH
    Select Case mintMode
        Case 0
            strSql = "Zl_��׼·������_Insert(" & mlngStPathID & "," & mlng��� & ",'" & Trim(txtTitle.Text) & "','" & Trim(rtfContent.Text) & "')"
        Case 1
            strSql = "Zl_��׼·������_Update(" & mlngStPathID & "," & mlng��� & ",'" & Trim(txtTitle.Text) & "','" & Trim(rtfContent.Text) & "')"
        Case 2
            strSql = "Zl_��׼·������_Delete(" & mlngStPathID & "," & mlng��� & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    If chkContinual.Value = 0 Then
        Unload Me
    Else '��������-��ʼ����������

        Call GetCourseItems
        'ɾ�����Զ��˳�
        If mrsCourseItems.RecordCount = 0 And mintMode = 2 Then Unload Me
        '�������
        If mlng��� <= mrsCourseItems.RecordCount Then
            mlng��� = mlng��� + IIf(mintMode <> 2, 1, 0)
        Else
            mlng��� = mrsCourseItems.RecordCount + IIf(mintMode <> 2, 1, 0)
        End If
        '�޸ĵ����һ��ʱ�˳�
        If mlng��� > mrsCourseItems.RecordCount And mintMode = 1 Then Unload Me
        cboNo.Text = mlng���
        
        If mintMode = 0 Then '����ʱ�������
            txtTitle.Text = ""
            rtfContent.Text = ""
        End If
        
        Call InitcboNo '���������б�
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)
'���ܣ��س�����click�¼�
    If KeyAscii = vbKeyReturn Then
        Call cmdUpdate_Click
    End If
End Sub

Private Sub Form_Activate()

    If mintMode <> 2 Then
        txtTitle.SetFocus
        txtTitle.SelStart = 0
        txtTitle.SelLength = Len(txtTitle.Text)
    Else
        cboNo.SetFocus
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'���ܣ��س���λ��һ���ؼ�
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
'���ܣ����ݲ������ͳ�ʼ������
    Select Case mintMode
        Case 0
            chkContinual.Caption = "��������"
            Me.Caption = "���Ӷ���"
        Case 1
            chkContinual.Caption = "�����޸�"
            Me.Caption = "�޸Ķ���"
        Case 2
            chkContinual.Caption = "����ɾ��"
            Me.Caption = "ɾ������"
            txtTitle.Locked = True
            rtfContent.Locked = True
    End Select

    Call InitcboNo
     
End Sub

Private Sub GetCourseItems()
'���ܣ���ȡ��ǰ��׼·��������·��������Ŀ
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select a.���, a.����, a.���� From ��׼·������ A Where ��׼·��id = [1] Order By a.���"
    Set mrsCourseItems = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngStPathID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitcboNo()
'���ܣ���ʼ�������б�
    Dim i As Long
    
    Call GetCourseItems
    Call cboNo.Clear
    With mrsCourseItems
        If .RecordCount = 0 Then cboNo.Text = mlng���: Exit Sub
        .MoveFirst
        For i = 1 To .RecordCount
            cboNo.AddItem !��� & "-" & !����
            
            If !��� = mlng��� Then
                cboNo.ListIndex = cboNo.NewIndex
            End If
            .MoveNext
        Next
        cboNo.Text = mlng���
        If mintMode = 0 Then
            cboNo.AddItem .RecordCount + 1
        End If
    End With

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then Exit Sub
    '������ı䴰���С
    If Me.Width < 6150 Then Me.Width = 6150
    If Me.Height < 6650 Then Me.Height = 6650
    
End Sub

Private Sub rtfContent_GotFocus()
'���ܣ�����������ý���ʱ����ȫѡ
    rtfContent.SelStart = Len(rtfContent.Text)
End Sub

Private Sub txtTitle_GotFocus()
'���ܣ�����������ý���ʱ����ȫѡ
    txtTitle.SelStart = 0
    txtTitle.SelLength = Len(txtTitle.Text)
End Sub
