VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmIdentify�ɶ��ڽ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���������֤"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin ZL9BillEdit.BillEdit msf�������� 
      Height          =   1275
      Left            =   915
      TabIndex        =   48
      Top             =   4845
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   2249
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.ComboBox cbo��Ժ��� 
      Height          =   300
      ItemData        =   "frmIdentify�ɶ��ڽ�.frx":0000
      Left            =   915
      List            =   "frmIdentify�ɶ��ڽ�.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4035
      Width           =   2295
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   6720
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4470
      Width           =   255
   End
   Begin VB.ComboBox cbo��� 
      Height          =   300
      ItemData        =   "frmIdentify�ɶ��ڽ�.frx":0004
      Left            =   915
      List            =   "frmIdentify�ɶ��ڽ�.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   885
      Width           =   2295
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5925
      TabIndex        =   9
      Top             =   6390
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4635
      TabIndex        =   8
      Top             =   6390
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   12
      Top             =   300
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -465
      TabIndex        =   10
      Top             =   6210
      Width           =   8340
   End
   Begin VB.TextBox TxtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   915
      MaxLength       =   20
      TabIndex        =   1
      Top             =   510
      Width           =   2295
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4605
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   510
      Width           =   2385
   End
   Begin VB.CommandButton cmd�޸����� 
      Caption         =   "�޸�����"
      Height          =   350
      Left            =   225
      TabIndex        =   11
      Top             =   6390
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   315
      Left            =   930
      TabIndex        =   7
      Top             =   4440
      Width           =   6075
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   285
      Left            =   150
      TabIndex        =   47
      Top             =   4845
      Width           =   750
   End
   Begin VB.Label ��Ժ��� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "��Ժ���"
      Height          =   180
      Left            =   135
      TabIndex        =   46
      Top             =   4095
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "����(&F)"
      Height          =   180
      Left            =   240
      TabIndex        =   37
      Top             =   4515
      Width           =   630
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   915
      TabIndex        =   44
      Top             =   3645
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ƿ���λ"
      Height          =   180
      Index           =   17
      Left            =   180
      TabIndex        =   43
      Top             =   3690
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   15
      Left            =   4605
      TabIndex        =   42
      Top             =   3645
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ְ���"
      Height          =   180
      Index           =   16
      Left            =   3825
      TabIndex        =   41
      Top             =   3690
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   4605
      TabIndex        =   40
      Top             =   3240
      Width           =   2385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ʻ����"
      Height          =   180
      Index           =   15
      Left            =   3825
      TabIndex        =   39
      Top             =   3285
      Width           =   720
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "�������������벡�˵�IC�������س�,����ȡ���˵������Ϣ��"
      Height          =   180
      Left            =   720
      TabIndex        =   38
      Top             =   60
      Width           =   5130
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   555
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���˱��"
      Height          =   180
      Index           =   1
      Left            =   3825
      TabIndex        =   36
      Top             =   945
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   540
      TabIndex        =   35
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      Height          =   180
      Index           =   3
      Left            =   4185
      TabIndex        =   34
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���֤��"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   33
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   32
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   6
      Left            =   3825
      TabIndex        =   31
      Top             =   1740
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����Ч��"
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   30
      Top             =   2895
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   8
      Left            =   540
      TabIndex        =   29
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ͳ����"
      Height          =   180
      Index           =   9
      Left            =   3825
      TabIndex        =   28
      Top             =   2130
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Index           =   10
      Left            =   3825
      TabIndex        =   27
      Top             =   2895
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�ƿ�����"
      Height          =   180
      Index           =   11
      Left            =   3825
      TabIndex        =   26
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��λ����"
      Height          =   180
      Index           =   12
      Left            =   180
      TabIndex        =   25
      Top             =   3285
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   4605
      TabIndex        =   24
      Top             =   900
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   915
      TabIndex        =   23
      Top             =   1305
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4605
      TabIndex        =   22
      Top             =   1305
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   915
      TabIndex        =   21
      Top             =   1695
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   5
      Left            =   4605
      TabIndex        =   20
      Top             =   1695
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   915
      TabIndex        =   19
      Top             =   2085
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   7
      Left            =   4605
      TabIndex        =   18
      Top             =   2085
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   915
      TabIndex        =   17
      Top             =   2475
      Width           =   1020
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   4605
      TabIndex        =   16
      Top             =   2475
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   915
      TabIndex        =   15
      Top             =   2850
      Width           =   2295
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   11
      Left            =   4605
      TabIndex        =   14
      Top             =   2850
      Width           =   2385
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   915
      TabIndex        =   13
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   13
      Left            =   4185
      TabIndex        =   2
      Top             =   555
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Index           =   14
      Left            =   180
      TabIndex        =   4
      Top             =   945
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify�ɶ��ڽ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
Private mlng����ID As Long
Private mstrReturn As String
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private mstr����֢ As String '20051026 �¶�

Private Sub cbo��Ժ���_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cbo���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd�޸�����_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    Dim StrInput As String, strOutput As String
    
    If InitInfor_�ɶ��ڽ�.������_�ڽ� = 0 Then
        '�����޸�����
    
        strNewPassWord = frm�޸�����.ChangePassword(strOldPassWord, strOldPassWord)
        
        If strOldPassWord = strNewPassWord Then Exit Sub
        If strNewPassWord = "" Then Exit Sub
        '    a)  Port�����������ΪͨѶ�˿ںţ�0��1��2��3�ֱ������1��2��3��4;����Ϊ��I/O��ַ����0x378�������齫���������ӵ�����1��
        '    b)  OldPassword�����������Ϊԭ���룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֣�
        '    c)  NewPassword�����������Ϊ�����룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֡�
        StrInput = InitInfor_�ɶ��ڽ�.���ź�_�ڽ�
        StrInput = StrInput & vbTab & strOldPassWord
        StrInput = StrInput & vbTab & strNewPassWord
        
        If ҵ������_�ɶ��ڽ�(��������_�ڽ�, StrInput, strOutput) = False Then Exit Sub
        TxtEdit(1).Text = strNewPassWord
    Else
        '����ֻ�ܶ���
    End If
    If ReadCardInFo() = False Then Exit Sub
    Call LoadCtrlData
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    '������������Ķ�����������������
    If InitInfor_�ɶ��ڽ�.������_�ڽ� = 0 Then Exit Sub
    TxtEdit(1).Enabled = False
    TxtEdit(1).BackColor = TxtEdit(0).BackColor
    
    '����
    If ReadCardInFo() = False Then Exit Sub
    Call LoadCtrlData
    Me.cmd�޸�����.Caption = "����(&R)"
    
    
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    'Beging 20051026 �¶�
    Dim i As Long
    Dim vat����֢ As Variant
    
    With msf��������
        '�����������������б�������
        .Rows = 4
        .Cols = 1
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 4800
        .TextMatrix(0, 0) = "���ֱ���������"
        
        '���ø��е���ֵ����ȷ����Щ�пɲ������ɱ༭��Ǳ༭��
        .ColData(0) = 1   '�ı��������������ť
        'δ���õ��е���ֵ��Ϊ0 (Ĭ��), ��Щ�н�����ѡ�񵫲����޸�
    End With
    
    If mstr����֢ <> "" Then
        If InStr(mstr����֢, "|") > 0 Then
            vat����֢ = Split(mstr����֢, "|")
            msf��������.Rows = UBound(vat����֢) + 1
            For i = 0 To UBound(vat����֢) - 1
                msf��������.TextMatrix(i + 1, 0) = "[" & Split(vat����֢(i), ";")(0) & "]"      '& Split(vat����֢(i), ";")(1)
            Next
        End If
    End If
    'End 20051026 �¶�
End Sub


Private Sub msf��������_CommandClick()
    Dim str���� As String
    Select Case msf��������.ColData(msf��������.Col)
        Case 0
            str���� = msf��������.TextMatrix(msf��������.Row, msf��������.Col)
            str���� = BZXZ_�ɶ��ڽ�(str����)
            If str���� = "" Then Exit Sub
            msf��������.TextMatrix(msf��������.Row, msf��������.Col) = str����
    End Select
End Sub

Private Sub msf��������_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim str���� As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    str���� = msf��������.Text
    
    If str���� = "" And msf��������.Rows = msf��������.Row + 1 Then
        SendKeys "{Tab}"
    End If
    
    If str���� = "" And msf��������.Rows = msf��������.Row + 2 Then
        If msf��������.TextMatrix(msf��������.Row + 1, msf��������.Col) = "" Then
            SendKeys "{Tab}"
        End If
    End If
    
    'Cancel = True
    str���� = BZXZ_�ɶ��ڽ�(str����, 1)
    If str���� <> "" Then
        msf��������.Text = str����
        msf��������.TextMatrix(msf��������.Row, msf��������.Col) = str����
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        TxtEdit(Index).Tag = ""
        g�������_�ɶ��ڽ�.���˱�� = ""
        g�������_�ɶ��ڽ�.���� = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    mblnChange = True
    
    If Index = 1 Then
        '�����������
        '���ȡ������Ϣ
         SetOKCtrl False
         If ReadCardInFo = False Then Exit Sub
        '��ʼֵ
        Call LoadCtrlData
        SetOKCtrl True
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
Private Function ReadCardInFo() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ������Ϣ
    '--�����:
    '--������:
    '--��  ��:�ɹ�,����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim StrInput As String
     '��ȡ������Ϣ
        '   a)  Port�����������ΪͨѶ�˿ںţ�0��1��2��3�ֱ������1��2��3��4;����Ϊ��I/O��ַ����0x378�������齫���������ӵ�����1��
        '   b)  UserPassword�����������Ϊ�û����룬Ҫ�󳤶�Ϊ6���ַ�����ֻ�ܰ���0��9�����֣�
           
    ReadCardInFo = False
    StrInput = InitInfor_�ɶ��ڽ�.���ź�_�ڽ�
    If InitInfor_�ɶ��ڽ�.������_�ڽ� = 0 Then
        '��������������
        If Trim(TxtEdit(1)) = "" Then
            ShowMsgbox "������IC������!"
            If TxtEdit(1).Enabled Then TxtEdit(1).SetFocus
            Exit Function
        End If
        StrInput = StrInput & vbTab & TxtEdit(1).Text
    End If
    
    Err = 0
    On Error GoTo ErrHand:
    
    If ��ȡ�α���Ա��Ϣ_�ɶ��ڽ�(StrInput) = False Then Exit Function
    ReadCardInFo = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmdȷ��.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��֤���ݵĺϷ���
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    IsValid = False
    If Trim(TxtEdit(0).Text) = "" Then
        MsgBox "��û�н��������֤��", vbInformation, gstrSysName
        If TxtEdit(1).Enabled Then TxtEdit(1).SetFocus
        Exit Function
    End If
    
    If Trim(g�������_�ɶ��ڽ�.����) = "" Then
        MsgBox "��û���������֤��", vbInformation, gstrSysName
        If TxtEdit(1).Enabled Then TxtEdit(1).SetFocus
        Exit Function
    End If
    
    If cbo���.Text = "" Then
        ShowMsgbox "�������δѡ��"
        Exit Function
    End If
    If cbo��Ժ���.Text = "" And mbytType = 4 Then
        ShowMsgbox "��Ժ���δѡ��"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '�����¼ǰ��̬
        Else
            '��鲡��״̬
            gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=" & TYPE_�ɶ��ڽ� & " and ҽ����='" & g�������_�ɶ��ڽ�.ͳ���� & g�������_�ɶ��ڽ�.���˱�� & "'"
            Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("״̬") > 0 Then
                    MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '����
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�ɶ��ڽ� & " and  ҽ����='" & g�������_�ɶ��ڽ�.ͳ���� & g�������_�ɶ��ڽ�.���˱�� & "'"
        
        zldatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
        End If
        rsTemp.Close
        mstrReturn = mlng����ID
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str��� As String
    Dim int��ǰ״̬ As Integer
    Dim lng��Ժ����ID As Long
    

    lng����ID = IIf(Val(Me.txt����.Tag) = 0, 0, Val(Me.txt����.Tag))
    
    If lng����ID <> 0 Then
        gstrSQL = "Select * From ���ղ��� where id=" & lng����ID
        zldatabase.OpenRecordset rsTemp, gstrSQL, "��ȡ����"
        g�������_�ɶ��ڽ�.���ֱ��� = Nvl(rsTemp!����)
        g�������_�ɶ��ڽ�.�������� = Nvl(rsTemp!����)
    Else
        g�������_�ɶ��ڽ�.���ֱ��� = ""
        g�������_�ɶ��ڽ�.�������� = ""
        If mbytType = 1 Or mbytType = 4 Then
            ShowMsgbox "�����ѡ����!"
            Exit Sub
        End If
    End If
        
    If IsValid = False Then Exit Sub
    'Beging 20051026 ��������
    If mbytType = 1 Or mbytType = 4 Then
        Dim str�������� As String, str�������ֱ��� As String, i As Long
        
        For i = 1 To msf��������.Rows - 1
            str�������ֱ��� = msf��������.TextMatrix(i, 0)
            If str�������ֱ��� <> "" Then
                If InStr(str�������ֱ���, "]") > 0 And InStr(str�������ֱ���, "[") > 0 And InStr(str�������ֱ���, "]") - InStr(str�������ֱ���, "[") > 1 Then
                    str�������ֱ��� = Mid(str�������ֱ���, InStr(str�������ֱ���, "[") + 1, InStr(str�������ֱ���, "]") - InStr(str�������ֱ���, "[") - 1)
                    str�������� = str�������� & str�������ֱ��� & "|"
                End If
            End If
        Next
    End If
    'End 20051026 ��������
    g�������_�ɶ��ڽ�.������� = Split(cbo���.Text, "-")(0)
    If mbytType = 4 Then
    g�������_�ɶ��ڽ�.��Ժ��� = Split(cbo��Ժ���.Text, "-")(0)
    End If
    int��ǰ״̬ = 0
    If mbytType = 4 Then
        '��ȷ����ǰ״̬,��Ϊ��ǰ״̬�ǲ��ܸı��
        gstrSQL = "Select * from �����ʻ� where ����=" & TYPE_�ɶ��ڽ� & " and  ҽ����='" & g�������_�ɶ��ڽ�.ͳ���� & g�������_�ɶ��ڽ�.���˱�� & "'"
        
        zldatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng����ID = Nvl(rsTemp!����ID, 0)
            int��ǰ״̬ = Nvl(rsTemp!��ǰ״̬, 0)
            '>>Beging �¶� 20050601
            lng��Ժ����ID = Nvl(rsTemp!����ID, 0)
            '>> End
        End If
        rsTemp.Close
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤(�������);7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��(ͳ���������|�ƿ�����|����Ч����);16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    With g�������_�ɶ��ڽ�
        
        strIdentify = .����                               '0����
        strIdentify = strIdentify & ";" & .ͳ���� & .���˱��              '1ҽ����
        strIdentify = strIdentify & ";"                     '2����
        strIdentify = strIdentify & ";" & .����               '3����
        strIdentify = strIdentify & ";" & Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)              '4�Ա�
        strIdentify = strIdentify & ";" & .��������                '5��������
        strIdentify = strIdentify & ";" & .���֤��           '6���֤
        strIdentify = strIdentify & ";" & IIf(.��λ���� = "", "", "(" & .��λ���� & ")")            '7.��λ����(����)
        strAddition = ";0"                                          '8.���Ĵ���
        strAddition = strAddition & ";"                             '9.˳���
        strAddition = strAddition & ";" & .�������                 '10��Ա���
        strAddition = strAddition & ";" & .�ʻ����                 '11�ʻ����
        
        strAddition = strAddition & ";" & int��ǰ״̬               '12��ǰ״̬
            'beging �¶� 20050601 ��Ժʱ,�ǳ�Ժ����,���ܽ���Ժ���ֳ����
        If mbytType = 4 Then
            strAddition = strAddition & ";" & lng��Ժ����ID                 '13����ID
        Else
            strAddition = strAddition & ";" & lng����ID                 '13����ID
        End If
            'End
        strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
        strAddition = strAddition & ";" & .ͳ���� & "|" & .�ƿ����� & "|" & .����Ч�� & "|" & .�ƿ���λ & "|" & .��ְ���    '15����֤��
        strAddition = strAddition & ";" & .��������                     '16�����
        strAddition = strAddition & ";" & .�������                            '17�Ҷȼ�
        strAddition = strAddition & ";" & .�ʻ����                             '18�ʻ������ۼ�
        strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
        strAddition = strAddition & ";0"                            '20���깤���ܶ�
        strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    End With
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_�ɶ��ڽ�)
    
    g�������_�ɶ��ڽ�.lng����ID = mlng����ID
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
        'Beging �¶� 20050601 �����Ժ����ID
        If mbytType = 4 Then
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ɶ��ڽ� & ",'��Ժ����ID','" & lng����ID & "')"
            Call zldatabase.ExecuteProcedure(gstrSQL, "���³�Ժ����")
            'beging 20051026 �¶�
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ɶ��ڽ� & ",'��Ժ��������','''" & str�������� & "''')"
            Call zldatabase.ExecuteProcedure(gstrSQL, "���³�Ժ��������")
        End If
            '
        If mbytType = 1 Then
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ɶ��ڽ� & ",'��������','''" & str�������� & "''')"
            Call zldatabase.ExecuteProcedure(gstrSQL, "������������")
            gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ɶ��ڽ� & ",'�������','''" & "" & "''')"
            Call zldatabase.ExecuteProcedure(gstrSQL, "���¸������")
            'End 20051026 �¶�
        End If
        'end
        
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    Dim rsTmp As New ADODB.Recordset
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    
    DebugTool "���������֤,����ʼ���������Ϣ"
    
    If LoadBaseData = False Then
        DebugTool "����ʧ��(�����֤)"
        Exit Function
    End If
    DebugTool "����ɹ�(�����֤)"
    
    'Beging 20051026 �¶�
    gstrSQL = "Select * from �����ʻ� Where ����ID=" & lng����ID & " And ����=" & TYPE_�ɶ��ڽ�
    Call zldatabase.OpenRecordset(rsTmp, gstrSQL, "ȡ����֢")
    If rsTmp.EOF = False Then
        If mbytType = 4 Then
            mstr����֢ = Nvl(rsTmp!��Ժ��������)
        Else
            mstr����֢ = Nvl(rsTmp!��������)
        End If
    End If
    'End 20051026 �¶�
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function
Private Function LoadBaseData() As Boolean
    '���ػ�������
    Dim rsTemp As New ADODB.Recordset
    LoadBaseData = False
    On Error GoTo ErrHand:
      
    If mbytType = 0 Or mbytType = 3 Or mbytType = 2 Then
        cbo���.AddItem "0-��ͨ����"
    Else
        cbo���.AddItem "1-��ͨסԺ"
    End If
    cbo���.ListIndex = cbo���.NewIndex
    If mbytType = 4 Then
        cbo��Ժ���.AddItem "0-������Ժ"
        cbo��Ժ���.AddItem "1-��ת��Ժ"
        cbo��Ժ���.AddItem "2-δ����Ժ"
        cbo��Ժ���.AddItem "3-����"
        cbo��Ժ���.AddItem "4-�Զ���Ժ"
        cbo��Ժ���.AddItem "5-ת��ͳ������ڵ�ҽԺ"
        cbo��Ժ���.AddItem "6-ת��ͳ��������ҽԺ"
        cbo��Ժ���.ListIndex = 0
    Else
        cbo��Ժ���.Enabled = False
    End If
    If mbytType = 0 Or mbytType = 3 Or mbytType = 2 Then
       msf��������.TabIndex = 47
    End If
    LoadBaseData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    With g�������_�ɶ��ڽ�
        TxtEdit(0) = .����
        lblEdit(1) = .���˱��
        lblEdit(2) = .����
        lblEdit(3) = Decode(.�Ա�, "1", "��", "2", "Ů", .�Ա�)
        lblEdit(4) = .���֤��
        lblEdit(5) = .�������
        lblEdit(6) = .��������
        lblEdit(7) = .ͳ����
        lblEdit(8) = .����
        lblEdit(9) = .�ƿ�����
        lblEdit(10) = .����Ч��
        lblEdit(11) = .��������
        lblEdit(12) = .��λ����
        lblEdit(13) = Format(.�ʻ����, "####0.00;-#####0.00; ;")
        lblEdit(14) = .�ƿ���λ
        lblEdit(15) = .��ְ���
   End With
End Sub

Private Sub cmd����_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
            " From ���ղ��� A where A.����=" & TYPE_�ɶ��ڽ�
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txt����.Text)
    If rsTemp.State = 0 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
        zlControl.TxtSelAll txt����
    End If
    txt����.SetFocus
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
    txt����.ForeColor = &HC0&
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt����.Text = "" Or txt����.Tag <> "" Then
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt����.Text
    
    
    gstrSQL = "Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ⲡ','��ͨ��') ��� " & _
             "   FROM ���ղ��� A WHERE A.����=" & TYPE_�ɶ��ڽ� & " And (" & _
                zlCommFun.GetLike("A", "����", strText) & " or " & zlCommFun.GetLike("A", "����", strText) & " or " & zlCommFun.GetLike("A", "����", strText) & ")"
    Call zldatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_�ɶ��ڽ�, rsTemp, "ID", "ҽ������ѡ��", "��ѡ���ض���ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        zlControl.TxtSelAll txt����
        Exit Sub
    Else
        '�϶����м�¼����
        txt����.Text = rsTemp("����")
        txt����.Tag = rsTemp("ID")
        SendKeys "{TAB}"
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt����.Text = ""
        txt����.Tag = ""
    End If
End Sub

Function BZXZ_�ɶ��ڽ�(ByVal StrInput As String, Optional strLoad As String = 0) As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmpSQL As String
    
    On Error Resume Next
   
    
    If StrInput = "" And strLoad = 1 Then Exit Function
    
    If StrInput = "" Then
        strTmpSQL = "Select ID,����,���� from ���ղ���"
    Else
        strTmpSQL = "Select ID,����,���� from ���ղ���" & _
                 " Where ���� Like '%" & StrInput & "%' OR " & _
                 "���� like '%" & StrInput & "%' Or " & _
                 "lower(����) like lower('%" & StrInput & "%')"
    End If
    
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "����", True, , , , False, gcnOracle)
    If rsTmp Is Nothing Then Exit Function
    BZXZ_�ɶ��ڽ� = "[" & rsTmp!���� & "]" & rsTmp!����
End Function

