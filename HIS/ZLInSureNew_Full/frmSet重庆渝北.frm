VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSet�����山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���в�������"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmSet�����山.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComCtl2.UpDown upDown 
      Height          =   300
      Left            =   1890
      TabIndex        =   16
      Top             =   3555
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   3
      BuddyControl    =   "txt��ʾ"
      BuddyDispid     =   196622
      OrigLeft        =   2100
      OrigTop         =   3600
      OrigRight       =   2340
      OrigBottom      =   3945
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt��ʾ 
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "3"
      Top             =   3555
      Width           =   360
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "��"
      Height          =   285
      Left            =   5280
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3090
      Width           =   255
   End
   Begin VB.TextBox txt��Ŀ 
      Height          =   300
      Left            =   1455
      TabIndex        =   12
      Top             =   3075
      Width           =   4095
   End
   Begin VB.ComboBox cbo������Ŀ 
      Height          =   300
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2685
      Width           =   2415
   End
   Begin VB.Frame fra 
      Caption         =   "��������ȷ��"
      Height          =   630
      Left            =   180
      TabIndex        =   20
      Top             =   1890
      Width           =   5655
      Begin VB.TextBox Txt�޶� 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1260
         TabIndex        =   8
         Text            =   "200.00"
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�������۴���                  Ԫ����������Ϣ"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   3960
      End
   End
   Begin VB.Frame fraҽ�������� 
      Caption         =   "ҽԺǰ��ҽ��������"
      Height          =   1605
      Left            =   165
      TabIndex        =   19
      Top             =   195
      Width           =   4155
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   1
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1110
         Width           =   1635
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "����(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�û���(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   0
         Top             =   390
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   2
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   4
         Top             =   1170
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4560
      TabIndex        =   18
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4560
      TabIndex        =   17
      Top             =   300
      Width           =   1100
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���ܴ���ʱ��С��       ����ʾ!"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   3630
      Width           =   2700
   End
   Begin VB.Label lbl 
      Caption         =   "Ĭ����Ŀ����"
      Height          =   285
      Index           =   1
      Left            =   255
      TabIndex        =   11
      Top             =   3135
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�����ʻ�֧��"
      Height          =   180
      Left            =   255
      TabIndex        =   9
      Top             =   2760
      Width           =   1080
   End
End
Attribute VB_Name = "frmSet�����山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Dim mblnFirst As Boolean
Private Enum enum�ı�
    textҽ���û� = 0
    Textҽ������ = 1
    Textҽ�������� = 2
End Enum

 
Public Function ��������() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select * From ������Ŀ "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.cbo������Ŀ.Clear
        Do While Not .EOF
             cbo������Ŀ.AddItem !����
             cbo������Ŀ.ItemData(cbo������Ŀ.NewIndex) = Nvl(!ID)
            .MoveNext
        Loop
    End With
    
    
    frmSet�����山.Show vbModal, frmҽ�����

    �������� = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo������Ŀ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag) = False Then
        Exit Sub
    End If
    MsgBox "���ӳɹ���", vbInformation, gstrSysName
End Sub



Private Sub cmd����_Click()
        '���˺�:20040706
        Dim strCode As String
        Dim STRNAME As String
        
        On Error Resume Next
        If frm������Ŀѡ�������山.GetCode(Me, strCode, STRNAME, True) = True Then
            Me.txt��Ŀ.Text = strCode & "-" & STRNAME
            Me.txt��Ŀ.Tag = strCode
        End If
    
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    
    gstrSQL = "Select * From ���ղ��� where ����=" & TYPE_�����山
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!������)
            Case "ҽ���û���"
                  txtEdit(textҽ���û�).Text = Nvl(!����ֵ)
            Case "ҽ���û�����"
                  txtEdit(Textҽ������).Text = Nvl(!����ֵ)
            Case "ҽ��������"
                  txtEdit(Textҽ��������).Text = Nvl(!����ֵ)
            Case "������Ŀ����"
                  txt��Ŀ.Text = Nvl(!����ֵ)
                  txt��Ŀ.Tag = txt��Ŀ.Text
            Case "������������"
                 Txt�޶�.Text = Nvl(!����ֵ)
            Case "�����ʻ�"
                Dim i As Long
                For i = 0 To cbo������Ŀ.ListCount - 1
                    If cbo������Ŀ.ItemData(i) = Val(Nvl(!����ֵ)) Then
                        cbo������Ŀ.ListIndex = i
                        Exit For
                    End If
                Next
            Case "���ܴ�����������"
                txt��ʾ.Text = Nvl(!����ֵ)
            End Select
            .MoveNext
        Loop
    End With
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Textҽ������ Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Textҽ�������� Or Index = Textҽ������ Or Index = textҽ���û� Then
        '�رն�ҽ�������������ӣ���Ϊ�ڲ����������ʱ��Ҫ���´�
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Textҽ��������).Text, txtEdit(textҽ���û�).Text, txtEdit(Textҽ������).Tag, False) = False Then
            If MsgBox("ҽ�������������������ӣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    'ɾ���Ѿ�����
    gstrSQL = "zl_���ղ���_Delete(" & TYPE_�����山 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '������������
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����山 & ",null,'ҽ���û���','" & txtEdit(textҽ���û�).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����山 & ",null,'ҽ���û�����','" & txtEdit(Textҽ������).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����山 & ",null,'ҽ��������','" & txtEdit(Textҽ��������).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����山 & ",null,'������������','" & Val(Txt�޶�.Text) & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If Me.cbo������Ŀ.ListIndex < 0 Then
        gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����山 & ",null,'�����ʻ�','" & 0 & "',5)"
    Else
        gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����山 & ",null,'�����ʻ�','" & Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex) & "',5)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����山 & ",null,'������Ŀ����','" & txt��Ŀ.Tag & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gstrSQL = "zl_���ղ���_Insert(" & TYPE_�����山 & ",null,'���ܴ�����������','" & Val(txt��ʾ.Text) & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub txt��ʾ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub Txt�޶�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt��Ŀ_Change()
    txt��Ŀ.Tag = ""
End Sub

Private Sub txt��Ŀ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim rsTemp As New ADODB.Recordset
    Dim strLeft As String
    Dim strTemp As String
    Dim blnReturn As Boolean
    If txt��Ŀ.Text = "" Then Exit Sub
    strLeft = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    strTemp = "'" & strLeft & txt��Ŀ.Text & "%'"
    
    gstrSQL = " select  ��Ʒ���� as ҽ������,  ҽԺ�������, ҩƷͨ��������, ҩƷͨ��Ӣ����,��Ʒ��, ��Ʒ������, ������Ŀ���㷽ʽ, ������ʶ, ҽ����ʶ, �Ƿ񴦷���ҩ, ҩƷ��Ӧ֢, ����ҽ��, ����Ȩ��, ����, ��װ���, " & _
             "         ��С��װ��λ, ��С������λ, ÿ���������, ָ���۸�, �б�۸�, ����֧���޼�1, ����֧���޼�2, ����֧���޼�3, ʵ��ִ�м۸�, �Ը�����1, �Ը�����2, �Ը�����3, �Ը�����4, �Ը�����5, �Ը�����6, �Ը�����7, �Ը�����8,  " & _
             "         �Ը�����9, �Ը�����10, �Ը�����11, �Ը�����12, ҽԺʹ��״̬, ����ʹ��״̬, ��׼���,  " & _
             "         ���������1, ���������2, ���������3, ƴ��������1, ƴ��������2, ƴ��������3, ��ע, ҽ���������,������׼���, ҽ�ƻ������, " & _
             "          �޸�ʱ��, Ŀ¼����  " & _
             "  from ҽ��������ĿĿ¼" & _
             "  where ҽԺ�������='61' and ( ��Ʒ���� like " & strTemp & " Or ��Ʒ�� like " & strTemp & " Or " & _
             "        ���������1 like " & UCase(strTemp) & " Or " & _
             "        ƴ��������1 like " & UCase(strTemp) & ")"
    
    If gcnOracle_CQYB.State = adStateOpen Then
        rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    Else
        'ǿ��ʹ��¼��Ϊ��״̬
        gstrSQL = "Select ����  ҽ������,����,���� FROM ������Ŀ Where Rownum<1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    End If
                   
    If rsTemp.RecordCount > 0 Then
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_�����山, rsTemp, "ҽ������", "ҽ����Ŀѡ��", "��ѡ���Ӧ��ҽ����Ŀ��")
        Else
            blnReturn = True
        End If
    Else
        MsgBox "�޴���Ŀ!"
        Exit Sub
    End If
    
    If blnReturn = False Then Exit Sub

    '�϶����м�¼����
    txt��Ŀ.Text = rsTemp("ҽ������") & "-" & Nvl(rsTemp!��Ʒ��)
    txt��Ŀ.Tag = rsTemp("ҽ������")

End Sub

Private Sub txt��Ŀ_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��Ŀ, KeyAscii, m�ı�ʽ
End Sub
