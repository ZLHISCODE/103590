VERSION 5.00
Begin VB.Form frmPreCompendEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ԥ����ٱ༭"
   ClientHeight    =   5325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5055
   Icon            =   "frmOutlineEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -135
      TabIndex        =   14
      Top             =   4725
      Width           =   5760
   End
   Begin VB.ListBox lstApply 
      Height          =   1740
      ItemData        =   "frmOutlineEdit.frx":058A
      Left            =   1710
      List            =   "frmOutlineEdit.frx":058C
      Style           =   1  'Checkbox
      TabIndex        =   9
      Top             =   2310
      Width           =   3075
   End
   Begin VB.CheckBox chkCopy 
      Caption         =   "�������(&R)"
      Height          =   210
      Left            =   675
      TabIndex        =   10
      Top             =   4155
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2550
      TabIndex        =   11
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3700
      TabIndex        =   12
      Top             =   4860
      Width           =   1100
   End
   Begin VB.TextBox txtExplain 
      Height          =   660
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1515
      Width           =   3420
   End
   Begin VB.TextBox txtName 
      Height          =   300
      Left            =   1380
      TabIndex        =   2
      Top             =   1125
      Width           =   3405
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   1380
      TabIndex        =   1
      Top             =   750
      Width           =   795
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   660
      TabIndex        =   0
      Top             =   600
      Width           =   5310
   End
   Begin VB.Label lblCopy 
      AutoSize        =   -1  'True
      Caption         =   "����ٵ������Ƿ���ں���������дʱֱ�����롣"
      Height          =   180
      Left            =   945
      TabIndex        =   13
      Top             =   4410
      Width           =   3960
   End
   Begin VB.Label lblApply 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ�÷�Χ(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   8
      Top             =   2310
      Width           =   990
   End
   Begin VB.Label lblExplain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   7
      Top             =   1575
      Width           =   630
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmOutlineEdit.frx":058E
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblNote 
      Caption         =   "�ڲ����ļ�����ǰ���ù����õĲ��������Ŀ���Ա��ڶ�������ļ��ظ�Ӧ�á�"
      Height          =   345
      Left            =   675
      TabIndex        =   5
      Top             =   135
      Width           =   4140
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   4
      Top             =   1185
      Width           =   630
   End
   Begin VB.Label lblCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   3
      Top             =   810
      Width           =   630
   End
End
Attribute VB_Name = "frmPreCompendEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1���ϼ�����ͨ��������ShowMe�������������塢�༭����ID,�༭״̬����Ϣ���ݽ��뱾����
'   2���༭״̬����Me.tag��ţ��ֱ�Ϊ"����"��"�޸�"�����ϼ�����ͨ��ShowMe����
'---------------------------------------------------
Private mlngItemID As Long       '���༭�ļ�¼ID���޸ġ�����ʱ���ϼ�����ͨ��ShowMe���ݽ���,����ʱΪ0��
Private mblnOK As Boolean        '�Ƿ���ɱ༭�˳�

'��ʱ����
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
Dim strApply As String

Public Function ShowMe(ByVal frmParent As Object, ByVal blnAdd As Boolean, Optional ByVal lngItemId As Long) As Long
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '���أ�ȷ�������������޸ĵ�ID��ȡ������0
    '---------------------------------------------------
    If blnAdd Then
        Me.Tag = "����"
    Else
        Me.Tag = "�޸�"
    End If
    mlngItemID = lngItemId
    
    With Me.lstApply
        .Clear
        .AddItem "1-���ﲡ��"
        .AddItem "2-סԺ����"
        .AddItem "3-�����¼"
        .AddItem "4-������"
        .AddItem "5-����֤������"
        .AddItem "6-֪���ļ�"
        .AddItem "7-��������"
        .AddItem "8-���Ʊ���"
        For lngCount = 1 To .ListCount
            .Selected(lngCount - 1) = True
        Next
        .Selected(2) = False
        .ListIndex = 0
    End With
    
    '��ȡ��Ϣ
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select �������, �����ı�, ��������, Nvl(�������, 0) As �������, ʹ��ʱ�� From �����ļ��ṹ Where �ļ�id Is Null And ID=[1]"
    Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption, mlngItemID)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txtCode.Text = !�������: Me.txtName.Text = "" & !�����ı�: Me.txtExplain.Text = "" & !��������
            Me.chkCopy.Value = !�������
            strApply = Left("" & !ʹ��ʱ��, 8)
            For lngCount = 1 To Len(strApply)
                Me.lstApply.Selected(lngCount - 1) = IIf(Val(Mid(strApply, lngCount, 1)) = 0, False, True)
            Next
        End If
        Me.txtCode.MaxLength = 3
        Me.txtName.MaxLength = 30
        Me.txtExplain.MaxLength = 200
    End With
    If Me.Tag = "����" Then
        gstrSQL = "Select nvl(max(�������),0) as ������� From �����ļ��ṹ Where �ļ�id Is Null"
        Set rsTemp = OpenSQLRecord(gstrSQL, Me.Caption)
        Me.txtCode.Text = Val(Format(Val(rsTemp!�������) + 1, String(Me.txtCode.MaxLength, "0")))
    End If
    
    '��ʾ����
    Me.Show vbModal, frmParent
    If mblnOK Then
        ShowMe = mlngItemID
    Else
        ShowMe = 0
    End If
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = 0
End Function

Private Sub chkCopy_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chkCopy_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim(Me.txtCode.Text) = "" Then MsgBox "�������ţ�", vbInformation, gstrSysName: Me.txtCode.SetFocus: Exit Sub
    If Trim(Me.txtName.Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: Me.txtName.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txtName.Text), vbFromUnicode)) > Me.txtName.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txtName.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txtName.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txtExplain.Text), vbFromUnicode)) > Me.txtExplain.MaxLength Then
        MsgBox "˵�����������" & Me.txtExplain.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txtExplain.SetFocus: Exit Sub
    End If
    
    Dim RS As New ADODB.Recordset, strS As String, lngSum As Long
    '���ݱ���
    With Me.lstApply
        strApply = ""
        For lngCount = 1 To .ListCount
            strApply = strApply & IIf(.Selected(lngCount - 1) = True, "1", "0")
        Next
    End With
    If Me.Tag = "����" Then
        gstrSQL = Trim(Me.txtCode.Text) & ",'" & Trim(Me.txtName.Text) & "','" & Trim(Me.txtExplain.Text) & "'," & IIf(Me.chkCopy.Value = 1, 1, 0) & ",'" & strApply & "'"
        mlngItemID = zlDatabase.GetNextId("�����ļ��ṹ")
        gstrSQL = "Zl_����Ԥ�����_Insert(" & mlngItemID & "," & gstrSQL & ")"
    Else
        gstrSQL = "Select count(A.ID) From �����ļ��ṹ A, �����ļ��б� B Where A.Ԥ�����id = [1] And A.�ļ�id = B.ID"
        strS = ""
        With Me.lstApply
            For lngCount = 1 To .ListCount
                If .Selected(lngCount - 1) = False Then
                    If strS = "" Then
                        strS = " B.���� = " & lngCount & " "
                    Else
                        strS = strS & " or B.���� = " & lngCount & " "
                    End If
                End If
            Next
            If strS <> "" Then gstrSQL = gstrSQL & " And (" & strS & ")"
        End With
        Set RS = OpenSQLRecord(gstrSQL, Me.Caption, mlngItemID)
        If Not RS.EOF Then
            lngSum = RS(0)
            If lngSum > 0 Then
                If MsgBox("��Ԥ������Ѿ�����������Ĳ�����ʹ�ã��Ƿ������", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
            End If
        End If
        RS.Close
        gstrSQL = Trim(Me.txtCode.Text) & ",'" & Trim(Me.txtName.Text) & "','" & Trim(Me.txtExplain.Text) & "'," & IIf(Me.chkCopy.Value = 1, 1, 0) & ",'" & strApply & "'"
        gstrSQL = "Zl_����Ԥ�����_Update(" & mlngItemID & "," & gstrSQL & ")"
    End If
    
    Err = 0: On Error GoTo errHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSQL): gcnOracle.Execute gstrSQL, , adCmdStoredProc: Call SQLTest
    mblnOK = True: Me.Hide
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lstApply_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub lstApply_ItemCheck(Item As Integer)
    If Item = 2 Then Me.lstApply.Selected(Item) = False
End Sub

Private Sub lstApply_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtCode_Change()
    txtCode = Val(txtCode)
End Sub

Private Sub txtCode_GotFocus()
    Me.txtCode.SelStart = 0: Me.txtCode.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtExplain_Change()
    ValidControlText txtExplain
End Sub

Private Sub txtName_Change()
    ValidControlText txtName
End Sub

Private Sub txtName_GotFocus()
    Me.txtName.SelStart = 0: Me.txtName.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtExplain_GotFocus()
    Me.txtExplain.SelStart = 0: Me.txtExplain.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtExplain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
