VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frm���մ���༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���մ���༭"
   ClientHeight    =   6345
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   7020
   Icon            =   "frm���մ���༭.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6735
      Left            =   5535
      TabIndex        =   28
      Top             =   -300
      Width           =   30
   End
   Begin VB.ComboBox cmb������� 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1335
      Width           =   1425
   End
   Begin VB.CheckBox chkҽ�� 
      Caption         =   "ҽ����Ŀ(&I)"
      Height          =   225
      Left            =   3990
      TabIndex        =   8
      Top             =   1365
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   5865
      TabIndex        =   27
      Top             =   5610
      Width           =   1100
   End
   Begin VB.Frame frmRule 
      Caption         =   "ͳ��֧���������"
      Height          =   3750
      Left            =   180
      TabIndex        =   13
      Top             =   2445
      Width           =   5130
      Begin ZL9BillEdit.BillEdit mshbill 
         Height          =   1695
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2990
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
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
      Begin VB.OptionButton opt�㷨 
         Caption         =   "������õ��μ��㷨(&T)"
         Height          =   240
         Index           =   3
         Left            =   225
         TabIndex        =   24
         Top             =   1560
         Width           =   2265
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   3705
         MaxLength       =   16
         TabIndex        =   23
         Top             =   1200
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   21
         Top             =   1200
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1575
         MaxLength       =   16
         TabIndex        =   19
         Top             =   870
         Width           =   630
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   3735
         MaxLength       =   16
         TabIndex        =   16
         Top             =   270
         Width           =   630
      End
      Begin VB.OptionButton opt�㷨 
         Caption         =   "סԺ�ն�����㷨(&Z)"
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   17
         Top             =   630
         Width           =   2265
      End
      Begin VB.OptionButton opt�㷨 
         Caption         =   "�ܶ�������㷨(&B)"
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   14
         Top             =   315
         Value           =   -1  'True
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��׼��������        ��"
         Height          =   180
         Index           =   6
         Left            =   2595
         TabIndex        =   22
         Top             =   1260
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ÿ����׼����        Ԫ"
         Height          =   180
         Index           =   5
         Left            =   465
         TabIndex        =   20
         Top             =   1260
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ÿ�ջ�������        Ԫ"
         Height          =   180
         Index           =   4
         Left            =   465
         TabIndex        =   18
         Top             =   930
         Width           =   1980
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ͳ��֧������        %"
         Height          =   180
         Index           =   3
         Left            =   2625
         TabIndex        =   15
         Top             =   330
         Width           =   1890
      End
   End
   Begin VB.Frame fraKind 
      Caption         =   "����"
      Height          =   630
      Left            =   195
      TabIndex        =   9
      Top             =   1710
      Width           =   5160
      Begin VB.OptionButton opt���� 
         Caption         =   "����(&W)"
         Height          =   180
         Index           =   3
         Left            =   4050
         TabIndex        =   12
         Top             =   270
         Width           =   945
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ҽ��(&D)"
         Height          =   180
         Index           =   2
         Left            =   2145
         TabIndex        =   11
         Top             =   285
         Width           =   945
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ҩƷ(&M)"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5775
      TabIndex        =   25
      Top             =   225
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5775
      TabIndex        =   26
      Top             =   735
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   5
      Top             =   937
      Width           =   1425
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   3
      Top             =   536
      Width           =   4080
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   1
      Top             =   135
      Width           =   1425
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "�������(&F)"
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   1398
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&U)"
      Height          =   180
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&S)"
      Height          =   180
      Index           =   2
      Left            =   495
      TabIndex        =   4
      Top             =   997
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   1
      Left            =   495
      TabIndex        =   2
      Top             =   596
      Width           =   630
   End
End
Attribute VB_Name = "frm���մ���༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    text���� = 0
    Text���� = 1
    Text���� = 2
    Text���� = 3
    Text���� = 4
    Text��׼ = 5
    Text���� = 6

    CheckҩƷ = 1
    Checkҽ�� = 2
    Check���� = 3
    
    Check���� = 1
    CheckסԺ�� = 2
    chk���õ��� = 3
End Enum
Private Enum mColHead
    ���� = 0
    ����
    ����
    ����
End Enum

Dim mlng���� As Long
Dim mstrID As String         '��ǰ�༭��ҽ������ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub chkҽ��_Click()
    mblnChange = True
End Sub

Private Sub chkҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmb�������_Click()
    mblnChange = True
End Sub

Private Sub cmb�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    If IsValid() = False Then Exit Sub
    If Save��Ŀ() = False Then Exit Sub
    
    If mstrID = "" Then
        '��������
        txtEdit(text����).Text = zlDatabase.GetMax("����֧������", "����", 6, " where ����=" & mlng����)
        For lngIndex = Text���� To Text����
            txtEdit(lngIndex).Text = ""
        Next
        chkҽ��.Value = 1
        mblnChange = False
        txtEdit(text����).SetFocus
    Else
        mblnChange = False
        Unload Me
    End If
End Sub

Private Function Save��Ŀ() As Boolean
    Dim lngID As Long, lng���� As Long, lng�㷨 As Long
    Dim dblͳ��ȶ� As Double, dbl��׼���� As Double, dbl��׼���� As Double
    Dim lngIndex As Long, lst As ListItem
    
    On Error GoTo errHandle
    
    For lngIndex = 1 To 3
        If opt����(lngIndex).Value = True Then
            lng���� = lngIndex
            Exit For
        End If
    Next
    If opt�㷨(1).Value = True Then
        '������
        lng�㷨 = 1
        dblͳ��ȶ� = Val(txtEdit(Text����).Text)
        
    Else
        If opt�㷨(3).Value = True Then
            lng�㷨 = 3
        Else
            '��סԺ��
            lng�㷨 = 2
            dblͳ��ȶ� = Val(txtEdit(Text����).Text)
            dbl��׼���� = Val(txtEdit(Text��׼).Text)
            dbl��׼���� = Val(txtEdit(Text����).Text)
        End If
    End If
    
    
    If mstrID = "" Then
        '����
        lngID = zlDatabase.GetNextID("����֧������")
        gstrSQL = "zl_����֧������_INSERT(" & lngID & "," & mlng���� & ",'" & Trim(txtEdit(text����).Text) & "','" & _
                Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng���� & "," & lng�㷨 & "," & _
                 dblͳ��ȶ� & "," & dbl��׼���� & "," & dbl��׼���� & "," & GetTextFromCombo(cmb�������, False) & "," & chkҽ��.Value & ")"
    Else
        gstrSQL = "zl_����֧������_Update(" & mstrID & ",'" & Trim(txtEdit(text����).Text) & "','" & _
                Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng���� & "," & lng�㷨 & "," & _
                 dblͳ��ȶ� & "," & dbl��׼���� & "," & dbl��׼���� & "," & GetTextFromCombo(cmb�������, False) & "," & chkҽ��.Value & ")"
    End If
    gcnOracle.BeginTrans
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If lng�㷨 = 3 Then
        If SaveGrdData(IIf(mstrID = "", lngID, Val(mstrID))) = False Then GoTo errHandle:
    End If
    gcnOracle.CommitTrans
    
    '����������
    If mstrID = "" Then
        Set lst = frm���մ���.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text����), "Class", "Class")
    Else
        Set lst = frm���մ���.lvwItem.SelectedItem
        lst.Text = Trim(txtEdit(text����).Text)
    End If
    lst.SubItems(1) = Trim(txtEdit(Text����).Text)
    lst.SubItems(2) = Trim(txtEdit(Text����).Text)
    lst.SubItems(3) = Choose(lng����, "ҩƷ", "ҽ��", "����")
    lst.SubItems(4) = IIf(lng�㷨 = 1, "�ܶ����", "סԺ�պ˶�")
    lst.SubItems(5) = Mid(cmb�������.Text, 3)
    lst.SubItems(6) = IIf(chkҽ��.Value = 1, "��", "��")
    lst.Tag = dblͳ��ȶ� & ";" & dbl��׼���� & ";" & dbl��׼����
    
    Save��Ŀ = True
    mblnOK = True
    Exit Function

errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
'����:���������й�ҽ�����������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim lngIndex As Integer
    For lngIndex = text���� To Text����
        If txtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
                txtEdit(lngIndex).SetFocus
                zlControl.TxtSelAll txtEdit(lngIndex)
                Exit Function
            End If
            
            If lngIndex = text���� Or lngIndex = Text���� Then
                If Len(Trim(txtEdit(lngIndex).Text)) = 0 Then
                    txtEdit(lngIndex).Text = ""
                    MsgBox "��������ƶ�����Ϊ�ա�", vbExclamation, gstrSysName
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
            End If
            
            If lngIndex >= Text���� Then
                If IsNumeric(txtEdit(lngIndex).Text) = False Then
                    MsgBox "������Ϸ�����ֵ��", vbInformation, gstrSysName
                    zlControl.TxtSelAll txtEdit(lngIndex)
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
                        
                If Val(txtEdit(lngIndex).Text) < 0 Then
                    MsgBox "��ֵ����С��0��", vbInformation, gstrSysName
                    zlControl.TxtSelAll txtEdit(lngIndex)
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
                
                If lngIndex = Text���� Then
                    If Val(txtEdit(Text����).Text) > 100 Then
                        MsgBox "ͳ��֧���������ܳ���100��", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(Text����)
                        txtEdit(lngIndex).SetFocus
                        Exit Function
                    End If
                Else
                    If Val(txtEdit(lngIndex).Text) > 10000 Then
                        MsgBox "��ֵ���ܳ���10000��", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(lngIndex)
                        txtEdit(lngIndex).SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    '����׼��������׼��������
    If opt�㷨(CheckסԺ��).Value = True Then
        If Val(txtEdit(Text��׼).Text) = 0 And Val(txtEdit(Text����).Text) <> 0 Then
            MsgBox "��׼����Ϊ0����׼����Ҳ��Ϊ0��", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text����)
            txtEdit(Text����).SetFocus
            Exit Function
        End If
        If Val(txtEdit(Text��׼).Text) <> 0 And Val(txtEdit(Text����).Text) = 0 Then
            MsgBox "��׼����Ϊ0����׼����Ҳ��Ϊ0��", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text��׼)
            txtEdit(Text��׼).SetFocus
            Exit Function
        End If
        If Val(txtEdit(Text��׼).Text) <> 0 And Val(txtEdit(Text����).Text) > Val(txtEdit(Text��׼).Text) Then
            MsgBox "��������ܴ�����׼���", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(Text����)
            txtEdit(Text����).SetFocus
            Exit Function
        End If
    End If
    Dim i As Long
    If opt�㷨(chk���õ���).Value = True Then
        With mshBill
            For i = 1 To .Rows - 1
                If i <> 1 Then
                    If Val(.TextMatrix(i - 1, mColHead.����)) <> Val(.TextMatrix(i, mColHead.����)) Then
                        MsgBox "�ڵ�" & i & "�е����޲������ڵ�" & i - 1 & "�е�����,������!", vbInformation + vbDefaultButton1, gstrSysName
                        mshBill.Row = i
                        mshBill.SetFocus
                        Exit Function
                    End If
               End If
                If (Val(.TextMatrix(i, mColHead.����)) <> 0 Or Val(.TextMatrix(i, mColHead.����)) <> 0) _
                    And Val(.TextMatrix(i, mColHead.����)) = 0 Then
                    MsgBox "�ڵ�" & i & "�еı�����������,������!", vbInformation + vbDefaultButton1, gstrSysName
                    mshBill.Row = i
                    mshBill.SetFocus
                    Exit Function
                End If
                If Val(.TextMatrix(i, mColHead.����)) = Val(.TextMatrix(i, mColHead.����)) And Val(.TextMatrix(i, mColHead.����)) <> 0 Then
                    MsgBox "�ڵ�" & i & "�е����޵�������,������!", vbInformation + vbDefaultButton1, gstrSysName
                    mshBill.Row = i
                    mshBill.SetFocus
                    Exit Function
                End If
            Next
        End With
    End If
    If chkҽ��.Value = 0 Then
        If MsgBox("���������������ҽ������Ӱ�쵽������������ҽ����Ŀ��" & vbCrLf & "�Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            chkҽ��.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function


Private Sub MshBill_AfterAddRow(Row As Long)
        If Row = mshBill.Rows - 1 Then
            mshBill.TextMatrix(Row, mColHead.����) = mshBill.TextMatrix(Row - 1, mColHead.����)
        End If
End Sub

Private Sub MshBill_AfterDeleteRow()
   If mshBill.Row = mshBill.Rows - 1 Then Exit Sub
    Call ReSet����(mshBill.Row - 1)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
''   If Row = MshBill.Rows - 1 Then Exit Sub
''    Call ReSet����(Row)
End Sub

Private Sub opt�㷨_Click(Index As Integer)
    Dim bln���� As Boolean
    
    mblnChange = True
    txtEdit(Text����).Enabled = (opt�㷨(Check����).Value = True)
    lblEdit(Text����).Enabled = txtEdit(Text����).Enabled
    
    txtEdit(Text����).Enabled = (opt�㷨(CheckסԺ��).Value = True)
    txtEdit(Text��׼).Enabled = txtEdit(Text����).Enabled
    txtEdit(Text����).Enabled = txtEdit(Text����).Enabled
    lblEdit(Text����).Enabled = txtEdit(Text����).Enabled
    lblEdit(Text��׼).Enabled = txtEdit(Text����).Enabled
    lblEdit(Text����).Enabled = txtEdit(Text����).Enabled
    
    mshBill.Active = (opt�㷨(chk���õ���).Value = True)
End Sub

Private Sub opt�㷨_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub opt����_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text���� Then
        txtEdit(Text����).Text = zlCommFun.SpellCode(txtEdit(Text����).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text����
          zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 'ʹ֮����
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = text���� Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
    If Index >= Text���� And Index <= Text��׼ Then
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
    End If
End Sub

Public Function �༭ҽ������(ByVal lng���� As Long, ByVal strID As String) As Boolean
'����:��������õ�ҽ���������ڽ���ͨѶ�ĳ���
'����:str���           ��ǰ�༭��ҽ�����ĵ����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mblnOK = False
    mlng���� = lng����
    mstrID = strID
    
    cmb�������.AddItem "1.���ﲡ��"
    cmb�������.AddItem "2.סԺ����"
    cmb�������.AddItem "3.���в���"
    cmb�������.ListIndex = 2
    rsTemp.CursorLocation = adUseClient
    If mstrID <> "" Then
        '�޸�ҽ������
        gstrSQL = "select ����,����,����,nvl(����,1) as ����,nvl(�㷨,1) as �㷨 " & _
                  ",ͳ��ȶ�,��׼����,��׼����,�Ƿ�ҽ��,nvl(�������,3) as ������� " & _
                  "from ����֧������ where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mstrID))
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "�ñ��մ����Ѿ���ɾ������ˢ�¡�", vbInformation, gstrSysName
            Exit Function
        End If
        txtEdit(text����).Text = rsTemp("����")
        txtEdit(Text����).Text = rsTemp("����")
        txtEdit(Text����).Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        Call SetComboByText(cmb�������, rsTemp("�������"), False)
        chkҽ��.Value = IIf(rsTemp("�Ƿ�ҽ��") = 1, 1, 0)
        opt����(rsTemp("����")).Value = True
        opt�㷨(rsTemp("�㷨")).Value = True
        Call opt�㷨_Click(rsTemp("�㷨"))
        If rsTemp("�㷨") = 1 Then
            '������õ���
            Call initGrd
            '1-����������Ŀ
            txtEdit(Text����).Text = Format(rsTemp("ͳ��ȶ�"), "0.00")
            opt�㷨_Click (1)
        ElseIf Nvl(rsTemp!�㷨, 0) = 2 Then
            '������õ���
            Call initGrd
            '2-סԺ�պ˶���Ŀ
            txtEdit(Text����).Text = Format(rsTemp("ͳ��ȶ�"), "0.00")
            txtEdit(Text��׼).Text = Format(rsTemp("��׼����"), "0.00")
            txtEdit(Text����).Text = Format(rsTemp("��׼����"), "0")
            opt�㷨_Click (2)
        Else '3-���õ��μ��㷨
            Call LoadGrd
            opt�㷨_Click (3)
        End If
    Else
        '����ҽ������
        txtEdit(text����).Text = zlDatabase.GetMax("����֧������", "����", 6, " where ����=" & mlng����)
        opt�㷨(1).Value = True
        Call opt�㷨_Click(1)
        '������õ���
        Call initGrd
    End If
    mblnChange = False
    frm���մ���༭.Show vbModal, frm���մ���
    �༭ҽ������ = mblnOK
End Function
Private Sub initGrd()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ʼ���õ��ε�Grid����
    '--�����:
    '--������:
    '--��  ��:
    '--��  ��:���˺�:20040615
    '-----------------------------------------------------------------------------------------------------------
   With mshBill
        .Active = True
        .ClearBill
        .Cols = 4
        .Rows = 2
        .msfObj.FixedCols = 1
        .TextMatrix(0, mColHead.����) = "����"
        .TextMatrix(0, mColHead.����) = "����"
        .TextMatrix(0, mColHead.����) = "����"
        .TextMatrix(0, mColHead.����) = "ʵ�ձ���"
        
        .ColWidth(mColHead.����) = 500
        .ColWidth(mColHead.����) = 1400
        .ColWidth(mColHead.����) = 1400
        .ColWidth(mColHead.����) = 1400
                
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(mColHead.����) = 5
        .ColData(mColHead.����) = 5
        .ColData(mColHead.����) = 4
        .ColData(mColHead.����) = 4

        .ColAlignment(mColHead.����) = flexAlignCenterCenter
        .ColAlignment(mColHead.����) = flexAlignCenterCenter
        .ColAlignment(mColHead.����) = flexAlignCenterCenter
        .ColAlignment(mColHead.����) = flexAlignCenterCenter
        .PrimaryCol = mColHead.����
        .LocateCol = mColHead.����
    End With
End Sub

Private Sub LoadGrd()
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�������
    '--�����:
    '--������:
    '--��  ��:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    '��ʼ���
    Call initGrd
    
    If mstrID = "" Then Exit Sub
    '��ʾ�޸��������
    gstrSQL = "Select * From ���൵�α��� where ����id=" & Val(mstrID) & " order by  ����"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then Exit Sub
    mshBill.Rows = rsTemp.RecordCount + 1
    lngRow = 1
    With mshBill
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, mColHead.����) = lngRow
            .TextMatrix(lngRow, mColHead.����) = Format(Nvl(rsTemp!����, 0), "####0.00;####0.00; ;")
            .TextMatrix(lngRow, mColHead.����) = Format(Nvl(rsTemp!����, 0), "####0.00;####0.00; ;")
            .TextMatrix(lngRow, mColHead.����) = Format(Nvl(rsTemp!����, 0), "####0.00;####0.00; ;")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
End Sub
Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EnterCell(Row As Long, COL As Long)
    With mshBill
        Select Case .COL
            Case mColHead.����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mColHead.����
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                If Trim(.TextMatrix(.Row, mColHead.����)) = "" Then
                    .AllowAddRow = False
                Else
                    .AllowAddRow = True
                End If


            Case mColHead.����
                .TxtCheck = True
                .MaxLength = 4
                .TextMask = ".1234567890"
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshBill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        
        Select Case .COL
            Case mColHead.����
                If .TextMatrix(.Row, .COL) = "" And strKey = "" And .Row <> 1 And Val(.TextMatrix(.Row, mColHead.����)) <> 0 Then
                    MsgBox "δ��������,���������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "���ޱ���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 1E+19 Then
                    MsgBox "����ֻ����0~9000900090009��Χ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If strKey <> "" Then
                    .Text = Format(strKey, "####0.00;####0.00; ;")
                ElseIf Trim(.TextMatrix(.Row, mColHead.����)) = "" Then
                    .Text = " "
                    .TextMatrix(.Row, mColHead.����) = " "
                End If
                
            Case mColHead.����
                If .TextMatrix(.Row, .COL) = "" And strKey = "" And .Row <> .Rows - 1 Then
                    MsgBox "δ��������,���������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "���ޱ���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 1E+19 Then
                    MsgBox "����ֻ����0~9000900090009��Χ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Val(strKey) <= Val(.TextMatrix(.Row, mColHead.����)) And (.Row <> .Rows - 1) Then
                    MsgBox "���޲���С�ڵ�������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If

                If Val(strKey) <= Val(.TextMatrix(.Row, mColHead.����)) And (.Row = .Rows - 1 And strKey <> "") Then
                    MsgBox "���޲���С�ڵ�������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                                
                If strKey <> "" Then
                    .Text = Format(strKey, "####0.00;####0.00; ;")
                    .TextMatrix(.Row, mColHead.����) = .Text
                    
                    If .Row <> .Rows - 1 Then
                        Call ReSet����(.Row)
                    End If
                ElseIf Trim(.TextMatrix(.Row, mColHead.����)) = "" Then
                    .Text = " "
                    .TextMatrix(.Row, mColHead.����) = " "
                End If
                If strKey = "" Then
                    .AllowAddRow = False
                Else
                    .AllowAddRow = True
                End If

                
            Case mColHead.����
                If .TextMatrix(.Row, .COL) = "" And strKey = "" And (Val(.TextMatrix(.Row, mColHead.����)) <> 0 Or Val(.TextMatrix(.Row, mColHead.����)) <> 0) Then
                    MsgBox "δ�������,���������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "��������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Val(strKey) > 100 Then
                    MsgBox "����ֻ����0~100��Χ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If (Val(.TextMatrix(.Row, mColHead.����)) = 0 And Val(.TextMatrix(.Row, mColHead.����)) = 0) Or .AllowAddRow = False Then
                    .AllowAddRow = False
                Else
                    .AllowAddRow = True
                End If
                If strKey <> "" Then
                    .Text = Format(strKey, "####0.00;####0.00; ;")
                End If
        End Select
        If .TextMatrix(.Row, mColHead.����) = "" Then
            '������ȷ������
            Call Set����
        End If
    End With
  
End Sub
Private Sub ReSet����(ByVal lngRow As Long)
    Dim i As Long
    Dim dbl��� As Double
    
    With mshBill
        For i = lngRow + 1 To .Rows - 1
            
            .TextMatrix(i, mColHead.����) = i
            dbl��� = Val(.TextMatrix(i, mColHead.����)) - Val(.TextMatrix(i, mColHead.����))
            .TextMatrix(i, mColHead.����) = Format(Val(.TextMatrix(i - 1, mColHead.����)), "####0.00;####0.00; ;")
            If dbl��� < 0 Then
            Else
                .TextMatrix(i, mColHead.����) = Format(Val(.TextMatrix(i, mColHead.����)) + IIf(dbl��� < 0, 0, dbl���), "####0.00;####0.00; ;")
            End If
        Next
'        If .TextMatrix(.Row, .Col) <> .Text Then
'            .Text = .TextMatrix(.Row, .Col)
'        End If
    End With
    
End Sub
Private Sub Set����()
    Dim lngRow As Long
    For lngRow = 1 To mshBill.Rows - 1
        mshBill.TextMatrix(lngRow, mColHead.����) = lngRow
    Next
End Sub
Private Function SaveGrdData(ByVal lngID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:������õ�������
    '--�����:
    '--������:
    '--��  ��:�ɹ�����true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim dbl���� As Double
    Dim dbl���� As Double
    Dim dbl����  As Double
    
    SaveGrdData = False
    Err = 0: On Error GoTo errHand:
    If mstrID <> "" Then
                gstrSQL = "zl_���൵�α���_Delete(" & _
                    lngID & ")"
                 Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    With mshBill
        For lngRow = 1 To .Rows - 1
            If Trim(.TextMatrix(lngRow, mColHead.����)) <> "" Or Trim(.TextMatrix(lngRow, mColHead.����)) <> "" Then
                '����:
                '   ����ID_IN    ���൵�α���.����id%Type,
                '   ����_IN        ���൵�α���.����%Type,
                '   ����_IN        ���൵�α���.����%Type,
                '   ����_IN        ���൵�α���.����%Type,
                '   ����_IN
                dbl���� = Val(.TextMatrix(lngRow, mColHead.����))
                dbl���� = Val(.TextMatrix(lngRow, mColHead.����))
                dbl���� = Val(.TextMatrix(lngRow, mColHead.����))
                
                gstrSQL = "zl_���൵�α���_InSert(" & _
                    lngID & "," & _
                    lngRow & "," & _
                    dbl���� & "," & _
                    dbl���� & "," & _
                    dbl���� & ")"
                
                 Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    SaveGrdData = True
    Exit Function
errHand:
    
End Function


