VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm�ӳ������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ӳ�������"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frm�ӳ�������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt�޼� 
      Height          =   300
      Left            =   6204
      MaxLength       =   16
      TabIndex        =   4
      Text            =   "800.00"
      Top             =   912
      Width           =   2196
   End
   Begin VB.ComboBox cbo���㷽�� 
      Height          =   276
      ItemData        =   "frm�ӳ�������.frx":030A
      Left            =   1128
      List            =   "frm�ӳ�������.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   924
      Width           =   2184
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6045
      TabIndex        =   6
      Top             =   5325
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7335
      TabIndex        =   7
      Top             =   5325
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Left            =   -500
      TabIndex        =   9
      Top             =   5070
      Width           =   10000
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   0
      TabIndex        =   8
      Top             =   645
      Width           =   10275
   End
   Begin ZL9BillEdit.BillEdit mshBill 
      Height          =   3804
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   8352
      _ExtentX        =   14737
      _ExtentY        =   6720
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
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����޼�(&X)"
      Height          =   180
      Index           =   1
      Left            =   5136
      TabIndex        =   3
      Top             =   972
      Width           =   1008
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "���㷽��(&J)"
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   1
      Top             =   972
      Width           =   1008
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   165
      Picture         =   "frm�ӳ�������.frx":0330
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   $"frm�ӳ�������.frx":09B1
      Height          =   480
      Left            =   780
      TabIndex        =   0
      Top             =   240
      Width           =   7668
   End
End
Attribute VB_Name = "frm�ӳ�������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnChange As Boolean

Dim mstrSql As String
Dim mblnReturn As Boolean
Dim mblnFirst As Boolean

Dim mstrPriv As String           'Ȩ�޴�

Private mintPreCol As Integer               'ǰһ�ε���ͷ��������
Private mintsort As Integer                 'ǰһ�ε���ͷ������
Private Enum marBillCol
    ��� = 0
    ��ͼ�
    ��߼�
    �ӳ���
    �ֶ�����޼�
    ˵��
End Enum

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Sub cbo���㷽��_Click()
 mblnChange = True
 SetCtlEnable
End Sub

Private Sub cbo���㷽��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub mshBill_AfterDeleteRow()
    '��������ֵ
    Call ReFormal
    mblnChange = True
    '���ÿؼ�����
    SetCtlEnable
End Sub
Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
    SetCtlEnable
End Sub
Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        SetInputFormat .Row
        Select Case .Col
            Case marBillCol.˵��
                ImeLanguage True
                .TxtCheck = False
                .MaxLength = 50
                .TxtSetFocus
            Case marBillCol.��ͼ�, marBillCol.��߼�, marBillCol.�ӳ���
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
        End Select
        SetCtlEnable
    End With
End Sub

Private Sub SetInputFormat(ByVal intRow As Integer)
    With mshBill
        If intRow <> 1 Then
            .ColData(marBillCol.��ͼ�) = 5               '��ֹ
        Else
            .ColData(marBillCol.��ͼ�) = 4               '���ı�����
        End If
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshBill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case marBillCol.˵��
                OS.OpenIme False
                If strKey = "" Then
                    .Text = " "
                    .TextMatrix(.Row, marBillCol.˵��) = " "
                End If
            Case marBillCol.��ͼ�
                
               If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "��ͼ۱���Ϊ������,���������룡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "��ͼ۱��������,���������룡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "��ͼ۱���С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(Val(strKey), mFMT.FM_�ɱ���)
                End If
                ReFormal

            Case marBillCol.��߼�
                
                If .Row - 1 > 1 Then
                    .TextMatrix(.Row, marBillCol.��ͼ�) = Format(Val(.TextMatrix(.Row - 1, marBillCol.��߼�)), mFMT.FM_�ɱ���)
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "��߼۱���Ϊ������,���������룡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "��߼۱��������,���������룡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "��߼۱���С��" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) < Val(.TextMatrix(.Row, marBillCol.��ͼ�)) And Val(strKey) <> 0 Then
                        MsgBox "��߼۱��������ͼ�", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(Val(strKey), mFMT.FM_�ɱ���)
                    .TextMatrix(.Row, .Col) = .Text
                ElseIf Val(.TextMatrix(.Row, marBillCol.��߼�)) = 0 Then
                    .TextMatrix(.Row, .Col) = " "
                    .Text = " "
                End If
                ReFormal
            Case marBillCol.�ӳ���
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�ӳ��ʱ���Ϊ������,���������룡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "�ӳ��ʱ��������,���������룡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) > 100 Then
                        MsgBox "�ӳ��ʱ���С��100%", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, GFM_VBJCL)
                    .Text = strKey
                    .TextMatrix(.Row, marBillCol.�ӳ���) = strKey
                End If
            Case marBillCol.�ֶ�����޼�
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�ֶ�����޼۱���Ϊ������,���������룡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "�ֶ�����޼۱��������,���������룡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

Private Sub mshBill_LeaveCell(Row As Long, Col As Long)
    OS.OpenIme False
End Sub

Private Sub mshBill_LostFocus()
    OS.OpenIme False
End Sub


'------------------------------------------------------------------
'------------------------------------------------------------------
'-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
' 0����ʾ���п���ѡ�񣬵������޸�
' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
'4:  ��ʾ����Ϊ�������ı����û�����
'5:  ��ʾ���в�����ѡ��
'-----------------------------------------------------------------
'-----------------------------------------------------------------

Private Function ValidData() As Boolean
    Dim intLop As Integer
    Dim dbl�ϴ���߼� As Double
    
    Dim blnStock As Boolean
    
    ValidData = False
    blnStock = False
    
    dbl�ϴ���߼� = 0
    If cbo���㷽��.ListIndex < 0 Then
        ShowMsgBox "���㷽������ѡ��!"
        If cbo���㷽��.Enabled Then cbo���㷽��.SetFocus
        Exit Function
    End If
    If txt�޼�.Text = "" Or Val(txt�޼�.Text) = 0 Then
        ShowMsgBox "������������޼�!"
        If txt�޼�.Enabled Then txt�޼�.SetFocus
        Exit Function
    End If
    If Abs(Val(txt�޼�)) > 10 ^ 11 - 1 Then
        ShowMsgBox "����޼۱�����(-" & 10 ^ 11 - 1 & " �� " & 10 ^ 11 - 1 & ")!"
        If txt�޼�.Enabled Then txt�޼�.SetFocus
        Exit Function
    End If
    With mshBill
            For intLop = 1 To .Rows - 1
                If .TextMatrix(intLop, marBillCol.��ͼ�) <> "" Or .TextMatrix(intLop, marBillCol.��߼�) <> "" Then           '�����з�����
                    If intLop = 1 Then
                        dbl�ϴ���߼� = Val(.TextMatrix(intLop, marBillCol.��߼�))
                    Else
                        If Val(.TextMatrix(intLop, marBillCol.��ͼ�)) <> dbl�ϴ���߼� Then
                            ShowMsgBox "�ڵ�" & intLop & "�е���ͼ۲�����" & intLop - 1 & "�е���߼�!"
                            Exit Function
                        End If
                        dbl�ϴ���߼� = Val(.TextMatrix(intLop, marBillCol.��߼�))
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.��ͼ�
                    End If
                    
                    
                    If Val(.TextMatrix(intLop, marBillCol.��ͼ�)) > Val(.TextMatrix(intLop, marBillCol.��߼�)) And Val(.TextMatrix(intLop, marBillCol.��߼�)) <> 0 Then
                        ShowMsgBox "�ڵ�" & intLop & "�е���ͼ۴�������߼�!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.��ͼ�
                            Exit Function
                    End If
                                        
                    
                    If Trim(Trim(.TextMatrix(intLop, marBillCol.�ӳ���))) = "" Then
                        ShowMsgBox "��" & intLop & "�мӳ���Ϊ���ˣ����飡"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.�ӳ���
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, marBillCol.��ͼ�)) > 9999999999# Then
                        ShowMsgBox "  ��" & intLop & "�е���ͼ۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.��ͼ�
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, marBillCol.��߼�)) > 9999999999# Then
                        ShowMsgBox "  ��" & intLop & "�е���߼۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.��߼�
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, marBillCol.�ӳ���)) > 100 Then
                        ShowMsgBox "  ��" & intLop & "�еļӳ��ʴ�����100%�����飡"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.��߼�
                        Exit Function
                    End If
                End If
                
                If LenB(StrConv(.TextMatrix(intLop, marBillCol.˵��), vbFromUnicode)) > 50 Then
                    MsgBox "��" & intLop & "��˵���г��ȴ���50���ַ��ˣ����������룡", vbInformation, gstrSysName
                    .SetFocus
                    .Row = intLop
                    .Col = marBillCol.˵��
                    Exit Function
                End If
            Next
    End With
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim str˵�� As String
    Dim dbl��ͼ� As Double
    Dim dbl��߼� As Double
    Dim dbl�ӳ��� As Double
    Dim byt���㷽�� As Byte
    Dim dbl�޼� As Double
    
    Dim strSQL As String
    Dim intRow As Integer
    
    SaveCard = False
    With mshBill
        On Error GoTo ErrHandle
        dbl�޼� = Val(txt�޼�.Text)
        byt���㷽�� = cbo���㷽��.ItemData(cbo���㷽��.ListIndex)
        gcnOracle.BeginTrans
        
        '���ԭ���Ĺ��ʷ���
        strSQL = "ZL_���ϼӳɷ���_DELETE()"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        '���ӹ̶������޼�ֵ
        If Trim(Me.txt�޼�.Text) <> "" Then
            strSQL = "ZL_���ϼӳɷ���_INSERT(0,null,null,null," _
                   & byt���㷽�� & "," _
                   & dbl�޼� & ",'���޼�')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, marBillCol.��ͼ�)) <> 0 Or Val(.TextMatrix(intRow, marBillCol.��߼�)) <> 0 Then           '�����з�����
                
                str˵�� = Trim(.TextMatrix(intRow, marBillCol.˵��))
                str˵�� = IIf(str˵�� = "", "Null", "'" & str˵�� & "'")
                dbl��ͼ� = Val(.TextMatrix(intRow, marBillCol.��ͼ�))
                dbl��߼� = Val(.TextMatrix(intRow, marBillCol.��߼�))
                dbl�ӳ��� = Val(.TextMatrix(intRow, marBillCol.�ӳ���))
                dbl�޼� = Val(.TextMatrix(intRow, marBillCol.�ֶ�����޼�))
                
                '�洢���̵Ĳ�������:
                'ZL_���ϼӳɷ���_INSERT(
                '  ���_In     In ���ϼӳɷ���.���%Type,
                '  ��ͼ�_In   In ���ϼӳɷ���.��ͼ�%Type,
                '  ��߼�_In   In ���ϼӳɷ���.��߼�%Type,
                '  �ӳ���_In   In ���ϼӳɷ���.�ӳ���%Type,
                '  ���㷽��_In In ���ϼӳɷ���.���㷽��%Type,
                '  �޼�_In     In ���ϼӳɷ���.�޼�%Type,
                '  ˵��_In     In ���ϼӳɷ���.˵��%Type
                
                strSQL = "ZL_���ϼӳɷ���_INSERT(" & _
                    intRow & "," & _
                    IIf(dbl��ͼ� = 0, "Null", dbl��ͼ�) & "," & _
                    IIf(dbl��߼� = 0, "Null", dbl��߼�) & "," & _
                    IIf(dbl�ӳ��� = 0, "Null", dbl�ӳ���) & "," & _
                    byt���㷽�� & "," & _
                    dbl�޼� & "," & _
                    str˵�� & ")"
                    
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
        gcnOracle.CommitTrans
        mblnReturn = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = 6
        .MsfObj.FixedCols = 1
        .TextMatrix(0, marBillCol.���) = "���"
        .TextMatrix(0, marBillCol.��ͼ�) = "��ͼ�"
        .TextMatrix(0, marBillCol.��߼�) = "��߼�"
        .TextMatrix(0, marBillCol.�ӳ���) = "�ӳ���"
        .TextMatrix(0, marBillCol.�ֶ�����޼�) = "�ֶ�����޼�"
        .TextMatrix(0, marBillCol.˵��) = "˵��"
        
        .ColWidth(marBillCol.���) = 600
        .ColWidth(marBillCol.��ͼ�) = 1400
        .ColWidth(marBillCol.��߼�) = 1400
        .ColWidth(marBillCol.�ӳ���) = 1000
        .ColWidth(marBillCol.�ֶ�����޼�) = 1400
        .ColWidth(marBillCol.˵��) = 2000
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
            
        .ColData(marBillCol.���) = 5
 
        .ColData(marBillCol.��ͼ�) = 4
        .ColData(marBillCol.��߼�) = 4
        .ColData(marBillCol.�ӳ���) = 4
        .ColData(marBillCol.�ֶ�����޼�) = 4
        .ColData(marBillCol.˵��) = 4
        
        .ColAlignment(marBillCol.��ͼ�) = flexAlignRightCenter
        .ColAlignment(marBillCol.��߼�) = flexAlignRightCenter
        .ColAlignment(marBillCol.�ӳ���) = flexAlignRightCenter
        .ColAlignment(marBillCol.�ֶ�����޼�) = flexAlignRightCenter
        .ColAlignment(marBillCol.˵��) = flexAlignLeftCenter
        
        .PrimaryCol = marBillCol.��߼�
        .LocateCol = marBillCol.��߼�
    End With
End Sub
Private Sub CmdHelp_Click()
        ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub

    mblnFirst = False
    err = 0
    On Error Resume Next
    '���ؿ�Ƭ��Ϣ
    If LoadCardInfor = False Then
        Unload Me
        Exit Sub
    End If
   mblnChange = False
    SetCtlEnable
End Sub
Private Sub Form_Load()
    mblnFirst = True
    Call initGrid
    With cbo���㷽��
        .Clear
        .AddItem "0-�������"
        .ItemData(.NewIndex) = 0
        .ListIndex = .NewIndex
        .AddItem "1-�ֶμ���"
        .ItemData(.NewIndex) = 1
    End With
     
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(0, g_�ɱ���)
        .FM_��� = GetFmtString(0, g_���)
        .FM_���ۼ� = GetFmtString(0, g_�ۼ�)
        .FM_���� = GetFmtString(0, g_����)
    End With
    
        
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

Private Sub cmdOk_Click()
    Dim intTmp As Integer

    '��֤�����ֵ�ĺϷ���
    If ValidData() = False Then Exit Sub
    
    '������������ֵ
    If SaveCard() = False Then Exit Sub
    mblnReturn = True
    Unload Me
End Sub

Private Function LoadCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------
    '����:����Ҫ�޸ĵĿ�Ƭ��Ϣ
    '����:���سɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------
    'ֻ�ж��޸Ĳ�������
    Dim intRow As Integer
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    LoadCardInfor = False
    mblnChange = False
    
    err = 0
    On Error GoTo ErrHand:
    
    strSQL = "select ���,��ͼ�,��߼�,�ӳ���,���㷽��,�޼�,˵�� from ���ϼӳɷ��� order by ���"
    
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    Call initGrid
    If Not rsTmp.EOF Then
        For intRow = 0 To cbo���㷽��.ListCount - 1
            If cbo���㷽��.ItemData(intRow) = Val(zlStr.Nvl(rsTmp!���㷽��)) Then
                cbo���㷽��.ListIndex = intRow
                Exit For
            End If
        Next
        txt�޼�.Text = Format(Val(zlStr.Nvl(rsTmp!�޼�)), mFMT.FM_�ɱ���)
    End If
    
    If cbo���㷽��.ListIndex < 0 Then cbo���㷽��.ListIndex = 0
    With mshBill
        .ClearBill
        .Rows = 2
        intRow = 1
        If Not rsTmp.EOF Then
            If rsTmp!��� = 0 Then rsTmp.MoveNext
            Do While Not rsTmp.EOF
                .TextMatrix(intRow, marBillCol.���) = zlStr.Nvl(rsTmp!���, 0)
                .TextMatrix(intRow, marBillCol.��ͼ�) = Format(Val(zlStr.Nvl(rsTmp!��ͼ�)), mFMT.FM_�ɱ���)
                .TextMatrix(intRow, marBillCol.��߼�) = Format(Val(zlStr.Nvl(rsTmp!��߼�)), mFMT.FM_�ɱ���)
                .TextMatrix(intRow, marBillCol.�ӳ���) = Format(Val(zlStr.Nvl(rsTmp!�ӳ���)), GFM_VBJCL)
                .TextMatrix(intRow, marBillCol.�ֶ�����޼�) = Format(Val(zlStr.Nvl(rsTmp!�޼�)), mFMT.FM_�ɱ���)
                .TextMatrix(intRow, marBillCol.˵��) = zlStr.Nvl(rsTmp!˵��)
                .Rows = .Rows + 1
                intRow = intRow + 1
                rsTmp.MoveNext
            Loop
        End If
    End With
    LoadCardInfor = True
    mblnChange = False
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub EditCard(ByVal frmMain As Object, ByVal strPriv As String, ByRef blnreturn As Boolean)

    '------------------------------------------------------------------------------------------------------
    '--����:�༭��Ƭ,��������ô��ڽ���ͨѶ�ĳ���
    '--�����:str��������: Ҫ�༭�ı�����ؼ���
    '         strPriv:Ȩ�޴�
    '--������:BlnReturn,����ֵ,true�������ӻ��޸ĳɹ�.false����δ�������޸�
    '--����:
    '------------------------------------------------------------------------------------------------------
    mstrPriv = strPriv
    mblnChange = False
    mblnReturn = False
    SetCtlEnable
    frm�ӳ�������.Show 1, frmMain
    blnreturn = mblnReturn
End Sub
Private Sub SetCtlEnable()
    '------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Enable����
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnSave As Boolean
    blnSave = Trim(mshBill.TextMatrix(1, marBillCol.��ͼ�)) <> "" Or Trim(mshBill.TextMatrix(1, marBillCol.��߼�)) <> ""
    Me.cmdOk.Enabled = blnSave And mblnChange = True
End Sub

Private Sub ReFormal()
    '------------------------------------------------------------------------------------------------------
    '--����:���㿪ʼ���ʺ���߼�
    '--����:
    '--����:
    '------------------------------------------------------------------------------------------------------
    Dim intRow As Integer
    Dim dbl��� As Double
    Dim dbl��ͼ� As Double
    Dim dbl��߼� As Double
    Dim dbl������߼� As Double
    
    With mshBill
        dbl��� = 0
        dbl������߼� = 0
        For intRow = 1 To .Rows - 1
            dbl��߼� = Val(.TextMatrix(intRow, marBillCol.��߼�))
            dbl��ͼ� = Val(.TextMatrix(intRow, marBillCol.��ͼ�))
            If dbl��߼� <> 0 Or dbl��ͼ� <> 0 Then
                .TextMatrix(intRow, marBillCol.���) = intRow
                If intRow <> 1 Then
                    dbl��� = dbl��߼� - dbl��ͼ�
                    If dbl������߼� <> dbl��ͼ� Then
                        '��������ͼ�,�����㵱ǰ����
                        .TextMatrix(intRow, marBillCol.��ͼ�) = Format(dbl������߼�, mFMT.FM_�ɱ���)
                        If dbl��߼� <> 0 Then
                            .TextMatrix(intRow, marBillCol.��߼�) = Format(dbl������߼� + dbl���, mFMT.FM_�ɱ���)
                        End If
                    End If
                    dbl������߼� = Val(.TextMatrix(intRow, marBillCol.��߼�))
                Else
                    dbl������߼� = dbl��߼�
                End If
            End If
        Next
    End With
End Sub
Private Sub txt�޼�_Change()
 mblnChange = True
 SetCtlEnable
End Sub

Private Sub txt�޼�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txt�޼�_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt�޼�, KeyAscii, m���ʽ)
End Sub

Private Sub txt�޼�_LostFocus()
    txt�޼�.Text = Format(Val(txt�޼�.Text), mFMT.FM_�ɱ���)
End Sub
