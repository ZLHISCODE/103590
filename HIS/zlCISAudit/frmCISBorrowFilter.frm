VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISBorrowFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4425
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6690
   Icon            =   "frmCISBorrowFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra��Χ 
      Height          =   4260
      Left            =   45
      TabIndex        =   24
      Top             =   75
      Width           =   5310
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   33
         Top             =   3105
         Width           =   1605
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   32
         Top             =   3465
         Width           =   1605
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   31
         Top             =   3825
         Width           =   3480
      End
      Begin VB.TextBox txtConver 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   3600
         TabIndex        =   30
         Top             =   2535
         Width           =   1545
      End
      Begin VB.TextBox txtConver 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   1695
         TabIndex        =   29
         Top             =   2535
         Width           =   1545
      End
      Begin VB.TextBox txtConver 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   3600
         TabIndex        =   28
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox txtConver 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   1695
         TabIndex        =   27
         Top             =   2160
         Width           =   1545
      End
      Begin VB.TextBox txtConver 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   26
         Top             =   1785
         Width           =   1545
      End
      Begin VB.TextBox txtConver 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   1695
         TabIndex        =   25
         Top             =   1785
         Width           =   1545
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1395
         Width           =   1605
      End
      Begin VB.CheckBox chk 
         Caption         =   "����׼����(&2)"
         Height          =   300
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   2130
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Caption         =   "�µǼǵ���(&1)"
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   1770
         Value           =   1  'Checked
         Width           =   1470
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   3570
         MaxLength       =   8
         TabIndex        =   3
         Top             =   285
         Width           =   1605
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   1
         Top             =   285
         Width           =   1605
      End
      Begin VB.CheckBox chk 
         Caption         =   "�Ѿܾ�����(&3)"
         Height          =   300
         Index           =   2
         Left            =   135
         TabIndex        =   18
         Top             =   2505
         Width           =   1485
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   8
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1035
         Width           =   1605
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   1665
         MaxLength       =   8
         TabIndex        =   5
         Top             =   660
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   0
         Left            =   1665
         TabIndex        =   11
         Top             =   1755
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   118751235
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   1
         Left            =   3570
         TabIndex        =   13
         Top             =   1755
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   118751235
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   2
         Left            =   1665
         TabIndex        =   15
         Top             =   2130
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   118751235
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   3
         Left            =   3570
         TabIndex        =   17
         Top             =   2130
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   118751235
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   4
         Left            =   1665
         TabIndex        =   19
         Top             =   2505
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   118751235
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   315
         Index           =   5
         Left            =   3570
         TabIndex        =   21
         Top             =   2505
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   118751235
         CurrentDate     =   36263
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   135
         X2              =   5130
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "סԺ��(&4)"
         Height          =   180
         Index           =   5
         Left            =   750
         TabIndex        =   36
         Top             =   3150
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������(&5)"
         Height          =   180
         Index           =   6
         Left            =   570
         TabIndex        =   35
         Top             =   3510
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������(&6)"
         Height          =   180
         Index           =   7
         Left            =   570
         TabIndex        =   34
         Top             =   3870
         Width           =   990
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�ܾ���(&S)"
         Height          =   180
         Index           =   4
         Left            =   795
         TabIndex        =   8
         Top             =   1470
         Width           =   810
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   3330
         TabIndex        =   12
         Top             =   1815
         Width           =   180
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   3330
         TabIndex        =   16
         Top             =   2190
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   3330
         TabIndex        =   2
         Top             =   345
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&No"
         Height          =   180
         Index           =   0
         Left            =   1425
         TabIndex        =   0
         Top             =   345
         Width           =   180
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   3330
         TabIndex        =   20
         Top             =   2565
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��׼��(&P)"
         Height          =   180
         Index           =   3
         Left            =   795
         TabIndex        =   6
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������(&H)"
         Height          =   180
         Index           =   2
         Left            =   795
         TabIndex        =   4
         Top             =   735
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5460
      TabIndex        =   23
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5460
      TabIndex        =   22
      Top             =   165
      Width           =   1100
   End
End
Attribute VB_Name = "frmCISBorrowFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################

Private mrsParam As New ADODB.Recordset
Private mblnDataChanged As Boolean
Private mblnOK As Boolean

'######################################################################################################################

Public Function ShowPara(ByVal frmMain As Object, ByRef rsParam As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnOK = False
    
    Set mrsParam = CopyRecordStruct(rsParam)
    Call CopyRecordData(rsParam, mrsParam)
                
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    If ExecuteCommand("��ȡ����") = False Then Exit Function
    
    Me.Show 1, frmMain
    
    If mblnOK Then
        mrsParam.Filter = ""
        Call DeleteRecordData(rsParam)
        Call CopyRecordData(mrsParam, rsParam)
        ShowPara = mblnOK
    End If
    
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    On Error GoTo errHand

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        chk(0).Value = 1
        chk(1).Value = 0
        chk(2).Value = 0
        
        Call chk_Click(0)
        Call chk_Click(1)
        Call chk_Click(2)
        
        dtp(0).Value = Format(zlDatabase.Currentdate, dtp(0).CustomFormat)
        dtp(1).Value = Format(zlDatabase.Currentdate, dtp(1).CustomFormat)
        dtp(2).Value = Format(zlDatabase.Currentdate, dtp(2).CustomFormat)
        dtp(3).Value = Format(zlDatabase.Currentdate, dtp(3).CustomFormat)
        dtp(4).Value = Format(zlDatabase.Currentdate, dtp(4).CustomFormat)
        dtp(5).Value = Format(zlDatabase.Currentdate, dtp(5).CustomFormat)
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ����"
        
        txt(0).Text = ParamRead(mrsParam, "��ʼ���ݺ�")
        txt(1).Text = ParamRead(mrsParam, "�������ݺ�")
        
        txt(9).Text = ParamRead(mrsParam, "������")
        txt(8).Text = ParamRead(mrsParam, "��׼��")
        txt(2).Text = ParamRead(mrsParam, "�ܾ���")
        
        chk(0).Value = Val(ParamRead(mrsParam, "�µǼǵ���"))
        Call chk_Click(0)
        If ParamRead(mrsParam, "�Ǽǿ�ʼ����") <> "" Then dtp(0).Value = Format(ParamRead(mrsParam, "�Ǽǿ�ʼ����"), dtp(0).CustomFormat)
        If ParamRead(mrsParam, "�Ǽǽ�������") <> "" Then dtp(1).Value = Format(ParamRead(mrsParam, "�Ǽǽ�������"), dtp(1).CustomFormat)
            
        chk(1).Value = Val(ParamRead(mrsParam, "����׼����"))
        Call chk_Click(1)
        If ParamRead(mrsParam, "��׼��ʼ����") <> "" Then dtp(2).Value = Format(ParamRead(mrsParam, "��׼��ʼ����"), dtp(2).CustomFormat)
        If ParamRead(mrsParam, "��׼��������") <> "" Then dtp(3).Value = Format(ParamRead(mrsParam, "��׼��������"), dtp(3).CustomFormat)
        
        chk(2).Value = Val(ParamRead(mrsParam, "�Ѿܾ�����"))
        Call chk_Click(2)
        If ParamRead(mrsParam, "�ܾ���ʼ����") <> "" Then dtp(4).Value = Format(ParamRead(mrsParam, "�ܾ���ʼ����"), dtp(4).CustomFormat)
        If ParamRead(mrsParam, "�ܾ���������") <> "" Then dtp(5).Value = Format(ParamRead(mrsParam, "�ܾ���������"), dtp(5).CustomFormat)
        
        txt(3).Text = ParamRead(mrsParam, "סԺ��")
        txt(4).Text = ParamRead(mrsParam, "��������")
        txt(5).Text = ParamRead(mrsParam, "��������")
        
        DataChanged = False
    '------------------------------------------------------------------------------------------------------------------
    Case "У������"
        
        If chk(0).Value = 0 And chk(1).Value = 0 And chk(2).Value = 0 Then
            ShowSimpleMsg "�Բ��𣬱���ѡ��һ�����ڷ�Χ��"
            chk(0).SetFocus
            Exit Function
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
        
        
        Call ParamWrite(mrsParam, "��ʼ���ݺ�", txt(0).Text)
        Call ParamWrite(mrsParam, "�������ݺ�", txt(1).Text)
        Call ParamWrite(mrsParam, "������", txt(9).Text)
        Call ParamWrite(mrsParam, "��׼��", txt(8).Text)
        Call ParamWrite(mrsParam, "�ܾ���", txt(2).Text)
        
        Call ParamWrite(mrsParam, "�µǼǵ���", chk(0).Value)
        Call ParamWrite(mrsParam, "�Ǽǿ�ʼ����", Format(dtp(0).Value, "yyyy-MM-dd"))
        Call ParamWrite(mrsParam, "�Ǽǽ�������", Format(dtp(1).Value, "yyyy-MM-dd") & " 23:59:59")
                    
        Call ParamWrite(mrsParam, "����׼����", chk(1).Value)
        Call ParamWrite(mrsParam, "��׼��ʼ����", Format(dtp(2).Value, "yyyy-MM-dd"))
        Call ParamWrite(mrsParam, "��׼��������", Format(dtp(3).Value, "yyyy-MM-dd") & " 23:59:59")
        
        Call ParamWrite(mrsParam, "�Ѿܾ�����", chk(2).Value)
        Call ParamWrite(mrsParam, "�ܾ���ʼ����", Format(dtp(4).Value, "yyyy-MM-dd"))
        Call ParamWrite(mrsParam, "�ܾ���������", Format(dtp(5).Value, "yyyy-MM-dd") & " 23:59:59")
        
        Call ParamWrite(mrsParam, "סԺ��", txt(3).Text)
        Call ParamWrite(mrsParam, "��������", txt(4).Text)
        Call ParamWrite(mrsParam, "��������", txt(5).Text)
                
    End Select
    
    ExecuteCommand = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'######################################################################################################################

Private Sub chk_Click(Index As Integer)
    Select Case Index
    Case 0
        dtp(0).Enabled = (chk(Index).Value = 1)
        dtp(1).Enabled = dtp(0).Enabled
    Case 1
        dtp(2).Enabled = (chk(Index).Value = 1)
        dtp(3).Enabled = dtp(2).Enabled
    Case 2
        dtp(4).Enabled = (chk(Index).Value = 1)
        dtp(5).Enabled = dtp(4).Enabled
    End Select
    
    txtConver(0).Visible = (chk(0).Value = 0)
    txtConver(1).Visible = (chk(0).Value = 0)
    txtConver(2).Visible = (chk(1).Value = 0)
    txtConver(3).Visible = (chk(1).Value = 0)
    txtConver(4).Visible = (chk(2).Value = 0)
    txtConver(5).Visible = (chk(2).Value = 0)
    
    DataChanged = True
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If DataChanged Then
        If ExecuteCommand("У������") = False Then Exit Sub
        If ExecuteCommand("�������") Then
            mrsParam.Filter = ""
            mblnOK = True
            DataChanged = False
        End If
    Else
        mrsParam.Filter = ""
        mblnOK = True
    End If
    
    Unload Me
End Sub

Private Sub dtp_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 2, 8, 9
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim strNo As String
    
    Select Case Index
    Case 0, 1
        If KeyCode = vbKeyReturn Then
            If Len(txt(Index).Text) < 8 And Len(txt(Index).Text) > 0 Then
                strNo = txt(Index).Text
                Call MakeNO(78, strNo)
                txt(Index).Text = strNo
            End If
        End If
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 0, 1
        
        Call txt_KeyDown(Index, 13, 0)
    
    Case 2, 8, 9
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
