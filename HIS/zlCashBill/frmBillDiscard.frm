VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillDiscard 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ʊ�ݱ���"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBillDiscard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraBack 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4275
      Left            =   120
      TabIndex        =   20
      Top             =   810
      Width           =   6375
      Begin VB.ComboBox cmb������ 
         Height          =   360
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1635
         Width           =   1830
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   1
         Left            =   1740
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1095
         Width           =   1485
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   2
         Left            =   3990
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1095
         Width           =   1485
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1410
         TabIndex        =   2
         Top             =   90
         Width           =   1815
      End
      Begin VB.OptionButton opt��Χ 
         Caption         =   "���ű���(&S)"
         Height          =   240
         Index           =   0
         Left            =   1425
         TabIndex        =   3
         Top             =   630
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton opt��Χ 
         Caption         =   "���ű���(&M)"
         Height          =   240
         Index           =   1
         Left            =   3180
         TabIndex        =   4
         Top             =   630
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   360
         Left            =   4755
         TabIndex        =   14
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   162332675
         CurrentDate     =   37007
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&G)"
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   11
         Top             =   1695
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��(&D)"
         Height          =   240
         Index           =   3
         Left            =   3375
         TabIndex        =   13
         Top             =   1710
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ������"
         Height          =   240
         Index           =   4
         Left            =   390
         TabIndex        =   1
         Top             =   150
         Width           =   960
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "���뷶Χ(&B)"
         Height          =   240
         Index           =   6
         Left            =   30
         TabIndex        =   5
         Top             =   1155
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   240
         Index           =   5
         Left            =   3330
         TabIndex        =   8
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label lbl˵�� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   2085
         Left            =   30
         TabIndex        =   15
         Top             =   2160
         Width           =   6300
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Index           =   1
         Left            =   1410
         TabIndex        =   6
         Top             =   1095
         Width           =   315
      End
      Begin VB.Label lbl 
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Index           =   2
         Left            =   3660
         TabIndex        =   9
         Top             =   1095
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   3630
      TabIndex        =   17
      Top             =   5430
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   4980
      TabIndex        =   18
      Top             =   5430
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   -270
      TabIndex        =   16
      Top             =   5160
      Width           =   7065
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   420
      Left            =   270
      TabIndex        =   19
      Top             =   5430
      Width           =   1200
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ʊ�ݱ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2760
   End
End
Attribute VB_Name = "frmBillDiscard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintƱ�� As gBillType
Private mstrPrivs As String
Private mstrID As String

Private mblnOK As Boolean
Private mblnChange As Boolean     'Ϊ��ʱ��ʾ�Ѹı���
Private mstrǰ׺ As String
Private mstr��С���� As String
Private mstr������ As String
Private mlngƱ�ݳ��� As Long
Private mblnIsBIll As Boolean '��ǰƱ���Ƿ�ΪƱ��

Private Sub InitContext()
    Dim dtCurrnet As Date
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    
    On Error GoTo errHandle
    dtCurrnet = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    dtpDate.Value = dtCurrnet
    dtpDate.MaxDate = dtCurrnet
    
    If mblnIsBIll Then
        lblTitle.Caption = "Ʊ�ݱ���"
        lbl(6).Caption = "���뷶Χ(&B)"
    Else
        lblTitle.Caption = IIf(mintƱ�� = gBillType.���￨, "ҽ�ƿ�����", "���ѿ�����")
        lbl(6).Caption = "���ŷ�Χ(&B)"
    End If
    
    txtEdit(0).Text = _
        Choose(mintƱ��, "�շ��վ�", "Ԥ���վ�", "�����վ�", "�Һ��վ�", "���￨", "���ѿ�", "��Ա��")
    
    mblnChange = True
    Select Case mintƱ��
        Case gBillType.�շ��վ�
            strWhere = " And B.��Ա����='�����շ�Ա'"
        Case gBillType.Ԥ���վ�
            strWhere = " And B.��Ա���� in ('Ԥ���տ�Ա','��Ժ�Ǽ�Ա')"
        Case gBillType.�����վ�
            strWhere = " And B.��Ա����='סԺ����Ա'"
        Case gBillType.�Һ��վ�
            strWhere = " And B.��Ա����='����Һ�Ա'"
        Case gBillType.���￨, gBillType.���ѿ�
            strWhere = " And B.��Ա���� in ('�����Ǽ���','��Ժ�Ǽ�Ա')"
        Case Else
            Exit Sub
    End Select
    strSQL = _
        "Select Distinct A.����" & vbNewLine & _
        "From ��Ա�� A,��Ա����˵�� B" & vbNewLine & _
        "Where A.ID=B.��ԱID " & strWhere & _
        "      And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        "      And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & vbNewLine & _
        "Order By A.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cmb������.Clear
    Do Until rsTemp.EOF
        cmb������.AddItem rsTemp("����")
        rsTemp.MoveNext
    Loop
    If cmb������.ListCount > 0 Then cmb������.ListIndex = 0

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmb������_Click()
    mblnChange = True
End Sub

Private Sub cmb������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub dtpDate_Change()
    mblnChange = True
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If ValidateContent() = False Then Exit Sub
    If MsgBox("һ������󣬱���" & IIf(mblnIsBIll, "����", "����") & "�Ͳ�����ʹ���ˡ�" & vbCrLf & _
        "�Ƿ�ȷ��Ҫ������", _
        vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If Save() = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Sub opt��Χ_Click(Index As Integer)
    mblnChange = True
    If opt��Χ(0).Value = True Then
        txtEdit(2).Enabled = False
        txtEdit(2).Text = txtEdit(1).Text
    Else
        txtEdit(2).Enabled = True
    End If
    Call ShowSum
End Sub

Private Sub opt��Χ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 And opt��Χ(0).Value = True Then txtEdit(2).Text = txtEdit(1).Text
    Call ShowSum
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyDelete _
        And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
    If (Index = 1 Or Index = 2) And (KeyAscii >= vbKey0 Or KeyAscii <= vbKey9) _
        And txtEdit(Index).SelLength = 0 Then
        If Len(txtEdit(Index)) >= mlngƱ�ݳ��� Then KeyAscii = 0
    End If
End Sub

Private Function ValidateContent() As Boolean
'����:����������ݵ��Ƿ���Ч
'����:��Ч�򷵻�True,���򷵻�False
    Dim lngCount As Long, i As Integer
    Dim strTemp As String, strName As String
    Dim strNOs As String, varPara() As Variant, strTable As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    strName = IIf(mblnIsBIll, "����", "����")
    '�ַ������
    For lngCount = 1 To 2
        txtEdit(lngCount).Text = Trim(txtEdit(lngCount).Text)
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            txtEdit(lngCount).SetFocus
            zlControl.TxtSelAll txtEdit(lngCount)
            Exit Function
        End If
        For i = 1 To Len(txtEdit(lngCount).Text)
            strTemp = Mid(txtEdit(lngCount), i, 1)
            If InStr("0123456789", strTemp) = 0 Then
                MsgBox strName & "�к��з������ַ���", vbExclamation, gstrSysName
                txtEdit(lngCount).SetFocus
                zlControl.TxtSelAll txtEdit(lngCount)
                Exit Function
            End If
        Next
        If Len(txtEdit(lngCount).Text) <> Len(txtEdit(lngCount).Tag) - Len(mstrǰ׺) Then
            MsgBox strName & "�ĳ��Ȳ��ԡ�", vbExclamation, gstrSysName
            txtEdit(lngCount).SetFocus
            zlControl.TxtSelAll txtEdit(lngCount)
            Exit Function
        End If
    Next
    
    If mstrǰ׺ & txtEdit(1).Text < txtEdit(1).Tag Then
        MsgBox "���ϵĿ�ʼ" & strName & "����������õĿ�ʼ" & strName & "��", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        zlControl.TxtSelAll txtEdit(1)
        Exit Function
    End If
    If txtEdit(2).Enabled = True Then
        If mstrǰ׺ & txtEdit(2).Text > txtEdit(2).Tag Then
            MsgBox "���ϵ���ֹ" & strName & "����С�����õ���ֹ" & strName & "��", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            zlControl.TxtSelAll txtEdit(2)
            Exit Function
        End If
    Else
        If mstrǰ׺ & txtEdit(1).Text > txtEdit(2).Tag Then
            MsgBox "���ϵ�" & strName & "����С�����õ���ֹ" & strName & "��", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            zlControl.TxtSelAll txtEdit(1)
            Exit Function
        End If
    End If
        
    If txtEdit(1).Text > txtEdit(2).Text Then
        MsgBox "���ϵĿ�ʼ" & strName & "����С�����ϵ���ֹ" & strName & "��", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        zlControl.TxtSelAll txtEdit(1)
        Exit Function
    End If
    If Val(txtEdit(2).Text) - Val(txtEdit(1).Text) + 1 > 10000 Then
        MsgBox "һ�����ϵ����������ܳ���һ���š�", vbExclamation, gstrSysName
        txtEdit(2).SetFocus
        zlControl.TxtSelAll txtEdit(2)
        Exit Function
    End If
    
    If mstr��С���� <> "" Then
        '����Ƿ����ʹ��
        If SplitCardNos(mstrǰ׺ & txtEdit(1).Text & "��" & mstrǰ׺ & txtEdit(2).Text, strNOs) = False Then Exit Function
        varPara = Array(mstrID)
        If FromStringListBulidSQL(0, strNOs, varPara, strTable, strName, 2) = False Then Exit Function
        If mintƱ�� = gBillType.���ѿ� Then
            strSQL = _
                "Select Distinct a.���� As ����" & vbNewLine & _
                "From ���ѿ�ʹ�ü�¼ A, (" & strTable & ") B" & vbNewLine & _
                "Where a.���� = b.���� And a.����id = [1]"
        Else
            strSQL = _
                "Select Distinct a.����" & vbNewLine & _
                "From Ʊ��ʹ����ϸ A, (" & strTable & ") B" & vbNewLine & _
                "Where a.���� = b.���� And a.����id = [1]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, varPara)
        strTemp = ""
        Do While Not rsTemp.EOF
            strTemp = strTemp & "," & Nvl(rsTemp!����)
            rsTemp.MoveNext
        Loop
        If strTemp <> "" Then
            strTemp = Mid(strTemp, 2)
            ShowMsgbox "����" & strName & "�Ѿ�ʹ�û��ѱ����ϣ����������ϣ�" & vbCrLf & strTemp
            zlControl.ControlSetFocus txtEdit(1)
            zlControl.TxtSelAll txtEdit(1)
            Exit Function
        End If
    End If
    If cmb������.Text = "" Then
        MsgBox "�����˲���Ϊ�ա�", vbExclamation, gstrSysName
        cmb������.SetFocus
        Exit Function
    End If
    
    ValidateContent = True
End Function

Private Function Save() As Boolean
'����:����༭������
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim strTemp As String
    Dim lngID As Long
    Dim strSQL As String
    
    On Error GoTo errHandle
    If mintƱ�� = gBillType.���ѿ� Then
        'Zl_���ѿ�ʹ�ü�¼_Damage
        strSQL = "Zl_���ѿ�ʹ�ü�¼_Damage("
        '  ����id_In   In ���ѿ�ʹ�ü�¼.����id%Type,
        strSQL = strSQL & "" & mstrID & ","
        '  ǰ׺_In     In Ʊ�����ü�¼.ǰ׺�ı�%Type,
        strSQL = strSQL & "'" & mstrǰ׺ & "',"
        '  ��ʼ����_In In ���ѿ�ʹ�ü�¼.����%Type,
        strSQL = strSQL & "'" & txtEdit(1).Text & "',"
        '  ��������_In In ���ѿ�ʹ�ü�¼.����%Type,
        strSQL = strSQL & "'" & txtEdit(2).Text & "',"
        '  ʹ��ʱ��_In In ���ѿ�ʹ�ü�¼.ʹ��ʱ��%Type := Null,
        strSQL = strSQL & "To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
        '  ʹ����_In   In ���ѿ�ʹ�ü�¼.ʹ����%Type := Null
        strSQL = strSQL & "'" & cmb������.Text & "')"
    Else
        'Zl_Ʊ��ʹ����ϸ_Damage
        strSQL = "Zl_Ʊ��ʹ����ϸ_Damage("
        '  ����id_In   In Ʊ��ʹ����ϸ.����id%Type,
        strSQL = strSQL & "" & mstrID & ","
        '  Ʊ��_In     In Ʊ��ʹ����ϸ.Ʊ��%Type,
        strSQL = strSQL & "" & mintƱ�� & ","
        '  ǰ׺_In     In Ʊ�����ü�¼.ǰ׺�ı�%Type,
        strSQL = strSQL & "'" & mstrǰ׺ & "',"
        '  ��ʼ����_In In Ʊ��ʹ����ϸ.����%Type,
        strSQL = strSQL & "'" & txtEdit(1).Text & "',"
        '  ��������_In In Ʊ��ʹ����ϸ.����%Type,
        strSQL = strSQL & "'" & txtEdit(2).Text & "',"
        '  ʹ��ʱ��_In In Ʊ��ʹ����ϸ.ʹ��ʱ��%Type := Null,
        strSQL = strSQL & "To_Date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
        '  ʹ����_In   In Ʊ��ʹ����ϸ.ʹ����%Type := Null
        strSQL = strSQL & "'" & cmb������.Text & "')"
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If gblnBillPrint Then
        Call gobjBillPrint.zlDiscardBill(mstrID, Val(txtEdit(0).Tag), _
            mstrǰ׺, txtEdit(1).Text, txtEdit(2).Text, dtpDate.Value, cmb������.Text)
    End If
    
    Save = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowSum()
'����:��ʾ������Ϣ
    Dim strTemp As String
    Dim strName1 As String, strName2 As String
    
    strName1 = IIf(mblnIsBIll, "����", "����")
    strName2 = IIf(mblnIsBIll, "Ʊ��", "��Ƭ")
    
    '���ϵĿ�ʼ����:
    '���ϵĽ�������:
    '���ϵ�Ʊ��������:
    '
    '���õĿ�ʼ����:
    '���õĽ�������:
    '�Ѿ�ʹ�õ���С����:
    '�Ѿ�ʹ�õ�������:
    
    strTemp = " ���ϵĿ�ʼ" & strName1 & "��" & lbl(1).Caption & txtEdit(1).Text & vbCrLf
    strTemp = strTemp & "  ���ϵĽ���" & strName1 & "��" & lbl(2).Caption & txtEdit(2).Text & vbCrLf
    If txtEdit(1).Text = "" Or txtEdit(2).Text = "" Then
        strTemp = strTemp & "  ���ϵ�" & strName2 & "��������" & vbCrLf & vbCrLf
    Else
        strTemp = strTemp & "  ���ϵ�" & strName2 & "��������" & Val(txtEdit(2).Text) - Val(txtEdit(1).Text) + 1 & vbCrLf & vbCrLf
    End If
    strTemp = strTemp & "  ���õĿ�ʼ" & strName1 & "��" & Replace(txtEdit(1).Tag, "&", "&&") & vbCrLf
    strTemp = strTemp & "  ���õĽ���" & strName1 & "��" & Replace(txtEdit(2).Tag, "&", "&&") & vbCrLf
    If mstr��С���� <> "" Then
        strTemp = strTemp & "  �Ѿ�ʹ�õ���С" & strName1 & "��" & Replace(mstr��С����, "&", "&&") & vbCrLf
        strTemp = strTemp & "  �Ѿ�ʹ�õ����" & strName1 & "��" & Replace(mstr������, "&", "&&") & vbCrLf
    End If
    
    lbl˵��.Caption = strTemp
End Sub

Public Function �༭Ʊ�ݱ���(frmParent As Object, ByVal strPrivs As String, _
    ByVal intƱ�� As gBillType, ByVal strID As String) As Boolean
    '����:��������õĲ����ش��ڽ���ͨѶ�ĳ���,�������ӽɿ��¼
    '����:
    '����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
        
    mstrPrivs = strPrivs
    mintƱ�� = intƱ��: mstrID = strID
    
    mblnIsBIll = CurrentIsBill(intƱ��)
    Call InitContext
    
    If mintƱ�� = gBillType.���ѿ� Then
        strSQL = _
            "Select ������,ǰ׺�ı�,��ʼ���� As ��ʼ����,��ֹ���� As ��ֹ����,��ǰ���� As ��ǰ����,ʹ�÷�ʽ" & vbNewLine & _
            "From ���ѿ����ü�¼" & vbNewLine & _
            "Where ID=[1]"
    Else
        strSQL = _
            "Select ������,ǰ׺�ı�,��ʼ����,��ֹ����,��ǰ����,ʹ�÷�ʽ" & vbNewLine & _
            "From Ʊ�����ü�¼" & vbNewLine & _
            "Where ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrID)
    
    mstrǰ׺ = Nvl(rsTemp!ǰ׺�ı�)
    lbl(1).Caption = Replace(mstrǰ׺, "&", "&&")
    lbl(2).Caption = lbl(1).Caption
    txtEdit(1).Tag = Nvl(rsTemp!��ʼ����)
    txtEdit(2).Text = Mid(Nvl(rsTemp!��ֹ����), Len(mstrǰ׺) + 1)
    mlngƱ�ݳ��� = Len(Mid(Nvl(rsTemp!��ֹ����), Len(mstrǰ׺) + 1))
    txtEdit(2).Tag = Nvl(rsTemp!��ֹ����)
    If IsNull(rsTemp!��ǰ����) Then
        txtEdit(1).Text = Mid(Nvl(rsTemp!��ʼ����), Len(mstrǰ׺) + 1)
    Else
        '�Ѿ�ʹ�ã��Ͱ����ֵ��һ
        txtEdit(1).Text = Mid(zlStr.Increase(Nvl(rsTemp!��ǰ����)), Len(mstrǰ׺) + 1)
    End If
    
    On Error Resume Next
    If Val(rsTemp!ʹ�÷�ʽ) = 2 Then    '����ʽ��,ֻ��ѡ��Ϊ������Ա:35846
        cmb������.Text = UserInfo.����
    Else
        cmb������.Text = Nvl(rsTemp!������)
    End If
    If Err <> 0 Then
        If Val(rsTemp!ʹ�÷�ʽ) = 2 Then
            cmb������.AddItem UserInfo.����
            cmb������.ListIndex = cmb������.NewIndex
        Else
            cmb������.AddItem Nvl(rsTemp!������)
            cmb������.ListIndex = cmb������.NewIndex
        End If
    End If
    If InStr(mstrPrivs, "���в���Ա") = 0 Then cmb������.Enabled = False
    On Error GoTo errHandle
    
    If mintƱ�� = gBillType.���ѿ� Then
        strSQL = _
            "Select Nvl(Min(����), ' ') As ��С����, Nvl(Max(����), ' ') As ������" & vbNewLine & _
            "From ���ѿ�ʹ�ü�¼" & vbNewLine & _
            "Where ����id =[1]"
    Else
        strSQL = _
            "Select Nvl(Min(����), ' ') As ��С����, Nvl(Max(����), ' ') As ������" & vbNewLine & _
            "From Ʊ��ʹ����ϸ" & vbNewLine & _
            "Where ����id =[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrID)
    
    mstr��С���� = Trim(Nvl(rsTemp!��С����))
    mstr������ = Trim(Nvl(rsTemp!������))
    Call opt��Χ_Click(0)
    
    mblnOK = False
    mblnChange = False
    frmBillDiscard.Show vbModal, frmParent
    �༭Ʊ�ݱ��� = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SplitCardNos(ByVal strCardNoRange As String, ByRef strCardNos As String) As Boolean
    '����:���ݴ���Ŀ��ŷ�Χ���ֽ����صĿ���
    '���:
    '   strCardNoRange-���ŷ�Χ
    '����:
    '   strCardNos-���ؿ�����(�ö��ŷ���)
    '����:�ֽ�ɹ�����True�����򷵻�False
    Dim varData As Variant, lngCount As Long
    Dim strCardStartNO As String, strCardEndNO As String, strCurNo As String
    Dim str���� As String

    varData = Split(strCardNoRange & "��", "��")
    strCardStartNO = varData(0): strCardEndNO = varData(1)
    If strCardEndNO = "" Then
        strCardNos = strCardStartNO
        SplitCardNos = True
        Exit Function
    End If
    If strCardStartNO > strCardEndNO Then Exit Function
    
    str���� = zlStr.ExpressValue(strCardEndNO & "-" & strCardStartNO & "+1")
    If InStr(UCase(str����), "E") > 0 Or Len(str����) > 4 Then '����̫���Ѿ���ɿ�ѧ���㷨
        ShowMsgbox "���ŷ�Χ���ܴ���10000����ֶ����ϣ�"
        Exit Function
    End If
    
    strCurNo = strCardStartNO
    strCardNos = strCardStartNO
    Do While True
        If strCurNo >= strCardEndNO Then Exit Do
        strCurNo = zlStr.Increase(strCurNo)
        strCardNos = strCardNos & "," & strCurNo
        
        lngCount = lngCount + 1
        If lngCount > 10000 Then
            ShowMsgbox "���ŷ�Χ���ܴ���10000����ֶ����ϣ�"
            Exit Function
        End If
    Loop
    SplitCardNos = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FromStringListBulidSQL(ByVal bytBulidType As Byte, ByVal strValues As String, _
    ByRef varPara As Variant, ByRef strBulitSQL As String, _
    ByVal strColumnAliaName As String, Optional intStartPara As Integer = 1) As Boolean
    '����:������ֵ(ֵ�б���ɵ�)�����Ĳ����ֽ�Ϊ���ж��������SQL,��:select ... From str2List Union ALL Selelct ..
    '���:strValues-ֵ,����ö��ŷ���
    '     strColumnAliaName-�б���
    '     bytType-0-�ַ���;1-������;
    '     intStartPara-�����Ĳ������
    '����:varPara-���صĲ���ֵ������
    '     strBulitSQL-���صĹ�����SQL��
    '����:�����ȡ�ɹ�,����true,���򷵻�False
    Dim varData As Variant, strTemp As String
    Dim i As Long, j As Long, strSQL As String
    Dim strTable As String, strColumnName As String
    
    On Error GoTo ErrHandler
    strColumnName = " a.Column_Value "
    If strColumnAliaName <> "" Then strColumnName = strColumnName & " As " & strColumnAliaName
    
    If bytBulidType = 0 Then
        strTable = "Table(f_str2list([0]))"
    Else
        strTable = "Table(f_Num2list([0]))"
    End If
    
    j = intStartPara
    ReDim Preserve varPara(0 To j - 1)
    
    varData = Split(strValues, ",")
    strTemp = ""
    For i = 0 To UBound(varData)
        If zlCommFun.ActualLen(strTemp & "," & varData(i)) > 4000 Then
            strSQL = strSQL & " Union ALL " & _
                " Select /*+cardinality(a,10) */" & strColumnName & _
                " From " & Replace(strTable, "[0]", "[" & j & "]") & " A"
            ReDim Preserve varPara(0 To j - 1)
            varPara(j - 1) = Mid(strTemp, 2)
            j = j + 1: strTemp = ""
        End If
        strTemp = strTemp & "," & varData(i)
    Next
    If strTemp <> "" Then
        strSQL = strSQL & " Union ALL " & _
            " Select /*+cardinality(a,10) */" & strColumnName & _
            " From " & Replace(strTable, "[0]", "[" & j & "]") & " A"
        ReDim Preserve varPara(0 To j - 1)
        varPara(j - 1) = Mid(strTemp, 2)
    End If
    
    If strSQL <> "" Then strSQL = Mid(strSQL, 11)
    strBulitSQL = strSQL
    FromStringListBulidSQL = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

