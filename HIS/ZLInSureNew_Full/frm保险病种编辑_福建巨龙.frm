VERSION 5.00
Begin VB.Form frm���ղ��ֱ༭_�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ղ��ֱ༭"
   ClientHeight    =   5295
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   4620
   Icon            =   "frm���ղ��ֱ༭_��������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra��Ŀ 
      Caption         =   "��׼ʹ����Ŀ"
      Height          =   2325
      Left            =   120
      TabIndex        =   10
      Top             =   2340
      Width           =   4365
      Begin VB.CommandButton cmdClear 
         Caption         =   "���(&L)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   3150
         TabIndex        =   14
         Top             =   1770
         Width           =   1100
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "ɾ��(&D)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   3150
         TabIndex        =   13
         Top             =   1320
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   3150
         TabIndex        =   12
         Top             =   270
         Width           =   1100
      End
      Begin VB.ListBox lst��Ŀ 
         Height          =   1860
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   2955
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "����"
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   4365
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   825
         MaxLength       =   20
         TabIndex        =   2
         Top             =   390
         Width           =   1995
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   825
         MaxLength       =   50
         TabIndex        =   4
         Top             =   780
         Width           =   3375
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   825
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1185
         Width           =   1095
      End
      Begin VB.OptionButton opt��� 
         Caption         =   "���Բ�(&M)"
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   1740
         Width           =   1155
      End
      Begin VB.OptionButton opt��� 
         Caption         =   "��ͨ��(&G)"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   7
         Top             =   1740
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opt��� 
         Caption         =   "���ֲ�(&S)"
         Height          =   180
         Index           =   2
         Left            =   2745
         TabIndex        =   9
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   840
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   1245
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   450
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   195
      TabIndex        =   17
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2160
      TabIndex        =   15
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3360
      TabIndex        =   16
      Top             =   4800
      Width           =   1100
   End
End
Attribute VB_Name = "frm���ղ��ֱ༭_��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    text���� = 0
    Text���� = 1
    Text���� = 2
End Enum

Dim mlng���� As Long
Dim mstrID As String         '��ǰ�༭��ҽ������ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub cmdADD_Click()
    Dim strID As String, str���� As String, str���� As String
    Dim blnReturn As Boolean
    Dim lngIndex As Long
    
    If frm�շ�ϸĿѡ��.ShowTree(strID, str����, str����) = True Then
        For lngIndex = 0 To lst��Ŀ.ListCount - 1
            '�Ѿ�����
            If lst��Ŀ.ItemData(lngIndex) = Val(strID) Then Exit Sub
        Next
        
        lst��Ŀ.AddItem "��" & str���� & "��" & str����
        lst��Ŀ.ItemData(lst��Ŀ.NewIndex) = Val(strID)
        
        mblnChange = True
    End If
End Sub

Private Sub CmdClear_Click()
    lst��Ŀ.Clear
    mblnChange = True
End Sub

Private Sub cmdDelete_Click()
    If lst��Ŀ.ListIndex < 0 Then Exit Sub
    
    lst��Ŀ.RemoveItem lst��Ŀ.ListIndex
    mblnChange = True
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
        txtEdit(text����).Text = GetMaxCode 'zlDatabase.GetMax("���ղ���", "����", 6, " where ����=" & mlng����)
        For lngIndex = Text���� To Text����
            txtEdit(lngIndex).Text = ""
        Next
        lst��Ŀ.Clear
        
        mblnChange = False
        txtEdit(text����).SetFocus
    Else
        mblnChange = False
        Unload Me
    End If
End Sub

Private Function GetMaxCode() As String
'���ܣ���ȡָ����ı�����������ֵ
'���أ��ɹ����� �¼�������; ���߷��� 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant
    Dim lngLengh As Long
    
    On Error GoTo ErrHand
    With rsTemp
        gstrSQL = "SELECT max(length(substr(����,1,instr(����,'@@')-1))) as �ֵ FROM ���ղ��� where ����=" & mlng����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF = True Then
            GetMaxCode = "1"
            Exit Function
        Else
            lngLengh = Nvl(rsTemp("�ֵ"), "1")
        End If
        
        gstrSQL = "SELECT MAX(LPAD(substr(����,1,instr(����,'@@')-1)," & lngLengh & ",' ')) as ���ֵ FROM ���ղ��� where ����=" & mlng����
        Call OpenRecordset(rsTemp, Me.Caption)
        If rsTemp.EOF Then
            GetMaxCode = Format(1, String(lngLengh, "0"))
            Exit Function
        End If
        
        varTemp = Nvl(rsTemp("���ֵ"), "0")
        If IsNumeric(varTemp) Then
            GetMaxCode = CStr(Val(varTemp) + 1)
            GetMaxCode = Format(GetMaxCode, String(lngLengh, "0"))
        Else
            GetMaxCode = Mid(varTemp, 1, Len(varTemp) - 1) & Chr(asc(Right(varTemp, 1)) + 1)
            GetMaxCode = Trim(GetMaxCode)
        End If
        .Close
    End With
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function Save��Ŀ() As Boolean
    Dim lngID As Long, lng��� As Long
    Dim lngIndex As Long, lst As ListItem
    Dim str��׼��Ŀ As String
    Dim strCode As String
    Dim rsTmp As New ADODB.Recordset
    
    For lngIndex = opt���.LBound To opt���.UBound
        If opt���(lngIndex).Value = True Then
            lng��� = lngIndex
            Exit For
        End If
    Next
    
    For lngIndex = 0 To lst��Ŀ.ListCount - 1
        str��׼��Ŀ = str��׼��Ŀ & lst��Ŀ.ItemData(lngIndex) & ":"
    Next
    
    On Error GoTo errHandle
    
    If mstrID = "" Then
        '����
        If CheckCode(txtEdit(text����)) = False Then Exit Function
        lngID = zlDatabase.GetNextId("���ղ���")
        '��ȡ���ձ���
        strCode = zlDatabase.GetMax("���ղ���", "����", 6, " Where ����=" & mlng����)
        gstrSQL = "zl_���ղ���_INSERT(" & lngID & "," & mlng���� & ",'" & strCode & "','" & _
                Trim(txtEdit(text����).Text) & "@@" & Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng��� & ",null,null,'" & str��׼��Ŀ & "')"
    Else
        '�޸�
        If CheckCode(txtEdit(text����), False) = False Then Exit Function
        '��ȡ���ձ���
        gstrSQL = "Select ���� From ���ղ��� Where ����=" & mlng���� & " And ID=" & mstrID
        Call OpenRecordset(rsTmp, "��ȡ��ǰ���ղ��ֵı���")
        strCode = rsTmp!����
        
        gstrSQL = "zl_���ղ���_Update(" & mstrID & ",'" & strCode & "','" & _
                Trim(txtEdit(text����).Text) & "@@" & Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng��� & ",null,null,'" & str��׼��Ŀ & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '����������
    If mstrID = "" Then
        Set lst = frm���ղ���.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text����), "Disease", "Disease")
    Else
        Set lst = frm���ղ���.lvwItem.SelectedItem
    End If
    lst.SubItems(1) = Trim(txtEdit(Text����).Text)
    lst.SubItems(2) = Trim(txtEdit(Text����).Text)
    lst.SubItems(3) = IIf(lng��� = 0, "��ͨ��", IIf(lng��� = 1, "���Բ�", "���ֲ�"))
    
    Save��Ŀ = True
    mblnOK = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckCode(ByVal strCode As String, Optional ByVal blnNew As Boolean = True) As Boolean
    Dim rsCode As New ADODB.Recordset
    '��Ϊ���볬����ֻ�н����������Ʊ����������У���������ʵ�ʱ�����Ǽ�¼�������û��޸ı���ʱ����Ҫ�жϱ����Ƿ��ظ�
    
    CheckCode = False
    gstrSQL = "Select 1 From ���ղ��� Where substr(����,1,instr(����,'@@')-1)='" & strCode & "'" & IIf(blnNew, "", " And ID<>" & mstrID)
    Call OpenRecordset(rsCode, "�жϱ����Ƿ��ظ�")
    
    If Not rsCode.EOF Then
        MsgBox "���ղ��ֱ����ظ���", vbInformation, gstrSysName
        txtEdit(text����).SetFocus
        Exit Function
    End If
    CheckCode = True
End Function

Private Function IsValid() As Boolean
'����:���������й�ҽ�����������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim lngIndex As Integer
    For lngIndex = text���� To Text����
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
    Next
    
    If lst��Ŀ.ListCount > 50 Then
        MsgBox "�ò��ֵ���׼ҽ����Ŀ̫�࣬���ܳ���50����", vbInformation, gstrSysName
        Exit Function
    End If
    IsValid = True
End Function

Private Sub opt���_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt���_KeyPress(Index As Integer, KeyAscii As Integer)
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
End Sub

Public Function �༭����(ByVal lng���� As Long, ByVal strID As String) As Boolean
'����:��������õ�ҽ���������ڽ���ͨѶ�ĳ���
'����:str���           ��ǰ�༭��ҽ�����ĵ����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mblnOK = False
    mlng���� = lng����
    mstrID = strID
    
    rsTemp.CursorLocation = adUseClient
    If mstrID <> "" Then
        '�޸�ҽ������
        gstrSQL = "select substr(����,1,instr(����,'@@')-1) ����,substr(����,instr(����,'@@')+2) ����,����,nvl(���,'0') as ��� from ���ղ��� where ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        txtEdit(text����).Text = rsTemp("����")
        txtEdit(Text����).Text = rsTemp("����")
        txtEdit(Text����).Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        opt���(rsTemp("���")).Value = True
        
        '�޸�ҽ������
        gstrSQL = "select A.ID,A.����,A.���� from �շ�ϸĿ A,������׼��Ŀ B where A.ID=B.�շ�ϸĿID and B.����ID=" & mstrID
        Call OpenRecordset(rsTemp, Me.Caption)
        
        Do Until rsTemp.EOF
            lst��Ŀ.AddItem "��" & rsTemp("����") & "��" & rsTemp("����")
            lst��Ŀ.ItemData(lst��Ŀ.NewIndex) = rsTemp("ID")
            rsTemp.MoveNext
        Loop
    Else
        '����ҽ������
        txtEdit(text����).Text = GetMaxCode 'zlDatabase.GetMax("���ղ���", "����", 6, " where ����=" & mlng����)
    End If
    
    
    mblnChange = False
    frm���ղ��ֱ༭_��������.Show vbModal, frm���ղ���
    �༭���� = mblnOK
End Function

