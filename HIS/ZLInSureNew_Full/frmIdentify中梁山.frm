VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentify����ɽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmIdentify����ɽ.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4020
      TabIndex        =   21
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4020
      TabIndex        =   22
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "���˻�����Ϣ"
      Height          =   4695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3705
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   5
         Left            =   240
         MaxLength       =   14
         TabIndex        =   20
         Top             =   3990
         Width           =   3195
      End
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2670
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1110
         Width           =   2085
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1500
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1320
         MaxLength       =   26
         TabIndex        =   13
         Top             =   2280
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
         Top             =   1890
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   23855107
         CurrentDate     =   36526
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   5
         Top             =   720
         Width           =   2085
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Index           =   1
         Left            =   3120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3090
         Width           =   255
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Index           =   0
         Left            =   3120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   4
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3060
         Width           =   2085
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "ͳ�ﱨ���ۼ�(&P)"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   3720
         Width           =   1350
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���(&K)"
         Height          =   180
         Index           =   16
         Left            =   240
         TabIndex        =   14
         Top             =   2730
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "��������(&B)"
         Height          =   180
         Index           =   15
         Left            =   240
         TabIndex        =   10
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�(&X)"
         Height          =   180
         Index           =   14
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "���֤��(&I)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����(&F)"
         Height          =   180
         Index           =   4
         Left            =   600
         TabIndex        =   16
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����֤��(&Z)"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmIdentify����ɽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum �ı�Enum
    Text���� = 0
    Text���� = 1
    Text����֤�� = 2
    Text���֤�� = 3
    Text���� = 4
    Textͳ�ﱨ���ۼ� = 5
End Enum

Private Enum ѡ��Enum
    Select���� = 0
    Select���� = 1
End Enum

Dim mstrIdentify As String
Dim mbytType As Byte
Dim mlng����ID As Long

Public Function ShowCard(Optional bytType As Byte, Optional lng����ID As Long) As String
'���ܣ�����ҽ�����˵������Ϣ
'������0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�,4-����
'���أ�
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23��������
    Dim rsTemp As New ADODB.Recordset
    Dim lng���ų��� As Long, lng����֤���� As Long
    
    If bytType <> 1 Then
        MsgBox "��ҽ��ֻ֧����Ժ�Ǽǡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mbytType = bytType
    mlng����ID = lng����ID
    mstrIdentify = ""
    
    cbo�Ա�.Clear
    gstrSQL = "select ����,���� from �Ա� order by ����"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Do Until rsTemp.EOF
        cbo�Ա�.AddItem rsTemp("����") & "." & rsTemp("����")
        rsTemp.MoveNext
    Loop
    
    cbo���.Clear
    gstrSQL = "select A.���,A.���� from ������Ⱥ A where A.����=" & TYPE_��������ɽ
    Call OpenRecordset(rsTemp, Me.Caption)
    Do Until rsTemp.EOF
        cbo���.AddItem rsTemp("���") & "." & rsTemp("����")
        cbo���.ItemData(cbo���.NewIndex) = rsTemp("���")
        rsTemp.MoveNext
    Loop
    cbo���.ListIndex = 0
    
    'ȱʡֵ
    lng���ų��� = 20
    lng����֤���� = 26
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=" & TYPE_��������ɽ & " and ����=0"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "���ų���"
                If IsNull(rsTemp("����ֵ")) = False Then lng���ų��� = Val(rsTemp("����ֵ"))
            Case "����֤����"
                If IsNull(rsTemp("����ֵ")) = False Then lng����֤���� = Val(rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    
    txtEdit(Text����).MaxLength = lng���ų���
    txtEdit(Text����֤��).MaxLength = lng����֤����
    
    dtp����.MaxDate = zlDatabase.Currentdate
    frmIdentify����ɽ.Show vbModal
    ShowCard = mstrIdentify
End Function

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strIdentify As String, strAddition As String
    Dim lng����ID As Long, lng���� As Long
    
    '���������ݵ���ȷ��
    If IsValid() = False Then
        Exit Sub
    End If
    
    '�õ��������
    If cbo���.ListIndex < 0 Then
        MsgBox "��ѡ�������", vbInformation, gstrSysName
        cbo���.SetFocus
        Exit Sub
    End If
    lng���� = 0
    
    '��鲡��״̬
    gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=" & TYPE_��������ɽ & " and ����=" & lng���� & " and ҽ����='" & Trim(txtEdit(Text����).Text) & "'"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("״̬") > 0 Then
            MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23��������
    strIdentify = Trim(txtEdit(Text����).Text)                         '0����
    strIdentify = strIdentify & ";" & Trim(txtEdit(Text����).Text)     '1ҽ���� ʹ����ͬ����
    strIdentify = strIdentify & ";"                                    '2����
    strIdentify = strIdentify & ";" & Trim(txtEdit(Text����).Text)     '3����
    strIdentify = strIdentify & ";" & Replace(GetTextFromCombo(cbo�Ա�, True), "'", "") '4�Ա�
    strIdentify = strIdentify & ";" & Format(dtp����.Value, "yyyy-MM-dd") '5��������
    strIdentify = strIdentify & ";" & Trim(txtEdit(Text���֤��).Text)    '6���֤
    strIdentify = strIdentify & ";" & "()"                                '7.��λ����(����)
    strAddition = ";" & lng����                                           '8.���Ĵ���
    strAddition = strAddition & ";"                                       '9.˳���
    strAddition = strAddition & ";"                                       '10��Ա���
    strAddition = strAddition & ";0"                                      '11�ʻ����
    strAddition = strAddition & ";0"                                      '12��ǰ״̬
    strAddition = strAddition & ";" & txtEdit(Text����).Tag               '13����ID
    strAddition = strAddition & ";" & cbo���.ItemData(cbo���.ListIndex) '14��ְ(1,2,3)
    strAddition = strAddition & ";" & Trim(txtEdit(Text����֤��).Text)    '15����֤��
    strAddition = strAddition & ";" & DateDiff("yyyy", dtp����.Value, dtp����.MaxDate) '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";0"                                      '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                                      '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";0"                                      '20����ͳ���ۼ�
    strAddition = strAddition & ";" & Val(txtEdit(Textͳ�ﱨ���ۼ�).Text) '21ͳ�ﱨ���ۼ�
    strAddition = strAddition & ";0"                                      '22סԺ�����ۼ�
    strAddition = strAddition & ";"                                       '23�������� (1����������)
    
    lng����ID = BuildPatiInfo(mbytType, strIdentify & strAddition, mlng����ID, TYPE_��������ɽ)
    '���ظ�ʽ:�м���벡��ID
    If lng����ID > 0 Then
        mstrIdentify = strIdentify & ";" & lng����ID & strAddition
    End If
    Unload Me
End Sub

Private Function IsValid() As Boolean
'���ܣ�������ݵ���ȷ��
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(txtEdit(lngIndex), txtEdit(lngIndex).MaxLength) = False Then
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    If Len(txtEdit(Text����).Text) <> txtEdit(Text����).MaxLength Then
        MsgBox "���ų��Ȳ���" & txtEdit(Text����).MaxLength & "λ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtEdit(Text����)
        txtEdit(Text����).SetFocus
        Exit Function
    End If
    If Trim(txtEdit(Text����).Text) = "" Then
        MsgBox "��������Ϊ�ա�", vbInformation, gstrSysName
        zlControl.TxtSelAll txtEdit(Text����)
        txtEdit(Text����).SetFocus
        Exit Function
    End If
    
    If IsNumeric(txtEdit(Textͳ�ﱨ���ۼ�).Text) = False Then
        MsgBox "ͳ�ﱨ���ۼ�����Ϸ�����ֵ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtEdit(Textͳ�ﱨ���ۼ�)
        txtEdit(Textͳ�ﱨ���ۼ�).SetFocus
        Exit Function
    End If
    
    If Val(txtEdit(Textͳ�ﱨ���ۼ�).Text) < 0 Or Val(txtEdit(Textͳ�ﱨ���ۼ�).Text) > 1000000 Then
        MsgBox "����С��0���Ҳ��ܳ���100��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtEdit(Textͳ�ﱨ���ۼ�)
        txtEdit(Textͳ�ﱨ���ۼ�).SetFocus
        Exit Function
    End If
    
    IsValid = True
End Function

Private Sub cmdSelect_Click(Index As Integer)
    Dim rsTemp As ADODB.Recordset
    
    Select Case Index
        Case Select����
            gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,B.����,B.�Ա�,B.��������,B.���֤��,C.��� as ����ID " & _
                    " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��" & _
                    " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
                    "  where A.����ID=B.����ID and A.����=" & TYPE_��������ɽ & _
                    "  and A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+)"
            
            Call Get�ʻ����
            zlControl.TxtSelAll txtEdit(Text����)
            txtEdit(Text����).SetFocus
        Case Select����
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A where A.����=" & TYPE_��������ɽ
            
            Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , txtEdit(Text����).Text)
            If Not rsTemp Is Nothing Then
                txtEdit(Text����).Text = rsTemp("����")
                txtEdit(Text����).Tag = rsTemp("ID")
                zlControl.TxtSelAll txtEdit(Text����)
            End If
            txtEdit(Text����).SetFocus
    End Select
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

Private Sub dtp����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = Text���� Then
        If KeyCode = vbKeyDelete Then
            txtEdit(Text����).Text = ""
            txtEdit(Text����).Tag = ""
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCode As String
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    If Index = Text���� Then
        If Len(txtEdit(Text����).Text) = txtEdit(Text����).MaxLength Or KeyAscii = vbKeyReturn Then
            strCode = Replace(Trim(txtEdit(Text����).Text), "'", "")
            
            If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) Then 'ˢ��
                str���� = " and A.����='" & strCode & "'"
            ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
                str���� = " and A.����ID=" & Mid(strCode, 2)
            ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��(��ס(��)Ժ�Ĳ���)
                str���� = " and B.סԺ��=" & Mid(strCode, 2)
            ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����(�������ﲡ��)
                str���� = " and B.�����=" & Mid(strCode, 2)
            Else '��������
                str���� = " and B.����='" & strCode & "'"
            End If
        
            gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,B.����,B.�Ա�,B.��������,B.���֤��,C.��� as ����ID " & _
                    " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��" & _
                    " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
                    "  where A.����ID=B.����ID and A.����=" & TYPE_��������ɽ & _
                    "  and A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+)" & str����
            
            Call Get�ʻ����
        End If
    ElseIf KeyAscii = asc("*") Then
        Call cmdSelect_Click(Select����)
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Index = Text���֤�� Then
            strCode = Get��������(txtEdit(Text���֤��).Text, 0)
            If IsDate(strCode) = True Then
                dtp����.Value = CDate(strCode)
            End If
        End If
        KeyAscii = 0  '��������
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cbo���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = Textͳ�ﱨ���ۼ� Then
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "######0.00;0.00;0.00;0.00")
    End If

End Sub

Private Sub Get�ʻ����()
'���Ѿ����ڵļ�¼�ж����ʻ���Ϣ
    Dim rs�ʻ� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    
    Set rs�ʻ� = frmPubSel.ShowSelect(Me, gstrSQL, 0, "�����ʻ�", , txtEdit(Text����).Text, "", False, True)
    If Not rs�ʻ� Is Nothing Then
    
        txtEdit(Text����).Text = rs�ʻ�("����")
        '�������õ�����
        txtEdit(Text����).Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txtEdit(Text���֤��).Text = IIf(IsNull(rs�ʻ�("���֤��")), "", rs�ʻ�("���֤��"))
        txtEdit(Text����).Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        txtEdit(Text����).Tag = IIf(IsNull(rs�ʻ�("����ID")), "", rs�ʻ�("����ID"))
        
        Call SetComboByText(cbo�Ա�, IIf(IsNull(rs�ʻ�("�Ա�")), "", rs�ʻ�("�Ա�")), True)
        txtEdit(Text����֤��).Text = ""
        If IsNull(rs�ʻ�("��������")) = False Then
            dtp����.Value = rs�ʻ�("��������")
        End If
        
        For lngIndex = 0 To cbo���.ListCount - 1
            If cbo���.ItemData(lngIndex) = rs�ʻ�("��ְID") Then
                cbo���.ListIndex = lngIndex
                Exit For
            End If
        Next
        
        '�ٶ����ʻ������Ϣ
        gstrSQL = "select * from �ʻ������Ϣ where ����=" & TYPE_��������ɽ & _
            " and ����ID=" & rs�ʻ�("ID") & " and ���=" & Year(dtp����.MaxDate)
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.EOF = False Then
            '�����ʻ����
            txtEdit(Textͳ�ﱨ���ۼ�).Text = Format(rsTemp("ͳ�ﱨ���ۼ�"), "######0.00;0.00;0.00;0.00")
        Else
            txtEdit(Textͳ�ﱨ���ۼ�).Text = "0.00"
        End If
        
    End If
End Sub

