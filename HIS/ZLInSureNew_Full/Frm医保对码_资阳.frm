VERSION 5.00
Begin VB.Form Frmҽ������_���� 
   Caption         =   "ҽ����Ŀ����"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   7125
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȷ�� 
      Appearance      =   0  'Flat
      Caption         =   "��    ��"
      Height          =   375
      Left            =   5160
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txt��־ 
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1200
      TabIndex        =   25
      Top             =   3080
      Width           =   1095
   End
   Begin VB.CommandButton cmd��ѯ 
      Appearance      =   0  'Flat
      Caption         =   "��   ѯ"
      Height          =   375
      Left            =   2400
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmd�걨 
      Appearance      =   0  'Flat
      Caption         =   "��    ��"
      Height          =   375
      Left            =   720
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox cmd�ѱ� 
      Height          =   300
      ItemData        =   "Frmҽ������_����.frx":0000
      Left            =   4560
      List            =   "Frmҽ������_����.frx":000D
      TabIndex        =   21
      Top             =   2595
      Width           =   1095
   End
   Begin VB.ComboBox cmd��� 
      Height          =   300
      ItemData        =   "Frmҽ������_����.frx":0023
      Left            =   4560
      List            =   "Frmҽ������_����.frx":0030
      TabIndex        =   20
      Top             =   160
      Width           =   1095
   End
   Begin VB.TextBox txt������Ŀ 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2595
      Width           =   2175
   End
   Begin VB.TextBox txt��� 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2115
      Width           =   2175
   End
   Begin VB.TextBox txt���� 
      Height          =   270
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1635
      Width           =   2175
   End
   Begin VB.TextBox txt��λ 
      Height          =   270
      Left            =   1200
      TabIndex        =   16
      Top             =   1635
      Width           =   2175
   End
   Begin VB.TextBox txt���� 
      Height          =   270
      Left            =   4560
      TabIndex        =   15
      Top             =   1155
      Width           =   2175
   End
   Begin VB.TextBox txt���� 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1155
      Width           =   2175
   End
   Begin VB.TextBox txtӢ������ 
      Height          =   270
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   675
      Width           =   2175
   End
   Begin VB.TextBox txt�������� 
      Height          =   270
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   675
      Width           =   2175
   End
   Begin VB.TextBox txt��� 
      Height          =   270
      Left            =   1200
      TabIndex        =   11
      Top             =   200
      Width           =   2175
   End
   Begin VB.Label lbl��־ 
      Caption         =   "���ñ�־"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lbl������� 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      Height          =   180
      Left            =   3720
      TabIndex        =   10
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lbl������Ŀ 
      AutoSize        =   -1  'True
      Caption         =   "������Ŀ"
      Height          =   180
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Left            =   3720
      TabIndex        =   7
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lbl��λ 
      AutoSize        =   -1  'True
      Caption         =   "�ۼ۵�λ"
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Left            =   3720
      TabIndex        =   5
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label LblӢ������ 
      AutoSize        =   -1  'True
      Caption         =   "Ӣ������"
      Height          =   180
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Lbl�������� 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   720
   End
   Begin VB.Label lbl��� 
      AutoSize        =   -1  'True
      Caption         =   "��    ��"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Frmҽ������_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrCode As String '�������,ҽ����ĿDetailCode
Private mrsDetail As ADODB.Recordset, mrsTMP As ADODB.Recordset
Private mblnOK As Boolean
Private mint���� As Integer
Private mint���� As Integer
Private mintID As Integer

Public Function GetCode(strCode As String, ByVal int���� As Integer, ByVal int���� As Integer) As Boolean
'���ܣ����һ���շ���Ŀ��ҽ������
'������strCode ����Ϊ��������������
'���أ��ɹ�����True
    Dim i As Integer, objItem As ListItem
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, mrsTMP As ADODB.Recordset
    
    mblnOK = False
    mint���� = int����
    
    On Error GoTo ErrH
    
    Set mrsTMP = New ADODB.Recordset
    mrsTMP.CursorLocation = adUseClient
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient

    mint���� = int����
    strSQL = "Select * from ҽ��֧����Ŀ Where ����=[1] And ����=[2] and ��Ŀ����=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ֧������", mint����, mint����, strCode)
    If rsTmp.EOF Then
        gstrSQL = "select A.ID,A.����,decode(A.���,'J','����','5','ҩƷ','6','ҩƷ','7','ҩƷ','����') as ���," & _
                  "A.���� As ��������,'' as Ӣ������, " & _
                  "zlspellcode(A.����) as ����,substrb(A.����,1,40) as ����,substrb(A.���㵥λ,1,20) as ���㵥λ, " & _
                  "B.�ּ�,substrb(substr(A.���,1,instr(A.���,'��')-1),1,20) as ���, " & _
                  "D.���� as ������Ŀ,A.�������� as �������,'δ�걨' as ��־ " & _
                  "from �շ�ϸĿ A,�շѼ�Ŀ B,������Ŀ D " & _
                  "where A.ID=B.�շ�ϸĿID and B.������ĿID=D.ID And " & _
                  "nvl(B.��ֹ����,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
                  "A.����=[1]"
        Else
        gstrSQL = "select A.ID,E.��Ŀ���� as ����,F.���� as ���," & _
                  "A.���� As ��������,'' as Ӣ������, " & _
                  "zlspellcode(A.����) as ����,substrb(A.����,1,40) as ����,substrb(A.���㵥λ,1,20) as ���㵥λ, " & _
                  "B.�ּ�,substrb(substr(A.���,1,instr(A.���,'��')-1),1,20) as ���, " & _
                  "D.���� as ������Ŀ,A.�������� as �������,decode(nvl(E.�Ƿ�ҽ��,0),1,'����','δ����') as ��־ " & _
                  "from �շ�ϸĿ A,�շѼ�Ŀ B,������Ŀ D,ҽ��֧����Ŀ E,����֧������ F " & _
                  "where A.ID=B.�շ�ϸĿID and B.������ĿID=D.ID And " & _
                  "nvl(B.��ֹ����,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
                  "A.����=[1] and A.ID=E.�շ�ϸĿID And E.����=F.���� And E.����ID=F.ID and F.����=[2] And E.����=" & mint����
    End If
    Set mrsTMP = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀѡ��", strCode, mint����)
    If Not mrsTMP.EOF Then
        mrsTMP.MoveFirst
        mintID = mrsTMP!ID
        cmd���.Text = mrsTMP!���
        txt���.Text = mrsTMP!����
        txt��������.Text = mrsTMP!��������
        txt����.Text = mrsTMP!����
        txt����.Text = IIf(IsNull(mrsTMP!����), "", mrsTMP!����)
        txt��λ.Text = IIf(IsNull(mrsTMP!���㵥λ), "", mrsTMP!���㵥λ)
        txt����.Text = mrsTMP!�ּ�
        txt���.Text = IIf(IsNull(mrsTMP!���), "", mrsTMP!���)
        txt������Ŀ.Text = IIf(IsNull(mrsTMP!������Ŀ), "", mrsTMP!������Ŀ)
        cmd�ѱ�.Text = IIf(IsNull(mrsTMP!�������), "", mrsTMP!�������)
        txt��־.Text = mrsTMP!��־
    End If
    
    Frmҽ������_����.Show 1
    '����ֵ
    If mblnOK = True Then
        strCode = mstrCode
    End If
    GetCode = mblnOK
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmd��ѯ_Click()
   Dim StrInput As String, strOutput As String
   Dim strTmpArr As Variant, strArr As Variant
   Dim rsTmp As New ADODB.Recordset
   Dim int����id As Integer
   
    
   
    StrInput = vbTab & g�������_��Ԫ����.��������
    StrInput = StrInput & vbTab & txt���.Text
    
    If ҵ������_��Ԫ����(��ȡ��Ŀ_����, StrInput, strOutput) = False Then Exit Sub
    
    strArr = Split(strOutput, "@$")
    strTmpArr = Split(strArr(0), "||")
    txt��־.Text = strTmpArr(1)

    If cmd�ѱ�.Text <> strTmpArr(4) Or cmd���.Text <> strTmpArr(2) Then
       If MsgBox("����Ŀ��ҽ�����ĵķ�������뱾�صĲ�һ�£��Ƿ����?", vbOKCancel) = vbOK Then
            cmd�ѱ�.Text = strTmpArr(4)
            cmd���.Text = strTmpArr(2)
            
            '���·������
            '$IF HIS9.19
            #If gverControl = 0 Then
                gstrSQL = "ZL_�շ�ϸĿ_UPDATE_����(" & mintID & ",'" & cmd�ѱ�.Text & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���·������")
            #Else
            '$ELSE  HIS+
                gstrSQL = "ZL_�շ���ĿĿ¼_UPDATE_����(" & mintID & ",'" & cmd�ѱ�.Text & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���·������")
            #End If
        End If
    End If
    
    gstrSQL = "select nvl(ID,0) as ID from ����֧������ where ����=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ֧������", mint����, CStr(cmd���.Text))
    int����id = rsTmp!ID
    
    gstrSQL = "ZL_ҽ��֧����Ŀ_Modify(" & mintID & "," & mint���� & "," & mint���� & "," & _
              int����id & ",'" & txt���.Text & "','" & txt����.Text & "','" & Format(zlDatabase.Currentdate, "YYYY-MM-DD") & "'," & IIf(txt��־.Text = "����", 1, 0) & ")"
    ExecuteProcedure_��Ԫ���� "����ҽ��֧����Ŀ"
    
End Sub

Private Sub cmdȡ��_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    '����ѡ����Ŀ����
    
    mstrCode = txt���.Text
    mblnOK = False
    Unload Me
End Sub

Private Sub cmd�걨_Click()
   Dim StrInput As String, strOutput As String
   Dim rsTmp As New ADODB.Recordset
   Dim int����id As Integer
     
    StrInput = vbTab & g�������_��Ԫ����.��������
    StrInput = StrInput & vbTab & txt���.Text & "||"
    StrInput = StrInput & cmd���.Text & "||"
    StrInput = StrInput & txt��������.Text & "||"
    StrInput = StrInput & txtӢ������.Text & "||"
    StrInput = StrInput & txt����.Text & "||"
    StrInput = StrInput & txt����.Text & "||"
    StrInput = StrInput & txt��λ.Text & "||"
    StrInput = StrInput & txt����.Text & "||"
    StrInput = StrInput & txt���.Text & "||"
    StrInput = StrInput & txt������Ŀ.Text & "||"
    StrInput = StrInput & cmd�ѱ�.Text
    
    StrInput = StrInput & vbTab & gstrUserName
    StrInput = StrInput & vbTab & Format(zlDatabase.Currentdate, "YYYY-M-DD")
    
    If ҵ������_��Ԫ����(�걨��Ŀ_����, StrInput, strOutput) = False Then Exit Sub
    
    gstrSQL = "select nvl(ID,0) as ID from ����֧������ where ����=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ֧������", mint����, CStr(cmd���.Text))
    int����id = rsTmp!ID
    
    gstrSQL = "ZL_ҽ��֧����Ŀ_Modify(" & mintID & "," & mint���� & "," & mint���� & "," & _
               int����id & ",'" & txt���.Text & "','" & txt����.Text & "','" & Format(zlDatabase.Currentdate, "YYYY-MM-DD") & "',0)"
    ExecuteProcedure_��Ԫ���� "����ҽ��֧����Ŀ"
    
    MsgBox "����Ŀ�Ѿ��ɹ����䵽ҽ�����ģ���֪ͨҽ��������ˣ�"
End Sub


Private Sub txt���_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer, objItem As ListItem
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, mrsTMP As ADODB.Recordset
    If KeyCode = vbKeyReturn Then
   
        mblnOK = False
        
        Set mrsTMP = New ADODB.Recordset
        mrsTMP.CursorLocation = adUseClient
        Set rsTmp = New ADODB.Recordset
        rsTmp.CursorLocation = adUseClient
    
        strSQL = "Select * from ҽ��֧����Ŀ Where ����=[1] And ����=[2] and ��Ŀ����=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ֧������", mint����, mint����, CStr(txt���.Text))
        If rsTmp.EOF Then
            gstrSQL = "select A.ID,A.����,decode(A.���,'J','����','1','����','5','ҩƷ','6','ҩƷ','7','ҩƷ','����') as ���," & _
                      "A.���� As ��������,'' as Ӣ������, " & _
                      "zlspellcode(A.����) as ����,substrb(A.����,1,40) as ����,substr(A.���㵥λ,1,20) as ���㵥λ, " & _
                      "B.�ּ�,substr(substr(A.���,1,instr(A.���,'��')-1),1,20) as ���, " & _
                      "D.���� as ������Ŀ,A.�������� as ������� ,'δ�걨' as ��־ " & _
                      "from �շ�ϸĿ A,�շѼ�Ŀ B,������Ŀ D " & _
                      "where A.ID=B.�շ�ϸĿID and B.������ĿID=D.ID And " & _
                      "nvl(B.��ֹ����,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
                      "A.����=[1]"
            Else
            gstrSQL = "select A.ID,E.��Ŀ���� as ����,F.���� as ���," & _
                      "A.���� As ��������,'' as Ӣ������, " & _
                      "zlspellcode(A.����) as ����,substrb(A.����,1,40) as ����,substr(A.���㵥λ,1,20) as ���㵥λ, " & _
                      "B.�ּ�,substr(substr(A.���,1,instr(A.���,'��')-1),1,20) as ���, " & _
                      "D.���� as ������Ŀ,A.�������� as �������,decode(nvl(E.�Ƿ�ҽ��,0),1,'����','δ����') as ��־ " & _
                      "from �շ�ϸĿ A,�շѼ�Ŀ B,�շѱ��� C,������Ŀ D,ҽ��֧����Ŀ E,����֧������ F " & _
                      "where A.ID=B.�շ�ϸĿID and B.������ĿID=D.ID And " & _
                      "nvl(B.��ֹ����,to_date('3000-01-01','YYYY-MM-DD'))=to_date('3000-01-01','YYYY-MM-DD') and " & _
                      "A.����=[1] and A.ID=E.�շ�ϸĿID And E.����=F.���� And E.����ID=F.ID and F.����=[2] And E.����=[3]"
        End If
        Set mrsTMP = zlDatabase.OpenSQLRecord(gstrSQL, "", CStr(txt���.Text), mint����, mint����)
            
        If Not mrsTMP.EOF Then
            mrsTMP.MoveFirst
            mintID = mrsTMP!ID
            cmd���.Text = mrsTMP!���
            txt���.Text = mrsTMP!����
            txt��������.Text = mrsTMP!��������
            txt����.Text = mrsTMP!����
            txt����.Text = IIf(IsNull(mrsTMP!����), "", mrsTMP!����)
            txt��λ.Text = IIf(IsNull(mrsTMP!���㵥λ), "", mrsTMP!���㵥λ)
            txt����.Text = mrsTMP!�ּ�
            txt���.Text = IIf(IsNull(mrsTMP!���), "", mrsTMP!���)
            txt������Ŀ.Text = IIf(IsNull(mrsTMP!������Ŀ), "", mrsTMP!������Ŀ)
            cmd�ѱ�.Text = IIf(IsNull(mrsTMP!�������), "", mrsTMP!�������)
            txt��־.Text = mrsTMP!��־
        End If
    End If
End Sub
