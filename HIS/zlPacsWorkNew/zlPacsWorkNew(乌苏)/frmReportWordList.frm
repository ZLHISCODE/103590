VERSION 5.00
Begin VB.Form frmReportWordList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ʾ�ʾ��"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   Icon            =   "frmReportWordList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Caption         =   "������"
      Height          =   735
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   6375
      Begin VB.CommandButton cmdWordTag 
         Caption         =   "�������"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdWordTag 
         Caption         =   "������"
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdWordTag 
         Caption         =   "����"
         Height          =   375
         Index           =   3
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCompound 
         Caption         =   "���"
         Height          =   375
         Left            =   5280
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.TextBox txtWord 
      Height          =   3735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   7935
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   8
      Top             =   510
      Width           =   5265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   5265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6960
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "ȫԺͨ��(&1)"
      Height          =   180
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   900
      Width           =   1305
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "����ͨ��(&2)"
      Height          =   180
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   900
      Width           =   1305
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "����ʹ��(&3)"
      Height          =   180
      Index           =   2
      Left            =   5400
      TabIndex        =   0
      Top             =   900
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.Label lbl�������� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&C)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   570
      Width           =   990
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ʾ�����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Width           =   990
   End
   Begin VB.Label lbl��Χ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ�÷�Χ(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   900
      Width           =   990
   End
End
Attribute VB_Name = "frmReportWordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngClassID As Long     '�ʾ����ID
Private mstrClassName As String '�ʾ��������
Private mlngWordID As Long      '�ʾ�ID
Private mlngDeptID As Long      '����ID
Private mstr��� As String

Public Sub zlShowMe(frmParent As Object, txtWordString As String, intWordPower As Integer, _
        lngClassID As Long, strClassName As String, lngDeptID As Long, _
        Optional ByVal lngWordID As Long)
'������ txtWordString ---�޸Ļ�����ӵĴʾ�����
'       intWordPower --- �޸Ĵʾ��Ȩ�ޣ�0-ȫԺ��1-���ң�2-���ˣ�
'       lngClassID --- �ʾ����ID
'       strClassName --- �ʾ���������
'       lngDeptID --- ����ID
'       lngWordID --- �ʾ��ID���޸Ĵʾ�ʱ��Ҫ�ṩ
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strWordName As String
    
    mlngClassID = lngClassID
    mstrClassName = strClassName
    mlngWordID = lngWordID
    mlngDeptID = lngDeptID
    
    Me.txt����.Text = mstrClassName
    Me.txt����.Tag = mlngClassID
    Me.txt����.MaxLength = 80       '.Fields("����").DefinedSize
    
    If lngWordID = 0 Then
        frmReportWordList.Caption = "����ʾ���ʾ�"
        mstr��� = zlDefaultWordCode(mlngClassID)
        Me.txtWord.Text = txtWordString
        Me.txt����.Text = strWordName
    Else
        frmReportWordList.Caption = "�޸�ʾ���ʾ�"
        
        '�Ӵʾ�ʾ���ж�ȡ�ʾ�����
        strSQL = "Select a.����,a.ͨ�ü�,a.���, b.���д���,b.�����ı� " & _
                 " From �����ʾ�ʾ�� a,�����ʾ���� b Where a.Id=[1] And a.Id=b.�ʾ�ID  order by ���д��� "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngWordID)
        Me.txtWord.Text = ""
        If Not rsTemp.EOF Then
            mstr��� = Nvl(rsTemp!���)
            Me.opt��Χ(Nvl(rsTemp!ͨ�ü�, 0)).value = True
            Me.txt����.Text = Nvl(rsTemp!����)
        End If
        While rsTemp.EOF = False
            Me.txtWord.Text = Me.txtWord.Text & Nvl(rsTemp!�����ı�)
            rsTemp.MoveNext
        Wend
    End If
    
    Select Case intWordPower
    Case 2: Me.opt��Χ(0).Enabled = False: Me.opt��Χ(1).Enabled = False
    Case 1: Me.opt��Χ(0).Enabled = False
    End Select
    
    frmReportWordList.Show 1, frmParent
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim ArraySQL() As String
    Dim lngCount As Long
    Dim i As Integer
    Dim strText As String
    Dim blnAdd As Boolean   'True-�����ʾ�ʾ����False-�޸Ĵʾ�ʾ��
    
    '����������ݵĺϷ���
    If Trim(Me.txt����.Text) = "" Then
        MsgBoxD Me, "���������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBoxD Me, "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtWord.Text)) = 0 Then
        MsgBoxD Me, "������ʾ�ʾ�����ݺ��ٱ��档", vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If mlngWordID = 0 Then      '����ʾ���ʾ�
        mlngWordID = zlDatabase.GetNextId("�����ʾ�ʾ��")
        blnAdd = True
    Else                        '�޸�ʾ���ʾ�
        blnAdd = False
    End If
    
    '����ʾ�ʾ������
    strSQL = mlngWordID & "," & Val(Me.txt����.Tag) & ",'" & mstr��� & "','" & Trim(Me.txt����.Text) & "'"
    If Me.opt��Χ(0).value Then
        strSQL = strSQL & ",0"
    ElseIf Me.opt��Χ(1).value Then
        strSQL = strSQL & ",1"
    Else
        strSQL = strSQL & ",2"
    End If
    strSQL = strSQL & "," & mlngDeptID & "," & UserInfo.ID
    strSQL = "Zl_�����ʾ�ʾ��_Edit(" & IIf(blnAdd = True, 1, 2) & "," & strSQL & ")"
    
    '�����ʾ���ɣ����ϱ�ǣ�#***#
    If InStr(txtWord.Text, "<<") > 0 Then
        '��һ��<<ǰ�����ݺ��Բ���
        strText = Mid(txtWord.Text, InStr(txtWord.Text, "<<"))
        
        If InStr(strText, "<<����>>") > 0 Then strText = Replace(strText, "<<����>>", "#***#<<����>>")
        If InStr(strText, "<<���>>") > 0 Then strText = Replace(strText, "<<���>>", "#***#<<���>>")
        If InStr(strText, "<<����>>") > 0 Then strText = Replace(strText, "<<����>>", "#***#<<����>>")
        If Mid(strText, 1, 5) = "#***#" Then strText = Mid(strText, 6)
    Else
        strText = "#***#" & txtWord.Text
    End If
    
    '��ȡSQL�������
    ReDim ArraySQL(1 To 2) As String
    ArraySQL(1) = strSQL
    
    'ǰ�ڴ���
    ArraySQL(2) = "Zl_�����ʾ����_Beforesave(" & mlngWordID & ")"
    
    '��ȡ����SQL����
    Call GetSaveSQL(ArraySQL, strText)
    
    '���ڴ���
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    strSQL = "Zl_�����ʾ����_Aftersave(" & mlngWordID & ")"
    ArraySQL(lngCount) = strSQL
    
    'ִ�б������
    err = 0: On Error GoTo errHand
    gcnOracle.BeginTrans
    For i = 1 To UBound(ArraySQL)
        strSQL = ArraySQL(i)
        Call zlDatabase.ExecuteProcedure(strSQL, "frmReportWordList")
    Next
    gcnOracle.CommitTrans
        
    Unload Me
    Exit Sub
errHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetSaveSQL(ByRef ArrySQL() As String, strText As String)
'��֯����ʾ���ɵ�SQL���
'������ ArrySQL --- SQL �������
'       strText --- Ҫ����Ĵʾ�ʾ������
    
    Dim strLine As String       'һ���ı����س�֮����ı�
    Dim lng��� As Long         '����CRLF���ֶ�
    Dim i As Integer
    On Error GoTo err
    
    lng��� = 1
    
    For i = 0 To UBound(Split(strText, "#***#"))
        strLine = Split(strText, "#***#")(i)
        Call GetPlainTextSaveSQL(ArrySQL, strLine, lng���)
    Next
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetPlainTextSaveSQL(ByRef ArraySQL() As String, ByVal strIn As String, ByRef lng��� As Long) As Boolean
'�Դ��ı���ȡ���䱣�浽�ʾ���ɵ�SQL��䣬���ȴ���4000���ַ��������д洢����ŵ���֮��
'������ ArraySQL --- SQL �������
'       strIn --- ��Ҫ������ı�
'       lng��� --- ���
    
    Dim lngLen As Long, strSub As String, i As Long, lngID As Long
    Dim lngCount As Long, lID As Long
    strIn = Replace(strIn, "'", "' || chr(39) || '")
    strIn = Replace(strIn, vbCrLf, "' || chr(13) || chr(10) || '")  '����strIn�ǲ�������vbCrlf�ġ�
    lngLen = Len(strIn)
    
    '����4000Ϊ��ֶδ洢��
    i = 0
    Do While (i * 2000 + 1 <= lngLen)
        lngCount = UBound(ArraySQL) + 1
        ReDim Preserve ArraySQL(1 To lngCount) As String

        strSub = Mid(strIn, i * 2000 + 1, 2000)

        gstrSQL = "Zl_�����ʾ����_Insert(" & mlngWordID & "," & lng��� & ",0,'" & strSub & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"
        
        ArraySQL(lngCount) = gstrSQL
       
        lng��� = lng��� + 1
        i = i + 1
    Loop
    GetPlainTextSaveSQL = True
End Function

Private Sub cmdWordTag_Click(Index As Integer)
'���롰����������������������������顱
    Dim strTag As String
    Dim strTemp
    
    On Error GoTo err
    Select Case Index
    Case 1
        strTag = "<<����>>"
    Case 2
        strTag = "<<���>>"
    Case 3
        strTag = "<<����>>"
    End Select
    
    txtWord.Text = Left(txtWord.Text, txtWord.SelStart) & strTag & Mid(txtWord.Text, txtWord.SelStart + 1)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

