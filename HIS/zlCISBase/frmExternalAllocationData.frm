VERSION 5.00
Begin VB.Form frmExternalAllocationData 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����Դ��ȡ�༭"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmd��֤ 
      Caption         =   "��֤(&V)"
      Height          =   350
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   5
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   4
      Top             =   3960
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "3.[����ID]��ֵΪ����������б��ֵ��"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   705
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2.[����ID]���ﲡ�˴�ֵΪ����ID,סԺ��ֵΪ��ҳID��"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   500
      Width           =   4410
   End
   Begin VB.Label lblTip 
      Caption         =   "1.SQL�еĲ�����ʽΪ�̶���[������]��������������̶�Ԥ�Ƶ�[����ID],[����ID],[����ID],[ҽ��ID]��"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmExternalAllocationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrReturn As String

Public Function ShowMe(ByVal frmMain As Form, ByVal strSQLText As String) As String
    mstrReturn = strSQLText
    Me.Show 1, frmMain
    
    ShowMe = mstrReturn
End Function

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    '���ܣ����ؽ����������

    mstrReturn = txtEdit.Text
    
    Unload Me
End Sub

Private Function TrueObject(ByVal strObject As String) As String
    '���ܣ�SQLObject�������Ӻ���,����ȥ���������е������ַ�
    Dim i As Integer
    'Ѱ�ҵ�һ�������ַ�λ��
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    'Ѱ�Һ����һ���������ַ�
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function

Private Function TransSpecialChar(ByRef strSql As String, Optional ByVal blnRestore As Boolean = False) As Boolean
    '���ܣ�ת��SQL�е������ַ����磺[]�ַ�������������ķ��ų�ͻ
    '���أ�True�ɹ���Falseʧ��

    Const STR_ORIGINAL As String = "[|]|(|)"
    Const STR_TRANS As String = "<��������>|<��������>|<������>|<������>"

    Dim strResult As String, strTmp As String
    Dim arrOriginal As Variant, arrTrans As Variant
    Dim i As Long, j As Long, lngBegin As Long
    Dim intLen As Integer
    
    If Trim(strSql = "") Then Exit Function
    
    On Error GoTo hErr
    
    strResult = strSql
    If blnRestore Then
        '��ԭ
        arrOriginal = Split(STR_TRANS, "|")
        arrTrans = Split(STR_ORIGINAL, "|")
    Else
        'ת��
        arrOriginal = Split(STR_ORIGINAL, "|")
        arrTrans = Split(STR_TRANS, "|")
    End If
    
    '���SQL�ַ����Ƿ����[]�ַ�
    i = 1
    lngBegin = 0
    Do While Mid(strResult, i) Like "*'*"
        If Mid(strResult, i, 1) = "'" Then
            If lngBegin <= 0 Then
                '��ʼ
                lngBegin = i
            Else
                '����
                lngBegin = 0
            End If
        Else
            If lngBegin > 0 Then
                '����''�ַ��ڲ����������ַ�������SQL�����ַ���
                strTmp = Mid(strResult, lngBegin + 1)
                If InStr(strTmp, "'") > 0 Then
                    strTmp = Left(strTmp, InStr(strTmp, "'") - 1)
                    strTmp = Replace(strTmp, arrTrans(0), arrOriginal(0))
                Else
                    strTmp = ""
                End If
                
                If Not (strTmp Like "*[[][0-9][]]*" Or strTmp Like "*[[][0-9][0-9][]]*") Then
                    For j = LBound(arrOriginal) To UBound(arrOriginal)
                        intLen = Len(arrOriginal(j))
                        If Mid(strResult, i, intLen) = arrOriginal(j) Then
                            strResult = Left(strResult, i - 1) & arrTrans(j) & Mid(strResult, i + intLen)
                        End If
                    Next
                End If
            End If
        End If
        
        i = i + 1
    Loop
    
    strSql = strResult
    TransSpecialChar = True
    Exit Function
    
hErr:
End Function

Private Function GetWithAsTables(ByVal strSql As String) As String
    '���ܣ���ȡWith as ֮��ı��������Զ��ŷָ�
    Dim lngL As Long, lngR As Long, lngS As Long, strTabs As String
    Dim strTmp As String, blnFirst As Boolean
        
    strSql = Replace(strSql, vbCrLf, " ")
    strSql = Replace(strSql, vbTab, " ")
    strSql = Replace(strSql, "  ", " ")
    strSql = Replace(strSql, "  ", " ")
    strSql = Replace(strSql, "AS (", "AS(")
    
    lngL = InStr(1, strSql, "WITH")
    If lngL = 0 Then
        Exit Function
    Else
        lngL = lngL + 4
        blnFirst = True
    End If
        
    Do
        lngR = InStr(lngL, strSql, " AS(")
        If lngR = 0 Then
            Exit Do
        Else
            If Not blnFirst Then
                lngL = InStrRev(strSql, ",", lngR) + 1
            End If
            
            strTmp = Trim(Mid(strSql, lngL, lngR - lngL))
            '11G R2֧�֣����磺with T��column alias 1,column alias 2,......��
            lngS = InStr(strTmp, "(")
            If lngS > 1 Then
                strTmp = Mid(strTmp, 1, strTmp - 1)
            End If
            
            strTabs = strTabs & "," & strTmp
        End If
        
        blnFirst = False
        lngL = lngR + Len(" AS(")
    Loop
    GetWithAsTables = Mid(strTabs, 2)
End Function

Private Function TrimChar(Str As String) As String
    '����:ȥ���ַ����������Ŀո�ͻس�(����ͷ�Ŀո�,�س�),��ȥ��TAB�ַ�,������������
    Dim strTmp As String
    
    If Trim(Str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(Str)
    
    strTmp = Replace(strTmp, "  ", " ")
    strTmp = Replace(strTmp, "  ", " ")

    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Private Function SQLObject(ByVal strSql As String, Optional ByVal strWithas As String) As String
'���ܣ�����SQL������õ��Ķ�����
'������strSQL=Ҫ������ԭʼSQL���
'���أ�SQL��������ʵ��Ķ�����,��"���ű�,���˷��ü�¼,ZLHIS.��Ա��"
'˵����1.��Oracle SELECT������
'      2.���SQL����еĶ�����ǰ����������ǰ׺,���ǰ׺���ᱻ��ȡ
'      3.��Ҫ����TrimChar;TrueObject��֧��
    Dim intB As Long, intE As Long, intL As Long, intR As Long
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Long, j As Long
    Dim lngTmp As Long
    Dim strTmp As String, strObjectSub As String
    
    On Error GoTo errH
    
    '��д����ȥ��������ַ�
    strAnal = UCase(TrimChar(strSql))
    If strWithas = "" Then
        strWithas = GetWithAsTables(strAnal)
    End If
    
    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    If TransSpecialChar(strAnal) = False Then Exit Function
    
    '�ȷֽ⴦��Ƕ���Ӳ�ѯ
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB 'ƥ�����������λ��
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                strTmp = Mid(strAnal, 1, intB - 1)
                lngTmp = 0
                If InStrRev(strTmp, " TABLE") > 0 Or InStrRev(strTmp, " TABLE ") > 0 Then
                    lngTmp = IIf(InStrRev(strTmp, " TABLE ") > 0, InStrRev(strTmp, " TABLE "), InStrRev(strTmp, " TABLE"))
                    strTmp = Mid(strTmp, lngTmp + 6)
                    strTmp = Trim(strTmp)
                End If
                If intE - intB - 1 <= 0 Then
                    '���ڷ��Ӳ�ѯ,�����Ż�����������,��ʹѭ������
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '�Ӳ�ѯ���
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '�����Ӳ�ѯ������ΪΪ���������
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "Ƕ�ײ�ѯ")
                    '�ݹ����
                    strObjectSub = SQLObject(strSub, strWithas)
                    If InStr(strObject & "," & strWithas & ",", "," & strObjectSub & ",") = 0 Then
                        strObject = strObject & "," & strObjectSub
                    End If
                ElseIf strTmp = "" And lngTmp <> 0 Then
                    'ȥ��Table��̬�ڴ��
                    strAnal = Replace(strAnal, Mid(strAnal, lngTmp + 1, intE - lngTmp + 1 + 1), "��̬�ڴ��")
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '��ƥ��������
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '�ֽ����(��ʱstrAnalΪ�򵥲�ѯ,���ܴ�Union������)
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '�ӵ�һ��From���沿�ݿ�ʼ
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "START WITH") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "START WITH") - 1)
        ElseIf InStr(strCur, "CONNECT BY") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "CONNECT BY") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        ElseIf InStr(strCur, "MINUS") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "MINUS") - 1)
        ElseIf InStr(strCur, "INTERSECT") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "INTERSECT") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject & "," & strWithas & ",", "," & strTrue & ",") = 0 And strTrue <> "Ƕ�ײ�ѯ" And strTrue <> "��̬�ڴ��" Then
                If InStr(strTrue, "'") = 0 And InStr(strTrue, "@") = 0 Then
                    strObject = strObject & "," & strTrue
                End If
            End If
        Next
    Next
    '���
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    err.Clear
End Function

Private Sub cmd��֤_Click()
    '���ܣ�����SQL����ȷ��
    Dim strSql As String
    Dim strObject As String
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    strSql = txtEdit.Text
    
    'SQL���
    '-----------------------------------------------
    strObject = SQLObject(strSql)
    If strObject = "" And InStr(UCase(strSql), "TABLE") = 0 And InStr(UCase(strSql), "@") = 0 Then
        MsgBox "���ܷ���SQL�������ѯ�����ݶ���,�����Ƿ���ȷ��д��", vbInformation, App.Title
        Exit Sub
    End If
    '-----------------------------------------------
    
    'SQLִ��
    '-----------------------------------------------
    strSql = UCase(strSql)
    
    strSql = Replace(strSql, "[����ID]", "0")
    strSql = Replace(strSql, "[����ID]", "0")
    strSql = Replace(strSql, "[����ID]", "0")
    strSql = Replace(strSql, "[ҽ��ID]", "0")
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    '-----------------------------------------------
    
    cmdȷ��.Enabled = True
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    txtEdit.Text = mstrReturn
    cmdȷ��.Enabled = False
End Sub

Private Sub txtEdit_Change()
    cmdȷ��.Enabled = False
End Sub
