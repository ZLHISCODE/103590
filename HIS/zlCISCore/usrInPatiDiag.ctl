VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl usrInPatiDiag 
   BackColor       =   &H8000000E&
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   LockControls    =   -1  'True
   ScaleHeight     =   1530
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdCons 
      Caption         =   "�ο���"
      Height          =   300
      Index           =   1
      Left            =   3090
      TabIndex        =   8
      Top             =   2505
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdCons 
      Caption         =   "�ο���"
      Height          =   300
      Index           =   0
      Left            =   735
      TabIndex        =   7
      Top             =   2505
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkDiff 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   1
      Left            =   3000
      TabIndex        =   6
      Top             =   1995
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CheckBox chkDiff 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   690
      TabIndex        =   5
      Top             =   1965
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtDiag 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   2715
      TabIndex        =   4
      Tag             =   "100"
      Text            =   "��ҽ���"
      Top             =   1230
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtDiag 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   735
      TabIndex        =   3
      Tag             =   "100"
      Text            =   "��ҽ���"
      Top             =   1305
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CheckBox chkWH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��ҽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2865
      TabIndex        =   2
      Top             =   915
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   225
      Index           =   0
      Left            =   810
      TabIndex        =   1
      Top             =   1650
      Width           =   350
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   225
      Index           =   1
      Left            =   3000
      TabIndex        =   0
      Top             =   1515
      Width           =   350
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfW 
      Height          =   1530
      Left            =   0
      TabIndex        =   9
      Top             =   315
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   2699
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   -2147483628
      BackColorSel    =   -2147483634
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483639
      GridColorFixed  =   16777215
      AllowBigSelection=   0   'False
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfH 
      Height          =   1530
      Left            =   15
      TabIndex        =   10
      Top             =   465
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   2699
      _Version        =   393216
      Rows            =   10
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   -2147483628
      BackColorSel    =   -2147483634
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483639
      GridColorFixed  =   16777215
      AllowBigSelection=   0   'False
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
End
Attribute VB_Name = "usrInPatiDiag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const STR_COMPART = "|';"
Private Const LAWLChar = "';`|,"""
Private mblnMode As Boolean 'Ϊ���Ǳ�ʾ���û����еı༭����ʱ�Ÿ�ֵ
Private rsTmp As New ADODB.Recordset
Private strSQL As String
Private mlng����id As Long

Private i As Long, j As Long
Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String
 
Private mblnLoaded As Boolean

Private Enum EnmDiag��ҽ
    x��� = 0
    x���� = 1
    x��� = 2
    x���� = 3
    x�ο� = 4
    x����ID = 5
End Enum

Private Enum EnmDiag��ҽ
    z��� = 0
    z���� = 1
    z��� = 2
    z���� = 3
    z�ο� = 4
    z֤ID = 5
    z����ID = 6
End Enum

Private mWestDiag As Boolean '��ҽ���

Private Sub ShowDiag(ByVal lng����ID As Long, ByVal blnEditMode As Boolean)
'ͳһ���ã���������
Dim rsTemp As New ADODB.Recordset

    mlng����id = lng����ID
    mDispMode = Not blnEditMode
    
    '���߼�Ӧ�ȳ�ʼ�ؼ�
    InitMe
    
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Sub
    
    '��鲡�������ǲ�����ȷ
    'strSQL = _
    '    "SELECT a.ID,B.����ID,B.��ҳID,c.����,c.�Ա�,c.����,d.���� ���ڿ���" & vbCrLf & _
    '    "  FROM ���˲������� a,���˲�����¼ b,������Ϣ c,���ű� d" & vbCrLf & _
    '    " WHERE a.Ԫ������ = 4 AND a.������¼id=b.Id AND b.����id =c.����id AND b.����id=d.Id AND" & vbCrLf & _
    '    "      a.Ԫ�ر��� IN" & vbCrLf & _
    '    "      (SELECT ����" & vbCrLf & _
    '    "         FROM ����Ԫ��Ŀ¼" & vbCrLf & _
    '    "        WHERE ���� = 4 AND ���� = '��Ժ��ϼ�¼��')" & vbCrLf & _
    '    " AND A.id=" & mlng����id
        
    strSQL = _
        "SELECT a.ID" & vbCrLf & _
        "  FROM ���˲������� a" & vbCrLf & _
        " WHERE a.Ԫ������ = 4 and " & vbCrLf & _
        "      a.Ԫ�ر��� IN" & vbCrLf & _
        "      (SELECT ����" & vbCrLf & _
        "         FROM ����Ԫ��Ŀ¼" & vbCrLf & _
        "        WHERE ���� = 4 AND ���� = '��Ժ��ϼ�¼��')" & vbCrLf & _
        " AND A.id=" & mlng����id
    If rsTemp.State = 1 Then rsTmp.Close
    Set rsTemp = Nothing
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "��Ժ��ϼ�¼��")
    If rsTemp.RecordCount < 1 Then
        SetErr -1, "�ò����������޷�������Ժ��ϼ�¼����"
        Exit Sub
    End If
    
    '��������
    ReadData
End Sub

Private Sub ReadData()
    Dim astrDiags() As String
'���������������
On Error GoTo ErrHandle

    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Sub
    
    msfW.Clear
    msfW.Rows = 2
    ReSetRowCode msfW
    SetSelColor msfW, 1
    msfW.Row = 1: msfW.Col = 2
    msfW_EnterCell
    
    msfH.Clear
    msfH.Rows = 2
    ReSetRowCode msfH
    SetSelColor msfH, 1
    msfH.Row = 1: msfH.Col = 2
    msfH_EnterCell
    If mDispMode = False Then
        If mWestDiag = True Then
            txtDiag(0).Visible = True
            chkDiff(0).Visible = True
            cmdCons(0).Visible = True
            cmdSel(0).Visible = True
            
            txtDiag(1).Visible = False
            chkDiff(1).Visible = False
            chkWH.Visible = False
            cmdCons(1).Visible = False
            cmdSel(1).Visible = False
        Else
            txtDiag(0).Visible = False
            chkDiff(0).Visible = False
            cmdCons(0).Visible = False
            cmdSel(0).Visible = False
            
            txtDiag(1).Visible = True
            chkDiff(1).Visible = True
            chkWH.Visible = True
            cmdCons(1).Visible = True
            cmdSel(1).Visible = True
        End If
    Else
            txtDiag(0).Visible = False
            chkDiff(0).Visible = False
            cmdCons(0).Visible = False
            cmdSel(0).Visible = False
            
            txtDiag(1).Visible = False
            chkDiff(1).Visible = False
            chkWH.Visible = False
            cmdCons(1).Visible = False
            cmdSel(1).Visible = False
    End If
    
    strSQL = " Select * from ������ϼ�¼ WHERE ������� in (2,12)  AND  ����ID=" & mlng����id & " ORDER BY ID"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "��Ժ��ϼ�¼��")
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        If mWestDiag Then
            rsTmp.Filter = "�������=2"
            If rsTmp.RecordCount > 0 Then
                msfW.Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    msfW.TextMatrix(i, EnmDiag��ҽ.x���) = CStr(i) & "��"
                    msfW.TextMatrix(i, EnmDiag��ҽ.x����) = "��ϣ�"
                    msfW.TextMatrix(i, EnmDiag��ҽ.x���) = zlCommFun.Nvl(rsTmp!�������)
                    msfW.TextMatrix(i, EnmDiag��ҽ.x����) = IIf(zlCommFun.Nvl(rsTmp!�Ƿ�����, 0) = 0, "", "��")
                    msfW.TextMatrix(i, EnmDiag��ҽ.x����ID) = IIf(zlCommFun.Nvl(rsTmp!����id, "") = "0", "", zlCommFun.Nvl(rsTmp!����id, ""))
                    msfW.RowData(i) = zlCommFun.Nvl(rsTmp!���ID, 0)
                    rsTmp.MoveNext
                Next
                msfW.Row = 1: msfW.Col = 2
                msfW_EnterCell
            End If
        Else
            If rsTmp.RecordCount > 0 Then
                msfH.Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    msfH.TextMatrix(i, EnmDiag��ҽ.z���) = CStr(i) & "��"
                    msfH.TextMatrix(i, EnmDiag��ҽ.z����) = IIf(rsTmp!������� = 2, "��ҽ", "��ҽ")
                    msfH.TextMatrix(i, EnmDiag��ҽ.z���) = zlCommFun.Nvl(rsTmp!�������)
                    msfH.TextMatrix(i, EnmDiag��ҽ.z����) = IIf(zlCommFun.Nvl(rsTmp!�Ƿ�����, 0) = 0, "", "��")
                    msfH.TextMatrix(i, EnmDiag��ҽ.z֤ID) = zlCommFun.Nvl(rsTmp!֤��ID)
                    msfH.TextMatrix(i, EnmDiag��ҽ.z����ID) = IIf(zlCommFun.Nvl(rsTmp!����id, "") = "0", "", zlCommFun.Nvl(rsTmp!����id, ""))
                    msfH.RowData(i) = zlCommFun.Nvl(rsTmp!���ID, 0)
                    rsTmp.MoveNext
                Next
                msfH.Row = 1: msfH.Col = 2
                msfH_EnterCell
            End If
        End If
    Else
        strSQL = "Select ���� From ���˲����ı��� Where ����ID=" & mlng����id
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "��Ժ��ϼ�¼��")
        If rsTmp.EOF Then Exit Sub
        astrDiags = Split(Nvl(rsTmp(0)), Chr(13) + Chr(10))
        msfW.Rows = UBound(astrDiags, 1) + 2
        For i = 1 To UBound(astrDiags, 1) + 1
            msfW.TextMatrix(i, EnmDiag��ҽ.x���) = CStr(i) & "��"
            msfW.TextMatrix(i, EnmDiag��ҽ.x����) = "��ϣ�"
            msfW.TextMatrix(i, EnmDiag��ҽ.x���) = IIf(astrDiags(i - 1) Like "#��*", Mid(astrDiags(i - 1), 3), astrDiags(i - 1))
            msfW.TextMatrix(i, EnmDiag��ҽ.x����) = ""
            msfW.TextMatrix(i, EnmDiag��ҽ.x����ID) = ""
            msfW.RowData(i) = 0
        Next
    End If
    Exit Sub
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function SaveData(lng����ID As Long, lng��ҳID As Long, lng����ID As Long, ReturnStrSQL As String, strError As String) As Boolean
'�ⲿ���ù����沢���ش����ַ����������
On Error GoTo ErrHandle
Dim strErr As String
Dim strZID As String '�õ�֤ID
Dim strIllnID As String '�õ�����ID

Dim strTmp As String
Dim lngRow As Long

    If mDispMode Then
        strError = "��ǰΪ��ʾģʽ���ܱ������ݣ�"
        SetErr -1, "��ǰΪ��ʾģʽ���ܱ������ݣ�"
        Exit Function
    End If
    
    If mWestDiag Then
        lngRow = 1
        Do While lngRow < msfW.Rows
            For j = 1 To Len(LAWLChar)
                If InStr(msfW.TextMatrix(lngRow, EnmDiag��ҽ.x���), Mid(LAWLChar, j, 1)) > 0 Then
                    strError = "����д��ڷǷ��ַ���"
                    SetErr -1, "����д��ڷǷ��ַ���"
                    msfW.Row = lngRow
                    msfW_EnterCell
                    Exit Function
                End If
            Next
            If Trim(msfW.TextMatrix(lngRow, EnmDiag��ҽ.x���)) = "" Then
                If lngRow = 1 Then
'                    strError = "��һ��������ݲ���Ϊ�գ�"
'                    SetErr -1, "��һ��������ݲ���Ϊ�գ�"
'                    msfW.Row = lngRow
'                    msfW_EnterCell
'                    Exit Function
                    lngRow = lngRow + 1
                Else
                    '����ɾ��
                    msfW.RemoveItem lngRow
                    msfW_EnterCell
                    ReSetRowCode msfW
                End If
            Else
                lngRow = lngRow + 1
            End If
        Loop
    Else
        lngRow = 1
        Do While lngRow < msfH.Rows
            For j = 1 To Len(LAWLChar)
                If InStr(msfH.TextMatrix(lngRow, EnmDiag��ҽ.z���), Mid(LAWLChar, j, 1)) > 0 Then
                    strError = "����д��ڷǷ��ַ���"
                    SetErr -1, "����д��ڷǷ��ַ���"
                    msfH.Row = lngRow
                    msfH_EnterCell
                    Exit Function
                End If
            Next
            If Trim(msfH.TextMatrix(lngRow, EnmDiag��ҽ.z���)) = "" Then
                If lngRow = 1 Then
'                    strError = "��һ��������ݲ���Ϊ�գ�"
'                    SetErr -1, "��һ��������ݲ���Ϊ�գ�"
'                    msfH.Row = lngRow
'                    msfH_EnterCell
'                    Exit Function
                    lngRow = lngRow + 1
                Else
                    '����ɾ��
                    msfH.RemoveItem lngRow
                    msfH_EnterCell
                    ReSetRowCode msfH
                End If
            Else
                lngRow = lngRow + 1
            End If
        Loop
    End If
    
    '�������'���ID'����ID'֤��ID'�Ƿ�����'�������;�������'���ID'����ID'֤��ID'�Ƿ�����'�������;
    strSQL = ""
    If mWestDiag Then
        For i = 1 To msfW.Rows - 1
            '�õ�����ID
            If IsNumeric(msfW.TextMatrix(i, EnmDiag��ҽ.x����ID)) Then
                strIllnID = CLng(msfW.TextMatrix(i, EnmDiag��ҽ.x����ID))
            Else
                strIllnID = "0"
            End If
                    '�������'���ID'����ID'֤��ID'�Ƿ�����'�������;
            strSQL = strSQL & "2''" & msfW.RowData(i) & "''" & strIllnID & "''0''" & IIf(Trim(msfW.TextMatrix(i, EnmDiag��ҽ.x����)) = "��", 1, 0) & "''" & msfW.TextMatrix(i, EnmDiag��ҽ.x���) & ";"
        Next
    Else
        For i = 1 To msfH.Rows - 1
            '�õ��������
            strTmp = IIf(msfH.TextMatrix(i, EnmDiag��ҽ.z����) = "��ҽ", "12", "2")
            '�õ�����ID
            If IsNumeric(msfH.TextMatrix(i, EnmDiag��ҽ.z����ID)) Then
                strIllnID = CLng(msfH.TextMatrix(i, EnmDiag��ҽ.z����ID))
            Else
                strIllnID = "0"
            End If
            '�õ�֤ID
            If IsNumeric(msfH.TextMatrix(i, EnmDiag��ҽ.z֤ID)) Then
                strZID = CLng(msfH.TextMatrix(i, EnmDiag��ҽ.z֤ID))
            Else
                strZID = "0"
            End If
                    '�������'���ID'����ID'֤��ID'�Ƿ�����'�������;
            strSQL = strSQL & strTmp & "''" & msfH.RowData(i) & "''" & strIllnID & "''" & strZID & "''" & IIf(Trim(msfH.TextMatrix(i, EnmDiag��ҽ.z����)) = "��", 1, 0) & "''" & msfH.TextMatrix(i, EnmDiag��ҽ.z���) & ";"
        Next
    End If
    '��Ժ�����ﹲ��һ������
    ReturnStrSQL = "ZL_����������ϼ�¼��_INSERT(" & _
                IIf(lng����ID < 1, "NULL", lng����ID) & "," & _
                IIf(lng��ҳID < 1, "NULL", lng��ҳID) & "," & _
                lng����ID & ",'" & _
                strSQL & "','" & _
                UserInfo.���� & "')"
    
    SaveData = True
    Exit Function
ErrHandle:
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State <> adStateOpen Then Exit Function
    strError = Err.Description
    Call SaveErrLog
End Function

Private Function LocalCheck�Ƿ�Ƿ�(txt As Control, ByVal strLawlChar As String) As Boolean
'����:����ǲ��ǰ���strLawlChar����ַ���,����оͷ���Ϊ�����ͷ��ط�
On Error GoTo ErrHandle
    Dim strSour As String
    
    If TypeOf txt Is TextBox Or TypeOf txt Is ComboBox Then
        If TypeOf txt Is ComboBox Then
            If txt.Style <> 0 Then
                '����ComboBoxΪѡ��������ֻ����������
                LocalCheck�Ƿ�Ƿ� = True
                Exit Function
            End If
        End If
        strSour = txt.Text
        If Len(strSour) > 0 Then
            For i = 1 To Len(strSour)   ' Len(strLawlChar)
                If InStr(strLawlChar, Mid(strSour, i, 1)) > 0 Then
                    txt.SelStart = i - 1
                    txt.SelLength = 1
                    MsgBox "�ı�������зǷ��ַ���", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                    Exit Function
                End If
            Next
            If VarType(txt.Tag) = vbLong Or VarType(txt.Tag) = vbInteger Then
                If zlCommFun.ActualLen(strSour) > txt.Tag And txt.Tag > 0 Then
                    MsgBox "����������ı�������", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                End If
            ElseIf VarType(txt.Tag) = vbString And IsNumeric(txt.Tag) Then
                If zlCommFun.ActualLen(strSour) > CLng(txt.Tag) And CLng(txt.Tag) > 0 Then
                    MsgBox "����������ı�������", vbInformation, gstrSysName
                    LocalCheck�Ƿ�Ƿ� = True
                End If
            End If
        End If
    End If
    Exit Function
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Function
    If gcnOracle Is Nothing Then Exit Function
    If gcnOracle.State <> adStateOpen Then Exit Function
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetSelColor(objMsf As MSHFlexGrid, ByVal lngRow As Long, Optional ByVal oleForeColor As OLE_COLOR = 0, Optional ByVal oleBackColor As OLE_COLOR = &HFFFFFF)
'����ѡ���е���ɫ
Dim lngSelCol As Long, lngSelRow As Long

objMsf.Redraw = False
lngSelCol = objMsf.Col
lngSelRow = objMsf.Row

For i = 1 To objMsf.Rows - 1
    objMsf.Row = i
    If i = lngRow Then
        For j = 0 To objMsf.Cols - 1
            objMsf.Col = j
            objMsf.CellFontBold = True
            objMsf.CellForeColor = oleForeColor
            objMsf.CellBackColor = oleBackColor
        Next
    Else
        For j = 0 To objMsf.Cols - 1
            objMsf.Col = j
            objMsf.CellFontBold = False
            objMsf.CellForeColor = 0
            objMsf.CellBackColor = RGB(255, 255, 255)
        Next
    End If
Next
objMsf.Col = lngSelCol
objMsf.Row = lngSelRow
objMsf.Refresh
objMsf.Redraw = True
End Sub

Private Sub SetMSFFormat(objMsf As MSHFlexGrid, ByVal strFormat As String, Optional blnReCreatCol As Boolean = True)
'����:���ñ��ĸ�ʽ
'strFormat��ʽ:    ����1,���1,���뷽ʽ1,�Ƿ���ʾ����1;����2,���2,���뷽ʽ2,�Ƿ���ʾ����2;����3,���3,���뷽ʽ3,�Ƿ���ʾ����3;....
Dim arrStrTmp() As String
Dim strTmp As String



arrStrTmp = Split(strFormat, ";")
If UBound(arrStrTmp) + 1 <= objMsf.Cols Then
    For i = 0 To UBound(arrStrTmp)
        'ȷ���Ƿ���ʾ����
        If IsNumeric(Split(arrStrTmp(i), ",")(3)) Then
            If CLng(Split(arrStrTmp(i), ",")(3)) > 0 Then
                '��ʾ
                objMsf.TextMatrix(0, i) = Split(arrStrTmp(i), ",")(0)
            Else
                '����ʾ
                objMsf.TextMatrix(0, i) = ""
            End If
        Else
            '��ʾ
            objMsf.TextMatrix(0, i) = Split(arrStrTmp(i), ",")(0)
        End If
        
        'ȷ���п�
        If IsNumeric(Split(arrStrTmp(i), ",")(1)) Then
            If CLng(Split(arrStrTmp(i), ",")(1)) >= 0 Then
                objMsf.ColWidth(i) = CLng(Split(arrStrTmp(i), ",")(1))
            Else
                objMsf.ColWidth(i) = 1440
            End If
        Else
            objMsf.ColWidth(i) = 1440
        End If
        
        'ȷ�����뷽ʽ
        If IsNumeric(Split(arrStrTmp(i), ",")(2)) Then
            If CLng(Split(arrStrTmp(i), ",")(2)) >= 0 Then
                objMsf.ColAlignment = CLng(Split(arrStrTmp(i), ",")(2))
            Else
                objMsf.ColAlignment = 4
            End If
        Else
            If InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignCenterBottom"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignCenterBottom
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignCenterCenter"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignCenterCenter
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignCenterTop"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignCenterTop
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignGeneral"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignGeneral
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignLeftBottom"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignLeftBottom
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignLeftCenter"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignLeftCenter
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignLeftTop"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignLeftTop
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignRightBottom"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignRightBottom
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignRightCenter"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignRightCenter
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignRightTop"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignRightTop
            Else
                objMsf.ColAlignment = 4
            End If
        End If
    Next
Else
    If blnReCreatCol Then
        objMsf.Cols = UBound(arrStrTmp) + 1
    End If
    For i = 0 To objMsf.Cols - 1
        'ȷ���Ƿ���ʾ����
        If IsNumeric(Split(arrStrTmp(i), ",")(3)) Then
            If CLng(Split(arrStrTmp(i), ",")(3)) > 0 Then
                '��ʾ
                objMsf.TextMatrix(0, i) = Split(arrStrTmp(i), ",")(0)
            Else
                '����ʾ
                objMsf.TextMatrix(0, i) = ""
            End If
        Else
            '��ʾ
            objMsf.TextMatrix(0, i) = Split(arrStrTmp(i), ",")(0)
        End If
        
        'ȷ���п�
        If IsNumeric(Split(arrStrTmp(i), ",")(1)) Then
            If CLng(Split(arrStrTmp(i), ",")(1)) >= 0 Then
                objMsf.ColWidth(i) = CLng(Split(arrStrTmp(i), ",")(1))
            Else
                objMsf.ColWidth(i) = 1440
            End If
        Else
            objMsf.ColWidth(i) = 1440
        End If
        
        'ȷ�����뷽ʽ
        If IsNumeric(Split(arrStrTmp(i), ",")(2)) Then
            If CLng(Split(arrStrTmp(i), ",")(2)) >= 0 Then
                objMsf.ColAlignment = CLng(Split(arrStrTmp(i), ",")(2))
            Else
                objMsf.ColAlignment = 4
            End If
        Else
            If InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignCenterBottom"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignCenterBottom
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignCenterCenter"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignCenterCenter
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignCenterTop"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignCenterTop
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignGeneral"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignGeneral
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignLeftBottom"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignLeftBottom
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignLeftCenter"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignLeftCenter
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignLeftTop"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignLeftTop
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignRightBottom"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignRightBottom
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignRightCenter"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignRightCenter
            ElseIf InStr(1, UCase(Split(arrStrTmp(i), ",")(2)), UCase("flexAlignRightTop"), vbTextCompare) > 0 Then
                objMsf.ColAlignment = AlignmentSettings.flexAlignRightTop
            Else
                objMsf.ColAlignment = 4
            End If
        End If
    Next
End If
End Sub

Private Sub ReSetRowCode(objMSH As MSHFlexGrid)
'���кŽ�����������
    For i = 1 To objMSH.Rows - 1
        objMSH.TextMatrix(i, 0) = CStr(i) & "��"
    Next
End Sub

Private Sub InitRow(lngRow As Long, ByVal bln��ҽ As Boolean)
'���н��г�ʼ��
If bln��ҽ Then
    ReSetRowCode msfW
    msfW.TextMatrix(lngRow, 1) = "��ϣ�"
    msfW.TextMatrix(lngRow, 2) = ""
    msfW.TextMatrix(lngRow, 3) = ""
    msfW.TextMatrix(lngRow, 4) = ""
    msfW.TextMatrix(lngRow, 5) = ""
    msfW.RowData(lngRow) = 0
Else
    ReSetRowCode msfH
    msfH.TextMatrix(lngRow, 1) = "��ҽ"
    msfH.TextMatrix(lngRow, 2) = ""
    msfH.TextMatrix(lngRow, 3) = ""
    msfH.TextMatrix(lngRow, 4) = ""
    msfH.TextMatrix(lngRow, 5) = ""
    msfH.TextMatrix(lngRow, 6) = ""
    msfH.RowData(lngRow) = 0
End If
End Sub

Private Sub chkDiff_Click(Index As Integer)
If Index = 0 Then
    '����
    msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x����) = IIf(chkDiff(Index).Value = 0, "", "��")
    If msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x����) = "��" Then
        chkDiff(0).Value = 1
        chkDiff(0).FontBold = True
        chkDiff(0).ForeColor = RGB(200, 0, 0)
    Else
        chkDiff(0).Value = 0
        chkDiff(0).FontBold = True
        chkDiff(0).ForeColor = 0
    End If
Else
    '����
    msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����) = IIf(chkDiff(Index).Value = 0, "", "��")
    If msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����) = "��" Then
        chkDiff(1).Value = 1
        chkDiff(1).FontBold = True
        chkDiff(1).ForeColor = RGB(200, 0, 0)
    Else
        chkDiff(1).Value = 0
        chkDiff(1).FontBold = True
        chkDiff(1).ForeColor = 0
    End If
End If
End Sub

Private Sub InitMe()
Dim strTmp As String

    '�ȳ�ʼ�����ؼ�
    msfW.ForeColorSel = 0
    msfW.BackColorSel = RGB(255, 255, 255)
    msfW.SelectionMode = flexSelectionFree
    msfW.Rows = 2
    
    msfH.ForeColorSel = msfW.ForeColorSel
    msfH.BackColorSel = msfW.BackColorSel
    msfH.SelectionMode = msfW.SelectionMode
    msfH.Rows = 2
    
    SetMSFFormat msfW, "���,600,flexAlignCenterCenter,0;�����,0,flexAlignCenterCenter,0;���,8000,flexAlignLeftCenter,0;����,800,4,0;�ο�,800,4,0;����ID,0,4,0"
    
    SetMSFFormat msfH, "���,600,flexAlignCenterCenter,0;�����,800,flexAlignCenterCenter,0;���,8000,flexAlignLeftCenter,0;����,800,4,0;�ο�,800,4,0;֤ID,0,4,0;����ID,0,4,0"
    
    '������ҽ����
    msfW.RowHeight(0) = 0
    msfW.Col = 0: msfW.Row = 1
    
    '������ҽ����
    msfH.RowHeight(0) = msfW.RowHeight(0)
    msfH.Col = 0: msfH.Row = 1
    
    '���ñ��������ʼ
    msfW.ColAlignment(2) = AlignmentSettings.flexAlignLeftCenter
    
    msfH.ColAlignment(2) = AlignmentSettings.flexAlignLeftCenter
    
    WestDiag = mWestDiag
    
    msfW.Col = 1
    msfW.Row = 1
    InitRow msfW.Rows - 1, True
    msfW_MouseDown 1, 0, msfW.ColWidth(0) + msfW.ColWidth(1) + 50, msfW.RowHeight(0) + 50
    msfW_EnterCell
    
    msfH.Col = 1
    msfH.Row = 1
    msfH_MouseDown 1, 0, msfH.ColWidth(0) + msfH.ColWidth(1) + 50, msfH.RowHeight(0) + 50
    InitRow msfH.Rows - 1, False
End Sub

Private Sub chkDiff_GotFocus(Index As Integer)
zlCommFun.OpenIme
End Sub

Private Sub chkDiff_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyReturn
        If Index = 0 Then
            If msfW.Row >= msfW.Rows - 1 Then
                If msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x���) <> "" Or msfW.RowData(msfW.Row) <> 0 Then
                    msfW.Rows = msfW.Rows + 1
                    msfW.Row = msfW.Rows - 1
                    InitRow msfW.Rows - 1, True
                Else
                    txtDiag(0).SetFocus
                    Exit Sub
                End If
            Else
                msfW.Row = msfW.Row + 1
            End If
            msfW.Col = EnmDiag��ҽ.x���
            SetSelColor msfW, msfW.Row
            msfW_EnterCell
            If mDispMode = False Then
                txtDiag(0).Visible = True
                chkDiff(0).Visible = True
                cmdCons(0).Visible = True
                cmdSel(0).Visible = True
                cmdCons(0).ZOrder
                txtDiag(0).ZOrder
                cmdSel(0).ZOrder
                chkDiff(0).ZOrder
            End If
            If txtDiag(0).Enabled And txtDiag(0).Visible And UserControl.Enabled Then
                txtDiag(0).SetFocus
            End If
        Else
            If msfH.Row >= msfH.Rows - 1 Then
                If msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z���) <> "" Or msfH.RowData(msfH.Row) <> 0 And (msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID) <> "") Then
                    msfH.Rows = msfH.Rows + 1
                    msfH.Row = msfH.Rows - 1
                    InitRow msfH.Rows - 1, False
                Else
                    If Trim(txtDiag(1).Text) = "" Then
                        txtDiag(1).SetFocus
                    End If
                    Exit Sub
                End If
            Else
                msfH.Row = msfH.Row + 1
            End If
            msfH.Col = EnmDiag��ҽ.z����
            SetSelColor msfH, msfH.Row
            msfH_EnterCell
            chkWH.Value = IIf(msfH.TextMatrix(msfH.Row - 1, EnmDiag��ҽ.z����) = "��ҽ", 0, 1)
            On Error Resume Next
            If mDispMode = False Then
                txtDiag(1).Visible = True
                cmdCons(1).Visible = True
                chkDiff(1).Visible = True
                cmdCons(1).ZOrder
                txtDiag(1).ZOrder
                cmdSel(1).ZOrder
                chkDiff(1).ZOrder
            End If
            If txtDiag(1).Enabled And txtDiag(1).Visible And UserControl.Enabled Then
                txtDiag(1).SetFocus
            End If
        End If
End Select
End Sub

Private Sub chkWH_Click()
If chkWH.Value = 0 Then
    msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����) = "��ҽ"
    chkWH.FontBold = True
    chkWH.ForeColor = RGB(0, 0, 180)
Else
    msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����) = "��ҽ"
    chkWH.FontBold = False
    chkWH.ForeColor = 0
End If
txtDiag(1).Text = ""
msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����ID) = "0"
msfH.RowData(msfH.Row) = 0
msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID) = "0"
End Sub

Private Sub chkWH_GotFocus()
zlCommFun.OpenIme
End Sub

Private Sub chkWH_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case vbKeyReturn
        txtDiag(1).SetFocus
End Select
End Sub

Private Sub cmdCons_Click(Index As Integer)
Dim clsTmp As New clsCISCore
    Select Case Index
        Case 0
            If IsNumeric(msfW.RowData(msfW.Row)) Then
                clsTmp.ShowDiagHelp 1, UserControl.Parent, msfW.RowData(msfW.Row)
            Else
                clsTmp.ShowDiagHelp 1, UserControl.Parent
            End If
        Case 1
            If IsNumeric(msfH.RowData(msfH.Row)) Then
                clsTmp.ShowDiagHelp 1, UserControl.Parent, msfH.RowData(msfH.Row)
            Else
                clsTmp.ShowDiagHelp 1, UserControl.Parent
            End If
    End Select
    Set clsTmp = Nothing
End Sub

Private Sub cmdCons_GotFocus(Index As Integer)
zlCommFun.OpenIme
End Sub

Private Sub cmdSel_Click(Index As Integer)
Dim strReturn As String
Dim strTmp As String
On Error GoTo ErrHandle
    
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Sub
    
    If Index = 1 And msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����) = "��ҽ" Then
        strReturn = msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z���)
        '��ʼ���ID
        frmDiagSel2.mlngID1 = msfH.RowData(msfH.Row)
        '��ʼ����ID
        If IsNumeric(msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����ID)) Then
            frmDiagSel2.mlngIllnID1 = CLng(msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����ID))
        End If
        '��ʼ֤����
        If IsNumeric(msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID)) Then
            frmDiagSel2.mlngID2 = CLng(msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID))
        End If
        If frmDiagSel2.ShowDiagSel(Me, strReturn, False) Then
            '���ظ�ʽ:  ������������;���ID;����ID;֤����;֤ID
            '�õ��������
            txtDiag(1).Text = Trim(Split(strReturn, ";")(3) & "  " & Split(strReturn, ";")(0))
            '�õ��Ϸ��ļ���ID
            strTmp = Trim(Split(strReturn, ";")(2))
            strTmp = IIf(IsNumeric(strTmp), CLng(strTmp), 0)
            strTmp = IIf(Trim(strTmp) = "0", "", strTmp)
            msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����ID) = strTmp
            '�õ����ID
            If IsNumeric(Trim(Split(strReturn, ";")(1))) Then
                msfH.RowData(msfH.Row) = CLng(Trim(Split(strReturn, ";")(1)))
            End If
            '֤ID
            If IsNumeric(Trim(Split(strReturn, ";")(4))) Then
                msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID) = CLng(Trim(Split(strReturn, ";")(4)))
            End If
        End If
    Else
        If Index = 0 Then
            '��ʼ���ID
            frmDiagSel2.mlngID1 = msfW.RowData(msfW.Row)
            '��ʼ����ID
            If IsNumeric(msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x����ID)) Then
                frmDiagSel2.mlngIllnID1 = CLng(msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x����ID))
            End If
        Else
            '��ʼ���ID
            frmDiagSel2.mlngID1 = msfH.RowData(msfH.Row)
            '��ʼ����ID
            If IsNumeric(msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����ID)) Then
                frmDiagSel2.mlngIllnID1 = CLng(msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����ID))
            End If
            '��ʼ֤����
            If IsNumeric(msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID)) Then
                frmDiagSel2.mlngID2 = CLng(msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID))
            End If
        End If
        strReturn = msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x���)
        If frmDiagSel2.ShowDiagSel(Me, strReturn, True) Then
            '���ظ�ʽ:  ������������;���ID;����ID;֤����;֤ID
            '�õ��������
            txtDiag(Index).Text = Trim(Split(strReturn, ";")(3) & "  " & Split(strReturn, ";")(0))
            '�õ��Ϸ��ļ���ID
            strTmp = Trim(Split(strReturn, ";")(2))
            strTmp = IIf(IsNumeric(strTmp), CLng(strTmp), 0)
            strTmp = IIf(Trim(strTmp) = "0", "", strTmp)
            If Index = 0 Then
                msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x����ID) = strTmp
            Else
                msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����ID) = strTmp
            End If
            '�õ����ID
            If IsNumeric(Trim(Split(strReturn, ";")(1))) Then
                If Index = 0 Then
                    msfW.RowData(msfW.Row) = CLng(Trim(Split(strReturn, ";")(1)))
                Else
                    msfH.RowData(msfH.Row) = CLng(Trim(Split(strReturn, ";")(1)))
                End If
            End If
            '֤ID
            If IsNumeric(Trim(Split(strReturn, ";")(4))) Then
                If Index <> 0 Then
                    msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID) = CLng(Trim(Split(strReturn, ";")(4)))
                End If
            End If
        End If
    End If
    On Error Resume Next
    txtDiag(Index).SetFocus
    Exit Sub
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msfH_GotFocus()
On Error Resume Next
zlCommFun.OpenIme
If mDispMode = False And txtDiag(1).Visible = True And txtDiag(1).Enabled And UserControl.Enabled Then txtDiag(1).SetFocus
txtDiag(1).ZOrder
msfH_EnterCell
End Sub

Private Sub msfH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
msfH.ToolTipText = txtDiag(1).Text
End Sub

Private Sub msfH_Scroll()
    If msfH.ColPos(0) < 0 Then
        msfH.Col = 0
    End If
    If msfW.ColPos(0) < 0 Then
        msfW.Col = 0
    End If
    chkDiff(0).Visible = False
    cmdCons(0).Visible = False
    txtDiag(0).Visible = False
    cmdSel(0).Visible = False
    chkDiff(1).Visible = False
    cmdCons(1).Visible = False
    txtDiag(1).Visible = False
    cmdSel(1).Visible = False
    chkWH.Visible = False
End Sub

Private Sub msfW_GotFocus()
On Error Resume Next
zlCommFun.OpenIme
If mDispMode = False And txtDiag(0).Visible And txtDiag(0).Enabled And UserControl.Enabled Then
    txtDiag(0).SetFocus
End If
txtDiag(0).ZOrder
cmdSel(0).ZOrder
End Sub

Private Sub msfW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
msfW.ToolTipText = txtDiag(0).Text
End Sub

Private Sub msfW_Scroll()
    If msfH.ColPos(0) < 0 Then
        msfH.Col = 0
    End If
    If msfW.ColPos(0) < 0 Then
        msfW.Col = 0
    End If
    chkDiff(0).Visible = False
    cmdCons(0).Visible = False
    txtDiag(0).Visible = False
    cmdSel(0).Visible = False
    chkDiff(1).Visible = False
    cmdCons(1).Visible = False
    txtDiag(1).Visible = False
    cmdSel(1).Visible = False
    chkWH.Visible = False
End Sub

Private Sub txtDiag_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrHandle
    'ֻҪ�����зǷ��ַ����˳�
    If InStr(LAWLChar, Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
        
    Select Case KeyAscii
        Case vbKeyReturn
            chkDiff(Index).SetFocus
        Case Asc("*")
            KeyAscii = 0
            cmdSel_Click Index
    End Select
    Exit Sub
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDiag_LostFocus(Index As Integer)
zlCommFun.OpenIme
End Sub

Private Sub txtDiag_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
txtDiag(Index).ToolTipText = txtDiag(Index).Text
End Sub

Private Sub UserControl_GotFocus()
    zlCommFun.OpenIme
End Sub

Private Sub UserControl_InitProperties()
    mWestDiag = True
    mDispMode = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDispMode = PropBag.ReadProperty("DispMode", False)
    mWestDiag = PropBag.ReadProperty("WestDiag", True)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", BorderStyleSettings.flexBorderNone)
    InitMe
End Sub

Private Sub UserControl_Resize()

    If msfH.ColPos(0) < 0 Then
        msfH.Col = 0
    End If
    If msfW.ColPos(0) < 0 Then
        msfW.Col = 0
    End If
    msfH.Left = 0
    msfH.Top = 0
    
    msfW.Left = msfH.Left
    msfW.Top = msfH.Top
    
    i = ScaleWidth - msfH.Left * 2
    msfW.Width = IIf(i < Screen.TwipsPerPixelX, Screen.TwipsPerPixelX, i)
    msfH.Width = msfW.Width
    
    i = ScaleHeight - Screen.TwipsPerPixelY * 2
    msfW.Height = IIf(i < Screen.TwipsPerPixelY, Screen.TwipsPerPixelY, i)
    msfH.Height = msfW.Height
    
    i = 0
    For j = 0 To msfW.Cols - 1
        If j <> 2 Then
            i = i + msfW.ColWidth(j)
        End If
    Next
    i = msfW.Width - Screen.TwipsPerPixelX * 6 - i - 15 * Screen.TwipsPerPixelX
    If i < 600 Then
        msfW.ColWidth(2) = 600
    Else
        msfW.ColWidth(2) = i
    End If
    txtDiag(0).Width = msfW.ColWidth(2) - Screen.TwipsPerPixelX * 3
    
    i = 0
    For j = 0 To msfH.Cols - 1
        If j <> 2 Then
            i = i + msfH.ColWidth(j)
        End If
    Next
    i = msfH.Width - Screen.TwipsPerPixelX * 6 - i - 15 * Screen.TwipsPerPixelX
    msfH.ColWidth(2) = IIf(i < 600, 600, i)
    txtDiag(1).Width = msfH.ColWidth(2) - Screen.TwipsPerPixelX * 3
    
    chkDiff(0).Visible = False
    cmdCons(0).Visible = False
    txtDiag(0).Visible = False
    cmdSel(0).Visible = False
    chkDiff(1).Visible = False
    cmdCons(1).Visible = False
    txtDiag(1).Visible = False
    cmdSel(1).Visible = False
    chkWH.Visible = False
End Sub

Private Sub msfH_EnterCell()
    '������
    'msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z�ο�) = CStr(msfH.RowData(msfH.Row))
    
    If msfH.Visible And mDispMode = False Then
        chkDiff(1).Visible = True
        cmdCons(1).Visible = True
        txtDiag(1).Visible = True
        cmdSel(1).Visible = True
        chkWH.Visible = True
    Else
        chkDiff(1).Visible = False
        cmdCons(1).Visible = False
        txtDiag(1).Visible = False
        cmdSel(1).Visible = False
        chkWH.Visible = False
    End If
    SetSelColor msfH, msfH.Row
    
    '������ҽ��ҽ������
    chkWH.Left = msfH.Left + msfH.ColWidth(0) + Screen.TwipsPerPixelY * 2
    chkWH.Top = msfH.Top + msfH.CellTop + Screen.TwipsPerPixelY * 0
    chkWH.Width = msfH.ColWidth(1) - Screen.TwipsPerPixelX * 2
    chkWH.Height = msfH.CellHeight - Screen.TwipsPerPixelY * 2
    If msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����) = "��ҽ" Then
        chkWH.Value = 1
        chkWH.FontBold = True
        chkWH.ForeColor = 0
    Else
        chkWH.Value = 0
        chkWH.FontBold = True
        chkWH.ForeColor = RGB(0, 0, 180)
    End If
    
    '����������ݵ��ı���
    txtDiag(1).Left = msfH.Left + msfH.ColWidth(0) + msfH.ColWidth(1) + Screen.TwipsPerPixelY * 2
    txtDiag(1).Top = msfH.Top + msfH.CellTop + Screen.TwipsPerPixelY * 0
    txtDiag(1).Width = msfH.ColWidth(2) - Screen.TwipsPerPixelX * 2
    txtDiag(1).Height = msfH.CellHeight - Screen.TwipsPerPixelY * 2
    
    mblnMode = True
    txtDiag(1).Text = msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z���)
    mblnMode = False
    
    'ѡ����
    cmdSel(1).Height = txtDiag(1).Height
    cmdSel(1).Top = txtDiag(1).Top
    cmdSel(1).Left = txtDiag(1).Left + txtDiag(1).Width - cmdSel(1).Width
    '�������︴ѡ��
    chkDiff(1).Left = msfH.Left + msfH.ColWidth(EnmDiag��ҽ.z���) + msfH.ColWidth(EnmDiag��ҽ.z����) + msfH.ColWidth(EnmDiag��ҽ.z���) + Screen.TwipsPerPixelY * 2
    chkDiff(1).Top = msfH.Top + msfH.CellTop + Screen.TwipsPerPixelY * 0
    chkDiff(1).Width = msfH.ColWidth(EnmDiag��ҽ.z����) - Screen.TwipsPerPixelX * 2
    chkDiff(1).Height = msfH.CellHeight - Screen.TwipsPerPixelY * 2
    
    If msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����) = "��" Then
        chkDiff(1).Value = 1
        chkDiff(1).FontBold = True
        chkDiff(1).ForeColor = RGB(200, 0, 0)
    Else
        chkDiff(1).Value = 0
        chkDiff(1).FontBold = True
        chkDiff(1).ForeColor = 0
    End If
    '���òο���ť
    cmdCons(1).Left = msfH.Left + msfH.ColWidth(EnmDiag��ҽ.z���) + msfH.ColWidth(EnmDiag��ҽ.z����) + msfH.ColWidth(EnmDiag��ҽ.z���) + msfH.ColWidth(EnmDiag��ҽ.z����) + Screen.TwipsPerPixelY * 2
    cmdCons(1).Top = msfH.Top + msfH.CellTop + Screen.TwipsPerPixelY * 0
    cmdCons(1).Width = msfH.ColWidth(EnmDiag��ҽ.z�ο�) - Screen.TwipsPerPixelX * 2
    cmdCons(1).Height = msfH.CellHeight - Screen.TwipsPerPixelY * 2
    
    chkWH.ZOrder
    chkDiff(1).ZOrder
    cmdCons(1).ZOrder
    txtDiag(1).ZOrder
    cmdSel(1).ZOrder
End Sub

Private Sub msfH_KeyPress(KeyAscii As Integer)
    If mDispMode = False Then
        txtDiag(1).Visible = True
        cmdSel(1).Visible = True
        chkDiff(1).Visible = True
        cmdCons(1).Visible = True
        chkWH.Visible = True
        cmdSel(1).ZOrder
    End If
    If txtDiag(1).Enabled And txtDiag(1).Visible Then txtDiag(1).SetFocus: txtDiag(1).SelStart = Len(txtDiag(1).Text)
End Sub

Private Sub msfH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Shift = 0 Then
    SetSelColor msfH, msfH.Row
    'ͨ�������������������
    If Y > msfH.RowPos(msfH.Row) + msfH.RowHeight(msfH.Row) Then
        If msfH.Row = msfH.Rows - 1 And (msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z���) <> "" Or msfH.RowData(msfH.Row) <> 0) Or (msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z֤ID) <> "") Then
            msfH.Rows = msfH.Rows + 1
            InitRow msfH.Rows - 1, False
            msfH.Row = msfH.Rows - 1
            msfH.Col = 1
            msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z����) = msfH.TextMatrix(msfH.Row - 1, EnmDiag��ҽ.z����)
            If chkWH.Enabled And chkWH.Visible Then chkWH.SetFocus
            SetSelColor msfH, msfH.Row
        End If
    End If
    UserControl_Resize
    msfH_EnterCell
ElseIf Button = 2 And mDispMode = False Then
    If msfH.MouseRow > 1 Then
        msfH.Row = msfH.MouseRow
        SetSelColor msfH, msfH.Row
        msfH_EnterCell
        If MsgBox("��Ҫɾ���к�Ϊ " & msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z���) & " �������", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            i = msfH.Row
            If i = msfH.Rows - 1 Then
                i = i - 1
            End If
            msfH.RemoveItem msfH.Row
            msfH.Row = i
            ReSetRowCode msfH
            SetSelColor msfH, msfH.Row
        End If
        msfH_EnterCell
    ElseIf msfH.MouseRow = 1 And (msfH.TextMatrix(1, EnmDiag��ҽ.z���) <> "" Or msfH.TextMatrix(1, EnmDiag��ҽ.z֤ID) <> "" Or msfH.RowData(1) <> 0) Then
        msfH.Row = 1
        SetSelColor msfH, msfH.Row
        msfH_EnterCell
        If MsgBox("��Ҫɾ���к�Ϊ " & msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z���) & " �������", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            msfH.TextMatrix(1, EnmDiag��ҽ.z���) = ""
            msfH.TextMatrix(1, EnmDiag��ҽ.z����) = ""
            msfH.TextMatrix(1, EnmDiag��ҽ.z֤ID) = ""
            msfH.RowData(1) = 0
            txtDiag(1).Text = ""
            chkWH.Value = 1
            chkWH_Click
            chkDiff(1).Value = 0
            chkDiff_Click 1
        End If
        msfH_EnterCell
    End If
End If
End Sub

Private Sub msfH_SelChange()
    msfH.Redraw = False
    msfH.ColSel = msfH.Col
    msfH.RowSel = msfH.Row
    msfH.Redraw = True
End Sub

Private Sub msfW_EnterCell()
    '������
'    msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x�ο�) = CStr(msfW.RowData(msfW.Row))
    
    If msfW.Visible And mDispMode = False Then
        chkDiff(0).Visible = True
        cmdCons(0).Visible = True
        txtDiag(0).Visible = True
        cmdSel(0).Visible = True
    Else
        chkDiff(0).Visible = False
        cmdCons(0).Visible = False
        txtDiag(0).Visible = False
        cmdSel(0).Visible = False
    End If
    SetSelColor msfW, msfW.Row
    

    '����������ݵ��ı���
    txtDiag(0).Left = msfW.Left + msfW.ColWidth(0) + msfW.ColWidth(1) + Screen.TwipsPerPixelY * 3
    txtDiag(0).Top = msfW.Top + msfW.CellTop + Screen.TwipsPerPixelY * 0
    txtDiag(0).Width = msfW.ColWidth(2) - Screen.TwipsPerPixelX * 3
    txtDiag(0).Height = msfW.CellHeight - Screen.TwipsPerPixelY * 2
    
    mblnMode = True
    txtDiag(0).Text = msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x���)
    mblnMode = False
    
    'ѡ����
    cmdSel(0).Height = txtDiag(0).Height
    cmdSel(0).Top = txtDiag(0).Top
    cmdSel(0).Left = txtDiag(0).Left + txtDiag(0).Width - cmdSel(0).Width
    '�������︴ѡ��
    chkDiff(0).Left = msfW.Left + msfW.ColWidth(0) + msfW.ColWidth(1) + msfW.ColWidth(2) + Screen.TwipsPerPixelY * 2
    chkDiff(0).Top = msfW.Top + msfW.CellTop + Screen.TwipsPerPixelY * 0
    chkDiff(0).Width = msfW.ColWidth(3) - Screen.TwipsPerPixelX * 2
    chkDiff(0).Height = msfW.CellHeight - Screen.TwipsPerPixelY * 2
    
    If msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x����) = "��" Then
        chkDiff(0).Value = 1
        chkDiff(0).FontBold = True
        chkDiff(0).ForeColor = RGB(200, 0, 0)
    Else
        chkDiff(0).Value = 0
        chkDiff(0).FontBold = True
        chkDiff(0).ForeColor = 0
    End If
    
    '���òο���ť
    cmdCons(0).Left = msfW.Left + msfW.ColWidth(0) + msfW.ColWidth(1) + msfW.ColWidth(2) + msfW.ColWidth(3) + Screen.TwipsPerPixelY * 2
    cmdCons(0).Top = msfW.Top + msfW.CellTop + Screen.TwipsPerPixelY * 0
    cmdCons(0).Width = msfW.ColWidth(4) - Screen.TwipsPerPixelX * 2
    cmdCons(0).Height = msfW.CellHeight - Screen.TwipsPerPixelY * 2
    
    
    
    chkDiff(0).ZOrder
    cmdCons(0).ZOrder
    txtDiag(0).ZOrder
    cmdSel(0).ZOrder
End Sub

Private Sub msfW_KeyPress(KeyAscii As Integer)
    If mDispMode = False Then
        txtDiag(0).Visible = True
        cmdSel(0).Visible = True
        chkDiff(0).Visible = True
        cmdCons(0).Visible = True
        cmdSel(0).ZOrder
    End If
    If txtDiag(0).Enabled And txtDiag(0).Visible Then txtDiag(0).SetFocus: txtDiag(0).SelStart = Len(txtDiag(0).Text)
End Sub

Private Sub msfW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Shift = 0 Then
    SetSelColor msfW, msfW.Row
    'ͨ�������������������
    If Y > msfW.RowPos(msfW.Row) + msfW.RowHeight(msfW.Row) Then
        If msfW.Row = msfW.Rows - 1 And (msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x���) <> "" Or msfW.RowData(msfW.Row) <> 0) Then
            msfW.Rows = msfW.Rows + 1
            InitRow msfW.Rows - 1, True
            msfW.Row = msfW.Rows - 1
            msfW.Col = 2
            If txtDiag(0).Enabled And txtDiag(0).Visible Then txtDiag(0).SetFocus
            SetSelColor msfW, msfW.Row
        End If
    End If
    UserControl_Resize
    msfW_EnterCell
ElseIf Button = 2 And mDispMode = False Then
    If msfW.MouseRow > 1 Then
        msfW.Row = msfW.MouseRow
        SetSelColor msfW, msfW.Row
        msfW_EnterCell
        If MsgBox("��Ҫɾ���к�Ϊ " & msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x���) & " �������", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            i = msfW.Row
            If i = msfW.Rows - 1 Then
                i = i - 1
            End If
            msfW.RemoveItem msfW.Row
            msfW.Row = i
            ReSetRowCode msfW
            SetSelColor msfW, msfW.Row
        End If
        msfW_EnterCell
    ElseIf msfW.MouseRow = 1 And (msfW.TextMatrix(1, EnmDiag��ҽ.x���) <> "" Or msfW.TextMatrix(1, EnmDiag��ҽ.x����) <> "" Or msfW.RowData(1) <> 0) Then
        msfW.Row = 1
        SetSelColor msfW, msfW.Row
        msfW_EnterCell
        If MsgBox("��Ҫɾ���к�Ϊ " & msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x���) & " �������", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            msfW.TextMatrix(1, EnmDiag��ҽ.x���) = ""
            msfW.TextMatrix(1, EnmDiag��ҽ.x����) = ""
            msfW.RowData(1) = 0
            txtDiag(0).Text = ""
            chkDiff(0).Value = 0
            chkDiff_Click 0
        End If
        msfW_EnterCell
    End If
End If
End Sub

Private Sub msfW_SelChange()
    msfW.Redraw = False
    msfW.ColSel = msfW.Col
    msfW.RowSel = msfW.Row
    msfW.Redraw = True
End Sub

Private Sub txtDiag_Change(Index As Integer)
    If mblnMode = False Then
        '��������п���
        If Index = 0 Then '��ҽ
            msfW.TextMatrix(msfW.Row, EnmDiag��ҽ.x���) = txtDiag(Index).Text
        Else               '��ҽ
            msfH.TextMatrix(msfH.Row, EnmDiag��ҽ.z���) = txtDiag(Index).Text
        End If
    End If
End Sub

Private Sub txtDiag_GotFocus(Index As Integer)
    If Index = 0 Then
        msfW.Col = 2
    Else
        msfH.Col = 2
    End If
zlControl.TxtSelAll txtDiag(Index)
zlCommFun.OpenIme True
End Sub

Private Sub txtDiag_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
    'ֻҪ�����д���;��˳�
    If InStr(LAWLChar, Chr(KeyCode)) > 0 Then
        KeyCode = 0
        Exit Sub
    End If
    
    '����ǲ������������¼����Զ������л�������
    If KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Then
        KeyCode = 0
        If Index = 0 Then
            If msfW.Row > 1 Then
                msfW.Row = msfW.Row - 1
                msfW_EnterCell
                SetSelColor msfW, msfW.Row
            End If
        Else
            If msfH.Row > 1 Then
                msfH.Row = msfH.Row - 1
                msfH_EnterCell
                SetSelColor msfH, msfH.Row
            End If
        End If
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        KeyCode = 0
        If Index = 0 Then
            If msfW.Row < msfW.Rows - 1 Then
                msfW.Row = msfW.Row + 1
                msfW_EnterCell
                SetSelColor msfW, msfW.Row
            End If
        Else
            If msfH.Row < msfH.Rows - 1 Then
                msfH.Row = msfH.Row + 1
                msfH_EnterCell
                SetSelColor msfH, msfH.Row
            End If
        End If
    ElseIf KeyCode = vbKeyDelete And Shift = 0 Then
        '��ʱ��ʾ�ǲ���Ҫ�����ǰ�е�����
        If MsgBox("��Ҫ�����ǰ�е�����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            KeyCode = 0
            If Index = 0 Then
                InitRow msfW.Row, True
                msfW_EnterCell
            Else
                InitRow msfH.Row, False
                msfH_EnterCell
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDiag_Validate(Index As Integer, Cancel As Boolean)
    Cancel = LocalCheck�Ƿ�Ƿ�(txtDiag(Index), LAWLChar)
End Sub

Private Sub chkWH_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
    If KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp Then
        If msfH.Row > 1 Then
            msfH.Row = msfH.Row - 1
            msfH_EnterCell
            SetSelColor msfH, msfH.Row
        End If
    ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If msfH.Row < msfH.Rows - 1 Then
            msfH.Row = msfH.Row + 1
            msfH_EnterCell
            SetSelColor msfH, msfH.Row
        End If
    End If
    Exit Sub
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub UserControl_Show()
On Error GoTo ErrHandle
Dim objCtl As Control
Dim i As Integer

    cmdSel(0).ToolTipText = "��*��ѡ����"
    cmdSel(1).ToolTipText = "��*��ѡ����"
    'ֻ������ʱ��ʾ
    UserControl_Resize
'    If mblnLoaded = False Then
'        InitMe
'    End If
    mblnLoaded = True
    Exit Sub
ErrHandle:
    If Ambient.UserMode = False Or InDesign = False Then SetErr Err.Number, Err.Description:    Exit Sub
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InDesign() As Boolean
'���ܣ��жϵ�ǰ���г����Ƿ���VB�Ĺ��̻�����
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Private Sub SetErr(lngErrNum As Long, strErr As String)
'���ô��������������
'���lngErrNum=-1 ��ʾ �ؼ��Լ�����Ĵ���
mReturnErrnumber = lngErrNum
mReturnErrDescription = strErr
End Sub

Public Property Get ReturnErrNumber() As Long
'�������һ�εĴ����
    ReturnErrNumber = mReturnErrnumber
End Property

Public Property Get ReturnErrDescription() As String
'�������һ�δ��������ַ���
    ReturnErrDescription = mReturnErrDescription
End Property

Public Property Get ID���˲���() As Long
'���ز��˲���ID
    ID���˲��� = mlng����id
End Property

Public Property Let ID���˲���(ByVal New_ID���˲��� As Long)
'���ò��˲���ID,�����ò����ǲ��Ǵ���
    mlng����id = New_ID���˲���
    ShowDiag mlng����id, Not mDispMode
End Property

Public Property Get DispMode() As Boolean
'�Ƿ�Ϊ��ʾģʽ
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    mDispMode = New_DispMode
    msfW_EnterCell
    msfH_EnterCell
    PropertyChanged "DispMode"
End Property

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Terminate()
    If rsTmp.State = adStateOpen Then rsTmp.Close
    Set rsTmp = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, BorderStyleSettings.flexBorderNone)
    Call PropBag.WriteProperty("WestDiag", mWestDiag, True)
    Call PropBag.WriteProperty("DispMode", mDispMode, False)
End Sub

Public Property Get WestDiag() As Boolean
    WestDiag = mWestDiag
End Property

Public Property Let WestDiag(ByVal New_WestDiag As Boolean)
    mWestDiag = New_WestDiag
    If New_WestDiag = True Then
        'ѡ������ҽ�����
        msfW.Visible = True
        msfW.ZOrder
        cmdCons(0).ZOrder
        chkDiff(0).ZOrder
        txtDiag(0).ZOrder
        cmdSel(0).ZOrder
        msfW_EnterCell
        
        msfH.Visible = False
        msfH.ZOrder 1
        txtDiag(1).Visible = False
        cmdSel(1).Visible = False
        chkDiff(1).Visible = False
        cmdCons(1).Visible = False
        chkWH.Visible = False
    Else
        'ѡ������ҽ�����
        msfH.Visible = True
        msfH.ZOrder
        cmdCons(1).ZOrder
        chkDiff(1).ZOrder
        txtDiag(1).ZOrder
        cmdSel(1).ZOrder
        chkWH.ZOrder
        msfH_EnterCell
        
        txtDiag(0).Visible = False
        cmdSel(0).Visible = False
        chkDiff(0).Visible = False
        cmdCons(0).Visible = False
        msfW.Visible = False
        msfW.ZOrder 1
    End If
    PropertyChanged "WestDiag"
End Property

Public Property Get Text() As String
'Ϊÿһ���ؼ������ı�ת������
Dim i As Long
Dim strTmp As String

'ͨ���û���������ݵõ�ת���ı�
    If mWestDiag Then '��ҽ���
        For i = 1 To msfW.Rows - 1
            If i = 1 Then
                If Trim(msfW.TextMatrix(i, EnmDiag��ҽ.x���)) = "" Then
                    Text = ""
                    Exit Property
                Else
                    strTmp = strTmp & msfW.TextMatrix(i, EnmDiag��ҽ.x���) & msfW.TextMatrix(i, EnmDiag��ҽ.x���) & "  " & msfW.TextMatrix(i, EnmDiag��ҽ.x����) & IIf(i = msfW.Rows - 1, "", vbCrLf)
                End If
            Else
                strTmp = strTmp & msfW.TextMatrix(i, EnmDiag��ҽ.x���) & msfW.TextMatrix(i, EnmDiag��ҽ.x���) & "  " & msfW.TextMatrix(i, EnmDiag��ҽ.x����) & IIf(i = msfW.Rows - 1, "", vbCrLf)
            End If
        Next
    Else
        For i = 1 To msfH.Rows - 1
            If i = 1 Then
                If Trim(msfH.TextMatrix(i, EnmDiag��ҽ.z���)) = "" Then
                    Text = ""
                    Exit Property
                Else
                    strTmp = strTmp & msfH.TextMatrix(i, EnmDiag��ҽ.z���) & msfH.TextMatrix(i, EnmDiag��ҽ.z���) & "  " & msfH.TextMatrix(i, EnmDiag��ҽ.z����) & IIf(i = msfH.Rows - 1, "", vbCrLf)
                End If
            Else
                strTmp = strTmp & msfH.TextMatrix(i, EnmDiag��ҽ.z���) & msfH.TextMatrix(i, EnmDiag��ҽ.z���) & "  " & msfH.TextMatrix(i, EnmDiag��ҽ.z����) & IIf(i = msfH.Rows - 1, "", vbCrLf)
            End If
        Next
    End If
    Text = strTmp
End Property
 
Private Sub UserControl_EnterFocus()
    On Error Resume Next
    UserControl.Parent.CallBack_GotFocus
End Sub

