VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frm����ѡ��_�ɶ��ڽ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ѡ��"
   ClientHeight    =   3645
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6870
   Icon            =   "frm����ѡ��_�ɶ��ڽ�.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CancelButton 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   5370
      TabIndex        =   1
      Top             =   3105
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   3975
      TabIndex        =   0
      Top             =   3090
      Width           =   1215
   End
   Begin ZL9BillEdit.BillEdit msf���Ӳ��� 
      Height          =   2730
      Left            =   135
      TabIndex        =   2
      Top             =   240
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   4815
      Enabled         =   -1  'True
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
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
End
Attribute VB_Name = "frm����ѡ��_�ɶ��ڽ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mlng����ID   As Long, mstr����֢ As String

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim vat����֢ As Variant
    
    With msf���Ӳ���
        '�����������������б�������
        .Rows = 4
        .Cols = 2
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 1000
        .ColWidth(1) = 4800
        .TextMatrix(0, 1) = "���ֱ���������"
        .TextMatrix(1, 0) = "�������"
        .PrimaryCol = 1
        
        '���ø��е���ֵ����ȷ����Щ�пɲ������ɱ༭��Ǳ༭��
        .ColData(1) = 1  '�ı��������������ť
        'δ���õ��е���ֵ��Ϊ0 (Ĭ��), ��Щ�н�����ѡ�񵫲����޸�
    End With
    
    If mstr����֢ <> "" Then
        If InStr(mstr����֢, "|") > 0 Then
            vat����֢ = Split(mstr����֢, "|")
            msf���Ӳ���.Rows = UBound(vat����֢) + 2
            For i = 0 To UBound(vat����֢) - 1
                msf���Ӳ���.TextMatrix(i + 1, 1) = "[" & Split(vat����֢(i), ";")(0) & "]" & Split(vat����֢(i), ";")(1)
            Next
        End If
    End If
    

End Sub

Private Sub msf���Ӳ���_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If Row < 2 Then Cancel = True
End Sub

Private Sub msf���Ӳ���_CommandClick()
    Dim str���� As String
    Select Case msf���Ӳ���.ColData(msf���Ӳ���.COL)
        Case 1
            str���� = msf���Ӳ���.TextMatrix(msf���Ӳ���.Row, msf���Ӳ���.COL)
            str���� = BZXZ_�ɶ��ڽ�(str����)
            If str���� = "" Then Exit Sub
            msf���Ӳ���.TextMatrix(msf���Ӳ���.Row, msf���Ӳ���.COL) = str����
    End Select
End Sub

Private Sub msf���Ӳ���_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim str���� As String
    If KeyCode <> vbKeyReturn Or msf���Ӳ���.COL = 0 Then Exit Sub
    str���� = msf���Ӳ���.Text
    
    If str���� = "" And msf���Ӳ���.Rows = msf���Ӳ���.Row + 1 Then
        SendKeys "{Tab}"
    End If
    
    If str���� = "" And msf���Ӳ���.Rows = msf���Ӳ���.Row + 2 Then
        If msf���Ӳ���.TextMatrix(msf���Ӳ���.Row + 1, msf���Ӳ���.COL) = "" Then
            SendKeys "{Tab}"
        End If
    End If
    
    'Cancel = True
    str���� = BZXZ_�ɶ��ڽ�(str����, 1)
    If str���� <> "" Then
        msf���Ӳ���.Text = str����
        msf���Ӳ���.TextMatrix(msf���Ӳ���.Row, msf���Ӳ���.COL) = str����
    End If
End Sub

Function BZXZ_�ɶ��ڽ�(ByVal StrInput As String, Optional strLoad As String = 0) As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmpSQL As String
    
    On Error Resume Next
   
    
    If StrInput = "" And strLoad = 1 Then Exit Function
    
    If StrInput = "" Then
        strTmpSQL = "Select ID,����,���� from ���ղ���"
    Else
        strTmpSQL = "Select ID,����,���� from ���ղ���" & _
                 " Where ���� Like '%" & StrInput & "%' OR " & _
                 "���� like '%" & StrInput & "%' Or " & _
                 "lower(����) like lower('%" & StrInput & "%')"
    End If
    
    Set rsTmp = frmPubSel.ShowSelect(Me, strTmpSQL, 0, "����", True, , , , False, gcnOracle)
    If rsTmp Is Nothing Then Exit Function
    BZXZ_�ɶ��ڽ� = "[" & rsTmp!���� & "]" & rsTmp!����
End Function

Function GetCode(lng����ID) As Boolean
    Dim rsTmp As New ADODB.Recordset
    mlng����ID = 0
    mstr����֢ = ""
    mlng����ID = lng����ID
    
    gstrSQL = "Select * from �����ʻ� Where ����ID=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ����֢", lng����ID, TYPE_�ɶ��ڽ�)
    mstr����֢ = Nvl(rsTmp!�������)
    
    frm����ѡ��_�ɶ��ڽ�.Show 1
    GetCode = True
End Function

Private Sub OKButton_Click()

    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim i As Integer
    Dim str���ֱ��� As String, str�������� As String, lng����ID  As Long, str������� As String
    Dim strͳ������� As String, strסԺ��ˮ�� As String
    Dim StrInput As String, strOutput As String
    Dim lng��Ժ���� As Long
    '>Beging ���ִ���
    lng��Ժ���� = 0
    If msf���Ӳ���.Rows < 1 Then
        MsgBox "�����������Ϣ��", vbInformation, gstrSysName
        Exit Sub
    End If
 
    
    For i = 1 To msf���Ӳ���.Rows - 1
        str���ֱ��� = msf���Ӳ���.TextMatrix(i, 1)
        If str���ֱ��� <> "" Then
            If InStr(str���ֱ���, "]") > 0 And InStr(str���ֱ���, "[") > 0 And InStr(str���ֱ���, "]") - InStr(str���ֱ���, "[") > 1 Then
                str�������� = Mid(str���ֱ���, InStr(str���ֱ���, "]") + 1)
                str���ֱ��� = Mid(str���ֱ���, InStr(str���ֱ���, "[") + 1, InStr(str���ֱ���, "]") - InStr(str���ֱ���, "[") - 1)
                '������ 20051029
                If str�������� = "ƽ��" Or str�������� = "�ʹ���" Then
                   gstrSQL = "Select id from ���ղ��� where ����=[1]"
                   Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�������ղ���", str��������)
                   lng��Ժ���� = rsTmp!ID
                End If
                str������� = str������� & str���ֱ��� & ";" & str�������� & "|"
                
                gstrSQL = "Select * from �����ʻ� where ����ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ʻ�", mlng����ID)
                strͳ������� = Split(rsTmp!����֤��, "|")(0)
                strסԺ��ˮ�� = rsTmp!˳���
                '����������ϴ�����
                StrInput = strͳ������� & vbTab & strסԺ��ˮ�� & vbTab & Rpad(Rpad(str���ֱ���, 20), 200)
                
                If ҵ������_�ɶ��ڽ�(����֢�����ϴ�_�ڽ�, StrInput, strOutput) Then Exit Sub
            End If
        End If
    Next
    
    If str������� <> "" Then
        gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ɶ��ڽ� & ",'�������','''" & str������� & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "���沢��֢")
    End If
    If lng��Ժ���� > 0 Then
       gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_�ɶ��ڽ� & ",'����id','''" & lng��Ժ���� & "''')"
       Call zlDatabase.ExecuteProcedure(gstrSQL, "������������")
    End If
    '>End ���ִ���
  
    Unload Me
End Sub


