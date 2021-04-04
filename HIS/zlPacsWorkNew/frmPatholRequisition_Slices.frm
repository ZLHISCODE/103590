VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatholRequisition_Slices 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ƭ����"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6075
   Icon            =   "frmPatholRequisition_Slices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5915.493
   ScaleMode       =   0  'User
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdSure 
      Cancel          =   -1  'True
      Caption         =   "ȷ��(&S)"
      Height          =   400
      Left            =   3120
      TabIndex        =   9
      Top             =   4748
      Width           =   900
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      Height          =   400
      Left            =   5040
      TabIndex        =   8
      Top             =   4748
      Width           =   900
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��(&D)"
      Height          =   400
      Left            =   4080
      TabIndex        =   7
      Top             =   4748
      Width           =   900
   End
   Begin VB.PictureBox picShow 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   3135
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtShow 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         DrawMode        =   1  'Blackness
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.TextBox txtDescription 
      Height          =   975
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   4755
   End
   Begin VB.TextBox txtRequestDoctor 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3120
      Width           =   1515
   End
   Begin MSComCtl2.DTPicker dtpRequestTime 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   3120
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm"
      Format          =   67633155
      CurrentDate     =   40646.4399652778
   End
   Begin zl9PACSWork.ucFlexGrid ufgData 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   5106
      GridRows        =   51
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      DataFontCharset =   134
      DataFontWeight  =   400
      ExtendLastCol   =   -1  'True
   End
   Begin VB.Label labRequestTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ�䣺"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   5
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label labRequestDoctor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����ҽʦ��"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3480
      TabIndex        =   4
      Top             =   3180
      Width           =   900
   End
   Begin VB.Label labDescription 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   900
   End
End
Attribute VB_Name = "frmPatholRequisition_Slices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mufgParentContextGrid As ucFlexGrid
Private mufgParentRequestGrid As ucFlexGrid

Private mlngCurRequestId As Long
Private mlngPatholAdviceId As Long
Private mlngRequestType As Long


Private mfrmOwner As Form

Private Const M_STR_REQUSLICES_COLS = "|�Ŀ�ID,hide,uncfg|�Ŀ��,hide,uncfg|�걾����,cbx<>,w2600,uncfg|��Ƭ��ʽ,cbx<1-����,2-����,3-����,4-��Ƭ,5-��Ⱦ,6-��Ƭ>,uncfg,w1300|��Ƭ����,w1200,uncfg|"
Private Const M_STR_REQUSLICES_CONVERTFORMAT = "��Ƭ��ʽ:1-����,2-����,3-����,4-��Ƭ,5-��Ⱦ,6-��Ƭ"

Public blnIsOk As Boolean


Public Function ShowSlicesRequestWindow(ufgParentRequestGrid As ucFlexGrid, ufgParentContextGrid As ucFlexGrid, _
    ByVal lngPatholAdviceId As Long, ByVal lngRequestId As Long, ByVal lngRequestType As Long, owner As Form) As Boolean
'��ʾ��Ƭ���봰��
    Set mufgParentRequestGrid = ufgParentRequestGrid
    Set mufgParentContextGrid = ufgParentContextGrid
    
    Set mfrmOwner = owner

    mlngPatholAdviceId = lngPatholAdviceId
    mlngCurRequestId = lngRequestId
    mlngRequestType = lngRequestType
    blnIsOk = False

    Call Me.Show(1, owner)
End Function

Private Sub InitRequisitionSlicesList()
On Error GoTo errHandle
'��ʼ��������Ƭ�б�
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    
    ufgData.IsKeepRows = True
    
    '��ֹ�Ҽ������б����ô���
    ufgData.IsEjectConfig = False
     ufgData.GridRows = 31
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    '������
    ufgData.DefaultColNames = M_STR_REQUSLICES_COLS
    ufgData.ColNames = M_STR_REQUSLICES_COLS
    ufgData.ColConvertFormat = M_STR_REQUSLICES_CONVERTFORMAT
    
   
    
    strSql = "select �Ŀ�ID,��� as �Ŀ�� ,�걾����,'-'||ȡ��λ�� ȡ��λ�� from ����ȡ����Ϣ where ����ҽ��ID=[1] and ȷ��״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�õ��걾����", mlngPatholAdviceId)
    
    If rsTemp.RecordCount < 1 Then Exit Sub
    For i = 1 To rsTemp.RecordCount
    
        strTemp = strTemp & "|" & "#" & Nvl(rsTemp!�Ŀ�id) & "!" & Nvl(rsTemp!�Ŀ��) & ";" & Nvl(rsTemp!�Ŀ��) & "-" & Nvl(rsTemp!�걾����) & IIf(Len(Nvl(rsTemp!ȡ��λ��)) <> 1, Nvl(rsTemp!ȡ��λ��), "")
        
        rsTemp.MoveNext
    Next i
    
    ufgData.ComboxListFormat(ufgData.GetColIndex("�걾����")) = Mid(strTemp, 2, Len(strTemp))
    
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub LoadMaterialInf()
'����Ŀ���Ϣ
    Dim strSql As String

    strSql = "select �Ŀ�ID,��� as �Ŀ��,�걾����,'0' as ��Ƭ��ʽ,1 as ��Ƭ���� from ����ȡ����Ϣ where ����ҽ��ID=[1] and ȷ��״̬=1"

    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngPatholAdviceId)

    ufgData.RefreshData
End Sub


Private Sub SaveSlicesRequest()
'������Ƭ����
On Error GoTo errHandle
    Dim lngNewRow As Long
    Dim i As Integer
    Dim lngSlicesType As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset


    If mlngCurRequestId <= 0 Then
        '��Ӽ��������Ϣ
        strSql = "select Zl_��������_����([1],[2],[3],[4],[5],[6],[7]) as ����ֵ from dual"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                mlngPatholAdviceId, _
                                                txtRequestDoctor.Text, _
                                                CDate(dtpRequestTime.value), _
                                                3, 0, _
                                                0, _
                                                txtDescription.Text)

        If rsData.RecordCount <= 0 Then
            Call err.Raise(0, "SaveSpeExamRequest", "δ�ɹ���ȡ�����������ID,����ʧ�ܡ�")
            Exit Sub
        End If

        '���ý�����Ϣ
        lngNewRow = mufgParentRequestGrid.NewRow

        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_����ID) = rsData!����ֵ
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_������) = txtRequestDoctor.Text
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_��������) = "����Ƭ"
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_����ʱ��) = dtpRequestTime.value
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_��������) = txtDescription.Text
        mufgParentRequestGrid.Text(lngNewRow, gstrRequisition_��ǰ״̬) = "������"

        mlngCurRequestId = Val(Nvl(rsData!����ֵ))

        '��λ��������
        Call mufgParentRequestGrid.LocateRow(lngNewRow)

        Call mufgParentContextGrid.ClearListData
    End If


       For i = 1 To ufgData.GridRows - 1

        If Trim(ufgData.Text(i, "��Ƭ��ʽ")) <> "" Then
            
            '�����Ƭ������Ŀ
            lngSlicesType = GetSlicesTypeCode(Nvl(Mid(ufgData.Text(i, "�걾����"), 1, InStr(ufgData.Text(i, "�걾����"), "!") - 1)))
        
            strSql = "select Zl_��������_��Ƭ��Ŀ_����([1],[2],[3],[4],[5],[6]) as ����ֵ from dual"
            Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                    mlngPatholAdviceId, _
                                                    Nvl(Mid(ufgData.Text(i, "�걾����"), 1, InStr(ufgData.Text(i, "�걾����"), "!") - 1)), _
                                                    mlngCurRequestId, _
                                                    lngSlicesType, _
                                                    Val(Nvl(ufgData.Text(i, "��Ƭ��ʽ"))), _
                                                    Val(Nvl(ufgData.Text(i, "��Ƭ����"))))
        
            If rsData.RecordCount <= 0 Then
                Call err.Raise(0, "SaveSpeExamRequest", "δ�ɹ���ȡ���������Ƭ��ĿID,����ʧ�ܡ�")
                Exit Sub
            End If
        
            '���ý�����Ϣ
            lngNewRow = mufgParentContextGrid.NewRow
        
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_ID) = rsData!����ֵ
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_�걾����) = Nvl(ufgData.DisplayText(i, "�걾����"))
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_�Ŀ��) = Nvl(Mid(ufgData.Text(i, "�걾����"), InStr(ufgData.Text(i, "�걾����"), "!") + 1, Len(ufgData.Text(i, "�걾����"))))
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_��Ƭ����) = GetSlicesTypeValue(lngSlicesType)
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_��Ƭ��ʽ) = Mid(Nvl(ufgData.Text(i, "��Ƭ��ʽ")), 3, 4)
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_��ǰ״̬) = "������"
            mufgParentContextGrid.Text(lngNewRow, gstrRequest_Slices_��Ƭ����) = Val(Nvl(ufgData.Text(i, "��Ƭ����")))
            
            Call mufgParentContextGrid.LocateRow(lngNewRow)
        End If
     Next i
     
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub



Private Sub cmdDel_Click()
'ɾ��ѡ��������
    ufgData.DelCurRow
End Sub

Private Sub cmdExit_Click()
'�˳���Ƭ����
    blnIsOk = False
    Call Me.Hide
End Sub


Private Function GetSlicesTypeValue(ByVal lngSlicesType As Long) As String
'��ȡ��Ƭ����ȡֵ
    Select Case lngSlicesType
        Case 0
            GetSlicesTypeValue = "ʯ����Ƭ"
        Case 1
            GetSlicesTypeValue = "������Ƭ"
        Case 2
            GetSlicesTypeValue = "ϸ����Ƭ"
    End Select
End Function


Private Function GetSlicesTypeCode(ByVal strMaterialId As String) As Long
'ȡ����Ƭ���ʹ���
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    '����Ǳ�����飬����Ҫ�жϵ�ǰ�Ŀ��Ƿ�Ϊ���࣬���Ϊ���࣬����Ƭ����Ϊʯ����Ƭ������Ϊ������Ƭ
    strSql = "select case ������� when 1 then case �Ƿ���� when 0 then 1 else 0 end when 2 then 2 else 0 end as ��Ƭ���� " & _
            "from ��������Ϣ a, ����ȡ����Ϣ b where a.����ҽ��ID=b.����ҽ��ID and b.�Ŀ�id=[1]"

    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strMaterialId)
    
    If rsData.RecordCount <= 0 Then
        Call err.Raise(0, "GetSlicesTypeCode", "���ܻ�ȡ��Ч����Ƭ���͡�")
        Exit Function
    End If
    
    GetSlicesTypeCode = rsData!��Ƭ����
End Function


Private Sub cmdSure_Click()
On Error GoTo errHandle

    Dim blnIsNumber As Boolean
    Dim blnIsRepetition As Boolean
    Dim blnIsContextRepetion As Boolean
    Dim lngSpecimenNameIndex As Long
    Dim lngSlicesTypeIndex As Long
    Dim lngCompletionStatus As Long
    Dim strOldSpecimenName As String
    Dim i As Integer
    

    For i = 1 To ufgData.GridRows - 1
        If Trim(ufgData.Text(i, "��Ƭ��ʽ")) <> "" Then
        
            '�ж���Ƭ���� �Ƿ�������  �Ƿ����0  �Ƿ���С��
             If IsNumeric(ufgData.Text(i, "��Ƭ����")) And Val(ufgData.Text(i, "��Ƭ����")) > 0 And InStr(ufgData.Text(i, "��Ƭ����"), ".") < 1 Then
                 blnIsNumber = True
             Else
                 blnIsNumber = False
                 Exit For
             End If
             
             
            '���걾���ƺ���Ƭ��ʽ �Ƿ��ظ�
            If InStr(strOldSpecimenName, ufgData.DisplayText(i, "�걾����") & "," & ufgData.Text(i, "��Ƭ��ʽ")) > 0 Then
                blnIsRepetition = True
                Exit For
            Else
                strOldSpecimenName = strOldSpecimenName & ufgData.DisplayText(i, "�걾����") & "," & ufgData.Text(i, "��Ƭ��ʽ") & "|"
                blnIsRepetition = False
            End If
            
            If mlngRequestType = 3 Then
                 '���걾���ƺ���Ƭ��ʽ ��������Ŀ���Ƿ���� ���ж��Ƿ����� ����ɵļ�¼
                 lngSpecimenNameIndex = mufgParentContextGrid.FindRowIndex(ufgData.DisplayText(i, "�걾����"), "�걾����", True)
                 lngSlicesTypeIndex = mufgParentContextGrid.FindRowIndex(Mid(ufgData.Text(i, "��Ƭ��ʽ"), InStr(ufgData.Text(i, "��Ƭ��ʽ"), "-") + 1, 10), "��Ƭ��ʽ", True)
                 lngCompletionStatus = mufgParentContextGrid.FindRowIndex("�����", "��ǰ״̬", True)
                 
                If lngSpecimenNameIndex > 0 And lngSlicesTypeIndex > 0 And lngCompletionStatus < 1 Then
                    blnIsContextRepetion = True
                    Exit For
                Else
                    blnIsContextRepetion = False
                End If
            Else
                blnIsContextRepetion = False
            End If

        End If
    Next i
  
    If Not blnIsNumber Then
        Call ShowProcessHint("��������Ч����Ƭ������")
        Exit Sub
    End If
    
    If blnIsRepetition Then
        Call ShowProcessHint("�걾���ƻ���Ƭ��ʽ�ظ���")
        Exit Sub
    End If
    
    If blnIsContextRepetion Then
        Call ShowProcessHint("������Ŀ�д����ظ����ݡ�")
        Exit Sub
    End If
    
    
    
    '�ж�������ϸ�б��Ƿ�Ϊ��Ƭ��Ŀ��ϸ�б�
    If mufgParentContextGrid.GetColIndex(gstrRequest_Slices_��Ƭ��ʽ) < 0 Then
        mufgParentContextGrid.ColNames = gstrRequest_Slices_Cols
        mufgParentContextGrid.ColConvertFormat = gstrRequest_SlicesConvertFormat
        
        '�л�������Ŀ��ƽ���
        Call mfrmOwner.ChangeControlFace(3)
    End If
    
    
    '������Ƭ����
    Call SaveSlicesRequest
    
    blnIsOk = True
    
    Call Me.Hide
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo errHandle

    picShow.Visible = False
    txtRequestDoctor.Text = UserInfo.����
    dtpRequestTime.value = zlDatabase.Currentdate
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitRequisitionSlicesList
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ShowProcessHint(ByVal strHint As String)
'��ʾ������Ϣ
On Error Resume Next

    txtShow.Text = strHint

    picShow.Visible = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If Trim(ufgData.Text(Row, "��Ƭ��ʽ")) = "" Then ufgData.Text(Row, "��Ƭ��ʽ") = "1-����"
    If Trim(ufgData.Text(Row, "��Ƭ����")) = "" Then ufgData.Text(Row, "��Ƭ����") = "1"
    
End Sub
