VERSION 5.00
Begin VB.Form frmPatholSlices_BatSure 
   Caption         =   "����ȷ��"
   ClientHeight    =   7008
   ClientLeft      =   72
   ClientTop       =   408
   ClientWidth     =   11484
   Icon            =   "frmPatholSlices_BatSure.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7008
   ScaleWidth      =   11484
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame framSureRecord 
      Caption         =   "��ȷ�ϼ�¼��"
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   9855
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   3615
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9615
         _ExtentX        =   16955
         _ExtentY        =   6371
         DefaultCols     =   ""
         IsKeepRows      =   0   'False
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   612
      ScaleWidth      =   9612
      TabIndex        =   9
      Top             =   5640
      Width           =   9615
      Begin VB.CommandButton cmdBatSure 
         Caption         =   "��ʼȷ��(&S)"
         Height          =   400
         Left            =   7080
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "�� ��(&B)"
         Height          =   400
         Left            =   8400
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label labRecordInf 
         Caption         =   "����Ƭ������0    ��ȷ��������0"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Frame framFilter 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9855
      Begin VB.OptionButton optUserCodeBar 
         Caption         =   "ʹ�������"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "ȷ ��(&S)"
         Height          =   400
         Left            =   5760
         TabIndex        =   2
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtSureCount 
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Text            =   "1"
         ToolTipText     =   "���������������ȷ�ϵĲ�Ƭ������"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSureNum 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "��δ��������ʱ����������ֱ�����롰����š����ҡ�"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "��Ƭ������"
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ȷ�Ϻ��룺"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   435
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPatholSlices_BatSure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mufgParentGrid As ucFlexGrid
Private mstrCurPatholNum As String

Public blnIsOk As Boolean


Public Sub ShowSlicesSureWindow(ufgParentGrid As ucFlexGrid, ByVal strPatholNum As String, owner As Form)
'��ʾ��Ƭȷ�ϴ���
    Set mufgParentGrid = ufgParentGrid
    
    mstrCurPatholNum = strPatholNum
    blnIsOk = False
    
    Call Me.Show(1, owner)
End Sub



Private Sub RefreshSureCount()
'ˢ��ȷ�ϵ�������Ϣ
    Dim i As Long
    Dim iNeedCount As Long
    Dim iSureCount As Long
    
    iNeedCount = 0
    iSureCount = 0
    
    For i = 1 To ufgData.GridRows - 1
        iNeedCount = iNeedCount + Val(ufgData.Text(i, gstrSlicesSure_����Ƭ��))
        iSureCount = iSureCount + Val(ufgData.Text(i, gstrSlicesSure_��ȷ����))
    Next i
    
    labRecordInf.Caption = "����Ƭ������" & iNeedCount & "    ��ȷ��������" & iSureCount
End Sub


Private Sub DecodeSureNum(ByVal strSureNum As String, ByRef strPatholNum As String, ByRef strSlicesId As String)
'�ֽ�ȷ�Ϻ���
    Dim lngFindSplitChar As Long
    
    If optUserCodeBar.value Then
        strPatholNum = ""
        strSlicesId = Trim(strSureNum)
    Else
        strPatholNum = Trim(strSureNum)
        strSlicesId = ""
    End If
    
'    lngFindSplitChar = InStr(1, strSureNum, "-")
'
'    If lngFindSplitChar > 0 Then
'        strPatholNum = Mid(strSureNum, 1, lngFindSplitChar - 1)
'        strSlicesId = Mid(strSureNum, lngFindSplitChar + 1, 20)
'    Else
'        strPatholNum = strSureNum
'        strSlicesId = ""
'    End If
End Sub



Private Sub SureSlices(ByVal strSureNum As String)
'ȷ����Ƭ�������
'strSureNum�����ʽΪ�������-��Ƭ�š�
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngFindRow As Integer
    Dim strPatholNum As String
    Dim strSlicesId As String
    
    strPatholNum = ""
    strSlicesId = ""
    
    Call DecodeSureNum(strSureNum, strPatholNum, strSlicesId)
    
    lngFindRow = ufgData.FindRowIndex(strSlicesId, gstrSlicesSure_ID)
    If lngFindRow > 0 Then GoTo errFindSlices
    
    
    If Trim(strPatholNum) = "" Then
        strSql = "select ����� from ��������Ϣ where ����ҽ��ID = (select ����ҽ��ID from ������Ƭ��Ϣ where ID=[1] and rownum = 1)"
        'If mblnMoved Then strSql = GetMovedDataSql(strSql)
        
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(strSlicesId))
        
        If rsData.RecordCount > 0 Then
            strPatholNum = Val(Nvl(rsData!�����))
        End If
    End If
    
    
    If Trim(strPatholNum) = "" Then
        Call MsgBoxD(Me, "����ĺ�����Ч�����ܸ��ݴ˺����ҵ���Ӧ�Ĳ���š�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    
    '������б����Ҳ�����ǰ¼��Ĳ���ţ�������ݿ��ж�ȡ�����Ϣ�����ص��б���
    lngFindRow = ufgData.FindRowIndex(strPatholNum, ufgData.KeyName)
    If lngFindRow <= 0 Then

        strSql = "Select a.id, c.�����, c.����ҽ��ID, e.����, 0 as ��ȷ����,c.�������, a.��Ƭ����,a.��Ƭ��ʽ,a.�Ŀ�ID,b.���, d.�걾����, " & _
                        " a.��ǰ״̬, case a.��ǰ״̬ when 2 then 0 else a.��Ƭ�� end as ����Ƭ��" & _
                        " From ������Ƭ��Ϣ A, ����ȡ����Ϣ B, ��������Ϣ C, ����걾��Ϣ D, ����ҽ����¼ E " & _
                        " Where a.�Ŀ�id = b.�Ŀ�id And b.����ҽ��ID = c.����ҽ��ID And c.ҽ��id = e.Id And b.�걾id = d.�걾id and c.�����=upper([1]) and a.��ǰ״̬<> 2" & _
                        " order by b.���,a.id"
        'If mblnMoved Then strSql = GetMovedDataSql(strSql)
        
        '��ѯ����
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatholNum)
        
        If rsData.RecordCount > 0 Then
            Set ufgData.AdoData = rsData
            Call ufgData.RefreshData(False)
        End If
    End If
    
errFindSlices:
    '������Ҫȷ�ϵ���Ƭ��¼
    lngFindRow = ufgData.FindRowIndex(strSlicesId, gstrSlicesSure_ID)
    If lngFindRow > 0 Then
        ufgData.Text(lngFindRow, gstrSlicesSure_��ȷ����) = ufgData.Text(lngFindRow, gstrSlicesSure_��ȷ����) + Val(txtSureCount.Text)
        
        If ufgData.Text(lngFindRow, gstrSlicesSure_��ȷ����) = ufgData.Text(lngFindRow, gstrSlicesSure_����Ƭ��) Then
            ufgData.CellColor(lngFindRow, ufgData.GetColIndex(gstrSlicesSure_��ȷ����)) = &HC0FFC0
        End If
        
        Call ufgData.LocateRow(lngFindRow)
    End If
    
End Sub





Private Sub InitSureList()

    '��������
    ufgData.GridRows = glngStandardRowCount
    '�����и�
    ufgData.RowHeightMin = glngStandardRowHeight
    
    '��ʼ��ȷ��������ʾ�б�
    ufgData.ColConvertFormat = gstrSlicesSureConvertFormat
    ufgData.DefaultColNames = gstrSlicesSureColsWithMaterialNum
    ufgData.ColNames = gstrSlicesSureColsWithMaterialNum

End Sub




Private Sub AdjustFace()
'�������沼��
    framFilter.Left = 120
    framFilter.Top = 120
'    framFilter.Width = Me.Width - 360
    
    
    
    framSureRecord.Left = 120
    framSureRecord.Top = framFilter.Top + framFilter.Height + 30
    framSureRecord.Width = Me.Width - 360
    framSureRecord.Height = Me.Height - framFilter.Height - picControl.Height - 680
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framSureRecord.Width - 240
    ufgData.Height = framSureRecord.Height - 360
    
    
    
    
    picControl.Left = 120
    picControl.Top = framSureRecord.Top + framSureRecord.Height + 120
    picControl.Width = Me.Width - 360
    
    
    cmdExit.Left = picControl.Width - cmdExit.Width
    cmdExit.Top = 0
    
    
    cmdBatSure.Left = cmdExit.Left - cmdBatSure.Width - 120
    cmdBatSure.Top = 0
    
    
    labRecordInf.Left = 0
    labRecordInf.Top = cmdBatSure.Top + 60
    
End Sub


Private Sub StartBatSure()
'��ʼ����ȷ��
    Dim i As Long
    Dim strSql As String
    Dim dtServicesTime As Date
    Dim strSurePatholNum As String
    Dim lngRowCheck As Long
    Dim blnUpdateParentGrid As Boolean
    
    
    dtServicesTime = zlDatabase.Currentdate
    
    blnUpdateParentGrid = False
    
    lngRowCheck = ufgData.GetColIndexWithRowCheck()
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, lngRowCheck) Then
            
            If strSurePatholNum <> ufgData.KeyValue(i) Then
                strSurePatholNum = ufgData.KeyValue(i)
                
                strSql = "Zl_������Ƭ_ȷ��('" & strSurePatholNum & "'," & zlStr.To_Date(dtServicesTime) & ")"

                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            If strSurePatholNum = mstrCurPatholNum Then
                blnUpdateParentGrid = True
            End If
            
            ufgData.Text(i, gstrSlicesSure_��ǰ״̬) = "�����"
            ufgData.Text(i, gstrSlicesSure_ȷ��״̬) = "��ȷ��"
        End If
    Next i
    
    '���µ��ý����б��е�ȷ��״̬
    If blnUpdateParentGrid And Not (mufgParentGrid Is Nothing) Then
        For i = 1 To mufgParentGrid.GridRows - 1
            If mufgParentGrid.Text(i, gstrSlicesSure_��ǰ״̬) = "�ѽ���" Then
                 mufgParentGrid.Text(i, gstrSlices_��ǰ״̬) = "�����"
                 mufgParentGrid.Text(i, gstrSlices_��Ƭʱ��) = dtServicesTime
            End If
        Next i
    End If
End Sub



Private Function IsAllowSure() As Boolean
'�ж�ȷ���б��У��Ƿ����������Ƭȷ��
    Dim i As Long
    Dim lngRowCheck As Long
    
    IsAllowSure = True
    
    lngRowCheck = ufgData.GetColIndexWithRowCheck()
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, lngRowCheck) Then
            If ufgData.Text(i, gstrSlicesSure_����Ƭ��) <> ufgData.Text(i, gstrSlicesSure_��ȷ����) Or _
                Val(ufgData.Text(i, gstrSlicesSure_����Ƭ��)) <= 0 Or _
                Val(ufgData.Text(i, gstrSlicesSure_��ȷ����)) <= 0 Then
                
                Call ufgData.LocateRow(i)
                IsAllowSure = False
                
                Exit Function
            End If
        End If
    Next i

End Function


Private Sub cmdBatSure_Click()
'ִ������ȷ��
On Error GoTo errHandle
    
    If Not ufgData.IsCheckedRow Then
        Call MsgBoxD(Me, "��ѡ����Ҫȷ�ϵ���Ƭ��Ϣ��", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    If Not IsAllowSure() Then
        Call MsgBoxD(Me, "��⵽Ҫȷ�ϵ���Ƭ����ʵ�ʵ�����Ƭ����һ�»����������󣬲��ܽ���ȷ�ϡ�", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call StartBatSure
    
    Call RefreshSureCount
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "����ȷ���Ѵ�����ɡ�", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
    blnIsOk = False
    Call Me.Hide
End Sub

Private Sub cmdSure_Click()
'��Ƭȷ��
On Error GoTo errHandle
    Call SureSlices(txtSureNum.Text)
    
    Call RefreshSureCount
    
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim strValue As String
    
    Call RestoreWinState(Me, App.ProductName)

    Call InitSureList
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub



Private Sub txtSureCount_GotFocus()
On Error Resume Next
    
    txtSureCount.SelStart = 0
    txtSureCount.SelLength = Len(txtSureCount.Text)
End Sub

Private Sub txtSureCount_KeyPress(KeyAscii As Integer)
'������»س��������ȷ��
    If KeyAscii = 13 Then
        Call cmdSure_Click
    End If
End Sub


Private Sub txtSureNum_GotFocus()
On Error Resume Next
    
    txtSureNum.SelStart = 0
    txtSureNum.SelLength = Len(txtSureNum.Text)
End Sub

Private Sub txtSureNum_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
'    '��ʱ���дſ���ȡ����
'
'    Dim blnCard As Boolean
'    Dim rsData As ADODB.Recordset
'
'     If KeyAscii = 13 Then
'
'        txtSureCount.SetFocus
'
'        Exit Sub
'    End If
'
'
'    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
'
'    '�ж��Ƿ�Ϊˢ��
'    blnCard = zlCommFun.InputIsCard(txtSureNum, KeyAscii, glngSys)
'    If blnCard And Len(txtSureNum.Text) = mbyt�ſ� - 1 And KeyAscii <> 8 Then
'
'        txtSureNum.Text = txtSureNum.Text & Chr(KeyAscii)
'        txtSureNum.SelStart = Len(txtSureNum.Text)
'
'        KeyAscii = 0
'
'        txtSureNum.SelStart = 0
'        txtSureNum.SelLength = Len(txtSureNum.Text)
'
'        Call SureSlices(txtSureNum.Text)
'        Call RefreshSureCount
'    End If

    '����ɨ��ȷ��
    If KeyAscii = 13 Then
    
        Call SureSlices(txtSureNum.Text)
        Call RefreshSureCount
       
        txtSureNum.SelStart = 0
        txtSureNum.SelLength = Len(txtSureNum.Text)
    
       Exit Sub
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    If ufgData.Text(Row, gstrSlicesSure_��ȷ����) = ufgData.Text(Row, gstrSlicesSure_����Ƭ��) Then
        ufgData.CellColor(Row, ufgData.GetColIndex(gstrSlicesSure_��ȷ����)) = &HC0FFC0
    End If
    
    Call RefreshSureCount
End Sub

