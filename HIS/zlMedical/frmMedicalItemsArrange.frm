VERSION 5.00
Begin VB.Form frmMedicalItemsArrange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����Ŀ����"
   ClientHeight    =   6135
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8910
   Icon            =   "frmMedicalItemsArrange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame4 
      Height          =   6135
      Left            =   45
      TabIndex        =   3
      Top             =   -45
      Width           =   7515
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   0
         Left            =   7095
         Picture         =   "frmMedicalItemsArrange.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "��ǰ�ƶ�"
         Top             =   180
         Width           =   345
      End
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   1
         Left            =   7095
         Picture         =   "frmMedicalItemsArrange.frx":0159
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "����ƶ�"
         Top             =   570
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   5895
         Left            =   75
         TabIndex        =   4
         Top             =   150
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   10398
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7650
      TabIndex        =   2
      Top             =   45
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7650
      TabIndex        =   1
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7650
      TabIndex        =   0
      Top             =   1380
      Width           =   1100
   End
End
Attribute VB_Name = "frmMedicalItemsArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean

Private Enum mCol
    ���� = 1
    ����
    ��λ
    ���
End Enum

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    If vData = False Then
        cmdOK.Tag = ""
    Else
        cmdOK.Tag = "Changed"
    
    End If
End Property

Private Property Get EditChanged() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
            
    EditChanged = (cmdOK.Tag = "Changed")
    
End Property


Public Function ShowEdit(ByVal frmMain As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
                
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    
    Call ReadData
    
            
    EditChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
        
    
    gstrSQL = "Select a.ID,a.����,a.����,A.���㵥λ As ��λ,DECODE(A.���,'C','����','���') AS ��� " & _
                    "From ������ĿĿ¼ a,�����Ŀ���� b where b.������Ŀid=a.ID And b.��������=1  Order By b.����˳�� "
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID"), 0)
            vsf.TextMatrix(vsf.Rows - 1, mCol.����) = zlCommFun.NVL(rs("����"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.����) = zlCommFun.NVL(rs("����"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.��λ) = zlCommFun.NVL(rs("��λ"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.���) = zlCommFun.NVL(rs("���"))
                        
            rs.MoveNext
        Loop
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand

    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "����", 2700, 1, "...", 1
        .NewColumn "����", 1200, 1
        .NewColumn "��λ", 900, 1
        .NewColumn "���", 900, 1
        
        .FixedCols = 1
        
        .SelectMode = True
    End With
        
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strSQL(ReDimArray(strSQL)) = "ZL_�����Ŀ����_DELETE(1)"
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            gstrSQL = "ZL_�����Ŀ����_INSERT("
            gstrSQL = gstrSQL & Val(vsf.RowData(lngLoop)) & ","
            gstrSQL = gstrSQL & lngLoop & ",1)"
            
            strSQL(ReDimArray(strSQL)) = gstrSQL
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����Ƿ����ظ�����Ŀ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey And vsf.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        If vsf.Row > 1 Then
            
            Call MoveItem(vsf.Row, -1)
            vsf.Row = vsf.Row - 1
            cmdOK.Tag = "Changed"
            
        End If
    ElseIf vsf.Row < vsf.Rows - 1 Then
        
        Call MoveItem(vsf.Row, 1)
        vsf.Row = vsf.Row + 1
        cmdOK.Tag = "Changed"
        
    End If
    
    vsf.ShowCell vsf.Row, vsf.Col
    vsf.SetFocus
End Sub

Private Function MoveItem(ByVal intCurRow As Integer, Optional ByVal intMove As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim intCol As Integer
    
    On Error GoTo errHand
    
    strTmp = CStr(vsf.RowData(intCurRow))
            
    vsf.RowData(intCurRow) = vsf.RowData(intCurRow + intMove)
    vsf.RowData(intCurRow + intMove) = Val(strTmp)
    
    For intCol = 0 To vsf.Cols - 1
        
        strTmp = vsf.TextMatrix(intCurRow, intCol)
        
        vsf.TextMatrix(vsf.Row, intCol) = vsf.TextMatrix(intCurRow + intMove, intCol)
        
        vsf.TextMatrix(intCurRow + intMove, intCol) = strTmp
        
    Next
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
       
    If EditChanged Then
    
        If SaveEdit Then
            mblnOK = True
            
            EditChanged = False
        End If
        
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Select Case Col
        Case mCol.����
        
            gstrSQL = GetPublicSQL(SQL.�����Ŀѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
            If ShowGrdSelect(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 4500) Then
                'ѡȡ��һ����Ŀ
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
                
                vsf.Cell(flexcpText, Row, mCol.���� + 1, Row, vsf.Cols - 1) = ""
                
                vsf.EditText = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.��λ) = zlCommFun.NVL(rs("��λ").Value)
                vsf.TextMatrix(Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                EditChanged = True
                
            End If

    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
    Dim rsData As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.����
                    
                    strText = UCase(vsf.EditText)
                    
                    gstrSQL = GetPublicSQL(SQL.�����Ŀ����ѡ��, strText)
                    
                    strText = strText & "%"
                    If ParamInfo.��Ŀ����ƥ�䷽ʽ = 0 Then strTmp = "%" & strText
                                
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText, strTmp, 1, 2)
                    If ShowGrdFilter(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀ����", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100) Then
                                                
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                        
                        vsf.EditText = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.��λ) = zlCommFun.NVL(rs("��λ").Value)
                        vsf.TextMatrix(Row, mCol.����) = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
                        vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        EditChanged = True
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                        vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        EditChanged = True
    End If
End Sub


