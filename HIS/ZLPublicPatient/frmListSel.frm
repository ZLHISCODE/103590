VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListSel 
   AutoRedraw      =   -1  'True
   Caption         =   "ѡ����"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   ControlBox      =   0   'False
   Icon            =   "frmListSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3525
      Left            =   150
      TabIndex        =   1
      Top             =   480
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   6218
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   2
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   7935
      TabIndex        =   7
      Top             =   0
      Width           =   7935
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ��һ����Ŀ,Ȼ����ȷ��"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   120
         Width           =   2430
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7935
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4110
      Width           =   7935
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   780
         MaxLength       =   6
         TabIndex        =   6
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6555
         TabIndex        =   4
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5325
         TabIndex        =   3
         Top             =   105
         Width           =   1100
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "����(&F)"
         Height          =   180
         Left            =   90
         TabIndex        =   5
         Top             =   210
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1065
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListSel.frx":014A
            Key             =   "Item"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'����
Private mblnOK As Boolean
Private mblnMutilSelect As Boolean
Private mrsSel As ADODB.Recordset
Private mrsReturn As New ADODB.Recordset
Private mstrKey  As String
Private mstrTitle As String
Private mblnHideCancel As Boolean

Private Const M_INT_AdLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const M_INT_AdDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const M_INT_AdDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName, mstrTitle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
End Sub

Public Function ShowSelect(rsSelect As ADODB.Recordset, ByVal strKey As String, _
    Optional ByVal strTitle As String, Optional ByVal strNote As String, _
    Optional ByVal blnMutilSelect As Boolean = False, Optional ByVal blnSerach As Boolean = False, _
    Optional ByVal strMshWidth As String = "", Optional ByVal blnHideCancel As Boolean) As Boolean
'���ܣ��๦��ѡ����
'������
'     frmParent=��ʾ�ĸ�����
'     rsSelect=ѡ��ļ�¼��
'     strKey=���ؼ��ֶ�
'     strTitle=ѡ������������
'     strNote=ѡ��˵��
'     blnMutilSelect=��ѡ��־����������ѡ����̶�Ϊ��һ����ѡ���У�
'     blnSerach=�Ƿ�֧�ֲ���
'     strMshWidth=�п��ַ���,��ʽΪ800|1000��Ϊ���򱣳�Ĭ���п�
'     blnHideCancel-True ����ȡ����ť;False-Ĭ�ϲ�����
    Dim lngIndex As Long
    Dim strValue As String
    Dim lngRow As Long, intCol As Integer
    Dim arrMshWidth() As String
    
    Set mrsSel = rsSelect
    mstrKey = strKey
    mstrTitle = strTitle
    mblnMutilSelect = blnMutilSelect
    mblnOK = False
    mblnHideCancel = blnHideCancel
    
    If rsSelect.RecordCount = 0 Then
        MsgBox "û�п�ѡ�������", vbInformation, gstrSysName
        Exit Function
    End If
    
    '������ͷ
    mshSelect.Clear
    mshSelect.TextMatrix(0, 0) = "Key"
    
    'װ������
    Set mshSelect.DataSource = rsSelect
    
    Call zlControl.MshSetColWidth(mshSelect, Me)
    
    '�����п�
    If strMshWidth <> "" Then
        arrMshWidth = Split(strMshWidth, "|")
        For intCol = 0 To UBound(arrMshWidth)
            mshSelect.ColWidth(intCol) = Val(arrMshWidth(intCol))
        Next
    End If
    
    For intCol = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(intCol).Name = mstrKey And Not (mstrKey Like "*����*") Then
            mshSelect.ColWidth(intCol) = 0
            Exit For
        End If
    Next
    
    mshSelect.Row = 1
    mshSelect.RowSel = 1
    mshSelect.Col = 0
    mshSelect.ColSel = mshSelect.Cols - 1
    
    lblFind.Visible = blnSerach
    txtFind.Enabled = blnSerach
    txtFind.Visible = blnSerach
    Me.Caption = strTitle
    Me.lblInfo.Caption = strNote
    Me.lblInfo.ToolTipText = Me.lblInfo.Caption
    cmdCancel.Visible = Not mblnHideCancel
    
    frmListSel.Show vbModal
    If mblnMutilSelect And mblnOK Then Set rsSelect = mrsReturn
    ShowSelect = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intCol As Integer
    Dim lngRow As Long
    Dim strFilter As String
    Dim strFields As String, strValues As String
    
    If EmptyContent Then Exit Sub
    
    For intCol = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(intCol).Name = mstrKey Then Exit For
    Next
    If intCol > mrsSel.Fields.Count - 1 Then intCol = 0
    
    If mblnMutilSelect = False Then
        If mrsSel.Fields(mstrKey).Type = adVarChar Then
            strFilter = mstrKey & "='" & mshSelect.TextMatrix(mshSelect.Row, intCol) & "'"
        Else
            strFilter = mstrKey & "=" & Val(mshSelect.TextMatrix(mshSelect.Row, intCol))
        End If
        mrsSel.Filter = strFilter
    Else
        '��ʼ����¼��
        strFields = ""
        For intCol = 0 To mrsSel.Fields.Count - 1
            strFields = strFields & "|" & mrsSel.Fields(intCol).Name & "," & adLongVarChar & "," & mrsSel.Fields(intCol).DefinedSize
        Next
        strFields = Mid(strFields, 2)
        Call Record_Init(mrsReturn, strFields)
        
        '���ݴ����¼��������Ӧ�ļ�¼��������
        strFields = ""
        For intCol = 0 To mrsSel.Fields.Count - 1
            strFields = strFields & "|" & mrsSel.Fields(intCol).Name
        Next
        strFields = Mid(strFields, 2)
        
        With mshSelect
            For lngRow = 1 To .Rows - 1
                If Trim(mshSelect.TextMatrix(lngRow, 0)) = "��" Then
                    mrsSel.MoveFirst
                    mrsSel.Move lngRow - 1
                    
                    strValues = ""
                    For intCol = 0 To mrsSel.Fields.Count - 1
                        strValues = strValues & "|" & mrsSel.Fields(intCol).Value
                    Next
                    strValues = Mid(strValues, 2)
                    
                    Call Record_Add(mrsReturn, strFields, strValues)
                End If
            Next
        End With
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With lblInfo
        .Left = 60
        .Top = 60
    End With
    
    With picInfo
        .Left = 0
        .Top = 0
        .width = Me.ScaleWidth
        .Height = lblInfo.Height + 120
    End With
    
    With mshSelect
        .Top = picInfo.Height
        .Left = 60
        .width = Me.ScaleWidth - 120
        .Height = Me.ScaleHeight - picInfo.Height - picCmd.Height
    End With
    
    'If Me.ScaleWidth - cmdCancel.Width * 1.3 >= cmdHelp.Left + cmdHelp.Width * 2 + 300 Then
    If mblnHideCancel Then
        cmdOK.Left = Me.ScaleWidth - cmdCancel.width * 1.3
    Else
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.width * 1.3
        cmdOK.Left = cmdCancel.Left - cmdOK.width * 1.1
    End If
    'End If
End Sub

Private Sub lvw_DblClick()
    Call cmdOK_Click
End Sub

Private Sub mshSelect_DblClick()
    If mblnMutilSelect = False Then
        Call cmdOK_Click
    Else
        If mshSelect.TextMatrix(mshSelect.Row, 0) = "" Then
            mshSelect.TextMatrix(mshSelect.Row, 0) = "��"
        Else
            mshSelect.TextMatrix(mshSelect.Row, 0) = ""
        End If
    End If
End Sub

Private Sub mshSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mshSelect_DblClick
End Sub

Private Sub txtFind_Change()
'���ܣ������û���������ݲ���ƥ�������
    Dim lngIndex As Long, lngRow As Long, lngCol As Long
    Dim strFind As String
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    strFind = strFind & "*"
    If EmptyContent Then Exit Sub
    
    With mshSelect
        For lngRow = 1 To .Rows - 1
            For lngCol = 1 To .Cols - 1
                If .TextMatrix(lngRow, lngCol) Like strFind Then
                    .Row = lngRow
                    .RowSel = lngRow
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                End If
            Next
        Next
    End With
End Sub

Private Function EmptyContent() As Boolean
    Dim intCol As Integer
    
    For intCol = 0 To mrsSel.Fields.Count - 1
        If mrsSel.Fields(intCol).Name = mstrKey Then Exit For
    Next
    If intCol > mrsSel.Fields.Count - 1 Then intCol = 0
    
    With mshSelect
        If .Rows - 1 = 1 Then
            If mrsSel.Fields(mstrKey).Type = adVarChar Then
                If .TextMatrix(1, intCol) = "" Then EmptyContent = True
            Else
                If Val(.TextMatrix(1, intCol)) = 0 Then EmptyContent = True
            End If
        End If
    End With
End Function

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues
    Dim intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intTYPE As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intTYPE = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intTYPE
                Case adDouble
                    lngLength = M_INT_AdDoubleDefault
                Case adVarChar
                    lngLength = M_INT_AdLongVarCharDefault
                Case adLongVarChar
                    lngLength = M_INT_AdLongVarCharDefault
                Case Else
                    lngLength = M_INT_AdDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intTYPE, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub


Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdOK_Click
End Sub
