VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm������Ϣ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   ControlBox      =   0   'False
   Icon            =   "frm������Ϣ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtEdit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3090
      TabIndex        =   6
      Top             =   1410
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   780
      Width           =   5325
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3105
      Width           =   5325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4140
      TabIndex        =   2
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2865
      TabIndex        =   1
      Top             =   3255
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshBill 
      Height          =   2100
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   930
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   3704
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      FixedCols       =   2
      BackColorSel    =   4194304
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label lbl 
      Caption         =   "����Ϊҽ�����˱��ν���������Ϣ��"
      Height          =   225
      Index           =   0
      Left            =   810
      TabIndex        =   5
      Top             =   450
      Width           =   4965
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   105
      Picture         =   "frm������Ϣ.frx":000C
      Stretch         =   -1  'True
      Top             =   165
      Width           =   525
   End
End
Attribute VB_Name = "frm������Ϣ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mblnOK As Boolean
Private mblnYesNo As Boolean
Private mstr���㷽ʽ As String
Private mdbl�ܷ� As Double
Private mblnChange As Boolean '�Ƿ���Ĺ�ֵ
Private mbytType    As Byte     'mbytType-0����Һ�,1סԺ
Private mblnLoad    As Boolean
Private Sub cmdCancel_Click()
    mblnOK = False
    If mblnYesNo = False Then
        mblnOK = True
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    mblnOK = True
    Dim str���㷽ʽ As String
'    If mstr���㷽ʽ <> "" And mblnChange = True Then
'            With mshBill
'                For i = 1 To .Rows - 1
'                    If Trim(.TextMatrix(i, 1)) <> "" And Trim(.TextMatrix(i, 1)) <> "�ֽ�" Then
'                        str���㷽ʽ = str���㷽ʽ & "||" & .TextMatrix(i, 1) & " |" & .TextMatrix(i, 2)
'                    End If
'                Next
'            End With
'            If str���㷽ʽ <> "" Then
'                '����Ԥ����¼
'                str���㷽ʽ = Mid(str���㷽ʽ, 3)
'                If mbytType = 0 Then
'                    gstrSQL = "zl_���˽����¼_Update(" & mlng����ID & ",'" & str���㷽ʽ & "',0)"
'                    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
'                Else
'                    gstrSQL = "zl_���˽����¼_Update(" & mlng����ID & ",'" & str���㷽ʽ & "',1)"
'                    Call zlDatabase.ExecuteProcedure(gstrSQL, "����Ԥ����¼")
'                End If
'            End If
'    End If
    Unload Me
End Sub

'Modified by ���� 20031218 ���������� ��������
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strArr
    Dim i As Long
    
    
    DebugTool "������ش���"
    strArr = Split(mstr���㷽ʽ, "|")
    gstrSQL = "Select Decode(A.��¼����,1,'��Ԥ��',11,'��Ԥ��',A.���㷽ʽ) ���㷽ʽ,Nvl(A.��Ԥ��,0) ��� " & _
                " From ����Ԥ����¼ A,�����ʻ� B " & _
                " Where A.����ID=B.����ID And A.����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ν��׽�����Ϣ", mlng����ID)
    DebugTool "������ش���-�򿪼�¼��"
    
    With mshBill
        .Clear
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 0) = "�����޶�"
        .TextMatrix(0, 1) = "���㷽ʽ"
        .TextMatrix(0, 2) = "���"
            
        .ColWidth(0) = 0
        .ColWidth(1) = 2000
        .ColWidth(2) = 1200
        
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
    End With
        
    mdbl�ܷ� = 0
    With rsTemp
        Do While Not .EOF
            For i = 0 To UBound(strArr)
                mshBill.RowData(.AbsolutePosition) = 0
                If Split(strArr(i), ":")(0) = Nvl(!���㷽ʽ) Then
                    If InStr(1, strArr(i), ":") <> 0 Then
                        mshBill.TextMatrix(.AbsolutePosition, 0) = Val(Split(strArr(i), ":")(1))
                        mshBill.RowData(.AbsolutePosition) = 1
                    End If
                    Exit For
                End If
            Next
            mshBill.TextMatrix(.AbsolutePosition, 1) = !���㷽ʽ
            mshBill.TextMatrix(.AbsolutePosition, 2) = Format(!���, "#####0.00;-#####0.00; ;")
            mdbl�ܷ� = mdbl�ܷ� + Nvl(!���, 0)
            If mshBill.Rows - 1 = .AbsolutePosition Then mshBill.Rows = mshBill.Rows + 1
            .MoveNext
        Loop
        If Trim(mshBill.TextMatrix(mshBill.Rows - 1, 0)) = "" Then mshBill.Rows = mshBill.Rows - 1
    End With
    DebugTool "������ش������"
End Sub

Public Function ShowME(Optional ByVal lng����ID As Long = 0, Optional blnYesNo As Boolean = False, Optional str���㷽ʽ As String = "", Optional bytType As Byte = 0) As Boolean
    'blnYesNO:�����Ƿ��ṩȷ����ȡ��ѡ��.
    'str���㷽ʽ-��ĳ����㷽ʽ���и���,��ʽ�����㷽ʽ:���ƶ�|���㷽ʽ:���ƶ�,��:�����ʻ�:20����ʾ�����ʻ����Ը��ģ������ܳ���20Ԫ
    'bytType-0-����Һ�,1סԺ
    
    mlng����ID = lng����ID
    mblnYesNo = blnYesNo
    mstr���㷽ʽ = str���㷽ʽ
    mbytType = bytType
    Me.cmdOK.Visible = blnYesNo
    If blnYesNo = False Then
        Me.cmdCancel.Caption = "ȷ��(&O)"
    End If
    DebugTool "�Ѿ�������㷽ʽshowme"
    frm������Ϣ.Show 1
    DebugTool "��ɽ��㷽ʽshowme"
    ShowME = mblnOK
End Function

Private Sub MshBill_DblClick()
    '���и�����ص�����
    With mshBill
        If .RowData(.Row) = 0 Then txtEdit.Visible = False: Exit Sub
        .COL = 2
        mblnLoad = True
        txtEdit.Left = .Left + .CellLeft + 15
        txtEdit.Top = .Top + .CellTop + 15
        txtEdit.Height = .CellHeight - 30
        txtEdit.Width = .CellWidth - 30
        txtEdit.Visible = True
        txtEdit.Tag = .TextMatrix(.Row, 0)
        txtEdit.Text = .TextMatrix(.Row, 2)
        txtEdit.SetFocus
    End With
End Sub

Private Sub txtEdit_Change()
    If mblnLoad Then Exit Sub
    mblnChange = True
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnLoad = False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m���ʽ
End Sub

Private Sub txtEdit_LostFocus()
    
    txtEdit.Text = Format(Abs(Val(txtEdit.Text)), "#####0.00;-####0.00; ;")
    If Val(txtEdit.Tag) < Val(txtEdit.Text) Then
        ShowMsgbox "�����ֵ���ܴ���" & Format(Val(txtEdit.Tag), "#####0.00;-####0.00; ;")
        Exit Sub
    End If
    If mdbl�ܷ� < Val(txtEdit.Text) Then
        ShowMsgbox "�����ֵ���ܴ����ܷ���" & Format(mdbl�ܷ�, "#####0.00;-####0.00; ;")
        Exit Sub
    End If
    mshBill.TextMatrix(mshBill.Row, 2) = txtEdit.Text
    Call ��������
End Sub
Private Sub ��������()
    Dim intRow As Integer
    Dim dblTemp As Double
    Dim i As Integer
    
    intRow = 0
    With mshBill
        dblTemp = mdbl�ܷ�
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 1)) = "�ֽ�" Then
                intRow = i
            Else
                dblTemp = dblTemp - Val(.TextMatrix(i, 2))
            End If
        Next
        If intRow <> 0 Then
            .TextMatrix(intRow, 2) = Format(dblTemp, "####0.00;-####0.00; ;")
        End If
    End With
End Sub

