VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm������Ŀѡ����� 
   AutoRedraw      =   -1  'True
   Caption         =   "ҽ����Ŀѡ��"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm������Ŀѡ�����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7845
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000A&
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7785
      TabIndex        =   10
      Top             =   4890
      Visible         =   0   'False
      Width           =   7845
      Begin MSComctlLib.ProgressBar prgs 
         Height          =   450
         Left            =   795
         TabIndex        =   11
         Top             =   45
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblInfor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1575
      Width           =   45
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshGrid 
      Height          =   3990
      Left            =   3045
      TabIndex        =   6
      Top             =   390
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   7038
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ�����.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������Ŀѡ�����.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   7
      Top             =   255
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7144
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   7845
      TabIndex        =   1
      Top             =   4980
      Width           =   7845
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   5
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   4
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "��ӡ�б�"
         Height          =   350
         Left            =   15
         TabIndex        =   3
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdRequery 
         Caption         =   "��Ŀ����"
         Height          =   350
         Left            =   1335
         TabIndex        =   2
         ToolTipText     =   "���������ط�����Ŀ��������Ϣ�Ͷ���ҽ�ƻ���"
         Top             =   180
         Width           =   1100
      End
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ����(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   9
      Top             =   0
      Width           =   2970
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "��Ŀ��ϸ(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   8
      Top             =   15
      Width           =   4710
   End
End
Attribute VB_Name = "frm������Ŀѡ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint���� As Integer
Private mstrCode As String
Private mstrName As String
Private mblnOK As Boolean

Private mLocalCode As String 'ָ�����
Private mblnFirst As Boolean
Private mbln���� As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(mshGrid.TextMatrix(mshGrid.Row, 0)) = "" Then
        MsgBox "û��ѡ����Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '����ѡ����Ŀ����
    mstrCode = mshGrid.TextMatrix(mshGrid.Row, 0) & Trim(mshGrid.TextMatrix(mshGrid.Row, 1))
    mstrName = mshGrid.TextMatrix(mshGrid.Row, 2)
    mblnOK = True
    Unload Me
End Sub

Private Function Loadtree() As Boolean
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim tmpNode As Node
    mblnOK = False
    
    On Error GoTo ErrHand:
    
    'װ������
    '
    tvwClass.Nodes.Clear
    Set tmpNode = tvwClass.Nodes.Add(, 4, "K1", "��1��ҩƷ", "Detail", "Detail")
    tmpNode.Sorted = True
    tmpNode.Selected = True
    
    Set tmpNode = tvwClass.Nodes.Add(, 4, "K2", "��2������", "Detail", "Detail")
    tmpNode.Sorted = True
    
    Set tmpNode = tvwClass.Nodes.Add(, 4, "K3", "��4������", "Detail", "Detail")
    tmpNode.Sorted = True
    
    'Call FillList
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    Loadtree = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Loadtree = False
End Function
Public Function GetCode(ByVal frmMain As Form, strCode As String, strName As String, Optional bln���� As Boolean = False) As Boolean
    '���ܣ���ȡ����
    '������strCode-����(���+����)
    '���أ��ɹ�����True
    mLocalCode = strCode
    frm������Ŀѡ�����.Show vbModal, frm������Ŀ
    
    '����ֵ
    If mblnOK = True Then
        strCode = mstrCode
        strName = mstrName
    End If
    GetCode = mblnOK
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetGrdColHead(Optional ByVal blnInit As Boolean = True)
    With mshGrid
        .Redraw = False
        If blnInit Then
            .Clear
            .Rows = 2
            .Cols = 15
            .TextMatrix(0, 0) = "���"
            .TextMatrix(0, 1) = "����"
            .TextMatrix(0, 2) = "����"
            .TextMatrix(0, 3) = "Ӣ������"
            .TextMatrix(0, 4) = "�շ����"
            .TextMatrix(0, 5) = "�շѵȼ�"
            .TextMatrix(0, 6) = "������"
            .TextMatrix(0, 7) = "��λ"
            .TextMatrix(0, 8) = "��׼�۸�"
            .TextMatrix(0, 9) = "֧����׼"
            .TextMatrix(0, 10) = "����"
            .TextMatrix(0, 11) = "���"
            .TextMatrix(0, 12) = "��ע"
            .TextMatrix(0, 13) = "���ʱ��"
            .TextMatrix(0, 14) = "ά����־"
        End If
        .ColWidth(0) = 0
        .ColWidth(1) = 1500
        .ColWidth(2) = 2000
        .ColWidth(3) = 1500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1400
        .ColWidth(11) = 1400
        .ColWidth(12) = 2000
        .ColWidth(13) = 1600
        .ColWidth(14) = 1000
        
        .ColAlignment(0) = 0
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        .ColAlignment(6) = 4
        .ColAlignment(8) = 4
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7
        .ColAlignment(11) = 1
        .ColAlignment(12) = 1
        .ColAlignment(13) = 4
        .ColAlignment(14) = 4
        .Redraw = True
End With

End Sub
Private Sub FillList()
    '���ܣ���ʾ��ǰ����µ�ҽ����ϸ
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, fld As ADODB.Field
    Dim str������ As String, blnColSet As Boolean
    Dim lngCol  As Long
    Dim varValue As Variant
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo errHandle
    With tvwClass.SelectedItem
        str������ = Mid(.Key, 2)
    End With
    
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = " select  ���,����,����,Ӣ������,�շ����,�շѵȼ�,������,��λ,��׼�۸�,֧����׼,����,���,��ע,���ʱ��,ά����־ " & _
             "  from ҽ���շ�Ŀ¼" & _
             "  where ���=" & Val(str������)
    
    rsTemp.Open gstrSQL, gcnOracle_����, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        '������ͷ
        Call SetGrdColHead
    Else
        Set mshGrid.DataSource = rsTemp
        Call SetGrdColHead(False)
    End If
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdPrint_Click()
    If gstrUserName = "" Then Call GetUserInfo
    subPrint 1
End Sub

Private Sub subPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwClass.SelectedItem
    Set objPrint.Body = mshGrid
    objPrint.Title.Text = "������Ŀ"
    
    objRow.Add "ҽ�����ࣺ" & nod.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & gstrUserName
    objRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Private Sub cmdRequery_Click()
    Dim strInput As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln���� As Boolean
    
    If MsgBox("���������ܻỨ�Ƚϳ���ʱ�䣬�Ƿ������" & vbCrLf & vbCrLf & "����ע�⣬������ֻ����ҽ����Ŀ��ϸ������������Ӧ��ϵ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
        
    MousePointer = vbHourglass
 
    picCmd.Enabled = False
    tvwClass.Enabled = False
        
    cmdRequery.Visible = False
    cmdCancel.Enabled = False
    cmdPrint.Visible = False
        
    With picBack
        .Left = 0
        .Width = ScaleWidth
        .Top = ScaleHeight - .Height
        picBack.Visible = True
    End With
    
    '���ط�����ĿĿ¼
    '1.����ҩƷ
    lblInfor.Caption = "ҩƷ"
    If ���ط�����ĿĿ¼_����(1, prgs) = False Then
        GoTo GoEnd:
    End If
    '2.��������
    lblInfor.Caption = "����"
    If ���ط�����ĿĿ¼_����(2, prgs) = False Then
        GoTo GoEnd:
    End If
    '3.���ط���
    lblInfor.Caption = "����"
    If ���ط�����ĿĿ¼_����(3, prgs) = False Then
        GoTo GoEnd:
    End If
    '4.���ط������
    lblInfor.Caption = "�������"
    If ���ط�����ĿĿ¼_����(4, prgs) = False Then
       GoTo GoEnd:
    End If
    '5.���ز���
    lblInfor.Caption = "����"
    If ���ط�����ĿĿ¼_����(5, prgs) = False Then
        GoTo GoEnd:
    End If
GoEnd:
    MousePointer = vbDefault
    picCmd.Enabled = True
    tvwClass.Enabled = True
    picBack.Visible = False
    cmdRequery.Visible = True
    cmdCancel.Enabled = True
    cmdPrint.Visible = True
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If Loadtree = False Then
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = tvwClass.Width
    
    On Error Resume Next
    
    tvwClass.Left = 0: tvwClass.Top = lblClass.Top + lblClass.Height
    tvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = tvwClass.Top
    picSplit.Left = tvwClass.Left + tvwClass.Width
    picSplit.Height = tvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If tvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    With picBack
        .Left = 0
        .Width = ScaleWidth
    End With
    With mshGrid
        .Top = tvwClass.Top
        .Left = lblDetail.Left
        .Width = lblDetail.Width
        .Height = tvwClass.Height
    End With
End Sub

Private Sub picBack_Resize()
    Err = 0
    On Error Resume Next
    With prgs
        .Left = lblInfor.Left + lblInfor.Width
        .Width = picBack.ScaleWidth - .Left
    End With
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
    cmdRequery.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshgrid_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvwClass.Width + x < 1000 Or mshGrid.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        mshGrid.Left = mshGrid.Left + x
        mshGrid.Width = mshGrid.Width - x
    End If
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
End Sub







