VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetExpense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ѱ�����"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   Icon            =   "frmSetExpense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picҩƷ 
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   3435
      TabIndex        =   18
      Top             =   2040
      Width           =   3495
      Begin MSComctlLib.TreeView tvwDetails 
         Height          =   3480
         Left            =   0
         TabIndex        =   19
         Tag             =   "1000"
         Top             =   240
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   6138
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgTvw"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   10080
      TabIndex        =   17
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����(&O)"
      Height          =   350
      Left            =   8760
      TabIndex        =   16
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   120
      TabIndex        =   15
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "���(&D)"
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      Top             =   7680
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Frame fraҩƷӦ�� 
      Caption         =   "Ӧ�÷�Χ"
      Height          =   1620
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   7335
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڱ����������ҩƷ(&5)"
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   13
         Top             =   1080
         Width           =   2955
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ����ͬ��������ҩƷ(&4)"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   12
         Top             =   720
         Width           =   3075
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڱ�Ʒ��������ҩƷ(&1)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2835
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ�������С�Ƭ������ҩƷ(&3)"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   3915
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ�������С�����ҩ��(&2)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2715
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "��Ӧ���ڱ����ҩƷ(&0)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2595
      End
   End
   Begin VB.ComboBox cbo���㷽�� 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.ComboBox cbo�ѱ� 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   4920
      Left            =   3840
      TabIndex        =   3
      Top             =   2400
      Width           =   7320
      _cx             =   12912
      _cy             =   8678
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSetExpense.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   2640
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetExpense.frx":0083
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetExpense.frx":061D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetExpense.frx":6E7F
            Key             =   "���U"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "ҩƷƷ��"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblNote 
      Caption         =   "ÿһ������Ŀ�ɰ�Ӧ�ս���Ϊ���(���16��)�����ò�ͬ��ʵ�ձ�����"
      Height          =   180
      Left            =   3840
      TabIndex        =   2
      Top             =   2040
      Width           =   6735
   End
   Begin VB.Label lblMeasure 
      Caption         =   "���㷽��"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ѡ��ѱ�"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "frmSetExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngId As Long            'ҩƷid
Private mstrGrade As String

Public Function ShowMe(objfrm As Object, ByVal lngId As Long, ByVal strgrade As String) As Boolean
    mlngId = lngId
    mstrGrade = strgrade
    Me.Show vbModal, objfrm
    
End Function
'
Private Sub LoadCharge()
    Dim rsTemp As ADODB.Recordset
    Dim intIndex As Integer

    gstrSql = "Select ���� From �ѱ� Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡ�ѱ�")

    cbo�ѱ�.Clear

    With rsTemp
        Do While Not .EOF
            cbo�ѱ�.AddItem !����

            If !���� = mstrGrade Then
                intIndex = cbo�ѱ�.ListCount - 1
            End If

            .MoveNext
        Loop
    End With

    If cbo�ѱ�.ListCount > 0 Then
        cbo�ѱ�.ListIndex = intIndex
    End If

End Sub

Private Sub cbo�ѱ�_Click()
    If tvwDetails.SelectedItem Is Nothing Then Exit Sub
    If tvwDetails.SelectedItem.Tag Like "*Ʒ��" Or tvwDetails.SelectedItem.Tag Like "*����" Then
        vsfList.Rows = 1
        Exit Sub
    End If
    
    If tvwDetails.SelectedItem Is Nothing Then Exit Sub
    If tvwDetails.SelectedItem.Tag Like "*Ʒ��" Or tvwDetails.SelectedItem.Tag Like "*����" Then
        vsfList.Rows = 1
        Exit Sub
    End If
    Call LoadChargeList(Mid(tvwDetails.SelectedItem.Key, 3))
End Sub

Private Sub cbo���㷽��_Click()
'    1-�ɱ��ۼ��ձ�������,���ֶ�
    If cbo���㷽��.ListIndex = 1 Then
        lblNote.Caption = "  ҩƷʵ�ս��=�ɱ���*(1+���ձ���)���������ҩƷ�����Դ����ã������ۡ�"
        vsfList.TextMatrix(0, 1) = "�ֶ����"
        vsfList.TextMatrix(0, 2) = "���ձ���(%)"
    '0-�ֶα�������
    Else
       lblNote.Caption = "    ÿһ������Ŀ�ɰ�Ӧ�ս���Ϊ���(���16��)�����ò�ͬ��ʵ�ձ�����"
       vsfList.TextMatrix(0, 1) = "Ӧ�շֶ����"
       vsfList.TextMatrix(0, 2) = "ʵ�ձ���(%)"
    End If
    
    If tvwDetails.SelectedItem Is Nothing Then Exit Sub
    If tvwDetails.SelectedItem.Tag Like "*Ʒ��" Or tvwDetails.SelectedItem.Tag Like "*����" Then
        vsfList.Rows = 1
        Exit Sub
    End If
    Call LoadChargeList(Mid(tvwDetails.SelectedItem.Key, 3))
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    If tvwDetails.SelectedItem.Tag Like "*����" Or tvwDetails.SelectedItem.Tag Like "*Ʒ��" Then Exit Sub
    If CheckData = True Then
        If optӦ����(0).Value = False Then
            For i = 1 To optӦ����.UBound
                If optӦ����(i).Value = True Then
                    If MsgBox("��ҩƷ�ķѱ�����Ӧ�÷�ΧΪ��" & optӦ����(i).Caption & "���Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
        Call SaveCharge
    End If
End Sub

Private Sub Form_Load()
    Dim objNode As Node
    
    'ByZT20030722
    If glngSys Like "8??" Then
        Caption = "��Ա�ȼ������շ�����"
    End If
    
    '���㷽��
    cbo���㷽��.AddItem "0-�ֶα�������", 0
    cbo���㷽��.AddItem "1-�ɱ��ۼ��ձ�������", 1
    cbo���㷽��.ListIndex = 0
    
    'ȡ�ѱ�
    Call LoadCharge
    '�����
    Call FullTreeView
    Call InitVsf    '��ʼ���ؼ�
    
    Call LoadChargeList(mlngId)
    For Each objNode In tvwDetails.Nodes
        If Mid(objNode.Key, 3) = mlngId Then
            objNode.Selected = True
            objNode.Expanded = True
        End If
    Next
End Sub


Private Function CheckData() As Boolean
    '������ݲ����п���
    Dim intRow As Integer
    Dim intCol As Integer
    With vsfList
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                If Trim(.TextMatrix(intRow, intCol)) = "" Then
                    MsgBox "��Ԫ����Ϊ�գ�", vbInformation, gstrSysName
                    CheckData = False
                    vsfList.SetFocus
                    .Row = intRow
                    .Col = intCol
                    Exit Function
                End If
            Next
        Next
        CheckData = True
    End With
End Function
Private Sub LoadChargeList(ByVal lngId As Long)
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String

    gstrSql = "Select �κ�, Ӧ�ն���ֵ, Ӧ�ն�βֵ, ʵ�ձ���, ���㷽�� " & _
        " From �ѱ���ϸ Where �ѱ� = [1] And �շ�ϸĿid=[2] And ���㷽��=[3]" & strSQL & " Order By �κ�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡ�ѱ���ϸ", cbo�ѱ�.Text, lngId, cbo���㷽��.ListIndex)

    vsfList.Rows = 2
    If rsTemp.RecordCount = 0 Then
        vsfList.TextMatrix(1, 0) = 1
        vsfList.TextMatrix(1, 1) = "0.00"
        vsfList.TextMatrix(1, 2) = "100.00"
        Exit Sub
    End If

    cbo���㷽��.ListIndex = IIf(rsTemp!���㷽�� = 0, 0, 1)

    With rsTemp
        vsfList.Rows = .RecordCount + 1
        cbo���㷽��.ListIndex = Val(.Fields("���㷽��").Value)     '����Click�¼�������ؿؼ�

        For i = 1 To .RecordCount
            If i > 16 Then Exit For
            vsfList.TextMatrix(i, 0) = i
            vsfList.TextMatrix(i, 1) = Format(.Fields("Ӧ�ն���ֵ").Value, "###########0.00;-##########0.00;0.00;0.00")
            vsfList.TextMatrix(i, 2) = Format(.Fields("ʵ�ձ���").Value, "###0.000;-##0.000;0.000;0.000")
            .MoveNext
        Next
    End With
    With vsfList
        .Cell(flexcpBackColor, 1, 1, 1, 1) = &H8000000F
    End With
End Sub

Private Sub FullTreeView()
    Dim NodeThis As Node
    Dim Intĩ�� As Integer
    Dim lng�ⷿID As Long
    Dim rs���ʷ��� As ADODB.Recordset
    Dim recdate As ADODB.Recordset
    
    'ҩƷ��;�����Ƿ�������
    gstrSql = " Select ����,���� From ������Ŀ��� " & _
              " Where Instr([1],����,1) > 0 " & _
              " Order by ����"
    Set rs���ʷ��� = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "567")
    
    If rs���ʷ���.RecordCount = 0 Then
        Exit Sub
    End If
    
    With tvwDetails
        .Nodes.Clear
        Do While Not rs���ʷ���.EOF
            .Nodes.Add , , "Root" & rs���ʷ���!����, rs���ʷ���!����, 1, 1
            .Nodes("Root" & rs���ʷ���!����).Tag = rs���ʷ���!����
            rs���ʷ���.MoveNext
        Loop
    End With
    '����
    gstrSql = "Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, '����' As ���" & vbNewLine & _
                "From ���Ʒ���Ŀ¼" & vbNewLine & _
                "Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & vbNewLine & _
                "Start With �ϼ�id Is Null" & vbNewLine & _
                "Connect By Prior ID = �ϼ�id"
    Set recdate = zlDatabase.OpenSQLRecord(gstrSql, "����")
    
    If recdate.EOF Then
        MsgBox "���ʼ��ҩƷ��;���ࣨҩƷ��;���ࣩ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With recdate
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set NodeThis = tvwDetails.Nodes.Add("Root" & !����, 4, "����K_" & !ID, !����, 1, 1)
            Else
                Set NodeThis = tvwDetails.Nodes.Add("����K_" & !�ϼ�ID, 4, "����K_" & !ID, !����, 1, 1)
            End If
            NodeThis.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
            .MoveNext
        Loop
    End With
    'Ʒ��
    gstrSql = "Select ID, ����id, ����, ����, Decode(���, 5, '����ҩ', 6, '�г�ҩ', 7, '�в�ҩ') ����, 'Ʒ��' As ���" & vbNewLine & _
                "From ������ĿĿ¼" & vbNewLine & _
                "Where ����id In (Select ID" & vbNewLine & _
                "               From ���Ʒ���Ŀ¼" & vbNewLine & _
                "               Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & vbNewLine & _
                "               Start With �ϼ�id Is Null" & vbNewLine & _
                "               Connect By Prior ID = �ϼ�id)"
    Set recdate = zlDatabase.OpenSQLRecord(gstrSql, "Ʒ��")
    With recdate
        Do While Not .EOF
            Set NodeThis = tvwDetails.Nodes.Add("����K_" & !����id, 4, "K_" & !ID, !����, 1, 1)
            NodeThis.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
            .MoveNext
        Loop
    End With
    '���
    gstrSql = "Select a.ҩƷid As ID, a.ҩ��id As �ϼ�id, b.����,b.��� as ����, b.����, b.���" & _
               " From ҩƷ��� A," & _
                "     (Select ID, ����id, ����, ���, Decode(���, '5', '����ҩ', '6', '�г�ҩ', '7', '�в�ҩ') ����, 'ҩƷ' As ���" & _
                      " From �շ���ĿĿ¼" & _
                      " Where ��� In ('5', '6', '7') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01') B" & _
               " Where a.ҩƷid = b.Id"
    Set recdate = zlDatabase.OpenSQLRecord(gstrSql, "����ѯ")
    
    With recdate
        Do While Not .EOF
            Set NodeThis = tvwDetails.Nodes.Add("K_" & !�ϼ�ID, 4, "M_" & !ID, IIf(IsNull(!����), "", !����), 1, 1)
            NodeThis.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
            .MoveNext
        Loop
    End With
    
    With tvwDetails
        If .Nodes.Count <> 0 Then
            .Nodes(1).Selected = True
            If .Nodes(1).Children <> 0 Then
                Intĩ�� = 1
                .Nodes(Intĩ��).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(2).Children <> 0 Then
                Intĩ�� = 2
                .Nodes(Intĩ��).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(3).Children <> 0 Then
                Intĩ�� = 3
                .Nodes(Intĩ��).Child.Selected = True
                .SelectedItem.Selected = True
            Else
                Intĩ�� = 0
                .Nodes(1).Selected = True
                .SelectedItem.Selected = True
            End If
            If Intĩ�� <> 0 Then .Nodes(Intĩ��).Expanded = True
        End If
    End With
    tvwDetails.Move 0, 0, picҩƷ.Width, picҩƷ.Height
End Sub

Private Sub InitVsf()
    '��ʼ��vsflexgrid
    With vsfList
        .Cols = 3
        .Rows = 1
        .Editable = flexEDNone
        .SelectionMode = flexSelectionFree
        
        .TextMatrix(0, 0) = "�ֶκ�"
        If cbo���㷽��.Text = "0-�ֶα�������" Then
            .TextMatrix(0, 1) = "Ӧ�շֶ����"
            .TextMatrix(0, 2) = "���ձ���(%)"
        ElseIf cbo���㷽��.Text = "1-�ɱ��ۼ��ձ�������" Then
            .TextMatrix(0, 1) = "�ֶ����"
            .TextMatrix(0, 2) = "���ձ���(%)"
        End If
    End With
End Sub

Private Sub optӦ����_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optӦ����.UBound
        If i = Index Then
            optӦ����(i).FontBold = True
        Else
            optӦ����(i).FontBold = False
        End If
    Next
End Sub

Private Sub tvwDetails_NodeClick(ByVal Node As MSComctlLib.Node)
    With tvwDetails
        If Node.Tag Like "*����" = True Or Node.Tag Like "*Ʒ��" Or Node.Tag = "5" Or Node.Tag = "6" Or Node.Tag = "7" Then
            vsfList.Rows = 1
            vsfList.Editable = flexEDNone
            Exit Sub
        End If
        Call GetDrugOtherInfo(Mid(Node.Key, 3))
        Call LoadChargeList(Mid(tvwDetails.SelectedItem.Key, 3))
        vsfList.SetFocus
    End With
End Sub

Private Sub GetDrugOtherInfo(ByVal lngItemId As Long)
    '��Ҫ����ҩƷĿ¼�����еõ���ǰҩƷ�ļ��ͺͲ���
    Dim rsTemp As ADODB.Recordset
    Dim str���� As String
    If lngItemId = 0 Then Exit Sub
    
    gstrSql = "Select Decode(A.���, '5', '����ҩ', '6', '�г�ҩ', '�в�ҩ') As ���, B.ҩƷ���� " & _
        " From �շ���ĿĿ¼ A, ҩƷ���� B, ҩƷ��� C " & _
        " Where A.ID = C.ҩƷid And B.ҩ��id = C.ҩ��id And A.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡҩƷ��Ϣ", lngItemId)
    
    If Not rsTemp.EOF Then
        optӦ����(2).Caption = "Ӧ�������С�" & rsTemp!��� & "��(&2)"
        optӦ����(3).Caption = "Ӧ�������С�" & rsTemp!ҩƷ���� & "����ҩƷ(&3)"
    End If
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        If .Editable = flexEDKbdMouse Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfList_EnterCell()
    With vsfList
        If .Col = 1 And .Row = 1 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfList
        If KeyCode = vbKeyReturn Then
            If .TextMatrix(.Row, .Col) = "" Then
                MsgBox "����Ϊ�գ�", vbInformation, gstrSysName
                vsfList.SetFocus
                KeyCode = 0
                Exit Sub
            End If
            If .Col <> .Cols - 1 Then
                .Col = .Col + 1
            ElseIf .Col < 16 And cbo���㷽��.ListIndex <> 1 Then
                .Rows = .Rows + 1
                .Row = .Row + 1
                .TextMatrix(.Row, 0) = .Row
                .Col = 1
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Rows > 2 Then
                .RemoveItem .Row
            Else
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
            End If
        End If
    End With
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfList
        If IsNumeric(Chr(KeyAscii)) = False And (KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete) And Chr(KeyAscii) <> "." And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
    End With
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If IsNumeric(.EditText) = False Then
            MsgBox "��������ȷ������", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If Row <> 1 And Val(.EditText) < Val(.TextMatrix(Row - 1, 1)) And Col = 1 Then
            MsgBox "Ӧ�ֶ�ֵ������С����", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If Row <> 1 And Val(.EditText) = Val(.TextMatrix(Row - 1, 1)) And Col = 2 Then
            MsgBox "�����б��ʲ�����ͬ��", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
    End With
End Sub


Private Sub SaveCharge()
    Dim str���� As String
    Dim curStart As Currency, curEnd As Currency, dblTax As Double
    Dim intRow As Long
    Dim blnTrans As Boolean
    Dim intӦ�� As Integer
    Dim lngId As Long
    
    On Error GoTo ErrHand
    lngId = Mid(tvwDetails.SelectedItem.Key, 3)
    If vsfList.Rows = 1 Then Exit Sub
'    If vsfList.Editable = flexEDNone Then Exit Sub
    With vsfList
        For intRow = 1 To .Rows - 1
            curStart = Val(.TextMatrix(intRow, 0))
            If intRow >= .Rows - 1 Then
                curEnd = Val("10000000000.00")
            Else
                curEnd = Val(.TextMatrix(intRow, 1)) - 0.01
            End If
            dblTax = .TextMatrix(intRow, 2)
            str���� = str���� & intRow & ":" & curStart & ":" & curEnd & ":" & dblTax & ";"
        Next
    End With
    
    gcnOracle.BeginTrans

    'ҩƷĿ¼�����÷ѱ�
    If optӦ����(0).Value = True Then
        intӦ�� = 0
    ElseIf optӦ����(1).Value = True Then
        intӦ�� = 1
    ElseIf optӦ����(2).Value = True Then
        intӦ�� = 2
    ElseIf optӦ����(3).Value = True Then
        intӦ�� = 3
    ElseIf optӦ����(4).Value = True Then
        intӦ�� = 4
    ElseIf optӦ����(5).Value = True Then
        intӦ�� = 5
    End If

    gstrSql = "zl_�ѱ���ϸ_update('" & cbo�ѱ�.Text & "'," & lngId & ",'" & str���� & "'," & Val(cbo���㷽��.Text) & "," & 3 & "," & intӦ�� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)

    gcnOracle.CommitTrans
    MsgBox "����ɹ���", vbInformation, gstrSysName
    Exit Sub
ErrHand:
    If Not blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub

