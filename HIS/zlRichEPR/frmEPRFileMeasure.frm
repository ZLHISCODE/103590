VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRFileMeasure 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "֪���ļ�Ҫ��"
   ClientHeight    =   5460
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7650
   Icon            =   "frmEPRFileMeasure.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkDelMsg 
      Caption         =   "ɾ������(&M)"
      Height          =   195
      Left            =   6255
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1725
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.PictureBox picHBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   30
      ScaleWidth      =   7935
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5430
      Width           =   7935
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   2385
      Left            =   -30
      ScaleHeight     =   2385
      ScaleWidth      =   7425
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2940
      Width           =   7425
      Begin MSComctlLib.ListView lvwItems 
         Height          =   1845
         Left            =   3585
         TabIndex        =   2
         Top             =   30
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   3254
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   3570
         TabIndex        =   3
         Top             =   1860
         Width           =   2640
      End
      Begin VB.PictureBox picVBar 
         BackColor       =   &H00808080&
         Height          =   4080
         Left            =   3435
         MousePointer    =   9  'Size W E
         ScaleHeight     =   4080
         ScaleWidth      =   30
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   -15
         Width           =   30
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   2040
         Left            =   165
         TabIndex        =   1
         Tag             =   "1000"
         Top             =   0
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   3598
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   0
         Top             =   1395
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
               Picture         =   "frmEPRFileMeasure.frx":058A
               Key             =   "close"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEPRFileMeasure.frx":0B24
               Key             =   "expend"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdFind 
         Height          =   300
         Left            =   6225
         Picture         =   "frmEPRFileMeasure.frx":10BE
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "���ҷ�����������Ŀ"
         Top             =   1860
         Width           =   360
      End
      Begin VB.CommandButton cmdSel 
         Height          =   300
         Index           =   0
         Left            =   6660
         Picture         =   "frmEPRFileMeasure.frx":1648
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��������Ŀ"
         Top             =   1860
         Width           =   360
      End
      Begin VB.CommandButton cmdSel 
         Height          =   300
         Index           =   1
         Left            =   7020
         Picture         =   "frmEPRFileMeasure.frx":1BD2
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "�������ѡ��"
         Top             =   1860
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Index           =   1
      Left            =   6255
      TabIndex        =   8
      Top             =   1980
      Width           =   1260
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      Height          =   350
      Left            =   6255
      TabIndex        =   10
      Top             =   600
      Width           =   1260
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "���(&A)"
      Height          =   350
      Index           =   0
      Left            =   6255
      TabIndex        =   7
      Top             =   2370
      Width           =   1260
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdItems 
      Height          =   2115
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   5925
      _cx             =   10451
      _cy             =   3731
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
      BackColorSel    =   16764057
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEPRFileMeasure.frx":215C
      ScrollTrack     =   -1  'True
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
      WordWrap        =   -1  'True
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
   Begin VB.Label lblMeasure 
      AutoSize        =   -1  'True
      Caption         =   "��ִ���������ƴ�ʩǰ���Ͳ��˻����Э��һ�£���д���ļ���"
      Height          =   180
      Left            =   765
      TabIndex        =   11
      Top             =   345
      Width           =   5040
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "�ļ�����:   001-����ͬ����"
      Height          =   180
      Left            =   765
      TabIndex        =   9
      Top             =   75
      Width           =   2340
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   195
      Picture         =   "frmEPRFileMeasure.frx":21BE
      Top             =   45
      Width           =   480
   End
End
Attribute VB_Name = "frmEPRFileMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const conColumn_ID = 0
Const conColumn_���� = 1
Const conColumn_���� = 2
Const conColumn_��� = 3

Private mlngFileID As Long        '�����ļ�ID
Private mblnOK As Boolean


Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileID As Long) As Boolean
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim lngCount As Long
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    mlngFileID = lngFileID
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ����, ���, ����, ͨ�� From �����ļ��б� Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        If .RecordCount = 0 Then MsgBox "�ļ���ʧ(���ܱ������û�ɾ��)��", vbInformation, gstrSysName: Exit Function
        Me.lblFile.Caption = "�ļ�����:   " & !��� & "-" & !����
    End With
    
    '---------------------------------------------------
    gstrSQL = "Select Distinct i.Id, i.����, i.����, k.���� As ���" & _
            " From ������Ŀ��� k, ������ĿĿ¼ i, ��������Ӧ�� a" & _
            " Where k.���� = i.��� And i.Id = a.������Ŀid And i.��� In ('C', 'D', 'E', 'F', 'G', 'K', 'L') And a.�����ļ�id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    Set Me.vgdItems.DataSource = rsTemp
    With Me.vgdItems
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftCenter
        Next
        .ColHidden(conColumn_ID) = True
        .ColWidth(conColumn_����) = 1000
        .ColWidth(conColumn_����) = 3650
        .ColWidth(conColumn_���) = 1000
    End With
    
    '---------------------------------------------------
    '�����������ݣ����Ʒ���
    gstrSQL = "Select Id, �ϼ�id, ����, ���� From ���Ʒ���Ŀ¼ Where ���� = 5 Start With �ϼ�id Is Null Connect By Prior Id = �ϼ�id"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    On Error GoTo 0
    With rsTemp
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        If Me.tvwClass.Nodes.Count > 0 Then
            Me.tvwClass.Nodes(1).Expanded = True
            Me.tvwClass.Nodes(1).Selected = True
            Call tvwClass_NodeClick(Me.tvwClass.Nodes(1))
        End If
    End With
    
    '---------------------------------------------------
    Me.Show vbModal, frmParent
    '---------------------------------------------------
    ShowMe = mblnOK: Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdEdit_Click(Index As Integer)
Dim strTemp As String
Dim objItem As ListItem
    
    If Index = 0 Then       '���
        strTemp = ""
        For Each objItem In Me.lvwItems.ListItems
            If objItem.Checked Then strTemp = strTemp & ";" & Mid(objItem.Key, 2)
        Next
        If strTemp = "" Then MsgBox "û��ѡ��������Ŀ��", vbInformation, gstrSysName: Exit Sub
        If Len(strTemp) > 4000 Then MsgBox "һ��ѡ����̫���������Ŀ��", vbInformation, gstrSysName: Exit Sub
        gstrSQL = "Zl_֪���ļ���Ŀ_Append(" & mlngFileID & ",'" & Mid(strTemp, 2) & "')"
        
        Err = 0: On Error GoTo errHand
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        With Me.vgdItems
            For Each objItem In Me.lvwItems.ListItems
                If objItem.Checked Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, conColumn_ID) = Mid(objItem.Key, 2)
                    .TextMatrix(.Rows - 1, conColumn_����) = objItem.Text
                    .TextMatrix(.Rows - 1, conColumn_����) = objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1)
                    .TextMatrix(.Rows - 1, conColumn_���) = objItem.SubItems(Me.lvwItems.ColumnHeaders("_���").Index - 1)
                End If
            Next
            If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
        End With
    Else                    'ɾ��
        With Me.vgdItems
            If Val(.TextMatrix(.Row, conColumn_ID)) = 0 Then MsgBox "�Ѿ�ɾ����ɣ�", vbInformation, gstrSysName: Exit Sub
            If Me.chkDelMsg.Value = vbChecked Then
                If MsgBox("���ɾ�������ƴ�ʩ��" & vbCrLf & "����" & .TextMatrix(.Row, conColumn_����), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            gstrSQL = "Zl_֪���ļ���Ŀ_Delete(" & mlngFileID & "," & Val(.TextMatrix(.Row, conColumn_ID)) & ")"
            Err = 0: On Error GoTo errHand
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            .RemoveItem .Row
        End With
    End If
    
    If Val(Me.lvwItems.Tag) = 0 Or Trim(Me.txtFind.Text) = "" Then
        If Me.tvwClass.Nodes.Count > 0 Then
            If Me.tvwClass.SelectedItem Is Nothing Then Me.tvwClass.Nodes(1).Selected = True
            Call tvwClass_NodeClick(Me.tvwClass.SelectedItem)
        End If
    Else
        Me.cmdFind.Tag = "1"
        Call cmdFind_Click
        Me.cmdFind.Tag = "0"
    End If
    
    mblnOK = True: Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdFind_Click()
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
    If Trim(Me.txtFind.Text) = "" Then MsgBox "û������������ݣ�", vbInformation, gstrSysName: Exit Sub
    
    gstrSQL = "Select Distinct i.Id, i.����, i.����, k.���� As ���" & _
            " From ������Ŀ��� k, ������ĿĿ¼ i, ������Ŀ���� n, (Select ������Ŀid From ��������Ӧ�� Where �����ļ�id = [3]) s" & _
            " Where k.���� = i.��� And i.��� In ('C', 'D', 'E', 'F', 'G', 'K', 'L') And i.Id = n.������Ŀid And" & _
            "       (i.���� like [1] or n.���� like [2] or n.���� like [2]) And (i.����ʱ�� > Sysdate Or i.����ʱ�� Is Null) And " & _
            "        i.Id = s.������Ŀid(+) And s.������Ŀid Is Null"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Me.txtFind.Text) & "%", gstrMatch & Trim(Me.txtFind.Text) & "%", mlngFileID)
    With rsTemp
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
'            objItem.Icon = "_" & !���: objItem.SmallIcon = objItem.Icon
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_���").Index - 1) = !���
            .MoveNext
        Loop
    End With
    If Me.lvwItems.ListItems.Count = 0 Then
        If Val(Me.cmdFind.Tag) = 0 Then MsgBox "û��ƥ���������Ŀ��", vbInformation, gstrSysName
        Me.txtFind.SetFocus
    Else
        Me.vgdItems.SetFocus
    End If
    Me.lvwItems.Tag = "1"
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click(Index As Integer)
Dim objItem As ListItem
    For Each objItem In Me.lvwItems.ListItems
        objItem.Checked = (Index = 0)
    Next
    Me.lvwItems.SetFocus
End Sub

Private Sub Form_Activate()
    Me.vgdItems.SetFocus
End Sub

Private Sub Form_Load()
    With Me.picHBar
        .ZOrder 0: .BackColor = Me.BackColor:
        .Top = Me.ScaleHeight - Me.picHBar.Height
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
    End With
    With Me.picVBar
        .ZOrder 0: .BackColor = Me.BackColor: .Top = 0: .Height = Me.picBack.ScaleHeight
    End With

    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "_����", "����", 1000
        .Add , "_����", "����", 2300
        .Add , "_���", "���", 600
    End With
    With Me.lvwItems
        .SortKey = .ColumnHeaders("_����").Index - 1
        .SortOrder = lvwAscending
    End With
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    With Me.picBack
        .Left = Me.ScaleLeft + Me.vgdItems.Left: .Width = Me.cmdClose.Left + Me.cmdClose.Width - .Left
        .Top = Me.vgdItems.Top + Me.vgdItems.Height + 90: .Height = Me.ScaleHeight - .Top - 90
    End With
    With Me.picVBar
        If .Left < 1000 Then .Left = 1000
        If .Left > Me.picBack.ScaleWidth - 1000 Then .Left = Me.picBack.ScaleWidth - 1000
        .Top = Me.picBack.ScaleTop: .Height = Me.picBack.ScaleHeight
    End With
    With Me.tvwClass
        .Top = Me.picBack.ScaleTop: .Height = Me.picBack.ScaleHeight
        .Left = Me.picBack.ScaleLeft: .Width = Me.picVBar.Left - .Left
    End With
    With Me.lvwItems
        .Top = Me.picBack.ScaleTop: .Height = Me.picBack.ScaleHeight - .Top - Me.txtFind.Height + 15
        .Left = Me.picVBar.Left + Me.picVBar.Width: .Width = Me.picBack.ScaleWidth - .Left
    End With
    With Me.cmdSel(1): .Top = Me.picBack.ScaleHeight - .Height: .Left = Me.picBack.ScaleWidth - .Width: End With
    With Me.cmdSel(0): .Top = Me.picBack.ScaleHeight - .Height: .Left = cmdSel(1).Left - .Width: End With
    With Me.cmdFind: .Top = Me.picBack.ScaleHeight - .Height: .Left = cmdSel(0).Left - .Width - 45: End With
    With Me.txtFind
        .Top = Me.picBack.ScaleHeight - .Height
        .Left = Me.picVBar.Left + Me.picVBar.Width: .Width = Me.cmdFind.Left - .Left
    End With
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwItems
        If .SortKey = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub picHBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngHMargin As Long
    If Button = 1 Then
        lngHMargin = Me.Height - Me.ScaleHeight
        Me.picHBar.Top = Me.picHBar.Top + y
        If Me.picHBar.Top < Me.vgdItems.Top + Me.vgdItems.Height + 900 Then
            Me.picHBar.Top = Me.vgdItems.Top + Me.vgdItems.Height + 900
        End If
        Me.Height = Me.picHBar.Top + Me.picHBar.Height + lngHMargin
    End If
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picVBar.Left = Me.picVBar.Left + x
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call Form_Resize
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select i.Id, i.����, i.����, k.���� As ���" & _
            " From ������Ŀ��� k, ������ĿĿ¼ i, (Select ������Ŀid From ��������Ӧ�� Where �����ļ�id = [2]) s" & _
            " Where k.���� = i.��� And i.��� In ('C', 'D', 'E', 'F', 'G', 'K', 'L') And i.����id = [1] And" & _
            "       (i.����ʱ�� > Sysdate Or i.����ʱ�� Is Null) And i.Id = s.������Ŀid(+) And s.������Ŀid Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Mid(Node.Key, 2)), mlngFileID)
    With rsTemp
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
'            objItem.Icon = "_" & !���: objItem.SmallIcon = objItem.Icon
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_����").Index - 1) = !����
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_���").Index - 1) = !���
            .MoveNext
        Loop
    End With
    Me.lvwItems.Tag = "0"
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click: Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub vgdItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then Call cmdEdit_Click(1)
End Sub
