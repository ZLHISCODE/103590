VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ��������"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7410
   Icon            =   "frmServiceSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7410
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picDept 
      Height          =   3795
      Left            =   2610
      ScaleHeight     =   3735
      ScaleWidth      =   4710
      TabIndex        =   1
      Top             =   15
      Width           =   4770
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   3300
         Left            =   15
         TabIndex        =   7
         Top             =   435
         Width           =   4680
         _cx             =   8255
         _cy             =   5821
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
         BackColorSel    =   16777152
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmServiceSelect.frx":000C
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
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
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
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   3
         Top             =   60
         Width           =   1335
      End
      Begin VB.CheckBox chkAllSelect 
         Caption         =   "ȫѡ"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   83
         Width           =   720
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   50
         TabIndex        =   4
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4920
      TabIndex        =   6
      Top             =   3870
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6210
      TabIndex        =   5
      Top             =   3870
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1740
      Top             =   225
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
            Picture         =   "frmServiceSelect.frx":0081
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSelect.frx":06CD
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSelect.frx":09E9
            Key             =   "Dept_No"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDept 
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   6694
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmServiceSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintRow  As Integer  '��¼��ǰ��
Private mintFind As Integer  '������¼��ѯ���ĸ�λ����
Private mblnChkFocus As Boolean  '"ȫѡ"��ѡ���ȡ����ʱΪTrue
Private mrs����  As Recordset    '���ż�¼��
Private mstrKey  As String       'ѡ�нڵ�����ʱ���
Private mblnFind As Boolean      'ͨ������(���롢����)��ѯʱΪTrue
Private mblnChang As Boolean
Private mintģ�� As Integer 'mintģ��=1Ϊ�洢�ⷿ��mintģ��=2Ϊ��������޸Ĵ洢�ⷿ

Private Sub SetColumns()
    Dim intCol As Integer
    
    With vsfList
        .Rows = 1
        .Cols = 7
        .ColDataType(0) = flexDTBoolean
        .Editable = flexEDKbdMouse
'        .ExtendLastCol = True
        .TextMatrix(0, 0) = "ѡ��"
        .TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "����"
        .TextMatrix(0, 4) = "����"
        .TextMatrix(0, 5) = "���ʱ���"
        .TextMatrix(0, 6) = "����"
        
        For intCol = 0 To .Cols - 1
            .ColKey(intCol) = .TextMatrix(0, intCol)
        Next
        .ColWidth(.ColIndex("ѡ��")) = 500
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("����")) = 2000
        .ColWidth(.ColIndex("���ʱ���")) = 0
        .ColWidth(.ColIndex("����")) = IIf(mblnFind, 1000, 0)
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("����")) = flexAlignLeftCenter
    End With
End Sub

Private Sub FillTree()
'��ȡ���ʸ�ֵ������
    Dim rs���� As Recordset
    Dim str���� As String
    
    gstrSql = "Select ����, ���� From �������ʷ��� Where Instr('3ABCDEF', ����) > 0"
    Set rs���� = zlDatabase.OpenSQLRecord(gstrSql, "��ѯ����")
    
    With tvwDept
        .Nodes.Clear
        .Nodes.Add , , "KRoot", "��������", "Root", "Root"
        .Nodes("KRoot").Sorted = True
        
        Do While Not rs����.EOF
            .Nodes.Add "KRoot", tvwChild, "K" & rs����!����, rs����!����, "Dept", "Dept"
            rs����.MoveNext
        Loop
        
        .Nodes.Item(1).Expanded = True
        .Nodes.Item(1).Selected = True
    End With
End Sub

Public Sub ShowMe(ByVal frmParent As Form, ByVal intRow As Integer, ByVal str������� As String, ByVal intģ�� As Integer, Optional strkey As String)
    Dim strTemp As String
    Dim strFind As String
    
    mintģ�� = intģ��
    mintRow = intRow
    mblnFind = (strkey <> "")
    If mblnFind Then
        strTemp = " And ( d.���� like [2] or d.���� like [2] or d.���� like [2]) "
    End If
    
    gstrSql = "Select Distinct d.Id, d.����, d.����, d.����, a.���� As ���ʱ���, c.�������� as ���� " & vbNewLine & _
            "From �������ʷ��� A, ��������˵�� C, ���ű� D " & vbNewLine & _
            "Where d.Id = c.����id And c.�������� = a.���� And (d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.����ʱ�� Is Null) And " & vbNewLine & _
            "Instr('3ABCDEF', a.����) > 0 And Instr([1], ',' || c.������� || ',') > 0 " & strTemp & vbNewLine & _
            "Order By d.id,d.����"
    Set mrs���� = zlDatabase.OpenSQLRecord(gstrSql, "��ȡ����", "," & str������� & ",", strkey & "%")
    
    frmServiceSelect.Show vbModal, frmParent
End Sub

Private Sub chkAllSelect_Click()
    Dim i As Integer
    
    With vsfList
        If mblnChkFocus Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ѡ��")) = IIf(chkAllSelect.Value = 1, -1, 0)
            Next
        End If
    End With
End Sub

Private Sub chkAllSelect_GotFocus()
    mblnChkFocus = True
End Sub

Private Sub chkAllSelect_LostFocus()
    mblnChkFocus = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If mblnFind Then
        Me.Width = 6000
        tvwDept.Visible = False
    Else
        Call FillTree
    End If
    
    Call SetColumns
    Call FillList("KRoot")
    mintFind = 0
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim str���� As String, str����ID As String
    
    'ѭ����ȡ�û���ѡ��Ŀ���
    With vsfList
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = -1 Then
                str���� = str���� & "," & .TextMatrix(lngRow, .ColIndex("����"))
                str����ID = str����ID & "," & .TextMatrix(lngRow, .ColIndex("ID"))
            End If
        Next
    End With
    
    If str���� <> "" Then
        str���� = Mid(str����, 2)
        str����ID = Mid(str����ID, 2)
    End If
    If mintģ�� = 1 Then '�洢�ⷿģ��
        With frmServiceSectOffice.msfServiceSectOffice
            If str���� <> "" Then .TextMatrix(.Row, 1) = "��"
            .Text = str����
            .TextMatrix(mintRow, 3) = .Text
            .TextMatrix(mintRow, 4) = str����ID
            If .Rows - 1 > .Row Then .Row = .Row + 1
        End With
    Else   '��������޸�ģ��
        With frmServiceDepartment.vsfDepartment
            .TextMatrix(.Row, .Col) = str����
            .TextMatrix(.Row, .Col + 1) = str����ID
        End With
    End If
    
    Unload Me
End Sub

Private Sub Form_Resize()
    If mblnFind Then
        picDept.Move 0, picDept.Top, Me.ScaleWidth
        cmdCancel.Move cmdCancel.Left - 800
        cmdOK.Move cmdOK.Left - 800
        vsfList.Width = picDept.ScaleWidth - 10
    End If
End Sub

Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If Col <> 0 Then
            Cancel = True
        End If
        mblnChkFocus = False
    End With
End Sub

Private Sub vsfList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim intRow As Integer
    Dim intCount As Integer
    
    With vsfList
        If Row > 0 And Col = 0 Then
            For intRow = 1 To .Rows - 1
                If Val(.TextMatrix(intRow, .ColIndex("ѡ��"))) = -1 Then
                    intCount = intCount + 1
                End If
            Next
            
            '�Ƿ�ȫѡ
            If mblnChkFocus = False Then
                If intCount = .Rows - 1 Then
                    chkAllSelect.Value = 1
                Else
                    chkAllSelect.Value = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfList_ChangeEdit()
    mblnChang = True
End Sub

Private Sub tvwDept_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrKey = Mid(Node.Key, 2) Then Exit Sub
    Call FillList(Mid(Node.Key, 2))
End Sub

Private Sub FillList(Optional ByVal strkey As String)
'��ȡ������Ϣ
    Dim i As Integer
    Dim intCount As Integer
    
    With vsfList
        mstrKey = strkey
        If strkey = "KRoot" Then
            .Rows = 1
            Do While Not mrs����.EOF
                For i = 1 To .Rows - 1
                    If mrs����!ID = .TextMatrix(i, .ColIndex("ID")) Then
                        mrs����.MoveNext
                        If mrs����.EOF Then
                            chkAllSelect.Value = IIf(intCount = .Rows - 1, 1, 0)
                            .Row = 1
                            Exit Sub
                        End If
                        i = 0
                    End If
                Next
                
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = IIf(IsNull(mrs����!ID), "", mrs����!ID)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(mrs����!����), "", mrs����!����)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(mrs����!����), "", mrs����!����)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(mrs����!����), "", mrs����!����)
                .TextMatrix(.Rows - 1, .ColIndex("���ʱ���")) = IIf(IsNull(mrs����!���ʱ���), "", mrs����!���ʱ���)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(mrs����!����), "", mrs����!����)
    
                If mintģ�� = 1 Then '�洢�ⷿģ��
                    If InStr(1, "," & frmServiceSectOffice.msfServiceSectOffice.TextMatrix(frmServiceSectOffice.msfServiceSectOffice.Row, 4) & ",", "," & mrs����!ID & ",") > 0 Then
                        .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = -1
                        intCount = intCount + 1
                    End If
                Else    '��������޸Ŀⷿģ��
                    If InStr(1, "," & frmServiceDepartment.vsfDepartment.TextMatrix(frmServiceDepartment.vsfDepartment.Row, 3) & ",", "," & mrs����!ID & ",") > 0 Then
                        .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = -1
                        intCount = intCount + 1
                    End If
                End If
                
                mrs����.MoveNext
            Loop
            chkAllSelect.Value = IIf(intCount = .Rows - 1, 1, 0)
            .Row = 1
        Else
            For i = 1 To .Rows - 1
                .RowHidden(i) = False
                If .TextMatrix(i, .ColIndex("���ʱ���")) <> strkey And strkey <> "Root" Then
                    .RowHidden(i) = True
                End If
            Next
        End If
    End With
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strFind As String
    Dim i As Integer
    Dim blnResult As Boolean
    Dim j As Integer
    Dim k As Integer
    
    blnResult = False
    With vsfList
        If KeyCode = vbKeyReturn And Trim(txtFind.Text) <> "" Then
            strFind = UCase(Trim(txtFind.Text))
            If mintFind > .Rows - 1 Then
                mintFind = 1
            Else
                mintFind = mintFind + 1
                If mintFind > .Rows - 1 Then
                    mintFind = 1
                End If
            End If
            
            For i = mintFind To .Rows - 1
                If IsNumeric(strFind) Then
                    If .TextMatrix(i, .ColIndex("����")) = strFind Then
                        .Row = i
                        .TopRow = i
                        mintFind = i
                        Call SelectNode
                        Exit Sub
                    End If
                    
                    If i = .Rows - 1 Then
                        For k = 1 To mintFind
                            If .TextMatrix(k, .ColIndex("����")) = strFind Then
                                .Row = k
                                .TopRow = k
                                mintFind = k
                                Call SelectNode
                                Exit Sub
                            End If
                        Next
                    End If
                Else
                    If .TextMatrix(i, .ColIndex("����")) Like "*" & strFind & "*" Then
                        .Row = i
                        .TopRow = i
                        mintFind = i
                        Call SelectNode
                        blnResult = True
                        Exit Sub
                    End If
                    
                    If i = .Rows - 1 Then
                        For k = 1 To mintFind
                            If .TextMatrix(k, .ColIndex("����")) Like "*" & strFind & "*" Then
                                .Row = k
                                .TopRow = k
                                mintFind = k
                                Call SelectNode
                                blnResult = True
                                Exit Sub
                            End If
                        Next
                    End If
                End If
            Next
            
            If blnResult = False Then
                For i = mintFind To .Rows - 1
                    If .TextMatrix(i, .ColIndex("����")) Like "*" & strFind & "*" Then
                        .Row = i
                        .TopRow = i
                        mintFind = i
                        Call SelectNode
                        blnResult = True
                        Exit Sub
                    End If
                Next
                
                For k = 1 To mintFind
                    If .TextMatrix(k, .ColIndex("����")) Like "*" & strFind & "*" Then
                        .Row = k
                        .TopRow = k
                        mintFind = k
                        Call SelectNode
                        blnResult = True
                        Exit Sub
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub SelectNode()
'ȷ�����Ҳ��ŵ�����
    Dim i As Integer
    
    With vsfList
        For i = 1 To tvwDept.Nodes.Count
            If .TextMatrix(.Row, .ColIndex("���ʱ���")) = Mid(tvwDept.Nodes(i).Key, 2) Then
                tvwDept.Nodes(i).Selected = True
                Call FillList(Mid(tvwDept.Nodes(i).Key, 2))
                Exit Sub
            End If
        Next
    End With
End Sub


