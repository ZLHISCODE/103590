VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmCompendWord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ʾ�ʾ��"
   ClientHeight    =   6540
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6165
   Icon            =   "frmCompendWord.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkChildren 
      Caption         =   "ͬ������¼�����(&2)"
      Height          =   195
      Index           =   1
      Left            =   2505
      TabIndex        =   5
      Top             =   5730
      Width           =   2055
   End
   Begin VB.CheckBox chkChildren 
      Caption         =   "ͬ��ѡ���¼�����(&1)"
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   5730
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox chkOnlySel 
      Caption         =   "����ʾ��ѡ�����(&S)"
      Height          =   195
      Left            =   165
      TabIndex        =   6
      Top             =   6135
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3750
      TabIndex        =   2
      Top             =   6060
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4890
      TabIndex        =   3
      Top             =   6060
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   5265
      Left            =   150
      TabIndex        =   1
      Top             =   390
      Width           =   5835
      _cx             =   10292
      _cy             =   9287
      Appearance      =   2
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmCompendWord.frx":000C
      Top             =   75
      Width           =   240
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "������ʵ�ʱ༭�����У���ǰ��ٹ�����ѡ�Ĵʾ�ʾ�����ࡣ"
      Height          =   180
      Left            =   435
      TabIndex        =   0
      Top             =   105
      Width           =   4860
   End
End
Attribute VB_Name = "frmCompendWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    ѡ�� = 0: ID: �ϼ�id: ����: ����: ˵��
End Enum

Private mlngCompendId As Long   '��ǰ���ID
Private mblnOK As Boolean       '�Ƿ�ȷ��

Dim lngCount As Long

'-----------------------------------------------------
'����Ϊ�ⲿ��������
'-----------------------------------------------------
Public Function ShowMe(ByVal frmParent As Form, ByVal lngCompendID As Long, bytFileType As Byte) As Boolean
    '���ܣ���ʾ���༭����
    '������ frmParent-������
    '       lngCompendId-���ID
    '       bytFileType-�ļ�����
    Dim rsTemp As New ADODB.Recordset
    mlngCompendId = lngCompendID
    
    'װ���ѡ��������
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select Decode(U.�ʾ����id, Null, 0, 1) As ѡ��, C.ID, C.�ϼ�id, C.����, C.����, C.˵��" & vbNewLine & _
            "From �����ʾ���� C, (Select �ʾ����id From ������ٴʾ� Where ���id = [1]) U" & vbNewLine & _
            "Where C.ID = U.�ʾ����id(+) And Substr(��Χ, [2], 1) = '1'" & vbNewLine & _
            "Order By C.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngCompendID, bytFileType)
    With Me.vfgList
        .Redraw = flexRDNone
        Set .DataSource = rsTemp
        .ColWidth(mCol.ѡ��) = 280
        .ColWidth(mCol.ID) = 0: .ColHidden(mCol.ID) = True
        .ColWidth(mCol.�ϼ�id) = 0: .ColHidden(mCol.�ϼ�id) = True
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, mCol.ѡ��)) = 1 Then
                .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked
            Else
                .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexUnchecked
            End If
            .TextMatrix(lngCount, mCol.ѡ��) = ""
        Next
        If .Rows > .FixedRows Then .Row = .FixedRows
        .Col = mCol.ѡ��
        .Redraw = flexRDDirect
    End With
    
    Me.Show vbModal, frmParent
    ShowMe = mblnOK
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub chkChildren_Click(Index As Integer)
    With Me.vfgList
        If .Visible And .Enabled Then .SetFocus
    End With
End Sub

Private Sub chkOnlySel_Click()
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If Me.chkOnlySel.Value = vbChecked Then
                If .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexUnchecked Then
                    .RowHidden(lngCount) = True
                End If
            Else
                .RowHidden(lngCount) = False
            End If
        Next
        If .Visible And .Enabled Then .SetFocus
    End With
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strClass As String
    
    Err = 0: On Error GoTo errHand
    strClass = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked Then
                strClass = strClass & "," & .TextMatrix(lngCount, mCol.ID)
            End If
        Next
    End With
    If strClass <> "" Then strClass = Mid(strClass, 2)
    
    gstrSQL = "Zl_������ٴʾ�_Update(" & mlngCompendId & ",'" & strClass & "')"
    zlDatabase.ExecuteProcedure gstrSQL, "����ʾ����"
    mblnOK = True: Unload Me
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vfgList_DblClick()
    With Me.vfgList
        If .Row < .FixedRows Then Exit Sub
        If .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexUnchecked Then
            .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexChecked
            If Me.chkChildren(0).Value = vbChecked Then
                For lngCount = .Row To .Rows - 1
                    If Val(.TextMatrix(lngCount, mCol.�ϼ�id)) = Val(.TextMatrix(.Row, mCol.ID)) Then
                        .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexChecked
                    End If
                Next
            End If
        Else
            .Cell(flexcpChecked, .Row, mCol.ѡ��) = flexUnchecked
            If Me.chkChildren(1).Value = vbChecked Then
                For lngCount = .Row To .Rows - 1
                    If Val(.TextMatrix(lngCount, mCol.�ϼ�id)) = Val(.TextMatrix(.Row, mCol.ID)) Then
                        .Cell(flexcpChecked, lngCount, mCol.ѡ��) = flexUnchecked
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii = vbKeySpace Then Call vfgList_DblClick
End Sub
