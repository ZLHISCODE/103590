VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTendFileElement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ҫ��¼��"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   Icon            =   "frmTendFileElement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   3615
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ȫ�����(&D)"
      Height          =   350
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   10
      Top             =   3840
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid VsfData 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6225
      _cx             =   10980
      _cy             =   6376
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTendFileElement.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
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
      AutoSizeMouse   =   0   'False
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
      Begin MSComCtl2.MonthView mthView 
         Height          =   2220
         Left            =   720
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   181927937
         CurrentDate     =   40899
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   1425
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
         Begin MSComCtl2.UpDown UD 
            Height          =   300
            Left            =   241
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txtInput"
            BuddyDispid     =   196618
            OrigLeft        =   120
            OrigRight       =   375
            OrigBottom      =   255
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CommandButton cmdDown 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            Picture         =   "frmTendFileElement.frx":006E
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   0
         ItemData        =   "frmTendFileElement.frx":03B0
         Left            =   3600
         List            =   "frmTendFileElement.frx":03C6
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   1
         ItemData        =   "frmTendFileElement.frx":03FE
         Left            =   4530
         List            =   "frmTendFileElement.frx":0414
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmTendFileElement.frx":044C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "��ʾ:"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1680
      TabIndex        =   8
      Top             =   3840
      Width           =   450
   End
End
Attribute VB_Name = "frmTendFileElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ColEnum
    COL_NULL
    COL_Group
    COL_Name
    COL_value
End Enum

Const mlngColEditor As Long = 3
Private mlngFileID As Long  '�ļ�ID
Private mlngFileFormatID As Long '�ļ���ʽID
Private mlngPageNo As Long 'ҳ��
Private mrsElement As New ADODB.Recordset '����ӵĻ���Ҫ����Ϣ
Private mblnOK As Boolean
Private mblnInit As Boolean
Private mintFace As Integer 'Ҫ�ر�ʾ 0-�ı�;1-����;2-����;3-��ѡ;4-��ѡ
Private mintType As String   'Ҫ������ 0-��ֵ;1-�ı�;2-����;3-�߼�
Private mblnShow As Boolean
Private mblnBlowup As Boolean
Private mblnStart As Boolean
Private mblnChange As Boolean

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngFileFormatID As Long, ByVal lngPageNo As Long, ByVal rsElement As ADODB.Recordset, _
    Optional ByVal bytSize As Byte = 0) As Boolean
 '------------------------------------------------------
 '���ܣ���ɲ�������¼��
 '������frmParent :���ô������
 '      lngFileID :�ļ�ID
 '      lngFileFormatID :�ļ���ʽID
 '      lngPageNo:ҳ��
 '      rsPartogram ������ӵı��ϡ����±�ǩ���ݣ�������,�滻��,����,����,С��,��λ,��ʾ��,��ֵ��,���
 '------------------------------------------------------
    mblnOK = False
    mlngFileID = lngFileID
    mlngFileFormatID = lngFileFormatID
    mlngPageNo = lngPageNo
    Set mrsElement = rsElement
    mblnStart = True
    mblnBlowup = (bytSize = 1)
    Me.FontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    
    If Not zlRefresh Then Exit Function

    Me.Show vbModal, frmParent

    ShowMe = mblnOK
End Function

Private Function zlRefresh() As Boolean
'---------------------------------------
'ˢ��������Ϣ
'---------------------------------------
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand

    mblnInit = False
    mblnShow = False
    mblnChange = False

    Call InitCons
    '��ȡ���ϱ�������
    gstrSQL = _
        " Select '' ��, '���ϱ�ǩ' ������, d.Ҫ������, a.����" & vbNewLine & _
        " From �����ļ��ṹ D, �����ļ��ṹ P, ���˻���Ҫ������ A" & vbNewLine & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ' And d.Ҫ������ = a.����(+) And a.�ļ�id(+) = [2] And A.ҳ��(+)=[3]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ȡҪ����Ϣ", mlngFileFormatID, mlngFileID, mlngPageNo)
    
    Call InitTabFormat(rsTemp)
    mblnInit = True
    Call vsfData_AfterRowColChange(0, COL_Name, VsfData.FixedRows, COL_value)
    zlRefresh = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitCons()
    '��������ؼ�
    mintType = -1
    mintFace = -1
    picInput.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    mthView.Visible = False
    UD.Visible = False
    cmdDown.Visible = False
End Sub

Private Sub InitTabFormat(ByVal rsTemp As ADODB.Recordset)
    Dim i As Integer, j As Integer
    With VsfData
        .Cols = 4
        .Rows = 2
        .FixedCols = 3
        .FixedRows = 1
        .MergeCells = flexMergeFixedOnly
        .MergeCol(COL_Group) = True
        .TextMatrix(0, COL_NULL) = ""
        .ColWidth(COL_NULL) = 255
        .TextMatrix(0, COL_Group) = "������"
        .ColWidth(COL_Group) = 0
        .TextMatrix(0, COL_Name) = "Ҫ������"
        .ColWidth(COL_Name) = 1500
        .TextMatrix(0, COL_value) = "Ҫ������"
        .ColWidth(COL_value) = 3400
        .RowHeightMin = 300
        .FontSize = 9 + 9 * IIf(mblnBlowup = True, 1, 0) / 3
        .ColHidden(COL_Group) = True
        .ExtendLastCol = True
        '������ݰ�
        If rsTemp.RecordCount = 0 Then
            .Rows = 2
        Else
            rsTemp.MoveFirst
            mrsElement.Filter = 0
            j = .FixedRows
            For i = 1 To rsTemp.RecordCount
                mrsElement.Filter = "������='" & NVL(rsTemp!Ҫ������) & "' And �滻��<>1"
                If mrsElement.RecordCount > 0 Then
                    If .Rows <= j Then .Rows = .Rows + 1
                    .TextMatrix(j, COL_NULL) = NVL(rsTemp!��)
                    .TextMatrix(j, COL_Group) = NVL(rsTemp!������)
                    .TextMatrix(j, COL_Name) = Replace(NVL(rsTemp!Ҫ������), ";", "")
                    .TextMatrix(j, COL_value) = Replace(NVL(rsTemp!����), "[ZLSOFTLPF]", "")
                    .MergeRow(j) = True
                    j = j + 1
                End If
                If i < rsTemp.RecordCount Then rsTemp.MoveNext
            Next i
        End If

        .COL = COL_value: .ROW = 1
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .FocusRect = flexFocusSolid

        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
    End With
End Sub

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub

    If Not (objVsf.Cell(flexcpPicture, intRow, COL_NULL) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, COL_NULL, objVsf.Rows - 1, COL_NULL) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, COL_NULL) = ils16.ListImages(1).Picture
End Sub


Private Sub cmdCancle_Click()
    mblnChange = False
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Dim arrData
    Dim i As Integer, j As Integer
    Dim strName As String, strBound As String
    Dim intIndex As Integer
    Dim CellRect As RECT
    Dim strValue As String

    If mblnShow = False Or mintFace = -1 Then Exit Sub

    CellRect.Left = picInput.Left
    CellRect.Top = picInput.Top + picInput.Height
    CellRect.Bottom = VsfData.CellHeight
    CellRect.Right = VsfData.CellWidth
    strValue = Trim(txtInput.Text)
    If mintType = 2 Then '��������
        With mthView
            If IsDate(strValue) Then
                .Value = Format(strValue, "YYYY-MM-DD")
            Else
                .Value = Format(zldatabase.Currentdate, "YYYY-MM-DD")
            End If
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Font.Name = VsfData.FontName
            .Font.Size = VsfData.FontSize
            If .Height + .Top > VsfData.Height Then
                .Top = VsfData.Height - .Height
            End If
            If .Width < CellRect.Right Then
                .Left = CellRect.Right + CellRect.Left - .Width
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    Else '�ı�����
        strName = VsfData.TextMatrix(VsfData.ROW, COL_Name)
        mrsElement.Filter = "������='" & strName & "'"
        strBound = NVL(mrsElement!��ֵ��)
        If Left(strBound, 1) = ";" Then strBound = Mid(strBound, 2)
        If strBound <> "" Then strBound = ";" & strBound
        intIndex = 0
        lstSelect(intIndex).Clear
        arrData = Split(strBound, ";")
        j = UBound(arrData)
        lstSelect(intIndex).AddItem 0 & "-"
        lstSelect(intIndex).ListIndex = 0
        For i = 1 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "��" Then
                    lstSelect(intIndex).AddItem i & "-" & Mid(arrData(i), 2)
                    lstSelect(intIndex).ListIndex = i
                Else
                    lstSelect(intIndex).AddItem i & "-" & arrData(i)
                End If
            End If
        Next
        
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(intIndex).List(i), InStr(1, lstSelect(intIndex).List(i), "-") + 1) & ",") <> 0 Then
                    lstSelect(intIndex).Selected(i) = True
                End If
            Next
        End If
        '��ʾ
        With lstSelect(intIndex)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top > VsfData.Height Then
                .Top = VsfData.Height - .Height
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim lngRow As Long, blnShow As Boolean
    With VsfData
        For lngRow = .FixedRows To .Rows - 1
            .TextMatrix(lngRow, COL_value) = ""
        Next lngRow
    End With
    
    mblnChange = True
    
    blnShow = mblnShow
    mblnShow = False
    Call VsfData_EnterCell
    mblnShow = blnShow
End Sub

Private Sub cmdOK_Click()
    Dim strPara As String, strSQL As String
    Dim strName As String, strValue As String
    Dim intRow As Integer
    If mblnChange = True Then
        For intRow = VsfData.FixedRows To VsfData.Rows - 1
            If InStr(1, strName, "[ZLSOFTLPF]" & VsfData.TextMatrix(intRow, COL_Name) & "[ZLSOFTLPF]") = 0 And Trim(VsfData.TextMatrix(intRow, COL_value)) <> "" Then
                strName = strName & "[ZLSOFTLPF]" & VsfData.TextMatrix(intRow, COL_Name) & "[ZLSOFTLPF]"
                strPara = strPara & "[ZLSOFTLPF]" & VsfData.TextMatrix(intRow, COL_Name) & ";" & VsfData.TextMatrix(intRow, COL_value)
            End If
        Next intRow
        If Left(strPara, 11) = "[ZLSOFTLPF]" Then strPara = Mid(strPara, 12)
        '����������Ϣ
        strSQL = "Zl_���˻���Ҫ������_Update("
        '�ļ�ID_IN IN ����Ҫ������.�ļ�ID%TYPE,
        strSQL = strSQL & mlngFileID & ","
        'ҳ��_In   In ���˻���Ҫ������.ҳ�� %Type
        strSQL = strSQL & mlngPageNo & ",'"
        'strPara IN Varchar2 --������ʽΪ��Ҫ������;Ҫ������|Ҫ������;Ҫ������
        strSQL = strSQL & strPara & "','" & gstrUserName & "')"
        Call zldatabase.ExecuteProcedure(strSQL, "Zl_���˻���Ҫ������_Update")
        mblnChange = False
        mblnOK = True
        Unload Me
    Else
        Call cmdCancle_Click
    End If
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then Exit Sub
    Me.Width = Me.Width + Me.Width * IIf(mblnBlowup = True, 1, 0) / 3
    Me.Height = Me.Height + Me.Height * IIf(mblnBlowup = True, 1, 0) / 3
    Me.FontSize = Me.FontSize + Me.FontSize * IIf(mblnBlowup = True, 1, 0) / 3
    lblInfo.FontSize = Me.FontSize
    mblnStart = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ���Ϊ���ݷָ�������¼�¼���ķָ�������˲�����¼��
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call InitCons
    End If
End Sub

Private Sub Form_Resize()

    With cmdDelete
        .Left = 120
        .Top = Me.ScaleHeight - .Height - 120
    End With

    With cmdCancle
        .Top = cmdDelete.Top
        .Left = Me.ScaleWidth - .Width - 120
    End With

    With cmdOk
        .Top = cmdCancle.Top
        .Left = cmdCancle.Left - .Width - 120
    End With

    With fraLine
        .Left = 60
        .Top = cmdOk.Top - 60
        .Width = Me.ScaleWidth - 120
    End With

    With lblInfo
        .Left = 120
        .Top = fraLine.Top - TextHeight("��") - 60
    End With

    With VsfData
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = lblInfo.Top - 60
    End With
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strText As String
    Dim intIndex As Integer
    If KeyCode = vbKeyReturn Then
        If mintFace = 2 Then '�ı�����
            If InStr(1, lstSelect(Index).Text, "-") <> 0 Then
                strText = Split(lstSelect(Index).Text, "-")(1)
            Else
                strText = ""
            End If
            txtInput.Text = strText
            lstSelect(Index).Visible = False
            If picInput.Visible = True Then picInput.SetFocus
        Else
            Call MoveNextCell
        End If
    ElseIf (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) And Shift = vbShiftMask Then
        Call MoveNextCell(KeyCode, 0)
    End If
    
End Sub

Private Sub mthView_DateDblClick(ByVal DateDblClicked As Date)
    txtInput.Text = Format(DateDblClicked, "YYYY-MM-DD")
    mthView.Visible = False
    If picInput.Visible = True Then picInput.SetFocus
End Sub

Private Sub mthView_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsDate(mthView.Value) Then Call mthView_DateDblClick(CDate(mthView.Value))
    End If
End Sub

Private Sub picInput_GotFocus()
    If txtInput.Visible = True Then
        txtInput.SetFocus
    End If
End Sub

Private Sub picInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> vbShiftMask Then
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            Call MoveNextCell(KeyCode, Shift)
        End If
    Else
        If KeyCode = vbKeyDown And mintFace = 2 Then
            Call cmdDown_Click
        End If
    End If
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
    If lstSelect(0).Visible = True And ((mintFace = 2 And mintType = 1) Or mintFace = 4) Then
        lstSelect(0).SetFocus
    ElseIf mthView.Visible = True And mintFace = 2 And mintType = 2 Then
        mthView.SetFocus
    ElseIf lstSelect(1).Visible = True And mintFace = 3 Then
        lstSelect(1).SetFocus
    End If
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picInput_KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strName As String, strBound As String, strInfo As String
    
    If mblnInit = False Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    Call AdjustRowFlag(VsfData, NewRow)
    strName = VsfData.TextMatrix(NewRow, COL_Name)
    If mrsElement Is Nothing Then Exit Sub
    
    '��ʾ������Ŀֵ����Ϣ
    mrsElement.Filter = 0
    mrsElement.Filter = "������='" & strName & "'"
    If mrsElement.RecordCount > 0 Then
        strBound = NVL(mrsElement!��ֵ��, "")
        If Left(strBound, 1) = ";" Then strBound = Mid(strBound, 2)
        If strBound <> "" Then
            If Val(NVL(mrsElement!����, 0)) = 0 Then
                strInfo = "��ֵ��:" & Split(strBound, ";")(0) & "��" & Split(strBound, ";")(1)
            Else
                strInfo = "��ֵ��;" & strBound
            End If
        End If
        If Val(NVL(mrsElement!��ʾ��, 0)) = 2 Then
            strInfo = strInfo & IIf(strInfo = "", "", Space(2)) & "[��SHIFT+������������]"
        End If
    End If
    
    lblInfo.Caption = "��ʾ��" & strInfo
    lblInfo.Tag = lblInfo.Caption
End Sub

Private Sub VsfData_DblClick()
    Call vsfdata_KeyDown(Asc("A"), 0)
End Sub

Private Sub VsfData_EnterCell()

    '��������ʾ�Ŀؼ�
    Select Case mintFace
        Case 0, 1, 2
            picInput.Visible = False
            If mintFace = 2 Then
                lstSelect(0).Visible = False
            End If
        Case 3
            lstSelect(1).Visible = False
        Case 4
            lstSelect(0).Visible = False
    End Select
    mthView.Visible = False
    UD.Visible = False
    cmdDown.Visible = False

    mintType = -1: mintFace = -1
    If mblnShow = False Or VsfData.COL <> COL_value Then Exit Sub

    Call ShowInput

    '��ȡ����
    Select Case mintFace
        Case 0, 1, 2
            picInput.SetFocus
        Case 3
            lstSelect(1).SetFocus
        Case 4
            lstSelect(0).SetFocus
    End Select
End Sub

Private Sub vsfdata_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngStart As Long
    Dim intLevel As Integer
    Dim strField As String, strKey As String, strValue As String
    On Error GoTo ErrHand

    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ShowInput(Optional ByVal intRow As Integer = -1) As Boolean
'��ʾ��Ӧ�� �༭�ؼ�
    Dim CellRect As RECT
    Dim arrData
    Dim intIndex As Integer, i As Integer, j As Integer, k As Integer
    Dim strItemName As String, strBound As String, strLen As String, strValue As String

    If intRow = -1 Then intRow = VsfData.ROW
    CellRect.Left = VsfData.CellLeft + VsfData.Left
    CellRect.Top = VsfData.CellTop + VsfData.Top
    CellRect.Bottom = VsfData.CellHeight + 0
    CellRect.Right = VsfData.CellWidth + 0

    mintType = -1
    mintFace = -1
    strItemName = VsfData.TextMatrix(intRow, COL_Name)
    strValue = VsfData.TextMatrix(intRow, COL_value)
    '������,�滻��,����,����,С��,��λ,��ʾ��,��ֵ��,����
    mrsElement.Filter = 0
    mrsElement.Filter = "������='" & strItemName & "'"
    If mrsElement.RecordCount = 0 Then Exit Function
    'ȷ����Ŀ����
    mintFace = Val(NVL(mrsElement!��ʾ��, 0))
    mintType = Val(NVL(mrsElement!����, 0))
    strBound = NVL(mrsElement!��ֵ��)
    If Left(strBound, 1) = ";" Then strBound = Mid(strBound, 2)
    strLen = Val(NVL(mrsElement!����)) & ";" & Val(NVL(mrsElement!С��))
    '����Ϊ�߼�����Ϊ �ļ�����
    If mintType = 3 Then mintFace = 2: mintType = 1
    Select Case mintFace
    Case 0, 1, 2
        With picInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            If .Height + .Top > VsfData.Height Then
                .Top = VsfData.Height - .Height
            End If
            .Visible = True
        End With
        '�ı���������Ŀ
        txtInput.Visible = True
        If Val(strLen) <> 0 Then
            txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
        Else
            txtInput.MaxLength = 0
        End If

        With txtInput
            .Top = 0
            .Text = strValue
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Tag = .Text
            If mintFace = 1 Then
                arrData = Split(strBound, ";")
                UD.Top = 0
                .Width = .Width - UD.Width
                UD.Left = .Width
                UD.Height = .Height
                UD.Min = 0: UD.Max = 10
                UD.Increment = 1
                If UBound(arrData) > 0 Then
                    UD.Min = Val(arrData(0))
                    UD.Max = Val(arrData(1))
                End If
                UD.Visible = True
            ElseIf mintFace = 2 Then
                cmdDown.Top = 0
                .Width = .Width - cmdDown.Width
                cmdDown.Left = .Width
                cmdDown.Height = .Height
                cmdDown.Visible = True
            End If
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '����9��ʱ��ȥ90,����Խ��۳��ı߾�ԽС,�Ա�֤�ı��������ʵ��һ��
        End With
    Case 3, 4
        intIndex = IIf(mintFace = 3, 1, 0)
        '��������
        lstSelect(intIndex).Clear
        If Left(strBound, 1) = ";" Then strBound = Mid(strBound, 2)
        If strBound <> "" Then strBound = ";" & strBound
        k = 1
        If intIndex = 0 Then
            lstSelect(intIndex).AddItem 0 & "-"
            lstSelect(intIndex).ListIndex = 0
        End If
        arrData = Split(strBound, ";")
        j = UBound(arrData)
        For i = k To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "��" Then
                    lstSelect(intIndex).AddItem lstSelect(intIndex).NewIndex + 1 & "-" & Mid(arrData(i), 2)
                    lstSelect(intIndex).ListIndex = lstSelect(intIndex).NewIndex
                Else
                    lstSelect(intIndex).AddItem lstSelect(intIndex).NewIndex + 1 & "-" & arrData(i)
                End If
            End If
        Next
        '��ѡ����¼�����ݵ������
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(intIndex).List(i), InStr(1, lstSelect(intIndex).List(i), "-") + 1) & ",") <> 0 Then
                    lstSelect(intIndex).Selected(i) = True
                End If
            Next
        End If
        '��ʾ
        With lstSelect(intIndex)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top > VsfData.Height Then
                .Top = VsfData.Height - .Height
            End If
            .Visible = True
            .Tag = strValue
        End With
    End Select

    ShowInput = True
End Function

Private Sub MoveNextCell(Optional KeyCode As Integer = vbKeyReturn, Optional Shift As Integer = 0)
'��������У��͵�Ԫ���ƶ�
    Dim intRow As Integer
    Dim strRetrun As String, strErrMsg As String
    Dim blnShow As Boolean
    
    If mintFace >= 0 And Shift = vbShiftMask And (KeyCode = vbKeyUp Or KeyCode = vbKeyDown) Then Exit Sub
    If mintFace >= 0 And KeyCode = vbKeyReturn Then
        '�������У��ͱ���
        If Not CheckInput(strRetrun, strErrMsg) Then
            lblInfo.Caption = "��ʾ��" & strErrMsg
            Exit Sub
        Else
            lblInfo.Caption = lblInfo.Tag
        End If
        '��ɸ�ֵ����
        VsfData.TextMatrix(VsfData.ROW, COL_value) = Replace(strRetrun, "[ZLSOFTLPF]", "")
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
toMoveNextCol:
        If VsfData.COL < mlngColEditor Then
            VsfData.COL = VsfData.COL + 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMoveNextCol
        Else
toMoveNextRow:
            '������һ��
            intRow = 1
            If VsfData.ROW + intRow < VsfData.Rows Then
                VsfData.ROW = VsfData.ROW + intRow
            Else
                blnShow = mblnShow
                mblnShow = False
                Call VsfData_EnterCell
                mblnShow = blnShow
            End If
            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMoveNextRow
            VsfData.COL = COL_value
        End If
    Else
toMovePrevCol:
        If VsfData.COL > mlngColEditor Then
            VsfData.COL = VsfData.COL - 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '������һ��
            intRow = 1
            If VsfData.ROW > VsfData.FixedRows Then
                VsfData.ROW = VsfData.ROW - intRow
            End If
            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMovePrevRow
            VsfData.COL = COL_value
        End If
    End If
    If VsfData.ColIsVisible(VsfData.COL) = False Then
        VsfData.LeftCol = VsfData.COL
    End If
    If VsfData.RowIsVisible(VsfData.ROW) = False Then
        VsfData.TopRow = VsfData.ROW
    End If
End Sub

Private Function CheckInput(strReturn As String, strInfo As String) As Boolean
    Dim i As Integer, j As Integer
    Dim strText As String, strOldText As String
    Dim intIndex As Integer
    Dim arrDate
    '���¼�����ݵĺϷ���(����Ҳ��Ϊ��һ���ַ�,���ǵ�������Ŀ�ȴ��ڲ���\�������Ϣ)
    '���ص�����,���һ�а󶨶����Ŀ,�Ե�������Ϊ�ָ���

    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�N����Ŀ,�ֹ�¼��
    Select Case mintFace
    Case 0, 1, 2
        strText = txtInput.Text
        strOldText = txtInput.Tag
    Case 3, 4   '���
        intIndex = IIf(mintFace = 3, 2, 1)
        If mintFace = 4 Then
            If InStr(1, lstSelect(intIndex - 1).Text, "-") <> 0 Then
                strText = Split(lstSelect(intIndex - 1).Text, "-")(1)
            Else
                strText = ""
            End If
        Else
            j = lstSelect(intIndex - 1).ListCount
            For i = 1 To j
                If lstSelect(intIndex - 1).Selected(i - 1) Then
                    strText = strText & "," & Split(lstSelect(intIndex - 1).List(i - 1), "-")(1)
                End If
            Next
            If strText <> "" Then strText = Mid(strText, 2)
        End If
        strOldText = lstSelect(intIndex - 1).Tag
    End Select

    If mintType = 0 Or (mintType = 1 And mintFace = 1) Then '��ֵ������Ҫ���
        If Not CheckValid(strText, strInfo) Then Exit Function
    ElseIf mintType = 2 Then '��������
        If strText <> "" Then
            If InStr(1, strText, "-") = 0 Then
                If IsNumeric(strText) = False Then
                    strInfo = "���ڲ��ܰ�����-��������ַ�,����!"
                    Exit Function
                End If
                If Len(strText) <> 8 Then
                     strInfo = "���ڸ�ʽֻ��Ϊ[YYYY-MM-DD]��[YYYYMMDD],����!"
                     Exit Function
                End If
                strText = Mid(strText, 1, 4) & "-" & Mid(strText, 5, 2) & "-" & Mid(strText, 7, 2)
            Else
                If Left(strText, 1) = "-" Or Right(strText, 1) = "-" Then
                    strInfo = "���ڿ�ʼ�ͽ�β���ܴ��ڡ�-���ַ�,����!"
                    Exit Function
                End If
            End If
            arrDate = Split(strText, "-")
            If UBound(arrDate) <> 2 Then
                strInfo = "���ڸ�ʽֻ��Ϊ[YYYY-MM-DD]��[YYYYMMDD],����!"
                Exit Function
            End If
            For intIndex = 0 To UBound(arrDate)
                If IsNumeric(CStr(arrDate(intIndex))) = False Then
                    strInfo = "���ڵ�������ֻ��Ϊ����,����!"
                    Exit Function
                End If
                If intIndex = 0 Then
                    If Len(CStr(arrDate(intIndex))) > 4 Then
                        strInfo = "������ݳ��Ȳ��ܳ���4λ,����!"
                        Exit Function
                    End If
                ElseIf intIndex = 1 Then
                    If Len(CStr(arrDate(intIndex))) > 2 Then
                        strInfo = "�����·ݳ��Ȳ��ܳ���2λ,����!"
                        Exit Function
                    End If
                Else
                    If Len(CStr(arrDate(intIndex))) > 2 Then
                        strInfo = "�����������Ȳ��ܳ���2λ,����!"
                        Exit Function
                    End If
                End If
            Next
            If Not IsDate(Format(strText, "YYYY-MM-DD")) Then
                strInfo = "¼������ڲ�����Ч������,����!"
                Exit Function
            End If
            strText = Format(strText, "YYYY-MM-DD")
        End If
    End If
    If strText <> strOldText Then mblnChange = True
    strReturn = strText
    CheckInput = True
End Function

Private Function CheckValid(strReturn As String, strInfo As String) As Boolean
    Dim blnCheck As Boolean
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String
    '�������

    On Error GoTo ErrHand

    strName = VsfData.TextMatrix(VsfData.ROW, COL_Name)
    strText = strReturn
    mrsElement.Filter = 0
    mrsElement.Filter = "������='" & strName & "'"
    If strText <> "" Then
        blnCheck = True
        '�����������Ŀ,�������Ĳ����������򲻼��
        If Val(NVL(mrsElement!����)) = 0 Then
            If Not IsNumeric(Trim(strText)) Then
                blnCheck = False
            End If
        End If

        If blnCheck Then
            If Val(NVL(mrsElement!����, 0)) = 0 Then
                strText = Val(strText)
                If Val(NVL(mrsElement!С��, 0)) <> 0 Then   '����ͨ���ؼ���MaxLength�����Ƶ�
                    If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                    If Len(strText) > Val(NVL(mrsElement!����)) Then
                        mrsElement.Filter = 0
                        strInfo = "[" & strName & "]¼������ݳ����˺Ϸ����ȣ�"
                        Exit Function
                    End If
                End If

                If Val(Replace(NVL(mrsElement!��ֵ��), ";", "")) <> 0 Then
                    dblMin = Val(Split(mrsElement!��ֵ��, ";")(0))
                    dblMax = Val(Split(mrsElement!��ֵ��, ";")(1))
                    If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                        mrsElement.Filter = 0
                        strInfo = "[" & strName & "]¼������ݲ���" & Format(dblMin, "#0.00") & "��" & Format(dblMax, "#0.00") & "����Ч��Χ��"
                        Exit Function
                    End If
                End If
                If Val(NVL(mrsElement!С��, 0)) > 0 Then
                    strText = Format(strText, "#0." & String(Val(NVL(mrsElement!С��, 0)), "0"))
                Else
                    strText = Format(strText, "#0")
                End If

            Else
                If LenB(StrConv(strText, vbFromUnicode)) > mrsElement!���� Then
                    strInfo = "[" & strName & "]¼������ݳ�������󳤶ȣ�" & mrsElement!���� & "��"
                    mrsElement.Filter = 0
                    Exit Function
                End If
            End If
        End If
    End If

    strReturn = strText
    CheckValid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
