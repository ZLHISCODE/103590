VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportOrder 
   Caption         =   "��Һ��¼ѡ��"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10440
   Icon            =   "frmImportOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10440
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkDay 
      Caption         =   "ֻ��ʾ���3��ķ���ҽ��"
      Height          =   200
      Left            =   60
      TabIndex        =   12
      Top             =   5940
      Width           =   2535
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "����(&M)"
      Height          =   350
      Left            =   9255
      TabIndex        =   11
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "����(&U)"
      Height          =   350
      Left            =   9255
      TabIndex        =   10
      Top             =   4365
      Width           =   1100
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   2670
      MousePointer    =   9  'Size W E
      TabIndex        =   9
      Top             =   0
      Width           =   45
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1050
      Top             =   2160
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
            Picture         =   "frmImportOrder.frx":6852
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportOrder.frx":6DEC
            Key             =   "Expend"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDel 
      Cancel          =   -1  'True
      Caption         =   "ɾ��(&D)"
      Height          =   350
      Left            =   9255
      TabIndex        =   7
      Top             =   3795
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   9255
      TabIndex        =   6
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9255
      TabIndex        =   5
      Top             =   5775
      Width           =   1100
   End
   Begin VB.Frame fraUD 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2760
      TabIndex        =   2
      Top             =   3435
      Width           =   7575
      Begin VB.Label lblDetail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "˵��"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   60
         TabIndex        =   3
         Top             =   30
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   3180
      Left            =   2745
      TabIndex        =   0
      Top             =   90
      Width           =   7605
      _cx             =   13414
      _cy             =   5609
      Appearance      =   1
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
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmImportOrder.frx":7386
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
      ExplorerBar     =   5
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
      Begin VB.CheckBox chkAll 
         Height          =   180
         Left            =   375
         TabIndex        =   4
         Top             =   30
         Width           =   195
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   240
         Top             =   1800
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
               Picture         =   "frmImportOrder.frx":7522
               Key             =   "δ����"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmImportOrder.frx":DD84
               Key             =   "�ѵ���"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfItem 
      Height          =   2400
      Left            =   2760
      TabIndex        =   1
      Top             =   3795
      Width           =   6315
      _cx             =   11139
      _cy             =   4233
      Appearance      =   1
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
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmImportOrder.frx":145E6
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
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   5730
      Left            =   45
      TabIndex        =   8
      Top             =   90
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   10107
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmImportOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng�ļ�ID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mintӤ�� As Integer
Private mstrImPortOrder As String '���ڲ��Ҹ���Ŀ�Ƿ���ҽ������
Private mstrImPortName As String
Private mstrDate As String  'ȱʡ��ȡ����ʱ��
Private mblnOK As Boolean
Private mblnLoadOver As Boolean
Private mrsItems As ADODB.Recordset
Private mrsFileData As ADODB.Recordset

Public Function ShowMe(frmParent As Object, ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intӤ�� As Integer, _
        ByVal strImPortOrder As String, Optional ByVal strDate As String = "") As Recordset
'strImPortOrder��ʽ:��Ŀ���,ҽ������;��Ŀ���,ҽ������
    Dim arrImport() As String, i As Integer
    mlng�ļ�ID = lng�ļ�ID
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mintӤ�� = intӤ��
    mstrImPortOrder = "": mstrImPortName = ""
    
    If strImPortOrder <> "" Then
        arrImport = Split(strImPortOrder, ";")
        For i = 0 To UBound(arrImport)
            mstrImPortOrder = mstrImPortOrder & "," & Split(arrImport(i), ",")(0)
            mstrImPortName = mstrImPortName & "," & Split(arrImport(i), ",")(1)
        Next
    End If
    If Left(mstrImPortOrder, 1) = "," Then mstrImPortOrder = Mid(mstrImPortOrder, 2)
    If Left(mstrImPortName, 1) = "," Then mstrImPortName = Mid(mstrImPortName, 2)
    
    If Not IsDate(strDate) Then strDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
    mstrDate = strDate
    mblnOK = False
    mblnLoadOver = False
    Set mrsItems = Nothing
    Set mrsFileData = Nothing
    
    Me.Show 1, frmParent
    If mblnOK = True Then
        Set ShowMe = mrsItems
    Else
        Set ShowMe = Nothing
    End If
    Set mrsItems = Nothing
End Function

Private Sub chkAll_Click()
    Dim lngRow As Long
    
    If mblnLoadOver = False Then Exit Sub
    With vsfList
        chkAll.Tag = "1"
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpData, lngRow, .ColIndex("ѡ��")) <> "" And Abs(Val(.TextMatrix(lngRow, .ColIndex("ѡ��")))) <> IIf(chkAll.Value = 0, 0, 1) Then
                .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(chkAll.Value = 0, 0, 1)
                Call vsfList_AfterEdit(lngRow, .ColIndex("ѡ��"))
            End If
        Next lngRow
        chkAll.Tag = ""
    End With
End Sub

Private Sub chkDay_Click()
    Call ShowTree
End Sub

Private Sub cmdCanCel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdDel_Click()
    'ɾ��
    Dim strKey As String
    Dim lngRow As Long
    With vsfItem
        If .ROW < .FixedRows Then Exit Sub
        strKey = .Cell(flexcpData, .ROW, .ColIndex("��Ŀ"))
        If strKey = "" Then Exit Sub
        '�Ƶ����һ����ɾ��
        .RowPosition(.ROW) = .Rows - 1
        .RemoveItem .Rows - 1
        If .Rows = .FixedRows Then
            .Rows = .Rows + 1
            .ROW = .FixedRows
        End If
    End With
    With vsfList
        For lngRow = .FixedRows To .Rows - 1
            If strKey = .Cell(flexcpData, lngRow, vsfList.ColIndex("ѡ��")) And Abs(Val(.TextMatrix(lngRow, .ColIndex("ѡ��")))) = 1 Then
                .TextMatrix(lngRow, .ColIndex("ѡ��")) = 0
                If chkAll.Value <> 0 Then
                    mblnLoadOver = False
                    chkAll.Value = 0
                    mblnLoadOver = True
                End If
                Exit For
            End If
        Next
    End With
End Sub

Private Sub cmdDown_Click()
    Dim strKey As String
    Dim lngRow As Long
    With vsfItem
        If .ROW < .FixedRows Then Exit Sub
        strKey = .Cell(flexcpData, .ROW, .ColIndex("��Ŀ"))
        If strKey = "" Then Exit Sub
        '�Ƶ����һ����ɾ��
        If .ROW < .Rows - 1 Then
            .RowPosition(.ROW) = .ROW + 1
            .ROW = .ROW + 1
            If .RowIsVisible(.ROW) = False Then .TopRow = .ROW
        End If
    End With
End Sub

Private Sub cmdOK_Click()
    mblnOK = SaveItems
    Unload Me
End Sub

Private Function SaveItems() As Boolean
'����ѡ�����Ϣ
    Dim lngRow As Long
    Dim strFileds As String
    
    On Error GoTo ErrHand
    '��ʼ����¼��
    strFileds = "key," & adLongVarChar & ",50|ҽ������," & adLongVarChar & ",100|����," & adDouble & ",16|�ܸ�����," & adDouble & ",16|ִ��Ƶ��," & adLongVarChar & ",20" & _
                "|��ҩĿ��," & adLongVarChar & ",10|��ҩ����," & adLongVarChar & ",1000|ҽ������," & adLongVarChar & ",100|��ҩ;��," & adLongVarChar & ",1000|��ʼִ��ʱ��," & adDate & ",16"
    Call Record_Init(mrsItems, strFileds)
    With vsfItem
        For lngRow = .FixedRows To .Rows - 1
            If Trim(.Cell(flexcpData, lngRow, .ColIndex("��Ŀ"))) <> "" Then
                mrsItems.AddNew
                mrsItems.Fields("key").Value = .Cell(flexcpData, lngRow, .ColIndex("��Ŀ"))
                mrsItems.Fields("����").Value = Val(NVL(.TextMatrix(lngRow, .ColIndex("����"))))
                mrsItems.Fields("ҽ������").Value = NVL(.TextMatrix(lngRow, .ColIndex("��Ŀ")))
                mrsItems.Fields("�ܸ�����").Value = Val(NVL(.TextMatrix(lngRow, .ColIndex("�ܸ�����"))))
                mrsItems.Fields("ִ��Ƶ��").Value = NVL(.TextMatrix(lngRow, .ColIndex("ִ��Ƶ��")))
                mrsItems.Fields("��ҩĿ��").Value = NVL(.TextMatrix(lngRow, .ColIndex("��ҩĿ��")))
                mrsItems.Fields("��ҩ����").Value = NVL(.TextMatrix(lngRow, .ColIndex("��ҩ����")))
                mrsItems.Fields("ҽ������").Value = NVL(.TextMatrix(lngRow, .ColIndex("ҽ������")))
                mrsItems.Fields("��ҩ;��").Value = NVL(.TextMatrix(lngRow, .ColIndex("��ҩ;��")))
                mrsItems.Fields("��ʼִ��ʱ��").Value = Format(NVL(.TextMatrix(lngRow, .ColIndex("��ʼִ��ʱ��"))), "YYYY-MM-DD HH:MM:SS")
                mrsItems.Update
            End If
        Next
    End With
    SaveItems = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowTree() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long, lngIndex As Long
    Dim objNode As Node, strMatch As String
    Dim strFileter As String
    Dim strBegin As String, strEnd As String
    
    On Error GoTo errH
        
    Screen.MousePointer = 11
    
    '��������ڼ�1��(������������)
    vsfList.Rows = vsfList.FixedRows
    vsfList.Rows = vsfList.Rows + 1
    
    lngIndex = 0
    If Not (chkDay.Value = 0) Then
        strEnd = Format(zlDatabase.Currentdate, "YYYY-MM-DD") & " 23:59:59"
        strBegin = Format(DateAdd("D", -2, CDate(strEnd)), "YYYY-MM-DD")
        strFileter = " And e.�״�ʱ�� between [4] and [5] "
    End If
    strSQL = " Select Rownum ���,�״�ʱ��" & vbNewLine & _
        " From (Select To_Date(To_Char(e.�״�ʱ��, 'YYYY-MM-DD'), 'YYYY-MM-DD') �״�ʱ��" & vbNewLine & _
        "       From ����ҽ������ e, ����ҽ����¼ a, ������ĿĿ¼ b, ����ҽ����¼ c, ������ĿĿ¼ d" & vbNewLine & _
        "       Where e.ҽ��id = a.Id " & strFileter & " And a.������� In ('5', '6', '7') And a.������Ŀid = b.Id And a.����id = [1] And a.��ҳid = [2] And" & vbNewLine & _
        "             a.Ӥ�� = [3] And c.������� = 'E' And c.ִ������ In (1,2,3,4) And d.Id = c.������Ŀid And a.���id = c.Id And" & vbNewLine & _
        "             Nvl(d.ִ�з���, 0) In (1, 2) And e.�״�ʱ�� is not null" & vbNewLine & _
        "       Group By To_Date(To_Char(e.�״�ʱ��, 'YYYY-MM-DD'), 'YYYY-MM-DD'))" & vbNewLine & _
        " Order By �״�ʱ�� DESC"
    If chkDay.Value = 0 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ����Ϣ", mlng����ID, mlng��ҳID, mintӤ��)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ����Ϣ", mlng����ID, mlng��ҳID, mintӤ��, CDate(strBegin), CDate(strEnd))
    End If
    '��Ӵʾ����
    tvw_s.Nodes.Clear
    tvw_s.Tag = ""
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!���, Format(rsTmp!�״�ʱ��, "YYYY-MM-DD"), "Close")
            objNode.Expanded = True
            objNode.ExpandedImage = "Expend"
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!���, tvwChild, Format(rsTmp!�״�ʱ��, "YYYY-MM-DD") & "=1", "��Һ��", "Close")
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!���, tvwChild, Format(rsTmp!�״�ʱ��, "YYYY-MM-DD") & "=2", "ע����", "Close")
            If Format(rsTmp!�״�ʱ��, "YYYY-MM-DD") = Format(mstrDate, "YYYY-MM-DD") Then lngIndex = objNode.Index
            rsTmp.MoveNext
        Loop
    Else
        Set objNode = tvw_s.Nodes.Add(, , "_", "����Һҽ��", "Close")
        objNode.ExpandedImage = "Expend"
        Screen.MousePointer = 0
        Exit Function
    End If
    
    If tvw_s.Nodes.Count > 0 And lngIndex > 0 Then
        tvw_s.Nodes(lngIndex).Selected = True
    End If
    If Not tvw_s.SelectedItem Is Nothing Then
        tvw_s.SelectedItem.Expanded = True
        tvw_s.SelectedItem.EnsureVisible
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End If
    
    Screen.MousePointer = 0
    ShowTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LocaItem(ByVal strKey As String)
' ����:����ҽ����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String, str���� As String
    Dim i As Long, lngRow As Long, blnSelAll As Boolean
    On Error GoTo ErrHand
    
    mblnLoadOver = False
    
    '��ȡ�Ѿ����������
    If mrsFileData Is Nothing Then
        gstrSQL = " Select b.Id, b.δ��˵��" & vbNewLine & _
            " From ���˻������� a, ���˻�����ϸ b" & vbNewLine & _
            " Where a.Id = b.��¼id And Instr([1], ',' || b.��Ŀ��� || ',') <> 0 And b.δ��˵�� Is Not Null And a.�ļ�id = [2]"
        Set mrsFileData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�Ѿ����������", "," & mstrImPortOrder & ",", mlng�ļ�ID)
    End If
    
    If strKey Like "_*" Then '����
        strBegin = Format(tvw_s.SelectedItem.Text, "YYYY-MM-DD")
        strEnd = Format(strBegin & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        str���� = ",1,2,"
    ElseIf strKey Like "*=*" Then
        strBegin = Format(Split(strKey, "=")(0), "YYYY-MM-DD")
        strEnd = Format(strBegin & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        str���� = "," & Split(strKey, "=")(1) & ","
    Else
        vsfList.Rows = vsfList.FixedRows
        chkAll.Value = 0
        mblnLoadOver = True
        Exit Sub
    End If
    
    gstrSQL = " Select a.Id, e.���ͺ�, a.����ҽ��, a.ҽ������, b.����, a.��������, b.���㵥λ ��λ, c.ҽ������ ��ҩ;��, e.�״�ʱ��, a.�ܸ�����, a.ִ��Ƶ��," & vbNewLine & _
            " Decode(a.��ҩĿ��, 2, '����', 'Ԥ��') As ��ҩĿ��, a.��ҩ����, a.ҽ������ " & vbNewLine & _
            " From ����ҽ������ e, ����ҽ����¼ a, ������ĿĿ¼ b, ����ҽ����¼ c, ������ĿĿ¼ d" & vbNewLine & _
            " Where e.ҽ��id = a.Id And a.������� In ('5', '6', '7') And a.������Ŀid = b.Id And a.����id = [1] And a.��ҳid = [2] And a.Ӥ�� = [3] And" & vbNewLine & _
            "      e.�״�ʱ�� Between [5] And [6] And c.������� = 'E' And c.ִ������ In (1,2,3,4) And d.Id = c.������Ŀid And a.���id = c.Id And" & vbNewLine & _
            "      Instr([4], ',' || Nvl(d.ִ�з���, 0) || ',') > 0" & vbNewLine & _
            " Order By e.�״�ʱ��, ���ͺ�"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "", mlng����ID, mlng��ҳID, mintӤ��, str����, CDate(strBegin), CDate(strEnd))
    
    blnSelAll = True
    vsfList.Redraw = flexRDNone
    vsfList.Rows = vsfList.FixedRows
    If Not rsTmp.EOF Then
        vsfList.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            vsfList.RowData(i) = Val(rsTmp!ID)
            vsfList.TextMatrix(i, vsfList.ColIndex("ѡ��")) = 0
            vsfList.TextMatrix(i, vsfList.ColIndex("����")) = NVL(rsTmp!����)
            vsfList.TextMatrix(i, vsfList.ColIndex("����")) = NVL(rsTmp!��������)
            vsfList.TextMatrix(i, vsfList.ColIndex("��λ")) = NVL(rsTmp!��λ)
            vsfList.TextMatrix(i, vsfList.ColIndex("�ܸ�����")) = NVL(rsTmp!�ܸ�����)
            vsfList.TextMatrix(i, vsfList.ColIndex("ִ��Ƶ��")) = NVL(rsTmp!ִ��Ƶ��)
            vsfList.TextMatrix(i, vsfList.ColIndex("��ҩĿ��")) = NVL(rsTmp!��ҩĿ��)
            vsfList.TextMatrix(i, vsfList.ColIndex("��ҩ����")) = NVL(rsTmp!��ҩ����)
            vsfList.TextMatrix(i, vsfList.ColIndex("ҽ������")) = NVL(rsTmp!ҽ������)
            vsfList.TextMatrix(i, vsfList.ColIndex("��ҩ;��")) = NVL(rsTmp!��ҩ;��)
            vsfList.TextMatrix(i, vsfList.ColIndex("��ʼִ��ʱ��")) = Format(NVL(rsTmp!�״�ʱ��), "YYYY-MM-DD HH:mm")
            vsfList.Cell(flexcpData, i, vsfList.ColIndex("ѡ��")) = Val(rsTmp!ID) & ":" & Val(rsTmp!���ͺ�)
            '����Ƿ�ͬ����
            mrsFileData.Filter = "δ��˵��='" & vsfList.Cell(flexcpData, i, vsfList.ColIndex("ѡ��")) & "'"
            vsfList.Cell(flexcpPicture, i, 0) = imgList.ListImages(IIf(mrsFileData.RecordCount > 0, 2, 1)).Picture
            vsfList.Cell(flexcpData, i, 0) = IIf(mrsFileData.RecordCount > 0, 2, 1)
            
            For lngRow = vsfItem.FixedRows To vsfItem.Rows - 1
                If vsfItem.Cell(flexcpData, lngRow, vsfItem.ColIndex("��Ŀ")) = vsfList.Cell(flexcpData, i, vsfList.ColIndex("ѡ��")) Then
                    vsfList.TextMatrix(i, vsfList.ColIndex("ѡ��")) = 1
                    Exit For
                End If
            Next lngRow
            If Abs(Val(vsfList.TextMatrix(i, vsfList.ColIndex("ѡ��")))) = 0 Then blnSelAll = False
            rsTmp.MoveNext
        Next
        vsfList.Cell(flexcpPictureAlignment, 1, 0, vsfList.Rows - 1, 0) = 4
        vsfList.ROW = vsfList.FixedRows
    Else
        vsfList.Rows = vsfList.FixedRows + 1
    End If
    vsfList.Redraw = flexRDDirect
    
    If rsTmp.RecordCount > 0 Then
        chkAll.Value = IIf(blnSelAll = True, 1, 0)
    Else
        chkAll.Value = 0
    End If
    mblnLoadOver = True
    Exit Sub
ErrHand:
    mblnLoadOver = True
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdUp_Click()
    Dim strKey As String
    Dim lngRow As Long
    With vsfItem
        If .ROW < .FixedRows Then Exit Sub
        strKey = .Cell(flexcpData, .ROW, .ColIndex("��Ŀ"))
        If strKey = "" Then Exit Sub
        '�Ƶ����һ����ɾ��
        If .ROW > .FixedRows Then
            .RowPosition(.ROW) = .ROW - 1
            .ROW = .ROW - 1
            If .RowIsVisible(.ROW) = False Then .TopRow = .ROW
        End If
    End With
End Sub

Private Sub Form_Load()
    Dim arrName() As String, i As Integer
    chkDay.Value = 1
    '�ɵ���ҽ��������:|����|ҽ������|�ܸ�����|ִ��Ƶ��|��ҩĿ��|��ҩ����|ҽ������|��ʼִ��ʱ��|��ҩ;��
    arrName = Split(mstrImPortName, ",")
    For i = 0 To UBound(arrName)
        '�����к������й̶���ʾ����������������˰󶨵�������ʾ
        If arrName(i) <> "����" And arrName(i) <> "ҽ������" Then
            vsfItem.ColHidden(vsfItem.ColIndex(arrName(i))) = False
        End If
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    chkDay.Left = 60
    chkDay.Top = Me.ScaleHeight - chkDay.Height - 60
    
    tvw_s.Left = 0
    tvw_s.Top = 0
    tvw_s.Height = chkDay.Top - 60
    
    fraLR.Left = tvw_s.Left + tvw_s.Width
    fraLR.Top = 0
    fraLR.Height = tvw_s.Height
    
    vsfList.Top = 0
    vsfList.Left = fraLR.Left + fraLR.Width
    vsfList.Height = Me.ScaleHeight - vsfItem.Height - fraUD.Height
    vsfList.Width = Me.ScaleWidth - fraLR.Width - tvw_s.Width
    
    fraUD.Top = vsfList.Top + vsfList.Height
    fraUD.Left = vsfList.Left
    fraUD.Width = vsfList.Width
    
    cmdOK.Left = Me.ScaleWidth - cmdOK.Width - 60
    cmdCancel.Left = cmdOK.Left
    cmdDel.Left = cmdOK.Left
    cmdUp.Left = cmdOK.Left
    cmdDown.Left = cmdOK.Left
    
    vsfItem.Top = fraUD.Top + fraUD.Height
    vsfItem.Left = vsfList.Left
    vsfItem.Width = cmdOK.Left - vsfItem.Left - 120
    
    cmdDel.Top = vsfItem.Top
    cmdCancel.Top = Me.ScaleHeight - cmdCancel.Height - 60
    cmdOK.Top = cmdCancel.Top - cmdOK.Height - 30
    cmdUp.Top = cmdDel.Top + cmdDel.Height + 60
    cmdDown.Top = cmdUp.Top + cmdUp.Height + 30
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsFileData Is Nothing Then Set mrsFileData = Nothing
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsfList.Width - X < Me.ScaleWidth / 2 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvw_s.Width = tvw_s.Width + X
        chkDay.Width = chkDay.Width + X
        
        vsfList.Left = vsfList.Left + X
        vsfList.Width = vsfList.Width - X
        
        fraUD.Left = fraUD.Left + X
        fraUD.Width = fraUD.Width - X
        
        vsfItem.Left = vsfItem.Left + X
        vsfItem.Width = cmdOK.Left - vsfItem.Left - 120
        
        Me.Refresh
    End If
End Sub


Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key <> "_" Then
        If tvw_s.Tag = Node.Key Then Exit Sub
        tvw_s.Tag = Node.Key
        Call LocaItem(Node.Key)
    End If
End Sub

Private Sub vsfItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then Call cmdDel_Click
End Sub

Private Sub vsfList_AfterEdit(ByVal ROW As Long, ByVal COL As Long)
    '�±��ֵ����
    Dim lngRow As Long, strKey As String
    Dim blnTrue As Boolean, lngCurRow As Long
    
    If COL = vsfList.ColIndex("ѡ��") Then
        strKey = vsfList.Cell(flexcpData, ROW, COL)
        With vsfItem
            For lngRow = .FixedRows To .Rows - 1
                '����Ŀ��Ӧ�������ӻ�ȡ��
                If .Cell(flexcpData, lngRow, .ColIndex("��Ŀ")) = strKey And strKey <> "" Then
                    lngCurRow = lngRow
                    blnTrue = True
                    Exit For
                End If
            Next lngRow
            If Abs(Val(vsfList.TextMatrix(ROW, COL))) = 1 And blnTrue = False Then '��ѡ:����
                '���ҵ�������
                lngCurRow = -1
                For lngRow = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, lngRow, .ColIndex("��Ŀ")) = "" Then
                        lngCurRow = lngRow
                        Exit For
                    End If
                Next lngRow
                If lngCurRow = -1 Then lngCurRow = .Rows: .Rows = .Rows + 1
                
                .TextMatrix(lngCurRow, .ColIndex("��Ŀ")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("����"))
                .TextMatrix(lngCurRow, .ColIndex("����")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("����"))
                .TextMatrix(lngCurRow, .ColIndex("�ܸ�����")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("�ܸ�����"))
                .TextMatrix(lngCurRow, .ColIndex("ִ��Ƶ��")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("ִ��Ƶ��"))
                .TextMatrix(lngCurRow, .ColIndex("��ҩĿ��")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("��ҩĿ��"))
                .TextMatrix(lngCurRow, .ColIndex("��ҩ����")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("��ҩ����"))
                .TextMatrix(lngCurRow, .ColIndex("ҽ������")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("ҽ������"))
                .TextMatrix(lngCurRow, .ColIndex("��ҩ;��")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("��ҩ;��"))
                .TextMatrix(lngCurRow, .ColIndex("��ʼִ��ʱ��")) = vsfList.TextMatrix(ROW, vsfList.ColIndex("��ʼִ��ʱ��"))
                .Cell(flexcpData, lngCurRow, .ColIndex("��Ŀ")) = strKey
                .ROW = lngCurRow: .TopRow = .ROW
            ElseIf Abs(Val(vsfList.TextMatrix(ROW, COL))) = 0 And blnTrue = True Then 'δ��ѡ:ȡ��
                '�Ƶ����һ����ɾ��
                .RowPosition(lngCurRow) = .Rows - 1
                .RemoveItem .Rows - 1
                If .Rows = .FixedRows Then
                    .Rows = .Rows + 1
                    .ROW = .FixedRows
                End If
            End If
        End With
        
        If chkAll.Tag = "" Then
            '����Ƿ�Ȩѡ��
            blnTrue = False
            For lngRow = vsfList.FixedRows To vsfList.Rows - 1
                blnTrue = True
                If Abs(Val(vsfList.TextMatrix(lngRow, vsfList.ColIndex("ѡ��")))) = 0 Then
                    blnTrue = False
                    Exit For
                End If
            Next lngRow
            mblnLoadOver = False
            chkAll.Value = IIf(blnTrue = True, 1, 0)
            mblnLoadOver = True
        End If
    End If
End Sub

Private Sub vsfList_DblClick()
    Dim intValue As Integer
    With vsfList
        If .ROW >= .FixedRows And .COL > .ColIndex("ѡ��") And Val(.RowData(.ROW)) > 0 Then
            intValue = Abs(Val(.TextMatrix(.ROW, .ColIndex("ѡ��"))))
            .TextMatrix(.ROW, .ColIndex("ѡ��")) = intValue - 1
            Call vsfList_AfterEdit(.ROW, .ColIndex("ѡ��"))
        End If
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then Call vsfList_DblClick
End Sub

Private Sub vsfList_StartEdit(ByVal ROW As Long, ByVal COL As Long, Cancel As Boolean)
    Cancel = Not ((COL = vsfList.ColIndex("ѡ��")) And Val(vsfList.RowData(ROW)) <> 0)
End Sub
