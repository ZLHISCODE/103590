VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNurseFileSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ļ�ѡ��"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6855
   Icon            =   "frmNurseFileSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6705
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   630
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   150
         Width           =   1350
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "˵����������ӡ����ӡ���ϲ��ļ�¼���ļ���"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2850
         TabIndex        =   7
         Top             =   210
         Width           =   3600
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�鿴"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   1
         Top             =   210
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ(&P)"
      Height          =   350
      Left            =   4440
      TabIndex        =   5
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5670
      TabIndex        =   6
      Top             =   4065
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgData 
      Left            =   1020
      Top             =   1785
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
            Picture         =   "frmNurseFileSelect.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNurseFileSelect.frx":D0B4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNurseFileSelect.frx":D64E
            Key             =   "��ͨ"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgFile 
      Height          =   1095
      Left            =   15
      TabIndex        =   4
      Top             =   615
      Width           =   6060
      _cx             =   10689
      _cy             =   1931
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
         Caption         =   "Check1"
         Height          =   180
         Left            =   1215
         TabIndex        =   3
         Top             =   165
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmNurseFileSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngPatiID As Long
Private mlngPageId As Long
Private mintBaby As Integer
Private mlng��� As Long
Private marrFile() As Variant
Private Enum mCol
    f��־ = 0: fѡ��: fID: f��ʽID: f�ļ�: f��ʼ����: f����ID: f����: f����
End Enum

Public Function ShowMe(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer) As Variant
    marrFile = Array()
    mlngPatiID = lngPatiID
    mlngPageId = lngPageId
    mintBaby = intBaby
    mlng��� = -2
    Me.Show 1
    ShowMe = marrFile
End Function

Private Function zlRefData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim intRow As Integer
    Dim lngID As Long
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo ErrHand
    '------------------------------------------------------------------------------------------------------------------
    '�����ļ�ˢ��
    
    With vfgFile
        .Clear
        .Rows = 2
        .Cols = 9
        .FixedCols = 1
        
        .TextMatrix(0, mCol.f��־) = ""
        .TextMatrix(0, mCol.fѡ��) = ""
        .TextMatrix(0, mCol.fID) = "ID"
        .TextMatrix(0, mCol.f��ʽID) = "��ʽID"
        .TextMatrix(0, mCol.f�ļ�) = "�ļ�"
        .TextMatrix(0, mCol.f��ʼ����) = "��ʼ����"
        .TextMatrix(0, mCol.f����ID) = "����id"
        .TextMatrix(0, mCol.f����) = "����"
        .TextMatrix(0, mCol.f����) = "����"
        
        Set .Cell(flexcpPicture, 1, mCol.f��־) = Nothing
        .TextMatrix(0, mCol.fѡ��) = ""
        .TextMatrix(1, mCol.fID) = ""
        .TextMatrix(1, mCol.f��ʽID) = ""
        .TextMatrix(1, mCol.f�ļ�) = ""
        .TextMatrix(1, mCol.f��ʼ����) = ""
        .TextMatrix(1, mCol.f����ID) = ""
        .TextMatrix(1, mCol.f����) = ""
        .TextMatrix(1, mCol.f����) = ""
        
        .ColWidth(mCol.f��־) = 270: .ColWidth(mCol.fѡ��) = 270
        .ColWidth(mCol.fID) = 0: .ColWidth(mCol.f��ʽID) = 0: .ColWidth(mCol.f�ļ�) = 2000: .ColWidth(mCol.f��ʼ����) = 1500
        .ColWidth(mCol.f����ID) = 0: .ColWidth(mCol.f����) = 1200: .ColWidth(mCol.f����) = 0
        chkAll.Top = 70
        chkAll.Left = 270 + (270 - chkAll.Width) \ 2
    End With
    
    intRow = vfgFile.FixedRows
    '--------------------------------------------------------------------------------------------------------------
    gstrSQL = "" & _
        " SELECT A.ID,A.��ʽID,A.����ID,A.Ӥ��,C.���� AS ����,A.�ļ�����,A.��ʼʱ��,B.����,b.���" & vbNewLine & _
        " FROM ���˻����ļ� A,�����ļ��б� B,���ű� C" & vbNewLine & _
        " WHERE A.��ʽID=B.ID AND A.����ID=C.ID And A.����ID=[1] And A.��ҳID=[2] " & IIf(mlng��� < 0, "", " And NVL(A.Ӥ��,0)=[3]") & _
        " And NOT EXISTS (SELECT 1 FROM ���˻����ļ� WHERE ����ID=A.����ID And ��ҳID=A.��ҳID And NVL(Ӥ��,0)=NVL(A.Ӥ��,0)  And ����ID=A.ID)" & _
        " ORDER BY A.Ӥ��,B.����,A.��ʼʱ�� "
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiID, mlngPageId, mlng���)
    
    With Me.vfgFile
        Do While Not rsTemp.EOF
            If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""
            If rsTemp!���� = -1 Then
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("����").Picture
            Else
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("��ͨ").Picture
            End If
            .RowData(.Rows - 1) = Val(NVL(rsTemp!Ӥ��))
            lngID = Val(NVL(rsTemp!ID))
            .TextMatrix(.Rows - 1, mCol.fID) = lngID
            .TextMatrix(.Rows - 1, mCol.f��ʽID) = NVL(rsTemp!��ʽID, 0)
            .TextMatrix(.Rows - 1, mCol.f�ļ�) = NVL(rsTemp!�ļ�����)
            .TextMatrix(.Rows - 1, mCol.f��ʼ����) = Format(NVL(rsTemp!��ʼʱ��), "yyyy-MM-dd")
            .TextMatrix(.Rows - 1, mCol.f����ID) = NVL(rsTemp!����ID)
            .TextMatrix(.Rows - 1, mCol.f����) = NVL(rsTemp!����)
            .TextMatrix(.Rows - 1, mCol.f����) = NVL(rsTemp!����)
            
            rsTemp.MoveNext
        Loop
    End With
    vfgFile.ColAlignment(mCol.fѡ��) = flexAlignCenterCenter
    vfgFile.Cell(flexcpChecked, intRow, mCol.fѡ��, vfgFile.Rows - 1, mCol.fѡ��) = flexTSUnchecked
    'ѡ����
    Call vfgFile.Select(intRow, mCol.fID)
    
    zlRefData = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboBaby_Click()
    If mlng��� = cboBaby.ItemData(cboBaby.ListIndex) Then Exit Sub
    mlng��� = cboBaby.ItemData(cboBaby.ListIndex)
    Call zlRefData
End Sub

Private Sub chkAll_Click()
    Dim lngRow As Long
    If chkAll.Tag = "OK" Then Exit Sub
    With vfgFile
        For lngRow = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, lngRow, mCol.fѡ��) = IIf(chkAll.Value = 0, flexTSUnchecked, flexTSChecked)
        Next lngRow
    End With
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim lngRow As Integer
    marrFile = Array()
    With vfgFile
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngRow, mCol.fѡ��) = flexTSChecked Then
                ReDim Preserve marrFile(UBound(marrFile) + 1)
                marrFile(UBound(marrFile)) = .TextMatrix(lngRow, mCol.fID) & "_" & Val(.RowData(lngRow)) & "_" & .TextMatrix(lngRow, mCol.f����)
            End If
        Next lngRow
    End With
    
    If UBound(marrFile) = -1 Then
        MsgBox "����ѡ����Ҫ��ӡ�ļ���", vbInformation, gstrSysName
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    cboBaby.Clear
    gstrSQL = "Select a.���,Decode(a.Ӥ������,Null,NVL(C.����,b.����) ||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������" & _
            " From ������Ϣ b,������ҳ C,������������¼ a Where b.����id=C.����id And A.����ID=C.����ID And A.��ҳID=C.��ҳID And C.����id=[1] And C.��ҳid=[2]  Order By a.���"
            
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡӤ����Ϣ", mlngPatiID, mlngPageId)
    If rs.RecordCount > 0 Then
        cboBaby.AddItem "����": cboBaby.ItemData(cboBaby.NewIndex) = -1
    End If
    cboBaby.AddItem "���˱���": cboBaby.ItemData(cboBaby.NewIndex) = 0
    If rs.BOF = False Then
        Do While Not rs.EOF
            cboBaby.AddItem rs("Ӥ������").Value: cboBaby.ItemData(cboBaby.NewIndex) = Val(NVL(rs("���").Value, 0))
            If cboBaby.ListIndex = -1 And Val(NVL(rs("���").Value, 0)) = mintBaby Then cboBaby.ListIndex = cboBaby.NewIndex
            rs.MoveNext
        Loop
    End If
    
    If cboBaby.ListIndex = -1 And cboBaby.ListCount > 0 Then
        If cboBaby.ItemData(0) = -1 Then
            cboBaby.ListIndex = 1
        Else
            cboBaby.ListIndex = 0
        End If
    End If
    cboBaby.Enabled = (cboBaby.ListCount > 1)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fra.Move 10, 0, Me.ScaleWidth - 30, fra.Height
    vfgFile.Move 10, fra.Height + 10, Me.ScaleWidth - 30, Me.ScaleHeight - vfgFile.Top - cmdPrint.Height - 200
    cmdCancle.Left = Me.ScaleWidth - cmdCancle.Width - 100
    cmdCancle.Top = vfgFile.Top + vfgFile.Height + 100
    cmdPrint.Top = cmdCancle.Top
    cmdPrint.Left = cmdCancle.Left - cmdPrint.Width - 100
End Sub

Private Sub vfgFile_AfterEdit(ByVal ROW As Long, ByVal COL As Long)
    Dim lngRow As Long
    Dim blnSelectAll As Boolean
    If COL <> mCol.fѡ�� Then Exit Sub
    
    blnSelectAll = True
    With vfgFile
        If vfgFile.Cell(flexcpChecked, ROW, COL) = 5 Then vfgFile.Cell(flexcpChecked, ROW, COL) = flexTSUnchecked
        For lngRow = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngRow, mCol.fѡ��) = flexTSUnchecked Then
                blnSelectAll = False
                Exit For
            End If
        Next lngRow
    End With
    chkAll.Tag = "OK"
    chkAll.Value = IIf(blnSelectAll, 1, 0)
    chkAll.Tag = ""
End Sub

Private Sub vfgFile_BeforeUserResize(ByVal ROW As Long, ByVal COL As Long, Cancel As Boolean)
    If COL < mCol.fID Then Cancel = True
End Sub

Private Sub vfgFile_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    Dim blnSelectAll As Boolean
    If KeyCode = vbKeySpace And Val(vfgFile.TextMatrix(vfgFile.ROW, mCol.fID)) > 0 Then
        If vfgFile.Cell(flexcpChecked, vfgFile.ROW, mCol.fѡ��) = flexTSUnchecked Then
            vfgFile.Cell(flexcpChecked, vfgFile.ROW, mCol.fѡ��) = flexTSChecked
        Else
            vfgFile.Cell(flexcpChecked, vfgFile.ROW, mCol.fѡ��) = flexTSUnchecked
        End If
        
        blnSelectAll = True
        With vfgFile
            For lngRow = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, lngRow, mCol.fѡ��) = flexTSUnchecked Then
                    blnSelectAll = False
                    Exit For
                End If
            Next lngRow
        End With
        chkAll.Tag = "OK"
        chkAll.Value = IIf(blnSelectAll, 1, 0)
        chkAll.Tag = ""
    End If
End Sub

Private Sub vfgFile_StartEdit(ByVal ROW As Long, ByVal COL As Long, Cancel As Boolean)
    If COL <> mCol.fѡ�� Then Cancel = True
End Sub
