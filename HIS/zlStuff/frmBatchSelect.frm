VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������ѡ��"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13200
   Icon            =   "frmBatchSelect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdGet 
      Caption         =   "��ȡ"
      Height          =   300
      Left            =   4785
      TabIndex        =   19
      Top             =   750
      Width           =   510
   End
   Begin VB.TextBox txtCostEnd 
      Height          =   300
      Left            =   4185
      TabIndex        =   18
      Top             =   750
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCostBegin 
      Height          =   300
      Left            =   3360
      TabIndex        =   16
      Top             =   750
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cboCost 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   750
      Width           =   2055
   End
   Begin VB.PictureBox picDrug 
      Height          =   5775
      Left            =   9120
      ScaleHeight     =   5715
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   1060
      Visible         =   0   'False
      Width           =   3975
      Begin VSFlex8Ctl.VSFlexGrid vsfDrug 
         Height          =   5830
         Left            =   0
         TabIndex        =   14
         Top             =   -120
         Width           =   3975
         _cx             =   7011
         _cy             =   10283
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
         BackColorSel    =   16769992
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
         FormatString    =   $"frmBatchSelect.frx":000C
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
         Editable        =   0
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
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   12120
      TabIndex        =   6
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "���(&A)"
      Height          =   300
      Left            =   11040
      TabIndex        =   5
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "���(&O)"
      Height          =   300
      Left            =   9960
      TabIndex        =   4
      Top             =   7920
      Width           =   975
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   7920
      Width           =   1335
   End
   Begin VB.TextBox txtSelect 
      Height          =   300
      Left            =   9120
      TabIndex        =   1
      Top             =   780
      Width           =   3975
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   360
      Top             =   1440
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
            Picture         =   "frmBatchSelect.frx":0081
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":061B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":6E7D
            Key             =   "���U"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSelectDrug 
      Height          =   6645
      Left            =   120
      TabIndex        =   8
      Top             =   1095
      Width           =   12975
      _cx             =   22886
      _cy             =   11721
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
      BackColorSel    =   16769992
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
      FormatString    =   $"frmBatchSelect.frx":7417
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
      Editable        =   0
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
   Begin VB.Frame fra������ 
      Caption         =   "������ʽ"
      Height          =   600
      Left            =   120
      TabIndex        =   9
      Top             =   100
      Width           =   12980
      Begin VB.ComboBox cbo������ʽ 
         Height          =   300
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   2580
      End
      Begin VB.TextBox txt������ 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblInfor 
         Caption         =   "���ݳɱ��ۣ������µļӳ������¼ӳɵ���"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   263
         Width           =   3420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3540
         TabIndex        =   12
         Top             =   285
         Width           =   225
      End
   End
   Begin VB.Label lblCost 
      AutoSize        =   -1  'True
      Caption         =   "�ɱ��۷�Χ"
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   810
      Width           =   900
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      Caption         =   "--"
      Height          =   180
      Left            =   3960
      TabIndex        =   17
      Top             =   810
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   7980
      Width           =   360
   End
   Begin VB.Label lblCalss 
      AutoSize        =   -1  'True
      Caption         =   "Ʒ�ּ���"
      Height          =   180
      Left            =   8355
      TabIndex        =   0
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "frmBatchSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintUnit As Integer '��ģ�������õ���ʾ��λ 0-ɢװ��λ,1-��װ��λ
Private Const mlngRowHeight As Long = 300 '����и����и�
Private mrsReturn As ADODB.Recordset        '����ѡ����������
Private mblnOk As Boolean   '��¼�Ƿ��ǵ����ȷ����ť
Private mrsFindName As ADODB.Recordset '��¼��ѯ���ݼ�
Private mstrMatch  As String '0-˫��ƥ�� 1-������ƥ��
Private mintType As Integer '������ʽ
Private mdbl���� As Double  '������ʽ�ı���
Private mint���� As Integer  'ֻ���ɱ���ʱ�����������
Private Const mstrCaption As String = "��������ѡ��"
Private mstr������ As String
'����λ
Private mFMT As g_FmtString

Private Enum vsfSelectDrugCol
    ����ID = 0
    ������Ϣ = 1
    ���ϱ���
    ��Ʒ��
    ͨ����
    ���
    ����
    ��λ
    ��λϵ��
    ��������
    ɢװ��λ
    ��װϵ��
    ��װ��λ
    ����
    ԭ�ۼ�
    �ۼ�
    �ɱ���
    ָ������
    ָ���ۼ�
    ������
End Enum

Public Sub ShowMe(ByVal frmParent As Form, ByRef rsTemp As ADODB.Recordset, ByRef blnOK As Boolean, ByRef intType As Integer, ByRef dbl���� As Double, Optional int���� As Integer, Optional str������ As String)
    mint���� = int����
    Me.Show vbModal, frmParent
    blnOK = mblnOk
    Set rsTemp = mrsReturn
    intType = mintType
    dbl���� = mdbl����
    str������ = mstr������
End Sub

Private Sub initVsflexgrid()
    With vsfSelectDrug
        .Editable = flexEDNone
        .Cols = vsfSelectDrugCol.������
        .Rows = 1
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMove '�ƶ�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��

        '�����п�
        .ColWidth(vsfSelectDrugCol.����ID) = 0
        .ColWidth(vsfSelectDrugCol.������Ϣ) = 3000
        .ColWidth(vsfSelectDrugCol.���ϱ���) = 0
        .ColWidth(vsfSelectDrugCol.��Ʒ��) = 0
        .ColWidth(vsfSelectDrugCol.ͨ����) = 0
        .ColWidth(vsfSelectDrugCol.����) = 1500
        .ColWidth(vsfSelectDrugCol.��������) = 0
        .ColWidth(vsfSelectDrugCol.��λ) = 500
        .ColWidth(vsfSelectDrugCol.ɢװ��λ) = 0
        .ColWidth(vsfSelectDrugCol.��װϵ��) = 0
        .ColWidth(vsfSelectDrugCol.��װ��λ) = 0
        
        .ColWidth(vsfSelectDrugCol.����) = 1000
        .ColWidth(vsfSelectDrugCol.�ۼ�) = 1500
        .ColWidth(vsfSelectDrugCol.ԭ�ۼ�) = 0
        .ColWidth(vsfSelectDrugCol.�ɱ���) = 1500
        .ColWidth(vsfSelectDrugCol.ָ������) = 1500
        .ColWidth(vsfSelectDrugCol.ָ���ۼ�) = 1500
        .ColWidth(vsfSelectDrugCol.��λϵ��) = 0
        '������ͷ
        .TextMatrix(0, vsfSelectDrugCol.����ID) = "����id"
        .TextMatrix(0, vsfSelectDrugCol.������Ϣ) = "������Ϣ"
        .TextMatrix(0, vsfSelectDrugCol.���ϱ���) = "���ϱ���"
        .TextMatrix(0, vsfSelectDrugCol.��������) = "��������"
        .TextMatrix(0, vsfSelectDrugCol.��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, vsfSelectDrugCol.ͨ����) = "ͨ����"
        .TextMatrix(0, vsfSelectDrugCol.���) = "���"
        .TextMatrix(0, vsfSelectDrugCol.����) = "����"
        .TextMatrix(0, vsfSelectDrugCol.��λ) = "��λ"
        
        .TextMatrix(0, vsfSelectDrugCol.ɢװ��λ) = "ɢװ��λ"
        .TextMatrix(0, vsfSelectDrugCol.��װϵ��) = "��װϵ��"
        .TextMatrix(0, vsfSelectDrugCol.��װ��λ) = "��װ��λ"
        
        .TextMatrix(0, vsfSelectDrugCol.����) = "����"
        .TextMatrix(0, vsfSelectDrugCol.�ۼ�) = "�ۼ�"
        .TextMatrix(0, vsfSelectDrugCol.�ɱ���) = "�ɱ���"
        .TextMatrix(0, vsfSelectDrugCol.ָ������) = "ָ������"
        .TextMatrix(0, vsfSelectDrugCol.ָ���ۼ�) = "ָ���ۼ�"

        .ColAlignment(vsfSelectDrugCol.����ID) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.������Ϣ) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.���ϱ���) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.���) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.����) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.��λ) = flexAlignCenterCenter
        .ColAlignment(vsfSelectDrugCol.����) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.�ۼ�) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.�ɱ���) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.ָ������) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.ָ���ۼ�) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub


'Private Sub setTvwInfo()
'    'Ϊ�����������
'    Dim objNode As Node
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'
'    gstrSQL = " Select ����,���� From ������Ŀ��� " & _
'              " Where Instr([1],����,1) > 0 " & _
'              " Order by ����"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, "4")
'
'    If rsTemp Is Nothing Then
'        Exit Sub
'    End If
'
'    With tvwDrug
'        .Nodes.Clear
'        Do While Not rsTemp.EOF
'            .Nodes.Add , , "Root" & rsTemp!����, rsTemp!����, 1, 1
'            .Nodes("Root" & rsTemp!����).Tag = rsTemp!����
'            rsTemp.MoveNext
'        Loop
'    End With
'
'
'    gstrSQL = "Select ID, �ϼ�id, ����, ����, Decode(����, 7, '����') ����, '����' As ���" & _
'                " From ���Ʒ���Ŀ¼" & _
'                " Where ���� ='7' And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                " Start With �ϼ�id Is Null" & _
'                " Connect By Prior ID = �ϼ�id"
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ѯ")
'    With rsTemp
'        Do While Not .EOF
'           If IsNull(!�ϼ�ID) Then
'                Set objNode = tvwDrug.Nodes.Add("Root" & !����, 4, "K_" & !Id, !���� & "-����", 1, 1)
'            Else
'                Set objNode = tvwDrug.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !Id, !���� & "-����", 1, 1)
'            End If
'            objNode.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'            .MoveNext
'        Loop
'    End With
'
'    If optVariety.Value = True Then
'        gstrSQL = "Select ID, ����id, ����, ����, Decode(���, 7, '����') ����, 'Ʒ��' As ���" & _
'                  "  From ������ĿĿ¼" & _
'                  "  Where ����id In (Select ID" & _
'                                   " From ���Ʒ���Ŀ¼" & _
'                                   " Where ���� ='7' And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                                   " Start With �ϼ�id Is Null" & _
'                                   " Connect By Prior ID = �ϼ�id)"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "Ʒ��")
'
'        With rsTemp
'            Do While Not .EOF
'                Set objNode = tvwDrug.Nodes.Add("K_" & !����id, 4, !��� & "K_" & !Id, !���� & "-Ʒ��", 1, 1)
'                objNode.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'                .MoveNext
'            Loop
'        End With
'    End If
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub cboCost_Click()
    If cboCost.Text = "�Զ���" Then
        txtCostBegin.Visible = True
        txtCostEnd.Visible = True
        lblTo.Visible = True
        cmdGet.Left = txtCostEnd.Left + txtCostEnd.Width + 5
    Else
        txtCostBegin.Visible = False
        txtCostEnd.Visible = False
        lblTo.Visible = False
        cmdGet.Left = txtCostBegin.Left
    End If
End Sub

Private Sub cbo������ʽ_Click()
    Dim intType As Integer
    Dim dbl���� As Double
    Dim dbl����ϵ��  As Double
    Dim intRow As Integer
    
    If cbo������ʽ.ListIndex < 0 Then Exit Sub
    Select Case cbo������ʽ.ItemData(cbo������ʽ.ListIndex)
        Case 1
            lblInfor.Caption = "���ݳɱ��ۣ������µļӳ������¼ӳɵ���"
            lbl����.Caption = "��"
            txt������.MaxLength = 3
        Case 2
            lblInfor.Caption = "�ڵ�ǰ�ۼۻ����ϰ��ձ�������"
            lbl����.Caption = "��"
            txt������.MaxLength = 3
        Case 3
            lblInfor.Caption = "�ڵ�ǰ�ۼۻ����ϰ��̶����Ӽ�����"
            lbl����.Caption = "Ԫ"
            txt������.MaxLength = 10
    End Select

    intType = cbo������ʽ.ItemData(cbo������ʽ.ListIndex)
    
    With vsfSelectDrug
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(1, vsfSelectDrugCol.����ID) = "" Then Exit Sub
        For intRow = 1 To .Rows - 1
            dbl���� = Val(txt������.Text)
            If Trim(txt������.Text) = "" Then
                .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.ԭ�ۼ�)), mFMT.FM_���ۼ�)
            Else
                Select Case intType
                    Case 1      '���ݳɱ��ۼӳ�
                        dbl���� = 1 + Val(dbl����) / 100
                        .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.�ɱ���)) * dbl����, mFMT.FM_���ۼ�)
                    Case 2      '�������ۼ۰�����
                        dbl���� = 1 + Val(dbl����) / 100
                        .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.ԭ�ۼ�)) * dbl����, mFMT.FM_���ۼ�)
                    Case 3      '�������ۼ۰��̶����Ӽ�
                        dbl���� = Val(dbl����)
                        .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format((Val(.TextMatrix(intRow, vsfSelectDrugCol.ԭ�ۼ�))) + dbl����, mFMT.FM_���ۼ�)
                End Select
            End If
            If Val(.TextMatrix(intRow, vsfSelectDrugCol.�ۼ�)) > Val(.TextMatrix(intRow, vsfSelectDrugCol.ָ���ۼ�)) And Val(.TextMatrix(intRow, vsfSelectDrugCol.ָ���ۼ�)) <> 0 Then
                .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.ָ���ۼ�)), mFMT.FM_���ۼ�)
            End If
        Next
    End With
End Sub

Private Sub cmdGet_Click()
    Dim dblBegin As Double
    Dim dblEnd As Double
    Dim strTemp As String
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If cboCost.Text = "�Զ���" Then
        If Trim(txtCostBegin.Text) = "" Then
            MsgBox "������Ҫ��ѯ�ɱ��۵Ŀ�ʼ�۸�", vbInformation, gstrSysName
            txtCostBegin.SetFocus
            Exit Sub
        End If
        If Trim(txtCostEnd.Text) = "" Then
            MsgBox "������Ҫ��ѯ�ɱ��۵Ľ����۸�", vbInformation, gstrSysName
            txtCostEnd.SetFocus
            Exit Sub
        End If
    End If
    
    vsfSelectDrug.Rows = 1
    
    If cboCost.Text <> "�Զ���" Then
        strTemp = cboCost.Text
        dblBegin = Mid(strTemp, 1, InStr(1, strTemp, "-") - 1)
        dblEnd = Mid(strTemp, InStr(1, strTemp, "-") + 1, InStr(1, strTemp, "(") - InStr(1, strTemp, "-") - 1)
    Else
        dblBegin = Trim(txtCostBegin.Text)
        dblEnd = Trim(txtCostEnd.Text)
    End If
    
    If mintUnit = 0 Then
        'ɢװ��λ
        strTemp = "And (b.ƽ���ɱ��� Between [1] And [2] Or d.�ɱ��� Between [1] And [2])"
    Else
        'ҩ�ⵥλ
        strTemp = "And (b.ƽ���ɱ��� Between [1]/d.����ϵ�� And [2]/d.����ϵ�� Or d.�ɱ��� Between [1]/d.����ϵ�� And [2]/d.����ϵ��)"
    End If
    
    gstrSQL = "Select Distinct a.Id As ����id, a.���� As ���ϱ���, a.���� As ͨ����, c.��Ʒ��, a.���, a.�Ƿ��� As ʱ��, a.����, a.���㵥λ, d.����ϵ��, d.��װ��λ," & vbNewLine & _
                "                Decode(Nvl(b.ƽ���ɱ���, 0), 0, d.�ɱ���, b.ƽ���ɱ���) As �ɱ���," & vbNewLine & _
                "                Decode(Nvl(b.ʵ������, 0), 0, e.�ּ�, b.ʵ�ʽ�� / b.ʵ������) As �ּ�, d.ָ��������, d.ָ�����ۼ�, d.��������" & vbNewLine & _
                "From �շ���ĿĿ¼ A, ҩƷ��� B, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) C, �������� D, �շѼ�Ŀ E" & vbNewLine & _
                "Where a.Id = b.ҩƷid(+) And a.Id = c.�շ�ϸĿid(+) And a.Id = d.����id And a.Id = e.�շ�ϸĿid And Sysdate Between e.ִ������ And" & vbNewLine & _
                "      e.��ֹ���� And (a.����ʱ�� = to_date('3000-01-01','yyyy-mm-dd') or a.����ʱ�� is null ) " & GetPriceClassString("E") & _
                " And a.��� = '4' " & strTemp
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ɱ��۷�Χ��ѯ", dblBegin, dblEnd)
        
    If rsData.RecordCount > 0 Then
        Call setVSFValue(rsData)
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtCostBegin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
        Else
           KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    Dim intType As Integer
    Dim dbl���� As Double
    Dim intRow As Integer
    Dim dbl����ϵ�� As Double
    
    If cbo������ʽ.ItemData(cbo������ʽ.ListIndex) = 3 Then
        Call zlControl.TxtCheckKeyPress(txt������, KeyAscii, m�����ʽ)
    Else
        Call zlControl.TxtCheckKeyPress(txt������, KeyAscii, m���ʽ)
    End If
    If KeyAscii <> 0 And KeyAscii = vbKeyReturn Then
        intType = cbo������ʽ.ItemData(cbo������ʽ.ListIndex)
        
        With vsfSelectDrug
            For intRow = 1 To .Rows - 1
                dbl���� = Val(txt������.Text)
                If Trim(txt������.Text) = "" Then
                    .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.ԭ�ۼ�)), mFMT.FM_���ۼ�)
                Else
                    Select Case intType
                        Case 1      '���ݳɱ��ۼӳ�
                            dbl���� = 1 + Val(dbl����) / 100
                            .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.�ɱ���)) * dbl����, mFMT.FM_���ۼ�)
                        Case 2      '�������ۼ۰�����
                            dbl���� = 1 + Val(dbl����) / 100
                            .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.ԭ�ۼ�)) * dbl����, mFMT.FM_���ۼ�)
                        Case 3      '�������ۼ۰��̶����Ӽ�
                            dbl���� = Val(dbl����)
                            .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format((Val(.TextMatrix(intRow, vsfSelectDrugCol.ԭ�ۼ�))) + dbl����, mFMT.FM_���ۼ�)
                    End Select
                End If
                If Val(.TextMatrix(intRow, vsfSelectDrugCol.�ۼ�)) > Val(.TextMatrix(intRow, vsfSelectDrugCol.ָ���ۼ�)) And Val(.TextMatrix(intRow, vsfSelectDrugCol.ָ���ۼ�)) <> 0 Then
                    .TextMatrix(intRow, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(intRow, vsfSelectDrugCol.ָ���ۼ�)), mFMT.FM_���ۼ�)
                End If
            Next
        End With
    End If
End Sub

Private Sub cbo������ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

'Private Sub cmdSelect_Click()
'    picDrug.Visible = True
'    tvwDrug.Visible = True
'    Call setTvwInfo
'End Sub

Private Sub cmdCal_Click()
    With vsfSelectDrug
        If MsgBox("ȷ��Ҫ��������Ѿ�ѡ������ģ�", vbYesNo, gstrSysName) = vbYes Then
            .Rows = 1
        End If
    End With
End Sub

Private Sub cmdOk_Click()
    Dim intRow As Integer
    Set mrsReturn = New ADODB.Recordset
    mintType = cbo������ʽ.ItemData(cbo������ʽ.ListIndex)
    mdbl���� = Val(txt������.Text)
    mstr������ = Trim(txt������.Text)
    With mrsReturn
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ͨ����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ʱ��", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 2, adFldIsNullable
        
        .Fields.Append "���㵥λ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����ϵ��", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "��װ��λ", adLongVarChar, 11, adFldIsNullable
        .Fields.Append "�ɱ���", adDouble, 18, adFldIsNullable
        .Fields.Append "ָ��������", adDouble, 18, adFldIsNullable
        .Fields.Append "ָ�����ۼ�", adDouble, 18, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    With vsfSelectDrug
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, vsfSelectDrugCol.����ID) = "" Then Exit For
            mrsReturn.AddNew
            mrsReturn!Id = .TextMatrix(intRow, vsfSelectDrugCol.����ID)
            mrsReturn!���� = .TextMatrix(intRow, vsfSelectDrugCol.���ϱ���)
            mrsReturn!��Ʒ�� = .TextMatrix(intRow, vsfSelectDrugCol.��Ʒ��)
            mrsReturn!ͨ���� = .TextMatrix(intRow, vsfSelectDrugCol.ͨ����)
            mrsReturn!��� = .TextMatrix(intRow, vsfSelectDrugCol.���)
            mrsReturn!ʱ�� = .TextMatrix(intRow, vsfSelectDrugCol.����)
            mrsReturn!���� = .TextMatrix(intRow, vsfSelectDrugCol.����)
            mrsReturn!�������� = .TextMatrix(intRow, vsfSelectDrugCol.��������)
            mrsReturn!���㵥λ = .TextMatrix(intRow, vsfSelectDrugCol.ɢװ��λ)
            mrsReturn!����ϵ�� = .TextMatrix(intRow, vsfSelectDrugCol.��װϵ��)
            mrsReturn!��װ��λ = .TextMatrix(intRow, vsfSelectDrugCol.��װ��λ)
            
            mrsReturn!�ɱ��� = Val(.TextMatrix(intRow, vsfSelectDrugCol.�ɱ���)) / Val(.TextMatrix(intRow, vsfSelectDrugCol.��λϵ��))
            mrsReturn!ָ�������� = Val(.TextMatrix(intRow, vsfSelectDrugCol.ָ������)) / Val(.TextMatrix(intRow, vsfSelectDrugCol.��λϵ��))
            mrsReturn!ָ�����ۼ� = Val(.TextMatrix(intRow, vsfSelectDrugCol.ָ���ۼ�)) / Val(.TextMatrix(intRow, vsfSelectDrugCol.��λϵ��))
            
            mrsReturn.Update
        Next
    End With
    mblnOk = True
    
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        vsfDrug.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim intUnitTemp As Integer
    '��ȡ���õĵ�λ
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, 1726, 1))
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    With cbo������ʽ
        .AddItem "���ݳɱ��۰��ӳɵ���"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "�����ۼ۰���������"
        .ItemData(.NewIndex) = 2
        .AddItem "�����ۼ۰��̶�������"
        .ItemData(.NewIndex) = 3
    End With
    
    With cboCost
        .AddItem "0-10(��10)"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "10-20(��20)"
        .ItemData(.NewIndex) = 2
        .AddItem "20-50(��50)"
        .ItemData(.NewIndex) = 3
        .AddItem "�Զ���"
        .ItemData(.NewIndex) = 4
    End With
    
    If mintUnit = 0 Then
        Me.Caption = "��������ѡ��(ɢװ��λ)"
    Else
        Me.Caption = "��������ѡ��(��װ��λ)"
    End If
    
    cmdGet.Left = txtCostBegin.Left
    
    mstrMatch = IIf(zlDatabase.GetPara("����ƥ��", , , 0) = "0", "%", "")
    mblnOk = False
    
    If mint���� = 1 Then
        txt������.Enabled = False
        txt������.BackColor = &H80000000
        cbo������ʽ.Enabled = False
    Else
        txt������.Enabled = True
        txt������.BackColor = &H80000005
        cbo������ʽ.Enabled = True
    End If
    
    Call initVsflexgrid
    
    Call RestoreWinState(Me, App.ProductName, mstrCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mstrCaption)
End Sub

Private Sub optClass_Click()
    picDrug.Visible = False
    lblCalss.Caption = "����"
End Sub

Private Sub optClassSub_Click()
    picDrug.Visible = False
    lblCalss.Caption = "����(������)"
End Sub

Private Sub optVariety_Click()
    picDrug.Visible = False
    lblCalss.Caption = "Ʒ��"
End Sub

'Private Sub tvwDrug_NodeClick(ByVal Node As MSComctlLib.Node)
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'    If Node.Key Like "Root" Then Exit Sub
'
'    gstrSQL = "select id,����,����,���㵥λ from ������ĿĿ¼ where  Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' and ����id=[1]"
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", Mid(Node.Key, InStr(1, Node.Key, "_") + 1))
'
'    Set vsfDetails.DataSource = rsTemp
'
'    Exit Sub
'errHandle:
'    If errcenter() = 1 Then Resume
'    Call saveerrlog
'End Sub

'Private Sub tvwDrug_DblClick()
'    '����������д���ֵ
'    Dim lngID As Long
'    Dim rsTemp As ADODB.Recordset
'    Dim intRow As Integer
'    Dim i As Integer
'    Dim blnDou As Boolean '�ظ�����
'    Dim dbl����ϵ�� As Double
'    Dim strUnit As String   '��λ
'    Dim intType As Integer '������ʽ
'    Dim dbl���� As Double   '������
'
'    On Error GoTo errHandle
'
'    intType = cbo������ʽ.ItemData(cbo������ʽ.ListIndex)
'    dbl���� = Val(txt������.Text)
'    With tvwDrug
'        If optVariety.Value = True Then
'            If InStr(1, .SelectedItem.Text, "-Ʒ��") <= 0 Then
'                Exit Sub
'            End If
'            gstrSQL = "Select Distinct a.����id, c.���� As ���ϱ���, c.���� As ͨ����, d.��Ʒ��, c.���, c.�Ƿ��� As ʱ��, c.����, c.���㵥λ,a.����ϵ��, a.��װ��λ," & _
'                                        " a.�ɱ���, e.�ּ�, a.ָ��������, a.ָ�����ۼ�,a.��������" & _
'                        " From �������� A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) D,�շѼ�Ŀ E" & _
'                        " Where a.����id = b.Id And a.����id = c.Id And c.Id = d.�շ�ϸĿid(+) and a.����id=e.�շ�ϸĿid and sysdate between e.ִ������ and e.��ֹ���� and b.id=[1] order by c.����"
'        Else
'            If InStr(1, .SelectedItem.Text, "-����") <= 0 Then
'                Exit Sub
'            End If
'                If optClassSub.Value = True Then '�������ӽڵ�
'                    gstrSQL = "(Select ID From ���Ʒ���Ŀ¼ Where ���� =7 Start With ID = [1] Connect By Prior ID = �ϼ�id) A,"
'                Else '������
'                    gstrSQL = "(select id from ���Ʒ���Ŀ¼ where ���� =7 and id=[1]) A,"
'                End If
'
'                gstrSQL = "Select Distinct c.����id, d.���� As ���ϱ���, d.���� As ͨ����, f.��Ʒ��, d.���, d.�Ƿ��� As ʱ��, d.����, d.���㵥λ, c.����ϵ��, c.��װ��λ," & _
'                                        "  c.�ɱ���, e.�ּ�, c.ָ��������, c.ָ�����ۼ�,c.�������� " & _
'                        " From " & gstrSQL & " ������ĿĿ¼ B, �������� C," & _
'                             " �շ���ĿĿ¼ D, �շѼ�Ŀ E, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) F" & _
'                        " Where a.Id = b.����id And b.Id = c.����id And c.����id = d.Id And d.Id = e.�շ�ϸĿid And e.�շ�ϸĿid = f.�շ�ϸĿid(+) And" & _
'                              " Sysdate Between e.ִ������ And e.��ֹ���� order by d.����"
'        End If
'        lngID = Mid(.SelectedItem.Key, InStr(1, .SelectedItem.Key, "K_") + 2)
'
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����", lngID)
'        If rsTemp.RecordCount = 0 Then
'            picDrug.Visible = False
'            Exit Sub
'        End If
'    End With
'
'    With vsfSelectDrug
'        For intRow = 0 To rsTemp.RecordCount - 1
'            blnDou = False
'            For i = 1 To .Rows - 1
'                If .TextMatrix(i, vsfSelectDrugCol.����ID) = rsTemp!����ID Then
'                    blnDou = True
'                End If
'            Next
'            If blnDou = False Then
'                .Rows = .Rows + 1
'                .RowHeight(.Rows - 1) = mlngRowHeight
'
'                Select Case mintUnit
'                    Case 0
'                        dbl����ϵ�� = 1
'                        strUnit = rsTemp!���㵥λ
'                    Case 1
'                        dbl����ϵ�� = rsTemp!����ϵ��
'                        strUnit = rsTemp!��װ��λ
'                End Select
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.����ID) = rsTemp!����ID
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.������Ϣ) = "[" & rsTemp!���ϱ��� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.���ϱ���) = rsTemp!���ϱ���
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��Ʒ��) = IIf(IsNull(rsTemp!��Ʒ��), "", rsTemp!��Ʒ��)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ͨ����) = IIf(IsNull(rsTemp!ͨ����), "", rsTemp!ͨ����)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��λ) = strUnit
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��λϵ��) = dbl����ϵ��
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ɢװ��λ) = rsTemp!���㵥λ
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��װ��λ) = rsTemp!��װ��λ
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��װϵ��) = rsTemp!����ϵ��
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.����) = IIf(rsTemp!ʱ�� = 1, "ʱ��", "����")
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��������) = NVL(rsTemp!��������)
'
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ɱ���) = Format(dbl����ϵ�� * rsTemp!�ɱ���, mFMT.FM_�ɱ���)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ������) = Format(dbl����ϵ�� * rsTemp!ָ��������, mFMT.FM_�ɱ���)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ���ۼ�) = Format(dbl����ϵ�� * rsTemp!ָ�����ۼ�, mFMT.FM_���ۼ�)
'                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ԭ�ۼ�) = Format(Val(NVL(rsTemp!�ּ�)) * dbl����ϵ��, mFMT.FM_���ۼ�)
'                If dbl���� = 0 Then
'                    .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�) * dbl����ϵ��, mFMT.FM_���ۼ�)
'                Else
'                    Select Case intType
'                        Case 1      '���ݳɱ��ۼӳ�
'                            dbl���� = 1 + Val(dbl����) / 100
'                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format(Val(NVL(rsTemp!�ɱ���)) * dbl���� * dbl����ϵ��, mFMT.FM_���ۼ�)
'                        Case 2      '�������ۼ۰�����
'                            dbl���� = 1 + Val(dbl����) / 100
'                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format(Val(NVL(rsTemp!�ּ�)) * dbl���� * dbl����ϵ��, mFMT.FM_���ۼ�)
'                        Case 3      '�������ۼ۰��̶����Ӽ�
'                            dbl���� = Val(dbl����)
'                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format((Val(NVL(rsTemp!�ּ�)) * dbl����ϵ��) + dbl����, mFMT.FM_���ۼ�)
'                    End Select
'                End If
'                If Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�)) > Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ���ۼ�)) And Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ���ۼ�)) <> 0 Then
'                    .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ���ۼ�)), mFMT.FM_���ۼ�)
'                End If
'            End If
'            rsTemp.MoveNext
'        Next
'        picDrug.Visible = False
'    End With
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim strҩ�� As String
    Dim lngRow As Long
    
    '��������
    On Error GoTo ErrHandle
    If strInput <> txtFind.Tag Then
        '��ʾ�µĲ���
        txtFind.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ���ϱ���, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.���='4' " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ���ϱ��� "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "ȡƥ�������ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If
    
    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst
        
        strҩ�� = mrsFindName!���ϱ��� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)
        
        For lngRow = 1 To vsfSelectDrug.Rows - 1
            lngFindRow = vsfSelectDrug.FindRow(strҩ��, lngRow, CLng(vsfSelectDrugCol.������Ϣ), True, True)
            If lngFindRow > 0 Then
                vsfSelectDrug.Select lngFindRow, 1, lngFindRow, vsfSelectDrug.Cols - 1
                vsfSelectDrug.TopRow = lngFindRow
                Exit For
            End If
        Next
        
        If lngFindRow > 0 Then  '��ѯ�����ݺ���ƶ�����һ�����˳����β�ѯ
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
        End If
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtSelect_GotFocus()
    If picDrug.Visible = True Then
        picDrug.Visible = False
    End If
End Sub

Private Sub txtSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim rsPinzhong As ADODB.Recordset
    Dim objNode As Node
    Dim lng����id As Long
    Dim i As Integer
    
    If KeyCode = vbKeyReturn Then
    
        On Error GoTo ErrHandle
        
        If Trim(txtSelect.Text) = "" Then Exit Sub
                
        gstrSQL = "Select Distinct a.id,a.����,a.����" & _
                  "  From ������ĿĿ¼ A, ������Ŀ���� B" & _
                    " Where a.Id = b.������Ŀid(+) And a.��� ='4' And Sysdate Between ����ʱ�� And ����ʱ�� And" & _
                         " (a.���� Like [1] Or a.���� Like [1] Or b.���� Like [1] Or b.���� Like [1])"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & UCase(txtSelect.Text) & mstrMatch)
        If rsTemp.RecordCount = 0 Then
            MsgBox "δ��ѯ��Ʒ�֣�", vbInformation, gstrSysName
            txtSelect.SetFocus
            txtSelect.SelStart = 1
            txtSelect.SelLength = Len(txtSelect.Text)
        Else
            picDrug.Visible = True
            vsfDrug.Visible = True
            Set vsfDrug.DataSource = rsTemp
            vsfDrug.SetFocus
            vsfDrug.Row = 1
        End If
        With vsfDrug
            For i = 0 To .Rows - 1
                .RowHeight(i) = mlngRowHeight
            Next
        End With
        
'        picDrug.Visible = True
'        tvwDrug.Visible = True
'
'        gstrSQL = " Select ����,���� From ������Ŀ��� " & _
'                  " Where Instr([1],����,1) > 0 " & _
'                  " Order by ����"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, "4")
'
'        If rsTemp Is Nothing Then
'            Exit Sub
'        End If
'
'        With tvwDrug
'            .Nodes.Clear
'            Do While Not rsTemp.EOF
'                .Nodes.Add , , "Root" & rsTemp!����, rsTemp!����, 1, 1
'                .Nodes("Root" & rsTemp!����).Tag = rsTemp!����
'                rsTemp.MoveNext
'            Loop
'        End With
'
'        If optVariety.Value = True Then 'Ʒ�ֱ�ѡ��
'            gstrSQL = "Select distinct a.Id, a.�ϼ�id, a.����, a.����, Decode(a.����, 7, '����') ����, '����' As ���" & _
'                        " From ���Ʒ���Ŀ¼ A" & _
'                        " Start With ID In (Select Distinct a.����id" & _
'                                          " From ������ĿĿ¼ A, ������Ŀ���� B" & _
'                                          " Where a.Id = b.������Ŀid(+) And a.��� = '4' And Sysdate Between ����ʱ�� And ����ʱ�� And" & _
'                                                " (a.���� Like [1] Or a.���� Like [1] Or b.���� Like [1] Or b.���� Like [1]))" & _
'                        " Connect By Prior a.�ϼ�id = a.Id" & _
'                        " order by a.id"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & UCase(txtSelect.Text) & mstrMatch)
'            If rsTemp.RecordCount = 0 Then Exit Sub
'
'            With rsTemp
'                Do While Not .EOF
'                   If IsNull(!�ϼ�ID) Then
'                        Set objNode = tvwDrug.Nodes.Add("Root" & !����, 4, "K_" & !Id, !���� & "-����", 1, 1)
'                    Else
'                        Set objNode = tvwDrug.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !Id, !���� & "-����", 1, 1)
'                    End If
'                    objNode.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'                    .MoveNext
'                Loop
'
'                rsTemp.MoveFirst
'                Do While Not rsTemp.EOF
'                    lng����id = rsTemp!Id
'                    gstrSQL = "Select Distinct a.Id, a.����id, a.����, a.����, Decode(a.���, '4','����') ����, 'Ʒ��' As ���" & _
'                                " From ������ĿĿ¼ A" & _
'                                " Where a.��� ='4' And a.����id=[1] and Sysdate Between a.����ʱ�� And a.����ʱ��"
'                    Set rsPinzhong = zlDatabase.OpenSQLRecord(gstrSQL, "Ʒ��", lng����id)
'
'                    Do While Not rsPinzhong.EOF
'                        Set objNode = tvwDrug.Nodes.Add("K_" & rsPinzhong!����id, 4, rsPinzhong!��� & "K_" & rsPinzhong!Id, rsPinzhong!���� & "-Ʒ��", 1, 1)
'                        objNode.Tag = rsPinzhong!���� & "-" & rsPinzhong!���
'                        rsPinzhong.MoveNext
'                    Loop
'                    rsTemp.MoveNext
'                Loop
'            End With
'        Else
'            gstrSQL = "Select ID, �ϼ�id, ����, ����, Decode(����, 7, '����') ����, '����' As ���" & _
'                        " From ���Ʒ���Ŀ¼" & _
'                        " Start With ID in (Select ID" & _
'                                         " From ���Ʒ���Ŀ¼" & _
'                                         " Where ���� = '7' And (Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' Or ����ʱ�� Is Null) And" & _
'                                               " (���� Like [1] Or ���� Like [1] Or ���� Like [1]))" & _
'                        " Connect By Prior �ϼ�id = ID" & _
'                        " order by id"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯƷ��", "%" & UCase(txtSelect.Text) & mstrMatch)
'            If rsTemp.RecordCount = 0 Then Exit Sub
'
'            With rsTemp
'                Do While Not .EOF
'                   If IsNull(!�ϼ�ID) Then
'                        Set objNode = tvwDrug.Nodes.Add("Root" & !����, 4, "K_" & !Id, !���� & "-����", 1, 1)
'                    Else
'                        Set objNode = tvwDrug.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !Id, !���� & "-����", 1, 1)
'                    End If
'                    objNode.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
'                    .MoveNext
'                Loop
'            End With
'        End If
'        tvwDrug.SetFocus
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDrug_DblClick()
    Dim lngId As Long
    Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandle
    
    With vsfDrug
        If Val(.TextMatrix(.Row, 0)) = 0 Then
            Exit Sub
        End If
        gstrSQL = "Select Distinct a.����id, c.���� As ���ϱ���, c.���� As ͨ����, d.��Ʒ��, c.���, c.�Ƿ��� As ʱ��, c.����, c.���㵥λ,a.����ϵ��, a.��װ��λ," & _
                                    " a.�ɱ���, e.�ּ�, a.ָ��������, a.ָ�����ۼ�,a.��������" & _
                    " From �������� A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) D,�շѼ�Ŀ E" & _
                    " Where a.����id = b.Id And a.����id = c.Id And c.Id = d.�շ�ϸĿid(+) and a.����id=e.�շ�ϸĿid and sysdate between e.ִ������ and e.��ֹ���� " & _
                    GetPriceClassString("E") & "and b.id=[1] And (c.����ʱ�� = to_date('3000-01-01','yyyy-mm-dd') or c.����ʱ�� is null ) order by c.����"
    
        lngId = Val(.TextMatrix(.Row, 0))

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����", lngId, gstrPriceClass)
        If rsTemp.RecordCount = 0 Then
            picDrug.Visible = False
            Exit Sub
        End If
    End With
    
    'Ϊ���ֵ
    Call setVSFValue(rsTemp)

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub setVSFValue(ByVal rsTemp As ADODB.Recordset)
    Dim lngId As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim blnDou As Boolean '�ظ�����
    Dim dbl����ϵ�� As Double
    Dim strUnit As String   '��λ
    Dim intType As Integer '������ʽ
    Dim dbl���� As Double   '������
    
    intType = cbo������ʽ.ItemData(cbo������ʽ.ListIndex)
    'Ϊ���ֵ
    With vsfSelectDrug
        For intRow = 0 To rsTemp.RecordCount - 1
            blnDou = False
            For i = 1 To .Rows - 1
                If .TextMatrix(i, vsfSelectDrugCol.����ID) = rsTemp!����ID Then
                    blnDou = True
                End If
            Next
            If blnDou = False Then
                .Rows = .Rows + 1
                .RowHeight(.Rows - 1) = mlngRowHeight
            
                Select Case mintUnit
                    Case 0
                        dbl����ϵ�� = 1
                        strUnit = rsTemp!���㵥λ
                    Case 1
                        dbl����ϵ�� = rsTemp!����ϵ��
                        strUnit = rsTemp!��װ��λ
                End Select
                                
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.����ID) = rsTemp!����ID
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.������Ϣ) = "[" & rsTemp!���ϱ��� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)

                .TextMatrix(.Rows - 1, vsfSelectDrugCol.���ϱ���) = rsTemp!���ϱ���
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��Ʒ��) = IIf(IsNull(rsTemp!��Ʒ��), "", rsTemp!��Ʒ��)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ͨ����) = IIf(IsNull(rsTemp!ͨ����), "", rsTemp!ͨ����)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��λ) = strUnit
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��λϵ��) = dbl����ϵ��
                
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ɢװ��λ) = rsTemp!���㵥λ
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��װ��λ) = rsTemp!��װ��λ
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��װϵ��) = rsTemp!����ϵ��
                
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.����) = IIf(rsTemp!ʱ�� = 1, "ʱ��", "����")
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.��������) = zlStr.nvl(rsTemp!��������)
                                
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ɱ���) = Format(dbl����ϵ�� * rsTemp!�ɱ���, mFMT.FM_�ɱ���)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ������) = Format(dbl����ϵ�� * rsTemp!ָ��������, mFMT.FM_�ɱ���)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ���ۼ�) = Format(dbl����ϵ�� * rsTemp!ָ�����ۼ�, mFMT.FM_���ۼ�)
                .TextMatrix(.Rows - 1, vsfSelectDrugCol.ԭ�ۼ�) = Format(Val(zlStr.nvl(rsTemp!�ּ�)) * dbl����ϵ��, mFMT.FM_���ۼ�)
                
                dbl���� = Val(txt������.Text)
                If Trim(txt������.Text) = "" Then
                    .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�) * dbl����ϵ��, mFMT.FM_���ۼ�)
                Else
                    Select Case intType
                        Case 1      '���ݳɱ��ۼӳ�
                            dbl���� = 1 + Val(dbl����) / 100
                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format(Val(zlStr.nvl(rsTemp!�ɱ���)) * dbl���� * dbl����ϵ��, mFMT.FM_���ۼ�)
                        Case 2      '�������ۼ۰�����
                            dbl���� = 1 + Val(dbl����) / 100
                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format(Val(zlStr.nvl(rsTemp!�ּ�)) * dbl���� * dbl����ϵ��, mFMT.FM_���ۼ�)
                        Case 3      '�������ۼ۰��̶����Ӽ�
                            dbl���� = Val(dbl����)
                            .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format((Val(zlStr.nvl(rsTemp!�ּ�)) * dbl����ϵ��) + dbl����, mFMT.FM_���ۼ�)
                    End Select
                End If
                If Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�)) > Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ���ۼ�)) And Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ���ۼ�)) <> 0 Then
                    .TextMatrix(.Rows - 1, vsfSelectDrugCol.�ۼ�) = Format(Val(.TextMatrix(.Rows - 1, vsfSelectDrugCol.ָ���ۼ�)), mFMT.FM_���ۼ�)
                End If
            End If
            rsTemp.MoveNext
        Next
        picDrug.Visible = False
    End With
End Sub

Private Sub vsfDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vsfDrug_DblClick
    End If
End Sub

Private Sub vsfSelectDrug_GotFocus()
    If picDrug.Visible = True Then
        picDrug.Visible = False
    End If
End Sub

Private Sub vsfSelectDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And vsfSelectDrug.Rows > 1 Then
        vsfSelectDrug.RemoveItem vsfSelectDrug.Row
    End If
End Sub
