VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckNumber 
   Caption         =   "���յ�����"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15855
   Icon            =   "frmCheckNumber.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   15855
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȫ�� 
      Cancel          =   -1  'True
      Caption         =   "ȫ��(&B)"
      Height          =   350
      Left            =   11520
      TabIndex        =   23
      Top             =   7335
      Width           =   1100
   End
   Begin VB.CommandButton cmdȫѡ 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   10200
      TabIndex        =   22
      Top             =   7335
      Width           =   1100
   End
   Begin VB.PictureBox pic�ѵ������ɫ 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   7440
      Width           =   260
   End
   Begin VB.Frame Frmline2 
      Height          =   120
      Left            =   120
      TabIndex        =   16
      Top             =   6960
      Width           =   15735
   End
   Begin VB.Frame fraFilter 
      Caption         =   " ��ȡ��������"
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2415
      Begin VB.TextBox txt����NO 
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   2205
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   2205
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   2205
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "��ȡ����"
         Height          =   300
         Left            =   1110
         TabIndex        =   7
         Top             =   5280
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   216334339
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   216334339
         CurrentDate     =   36263
      End
      Begin VB.Label lbl��ʾ 
         Caption         =   "(����NO�������NO�Ź���)"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label lbl����NO 
         Caption         =   "����NO"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label lbl����ʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   3300
         Width           =   720
      End
      Begin VB.Label lbl��ʼʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   2505
         Width           =   720
      End
      Begin VB.Label lbl����ʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label lblStore 
         AutoSize        =   -1  'True
         Caption         =   "���տⷿ"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   14640
      TabIndex        =   2
      Top             =   7335
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "������ⵥ��(&O)"
      Height          =   350
      Left            =   12960
      TabIndex        =   1
      Top             =   7335
      Width           =   1575
   End
   Begin VB.Frame Frmline1 
      Height          =   120
      Left            =   120
      TabIndex        =   0
      Top             =   645
      Width           =   15735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   5655
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
      Width           =   13140
      _cx             =   23177
      _cy             =   9975
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
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   25
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCheckNumber.frx":6852
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
   End
   Begin VB.Label lbl�ѵ������ɫ 
      AutoSize        =   -1  'True
      Caption         =   "�Ѳ������⹺��ⵥ"
      Height          =   180
      Left            =   5760
      TabIndex        =   18
      Top             =   7470
      Width           =   1620
   End
   Begin VB.Label lbl���������� 
      AutoSize        =   -1  'True
      Caption         =   "��ʾ���������0����ⵥ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   7410
      Width           =   2625
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   255
      Picture         =   "frmCheckNumber.frx":6C18
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   " ˵����������ȡ���������˳��ϸ�����յ��ݣ�����ҩƷ�⹺��ⵥ��"
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   10695
   End
End
Attribute VB_Name = "frmCheckNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const m�ѵ����ColColor As Long = &H8080FF
Private Const mδ�����ColColor As Long = &H0&

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��

Private mintListIndex As Integer    '�ⷿ����

Private mfrmMain As Form

Public Sub ShowCard(FrmMain As Form, ByVal intListIndex As Integer)

    mintListIndex = intListIndex
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, FrmMain
End Sub

Private Sub cbo�ⷿ_Click()
    If Val(Cbo�ⷿ.ListIndex) <> Val(Cbo�ⷿ.Tag) And vsfList.Rows > 1 Then
        If MsgBox("����ı�ⷿ����Ҫ������ȡ�������ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            vsfList.Rows = 1
            lbl����������.Caption = "��ʾ�������й����ܻ����0����ⵥ��"
        Else
            Cbo�ⷿ.ListIndex = Val(Cbo�ⷿ.Tag)
        End If
    End If
    Cbo�ⷿ.Tag = Val(Cbo�ⷿ.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFilter_Click()
'���������
    Dim rsTemp As ADODB.Recordset
    Dim lng�ⷿID As Long
    Dim str��Ӧ��id As String
    Dim int�������� As Integer
    
    On Error GoTo errHandle

    If Cbo�ⷿ.ListIndex = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If

    vsfList.Rows = 1
    '�ⷿid
    lng�ⷿID = Val(Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex))
    
    gstrSQL = "Select Distinct f.Id As ����id, f.�ⷿid,f.��ҩ��λid,c.ҩƷid,a.������,g.���� As ��ҩ��λ,f.No,f.������,f.��������," & vbNewLine & _
                    "                Nvl(f.�Ƿ�ϸ�,0) As �Ƿ�ϸ�,b.����,b.����, b.���, a.���� As ������,a.����,a.��ҩ����,b.���㵥λ," & vbNewLine & _
                    "                a.�ɱ���, a.���ۼ�, a.��������, a.Ч��, a.��׼�ĺ�, a.���ս���,b.�Ƿ���" & vbNewLine & _
                    "From ҩƷ������ϸ A, �շ���ĿĿ¼ B, ҩƷ��� C, ҩƷ���� D, ҩƷ���� E, ҩƷ���ռ�¼ F, ��Ӧ�� G" & vbNewLine & _
                    "Where a.ҩƷid = b.Id And b.Id = c.ҩƷid And c.ҩ��id = d.ҩ��id And d.ҩƷ���� = e.����(+) And f.Id = a.����id And f.�Ƿ�ϸ� = 0 And" & vbNewLine & _
                    "      f.��ҩ��λid = g.Id(+) And f.�ⷿid = [1] And f.�������� Between [2] And [3]" & IIf(Trim(txt����NO.Text) <> "", " and f.No=[4]", "") & vbNewLine & _
                    "Order By a.������ Desc, f.No"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ҩƷ�������", lng�ⷿID, _
    CDate(Format(dtp��ʼʱ��.Value, "yyyy-mm-dd")), CDate(Format(dtp����ʱ��.Value, "yyyy-mm-dd") & " 23:59:59"), Trim(txt����NO.Text))

   If rsTemp.RecordCount = 0 Then MsgBox "û�в�ѯ���ϸ�����յ��ݣ����飡", vbInformation, gstrSysName: Exit Sub
   
    With vsfList
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("����id")) = rsTemp!����id
            .TextMatrix(.Rows - 1, .ColIndex("�ⷿid")) = rsTemp!�ⷿid
            .TextMatrix(.Rows - 1, .ColIndex("��Ӧ��id")) = rsTemp!��ҩ��λID
            .TextMatrix(.Rows - 1, .ColIndex("ҩƷid")) = rsTemp!ҩƷid
            .TextMatrix(.Rows - 1, .ColIndex("��ҩ��λ")) = NVL(rsTemp!��ҩ��λ)
            .TextMatrix(.Rows - 1, .ColIndex("����NO")) = rsTemp!NO
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(NVL(rsTemp!��������), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�ϸ�")) = IIf(NVL(rsTemp!�Ƿ�ϸ�, 0) = 0, "�ϸ�", "���ϸ�")
            .TextMatrix(.Rows - 1, .ColIndex("ҩƷ")) = "[" & rsTemp!���� & "]" & rsTemp!���� & "(" & rsTemp!��� & ")"
            .TextMatrix(.Rows - 1, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(.Rows - 1, .ColIndex("��ҩ����")) = zlStr.FormatEx(NVL(rsTemp!��ҩ����, 0), mintNumberDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("��λ")) = NVL(rsTemp!���㵥λ)
            .TextMatrix(.Rows - 1, .ColIndex("�ɱ���")) = zlStr.FormatEx(NVL(rsTemp!�ɱ���, 0), mintCostDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("�ɱ����")) = zlStr.FormatEx(NVL(rsTemp!�ɱ���, 0) * NVL(rsTemp!��ҩ����, 0), mintMoneyDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("���ۼ�")) = zlStr.FormatEx(NVL(rsTemp!���ۼ�, 0), mintPriceDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("���۽��")) = zlStr.FormatEx(NVL(rsTemp!���ۼ�, 0) * NVL(rsTemp!��ҩ����, 0), mintMoneyDigit, True, True)
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(NVL(rsTemp!��������), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("Ч��")) = Format(NVL(rsTemp!Ч��), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("��׼�ĺ�")) = NVL(rsTemp!��׼�ĺ�)
            .TextMatrix(.Rows - 1, .ColIndex("���ս���")) = NVL(rsTemp!���ս���)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsTemp!����)
            .TextMatrix(.Rows - 1, .ColIndex("�Ƿ���")) = NVL(rsTemp!�Ƿ���, 0)
            
            If NVL(rsTemp!�Ƿ���, 0) = 0 Then .TextMatrix(.Rows - 1, .ColIndex("�������ۼ�")) = NVL(rsTemp!���ۼ�, 0)
            
            .Cell(flexcpForeColor, .Rows - 1, .ColIndex("��ҩ��λ"), .Rows - 1, .ColIndex("���ս���")) = IIf(NVL(rsTemp!������, 0) = 0, mδ�����ColColor, m�ѵ����ColColor)
            
            If InStr(";" & str��Ӧ��id & ";", ";" & rsTemp!��ҩ��λID & ";") = 0 Then
                str��Ӧ��id = IIf(str��Ӧ��id = "", "", str��Ӧ��id & ";") & rsTemp!��ҩ��λID
                int�������� = int�������� + 1
            End If
            
            rsTemp.MoveNext
        Loop
    End With
    
    lbl����������.Caption = "��ʾ�������й����ܻ����" & int�������� & "����ⵥ��"
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdȫ��_Click()
    Dim i As Integer
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
        Next
    End With
End Sub

Private Sub cmdȫѡ_Click()
    Dim i As Integer
    With vsfList
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 1
        Next
    End With
End Sub

Private Sub Form_Load()
    vsfList.AllowSelection = False '���ܶ�ѡ
    vsfList.Rows = 1
    Call initComboBox
    Call SetMedicalWH
    Call GetDrugDigit(Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex), "ҩƷ���չ���", 4, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    lbl����������.Caption = "��ʾ�������й����ܻ����0����ⵥ��"
End Sub

Private Sub SetMedicalWH()
    Dim i As Integer

    With mfrmMain.cboStock
        Cbo�ⷿ.Clear
        For i = 0 To .ListCount - 1
            Cbo�ⷿ.AddItem .List(i)
            Cbo�ⷿ.ItemData(Cbo�ⷿ.NewIndex) = .ItemData(i)
        Next
        Cbo�ⷿ.ListIndex = .ListIndex
    End With
        
    Cbo�ⷿ.Tag = IIf(mintListIndex = -1, 0, mintListIndex)
    Cbo�ⷿ.ListIndex = IIf(mintListIndex = -1, 0, mintListIndex)
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 7500 Then Me.Height = 7500
    If Me.Width < 15000 Then Me.Width = 15000
    
    Frmline1.Left = 0
    Frmline1.Top = Me.ScaleHeight / 12
    Frmline1.Width = Me.ScaleWidth

    Frmline2.Left = 0
    Frmline2.Top = Me.ScaleHeight * 43 / 48
    Frmline2.Width = Me.ScaleWidth

    fraFilter.Left = 50
    fraFilter.Top = Frmline1.Top + 200
    fraFilter.Height = Frmline2.Top - Frmline1.Top - 300
    
    cmdFilter.Left = 1200
    cmdFilter.Top = fraFilter.Height - cmdFilter.Height - 200
    
    vsfList.Left = fraFilter.Width + 100
    vsfList.Top = fraFilter.Top + 100
    vsfList.Width = Me.ScaleWidth - fraFilter.Left - fraFilter.Width - 150
    vsfList.Height = fraFilter.Height - 100
    
    lbl����������.Left = 100
    lbl����������.Top = Frmline2.Top + 350
    
    CmdCancel.Left = Me.Width - CmdCancel.Width - 200
    CmdCancel.Top = lbl����������.Top - 50
    
    cmdOK.Left = CmdCancel.Left - cmdOK.Width - 50
    cmdOK.Top = lbl����������.Top - 50
    
    cmdȫ��.Top = cmdOK.Top
    cmdȫ��.Left = cmdOK.Left - cmdȫ��.Width - 50
    
    cmdȫѡ.Top = cmdOK.Top
    cmdȫѡ.Left = cmdȫ��.Left - cmdȫѡ.Width - 50
    
    lbl�ѵ������ɫ.Top = lbl����������.Top + 30
    pic�ѵ������ɫ.Top = lbl����������.Top

    
End Sub


Private Sub cmdOK_Click()
    Dim i As Integer
    Dim int�������� As Integer
    Dim strNo As String
    Dim str��Ӧ��id As String
    Dim strDate As String
    Dim int��� As Integer
    Dim lng�ⷿID As Long
    Dim blnTrans As Boolean
    Dim arrSql As Variant
    Dim blnOK As Boolean
    Dim rsSort As New ADODB.Recordset   '����Ӧ������
    Dim intRow As Integer
    
    If vsfList.Rows < 2 Then Exit Sub
    If Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex) = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error GoTo ErrHand

    If MsgBox("�Ƿ�ȷ��������ⵥ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    With rsSort
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "��Ӧ��id", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    With vsfList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("ѡ��")) Like "*1" Then
                rsSort.AddNew
                rsSort!�к� = i
                rsSort!��Ӧ��id = Val(.TextMatrix(i, .ColIndex("��Ӧ��id")))
                
                rsSort.Update
            End If
        Next
    End With
    
    If rsSort.RecordCount = 0 Then
        MsgBox "û��ѡ��Ҫ����ĵ��ݣ������ٹ�ѡһ�����ݺ��ٱ��棡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    rsSort.Sort = "��Ӧ��id,�к�"
        
    arrSql = Array()
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    lng�ⷿID = Val(Cbo�ⷿ.ItemData(Cbo�ⷿ.ListIndex))
    
    With vsfList
        For i = 1 To rsSort.RecordCount
            intRow = rsSort!�к�

            If InStr(";" & str��Ӧ��id & ";", ";" & .TextMatrix(intRow, .ColIndex("��Ӧ��id")) & ";") = 0 Then
                str��Ӧ��id = IIf(str��Ӧ��id = "", "", str��Ӧ��id & ";") & .TextMatrix(intRow, .ColIndex("��Ӧ��id"))
                strNo = zlDatabase.GetNextNo(21, lng�ⷿID)
                int�������� = int�������� + 1
                int��� = 0
            End If

            gstrSQL = "Zl_ҩƷ������ϸ_������(" & Val(.TextMatrix(intRow, .ColIndex("����id"))) & "," & Val(.TextMatrix(intRow, .ColIndex("ҩƷid"))) & ")"
                
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            int��� = int��� + 1
            
            gstrSQL = "zl_ҩƷ�⹺_INSERT("
            'NO
            gstrSQL = gstrSQL & "'" & strNo & "'"
            '���
            gstrSQL = gstrSQL & "," & int���
            '�ⷿID
            gstrSQL = gstrSQL & "," & lng�ⷿID
            '�Է�����ID
            gstrSQL = gstrSQL & ",NULL"
            '��ҩ��λID
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("��Ӧ��id")))
            'ҩƷID
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("ҩƷID")))
            '����
            gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, .ColIndex("������")) & "'"
            '����
            gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, .ColIndex("����")) & "'"
            'Ч��
            gstrSQL = gstrSQL & "," & "to_date('" & .TextMatrix(intRow, .ColIndex("Ч��")) & "','yyyy-mm-dd')"
            'ʵ������
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("��ҩ����")))
            '�ɱ���
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("�ɱ���")))
            '�ɱ����
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("�ɱ����")))
            '����
            gstrSQL = gstrSQL & "," & 100
            '���ۼ�
            If .TextMatrix(intRow, .ColIndex("�Ƿ���")) = 0 Then
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("�������ۼ�")))
            Else
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("���ۼ�")))
            End If
            '���۽��
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("���۽��")))
            '���
            gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("���۽��"))) - Val(.TextMatrix(intRow, .ColIndex("�ɱ����")))
            'ժҪ
            gstrSQL = gstrSQL & ",'ҩƷ���չ�����'"
            '������
            gstrSQL = gstrSQL & ",'" & UserInfo.�û����� & "'"
            '��Ʊ��
            gstrSQL = gstrSQL & ",NULL"
            '��Ʊ����
            gstrSQL = gstrSQL & ",NULL"
            '��Ʊ���
            gstrSQL = gstrSQL & ",NULL"
            '��������
            gstrSQL = gstrSQL & "," & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS')"
            '���
            gstrSQL = gstrSQL & ",NULL"
            '��Ʒ�ϸ�֤
            gstrSQL = gstrSQL & ",NULL"
            '�˲���
            gstrSQL = gstrSQL & ",NULL"
            '�˲�����
            gstrSQL = gstrSQL & ",NULL"
            '����
            gstrSQL = gstrSQL & "," & 0
            '�Ƿ��˻�
            gstrSQL = gstrSQL & "," & 1
            '��������
            gstrSQL = gstrSQL & "," & "to_date('" & .TextMatrix(intRow, .ColIndex("��������")) & "','yyyy-mm-dd')"
            '��׼�ĺ�
            gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, .ColIndex("��׼�ĺ�")) & "'"
            '�������
            gstrSQL = gstrSQL & ",NULL"
            '����
            gstrSQL = gstrSQL & "," & 0
            '�ӳ���
            If Val(.TextMatrix(intRow, .ColIndex("�ɱ���"))) = 0 Then
                gstrSQL = gstrSQL & "," & 0
            Else
                gstrSQL = gstrSQL & "," & Val(.TextMatrix(intRow, .ColIndex("���ۼ�"))) / Val(.TextMatrix(intRow, .ColIndex("�ɱ���"))) - 1
            End If
            '��Ʊ����
            gstrSQL = gstrSQL & ",NULL"
            '�ƻ�id
            gstrSQL = gstrSQL & ",NULL"
            '�������
            gstrSQL = gstrSQL & "," & 0
            'ԭ����
            gstrSQL = gstrSQL & ",NULL"
            '�������
            gstrSQL = gstrSQL & ",NULL"
            '���ս���
            gstrSQL = gstrSQL & ",'" & .TextMatrix(intRow, .ColIndex("���ս���")) & "'"
            
            gstrSQL = gstrSQL & ")"
                
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            rsSort.MoveNext
        Next
    
    End With
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "�����⹺��ⵥ")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    MsgBox "���ι�����" & int�������� & "���⹺��ⵥ����ע��鿴��", vbInformation, gstrSysName
    vsfList.Rows = 1
    
    lbl����������.Caption = "��ʾ�������й����ܻ����0����ⵥ��"
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub initComboBox()
    With cbo��������
        .Clear
        .AddItem "����"
        .AddItem "һ������"
        .AddItem "һ������"
        .AddItem "��������"
        .AddItem "�Զ�������"
        .ListIndex = 0
    End With
End Sub
Private Sub cbo��������_Click()
    Dim dateCurrentDate As Date
    
    If cbo��������.Text = "�Զ�������" Then
        dtp��ʼʱ��.Enabled = True
        dtp����ʱ��.Enabled = True
        
    Else
        dtp��ʼʱ��.Enabled = False
        dtp����ʱ��.Enabled = False
    End If
    
    '����ѡ��ı�ʱ��
    dateCurrentDate = sys.Currentdate
    Select Case cbo��������.ListIndex
        Case 0
            dtp��ʼʱ��.Value = CDate(Format(dateCurrentDate, "yyyy-mm-dd") & " 00:00:00")
            dtp����ʱ��.Value = dateCurrentDate
        Case 1
            dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -7, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp����ʱ��.Value = dateCurrentDate
        Case 2
            dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -30, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp����ʱ��.Value = dateCurrentDate
        Case 3
            dtp��ʼʱ��.Value = CDate(Format(DateAdd("d", -90, dateCurrentDate), "yyyy-mm-dd") & " 00:00:00")
            dtp����ʱ��.Value = dateCurrentDate
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
End Sub

Private Sub txt����NO_GotFocus()
    Me.txt����NO.SelStart = 0: Me.txt����NO.SelLength = 100
End Sub
Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub
Private Sub vsfList_EnterCell()
    If vsfList.Col = vsfList.ColIndex("ѡ��") Then
        vsfList.Editable = flexEDKbdMouse
    Else
        vsfList.Editable = flexEDNone
    End If
End Sub
Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode <> vbKeyDelete Then Exit Sub
    
    With vsfList
        If .Rows < 2 Then Exit Sub
        If MsgBox("�Ƿ����ɾ������NOΪ " & .TextMatrix(.Row, .ColIndex("����NO")) & "��ҩƷΪ " & .TextMatrix(.Row, .ColIndex("ҩƷ")) & " �����ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        .RemoveItem .Row
        
    End With
End Sub
