VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPurchaseInputReturn 
   Caption         =   "����ⵥ���˻�"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12315
   Icon            =   "frmPurchaseInputReturn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   12315
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cboInputDate 
      Height          =   300
      Left            =   4935
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   60
      Width           =   1440
   End
   Begin VB.TextBox txtNO 
      Height          =   300
      Left            =   2400
      TabIndex        =   13
      Top             =   60
      Width           =   1440
   End
   Begin VB.TextBox txtLeechdom 
      Height          =   300
      Left            =   480
      TabIndex        =   12
      Top             =   60
      Width           =   1440
   End
   Begin VB.CheckBox chkAllSelect 
      Height          =   200
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   1320
      TabIndex        =   10
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8280
      TabIndex        =   9
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6960
      TabIndex        =   8
      Top             =   6000
      Width           =   1100
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "��ȡ(&G)"
      Height          =   300
      Left            =   11400
      TabIndex        =   6
      Top             =   60
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
      Height          =   315
      Left            =   7440
      TabIndex        =   2
      Top             =   53
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   214630403
      CurrentDate     =   40848
   End
   Begin MSComCtl2.DTPicker dtp����ʱ�� 
      Height          =   315
      Left            =   9960
      TabIndex        =   3
      Top             =   53
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy��MM��dd��"
      Format          =   214630403
      CurrentDate     =   40848
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2925
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   9615
      _cx             =   16960
      _cy             =   5159
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
      Rows            =   1
      Cols            =   44
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPurchaseInputReturn.frx":000C
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
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "���ʱ��"
      Height          =   180
      Left            =   4080
      TabIndex        =   15
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   9120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lbl��ʼ���� 
      Caption         =   "��ʼ����"
      Height          =   255
      Left            =   6645
      TabIndex        =   4
      Top             =   83
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "NO"
      Height          =   180
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   180
   End
   Begin VB.Label lblLeechdom 
      AutoSize        =   -1  'True
      Caption         =   "ҩƷ"
      Height          =   180
      Left            =   50
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmPurchaseInputReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngProvider As Long
Private mintUnit As Integer   '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mrsData As ADODB.Recordset  'Ҫ���ص����ݼ�
Private mlngStoreroom As Long   '�ⷿid
Private mstrSelectInfo As String '�Ѿ�ѡ���ҩƷ

Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��

Public Sub ShowMe(ByVal frmParent As Form, ByVal lngStoreroom As Long, ByVal lngProvider As Long, ByVal intUnit As Integer, ByVal intCostDigit As Integer, ByVal intPricedigit As Integer, ByVal intNumberDigit As Integer, ByVal intMoneyDigit As Integer, ByRef rsData As ADODB.Recordset)
    mlngStoreroom = lngStoreroom
    mintUnit = intUnit
    mlngProvider = lngProvider
    mintCostDigit = intCostDigit
    mintPriceDigit = intPricedigit
    mintNumberDigit = intNumberDigit
    mintMoneyDigit = intMoneyDigit
    
    Me.Show vbModal, frmParent
    Set rsData = mrsData
End Sub

Private Sub cboInputDate_Click()
    With cboInputDate
        If .Text = "�Զ���" Then
            lbl��ʼ����.Visible = True
            dtp��ʼʱ��.Visible = True
            Lbl��������.Visible = True
            dtp����ʱ��.Visible = True
        Else
            lbl��ʼ����.Visible = False
            dtp��ʼʱ��.Visible = False
            Lbl��������.Visible = False
            dtp����ʱ��.Visible = False
        End If
    End With
End Sub

Private Sub chkAllSelect_Click()
    Dim lngRow As Long
    
    With vsfList
        If .rows = 1 Then Exit Sub
        If chkAllSelect.Value = 1 Then
            For lngRow = 1 To .rows - 1
                If InStr(1, mstrSelectInfo, .TextMatrix(lngRow, .ColIndex("ҩƷid")) & "," & Val(.TextMatrix(lngRow, .ColIndex("����"))) & "|") = 0 Then
                    .TextMatrix(lngRow, 0) = "��"
                    mstrSelectInfo = mstrSelectInfo & .TextMatrix(lngRow, .ColIndex("ҩƷid")) & "," & Val(.TextMatrix(lngRow, .ColIndex("����"))) & "|"
                End If
            Next
        Else
            For lngRow = 1 To .rows - 1
                .TextMatrix(lngRow, 0) = ""
            Next
            mstrSelectInfo = ""
        End If
    End With
End Sub

Private Sub cmdAllCls_Click()
    With vsfList
        .rows = 1
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGetData_Click()
    Dim dbBeginDate As Date
    Dim dbEndDate As Date
    Dim rsTemp As ADODB.Recordset
    Dim rsSumNum As ADODB.Recordset
    Dim lngLeechdom As Long
    Dim strNo As String
    Dim lngRow As Long
    Dim strCurUnit As String
    Dim int����ϵ�� As Integer
    
    On Error GoTo errHandle
    
    If cboInputDate.Text = "������" Then
        dbBeginDate = CDate(Format(Date, "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "һ����" Then
        dbBeginDate = CDate(Format(DateAdd("d", -7, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "һ����" Then
        dbBeginDate = CDate(Format(DateAdd("M", -1, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "��������" Then
        dbBeginDate = CDate(Format(DateAdd("M", -3, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "һ����" Then
        dbBeginDate = CDate(Format(DateAdd("yyyy", -1, Date), "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(Date, "yyyy-mm-dd") & " 23:59:59")
    End If
    If cboInputDate.Text = "�Զ���" Then
        dbBeginDate = CDate(Format(dtp��ʼʱ��.Value, "yyyy-mm-dd") & " 00:00:00")
        dbEndDate = CDate(Format(dtp����ʱ��.Value, "yyyy-mm-dd") & " 23:59:59")
    End If
    
    strNo = Trim(txtNo.Text)
    
    gstrSQL = ""
    If Trim(txtLeechdom.Text) <> "" And txtLeechdom.Tag <> "" Then
        lngLeechdom = txtLeechdom.Tag
        gstrSQL = " And d.id = [5]"
    End If
    If Trim(txtNo.Text) <> "" Then
        gstrSQL = gstrSQL & " And a.no = [6]"
    End If
    '����30�����ʾ�û�
    If dbEndDate - dbBeginDate > 30 Then
        If MsgBox("��ѯʱ�䷶Χ̫�����Ƿ������", vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    gstrSQL = "Select Distinct a.No, a.ҩƷid, d.����, d.����, d.���, Nvl(d.�Ƿ���, 0) As �Ƿ���, d.���㵥λ, e.���ﵥλ, e.�����װ, e.סԺ��λ, e.סԺ��װ, e.ҩ�ⵥλ," & vbNewLine & _
            "                e.ҩ���װ, e.ָ��������, e.ָ�������, e.ҩƷ��Դ, e.����ҩ��, e.ҩ�ۼ���, e.ҩ�����, e.ҩ������, e.���Ч��, Nvl(e.���������, 0) As ���������," & vbNewLine & _
            "                Nvl(e.�б�ҩƷ, 0) As �б�ҩƷ, a.����, a.����, nvl(a.����,0) as ����, a.���, a.��׼�ĺ�, a.��������, a.ʵ������, a.�ɱ���, a.���ۼ�, a.�ɱ����, a.���۽��," & vbNewLine & _
            "                a.���, a.��������, a.Ч��, c.�������, c.��Ʊ��, c.��Ʊ����, c.��Ʊ���, c.��Ʊ����" & vbNewLine & _
            "From ҩƷ�շ���¼ A, ҩƷ��� B, Ӧ����¼ C, �շ���ĿĿ¼ D, ҩƷ��� E" & vbNewLine & _
            "Where a.ҩƷid + 0 = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) And a.�ⷿid = b.�ⷿid And a.Id = c.�շ�id(+) And b.ҩƷid = d.Id And" & vbNewLine & _
            "      d.Id = e.ҩƷid And b.�ⷿid = [1] And a.��ҩ��λid + 0 = [2] And b.���� = 1 And a.���� = 1 And Nvl(a.��ҩ��ʽ, 0) = 0 And" & vbNewLine & _
            "      (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0) And a.������� Between [3] And" & vbNewLine & _
            "      [4]" & gstrSQL & vbNewLine & _
            "order by no"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�⹺��¼", mlngStoreroom, mlngProvider, dbBeginDate, dbEndDate, lngLeechdom, strNo)
    mstrSelectInfo = ""
    vsfList.rows = 1
    chkAllSelect.Value = 0
    
    If rsTemp.RecordCount > 0 Then
        vsfList.rows = rsTemp.RecordCount + 1
        lngRow = 1
    End If
    Do While Not rsTemp.EOF
        With vsfList
            .TextMatrix(lngRow, .ColIndex("no")) = rsTemp!NO
            .TextMatrix(lngRow, .ColIndex("��������")) = "[" & rsTemp!���� & "]" & IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���")) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("ҩƷ��Դ")) = IIf(IsNull(rsTemp!ҩƷ��Դ), "", rsTemp!ҩƷ��Դ)
            .TextMatrix(lngRow, .ColIndex("����ҩ��")) = IIf(IsNull(rsTemp!����ҩ��), "", rsTemp!����ҩ��)
            .TextMatrix(lngRow, .ColIndex("��׼�ĺ�")) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
            .TextMatrix(lngRow, .ColIndex("��������")) = IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
            '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
            Select Case mintUnit
            Case 1
                strCurUnit = rsTemp!���㵥λ
                int����ϵ�� = 1
            Case 2
                strCurUnit = rsTemp!���ﵥλ
                int����ϵ�� = rsTemp!�����װ
            Case 3
                strCurUnit = rsTemp!סԺ��λ
                int����ϵ�� = rsTemp!סԺ��װ
            Case 4
                strCurUnit = rsTemp!ҩ�ⵥλ
                int����ϵ�� = rsTemp!ҩ���װ
            End Select
            .TextMatrix(lngRow, .ColIndex("��λ")) = strCurUnit
            
            gstrSQL = "Select Sum(ʵ������) As ʵ������" & vbNewLine & _
                "From ҩƷ���" & vbNewLine & _
                "Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And Nvl(����, 0) = [3]"
            Set rsSumNum = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ���", mlngStoreroom, rsTemp!ҩƷid, rsTemp!����)
            If rsSumNum.RecordCount > 0 Then
                .TextMatrix(lngRow, .ColIndex("�������")) = GetFormat(rsSumNum!ʵ������, mintNumberDigit)
            Else
                .TextMatrix(lngRow, .ColIndex("�������")) = 0
            End If
            .TextMatrix(lngRow, .ColIndex("�������")) = GetFormat(rsTemp!ʵ������, mintNumberDigit)
            .TextMatrix(lngRow, .ColIndex("�ɱ���")) = GetFormat(IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ��� * int����ϵ��), mintCostDigit)
            .TextMatrix(lngRow, .ColIndex("�ۼ�")) = GetFormat(IIf(IsNull(rsTemp!���ۼ�), 0, rsTemp!���ۼ� * int����ϵ��), mintPriceDigit)
            .TextMatrix(lngRow, .ColIndex("�ɱ����")) = GetFormat(rsTemp!�ɱ����, mintMoneyDigit)
            .TextMatrix(lngRow, .ColIndex("�ۼ۽��")) = GetFormat(rsTemp!���۽��, mintMoneyDigit)
            .TextMatrix(lngRow, .ColIndex("���")) = GetFormat(rsTemp!���, mintMoneyDigit)
            .TextMatrix(lngRow, .ColIndex("�������")) = IIf(IsNull(rsTemp!�������), "", rsTemp!�������)
            .TextMatrix(lngRow, .ColIndex("��Ʊ��")) = IIf(IsNull(rsTemp!��Ʊ��), "", rsTemp!��Ʊ��)
            .TextMatrix(lngRow, .ColIndex("��Ʊ����")) = IIf(IsNull(rsTemp!��Ʊ����), "", rsTemp!��Ʊ����)
            .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = GetFormat(IIf(IsNull(rsTemp!��Ʊ���), 0, rsTemp!��Ʊ���), mintMoneyDigit)
            .TextMatrix(lngRow, .ColIndex("��Ʊ����")) = Format(IIf(IsNull(rsTemp!��Ʊ����), "", rsTemp!��Ʊ����), "yyyy-mm-dd")
            .TextMatrix(lngRow, .ColIndex("ҩƷid")) = rsTemp!ҩƷid
            .TextMatrix(lngRow, .ColIndex("Ч��")) = Format(IIf(IsNull(rsTemp!Ч��), "", rsTemp!Ч��), "yyyy-mm-dd")
            .TextMatrix(lngRow, .ColIndex("ָ��������")) = GetFormat(rsTemp!ָ��������, mintCostDigit)
            .TextMatrix(lngRow, .ColIndex("ָ�������")) = GetFormat(rsTemp!ָ�������, mintCostDigit)
            .TextMatrix(lngRow, .ColIndex("�Ƿ���")) = rsTemp!�Ƿ���
            .TextMatrix(lngRow, .ColIndex("ҩ�����")) = IIf(IsNull(rsTemp!ҩ�����), 0, rsTemp!ҩ�����)
            .TextMatrix(lngRow, .ColIndex("ҩ������")) = IIf(IsNull(rsTemp!ҩ������), 0, rsTemp!ҩ������)
            .TextMatrix(lngRow, .ColIndex("���Ч��")) = IIf(IsNull(rsTemp!���Ч��), 0, rsTemp!���Ч��)
            .TextMatrix(lngRow, .ColIndex("���������")) = IIf(IsNull(rsTemp!���������), 0, rsTemp!���������)
            .TextMatrix(lngRow, .ColIndex("�б�ҩƷ")) = IIf(IsNull(rsTemp!�б�ҩƷ), 0, rsTemp!�б�ҩƷ)
            .TextMatrix(lngRow, .ColIndex("����ϵ��")) = int����ϵ��
            .TextMatrix(lngRow, .ColIndex("���㵥λ")) = IIf(IsNull(rsTemp!���㵥λ), "", rsTemp!���㵥λ)
            .TextMatrix(lngRow, .ColIndex("���ﵥλ")) = IIf(IsNull(rsTemp!���ﵥλ), "", rsTemp!���ﵥλ)
            .TextMatrix(lngRow, .ColIndex("�����װ")) = IIf(IsNull(rsTemp!�����װ), 1, rsTemp!�����װ)
            .TextMatrix(lngRow, .ColIndex("סԺ��λ")) = IIf(IsNull(rsTemp!סԺ��λ), "", rsTemp!סԺ��λ)
            .TextMatrix(lngRow, .ColIndex("סԺ��װ")) = IIf(IsNull(rsTemp!סԺ��װ), 1, rsTemp!סԺ��װ)
            .TextMatrix(lngRow, .ColIndex("ҩ�ⵥλ")) = IIf(IsNull(rsTemp!ҩ�ⵥλ), "", rsTemp!ҩ�ⵥλ)
            .TextMatrix(lngRow, .ColIndex("ҩ���װ")) = IIf(IsNull(rsTemp!ҩ���װ), 1, rsTemp!ҩ���װ)
            .TextMatrix(lngRow, .ColIndex("ҩ�ۼ���")) = IIf(IsNull(rsTemp!ҩ�ۼ���), "", rsTemp!ҩ�ۼ���)
            .TextMatrix(lngRow, .ColIndex("���")) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            lngRow = lngRow + 1
        End With
        rsTemp.MoveNext
    Loop
    
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdSave_Click()
    Dim lngRow As Long
    
    Call GetAssembled   '��ʼ�����ݼ�
    
    With vsfList
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) = "��" And .TextMatrix(lngRow, .ColIndex("ҩƷid")) <> "" Then
                mrsData.AddNew
                
                mrsData!ҩƷid = .TextMatrix(lngRow, .ColIndex("ҩƷid"))
                mrsData!�������� = .TextMatrix(lngRow, .ColIndex("��������"))
                mrsData!��� = .TextMatrix(lngRow, .ColIndex("���"))
                mrsData!���� = .TextMatrix(lngRow, .ColIndex("����"))
                mrsData!���� = .TextMatrix(lngRow, .ColIndex("����"))
                mrsData!���� = .TextMatrix(lngRow, .ColIndex("����"))
                mrsData!ҩƷ��Դ = .TextMatrix(lngRow, .ColIndex("ҩƷ��Դ"))
                mrsData!����ҩ�� = .TextMatrix(lngRow, .ColIndex("����ҩ��"))
                mrsData!��׼�ĺ� = .TextMatrix(lngRow, .ColIndex("��׼�ĺ�"))
                mrsData!�������� = .TextMatrix(lngRow, .ColIndex("��������"))
                mrsData!��λ = .TextMatrix(lngRow, .ColIndex("��λ"))
                mrsData!������� = .TextMatrix(lngRow, .ColIndex("�������"))
                mrsData!������� = .TextMatrix(lngRow, .ColIndex("�������"))
                mrsData!�ۼ� = .TextMatrix(lngRow, .ColIndex("�ۼ�"))
                mrsData!�ɱ��� = .TextMatrix(lngRow, .ColIndex("�ɱ���"))
                mrsData!�ɱ���� = .TextMatrix(lngRow, .ColIndex("�ɱ����"))
                mrsData!�ۼ۽�� = .TextMatrix(lngRow, .ColIndex("�ۼ۽��"))
                mrsData!��� = .TextMatrix(lngRow, .ColIndex("���"))
                mrsData!����ϵ�� = .TextMatrix(lngRow, .ColIndex("����ϵ��"))
                mrsData!������� = .TextMatrix(lngRow, .ColIndex("�������"))
                mrsData!��Ʊ�� = .TextMatrix(lngRow, .ColIndex("��Ʊ��"))
                mrsData!��Ʊ���� = .TextMatrix(lngRow, .ColIndex("��Ʊ����"))
                mrsData!��Ʊ��� = .TextMatrix(lngRow, .ColIndex("��Ʊ���"))
                mrsData!��Ʊ���� = .TextMatrix(lngRow, .ColIndex("��Ʊ����"))
                mrsData!Ч�� = .TextMatrix(lngRow, .ColIndex("Ч��"))
                mrsData!ָ�������� = .TextMatrix(lngRow, .ColIndex("ָ��������"))
                mrsData!ָ������� = .TextMatrix(lngRow, .ColIndex("ָ�������"))
                mrsData!�Ƿ��� = .TextMatrix(lngRow, .ColIndex("�Ƿ���"))
                mrsData!ҩ����� = .TextMatrix(lngRow, .ColIndex("ҩ�����"))
                mrsData!ҩ������ = .TextMatrix(lngRow, .ColIndex("ҩ������"))
                mrsData!���Ч�� = .TextMatrix(lngRow, .ColIndex("���Ч��"))
                mrsData!��������� = .TextMatrix(lngRow, .ColIndex("���������"))
                mrsData!�б�ҩƷ = .TextMatrix(lngRow, .ColIndex("�б�ҩƷ"))
                mrsData!ҩ�ۼ��� = .TextMatrix(lngRow, .ColIndex("ҩ�ۼ���"))
                mrsData!��� = .TextMatrix(lngRow, .ColIndex("���"))
                mrsData!���㵥λ = .TextMatrix(lngRow, .ColIndex("���㵥λ"))
                mrsData!���ﵥλ = .TextMatrix(lngRow, .ColIndex("���ﵥλ"))
                mrsData!�����װ = .TextMatrix(lngRow, .ColIndex("�����װ"))
                mrsData!סԺ��λ = .TextMatrix(lngRow, .ColIndex("סԺ��λ"))
                mrsData!סԺ��װ = .TextMatrix(lngRow, .ColIndex("סԺ��װ"))
                mrsData!ҩ�ⵥλ = .TextMatrix(lngRow, .ColIndex("ҩ�ⵥλ"))
                mrsData!ҩ���װ = .TextMatrix(lngRow, .ColIndex("ҩ���װ"))
                
                mrsData.Update
            End If
        Next
        Unload Me
    End With
End Sub

Private Sub GetAssembled()
    '��ʼ�����ݼ�
    Set mrsData = New ADODB.Recordset

    With mrsData
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 60, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "ҩƷ��Դ", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����ҩ��", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "��׼�ĺ�", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "�ɱ���", adDouble, 16, adFldIsNullable
        .Fields.Append "�ۼ�", adDouble, 16, adFldIsNullable
        .Fields.Append "�ɱ����", adDouble, 18, adFldIsNullable
        .Fields.Append "�ۼ۽��", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "����ϵ��", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "�������", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "��Ʊ��", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "��Ʊ����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��Ʊ���", adDouble, 18, adFldIsNullable
        .Fields.Append "��Ʊ����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ָ��������", adDouble, 16, adFldIsNullable
        .Fields.Append "ָ�������", adDouble, 16, adFldIsNullable
        .Fields.Append "�Ƿ���", adLongVarChar, 2, adFldIsNullable
        .Fields.Append "ҩ�����", adLongVarChar, 2, adFldIsNullable
        .Fields.Append "ҩ������", adLongVarChar, 2, adFldIsNullable
        .Fields.Append "���Ч��", adLongVarChar, 5, adFldIsNullable
        .Fields.Append "���������", adLongVarChar, 5, adFldIsNullable
        .Fields.Append "�б�ҩƷ", adDouble, 16, adFldIsNullable
        .Fields.Append "ҩ�ۼ���", adLongVarChar, 2, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "���㵥λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "���ﵥλ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "�����װ", adDouble, 11, adFldIsNullable
        .Fields.Append "סԺ��λ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "סԺ��װ", adDouble, 11, adFldIsNullable
        .Fields.Append "ҩ�ⵥλ", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "ҩ���װ", adDouble, 11, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
End Sub

Private Sub Form_Load()
    Set mrsData = Nothing
    
    dtp��ʼʱ��.Value = DateAdd("d", -2, zlDatabase.Currentdate)
    dtp����ʱ��.Value = DateAdd("d", -0, zlDatabase.Currentdate)
    Call InitComBox
End Sub

Private Sub Form_Resize()
    cmdGetData.Left = Me.ScaleWidth - cmdGetData.Width - 100
    With cmdAllCls
        .Top = Me.ScaleHeight - .Height - 100
        .Left = 200
    End With
    
    With cmdCancel
        .Top = cmdAllCls.Top
        .Left = Me.ScaleWidth - .Width - 200
    End With
    
    With CmdSave
        .Top = cmdAllCls.Top
        .Left = cmdCancel.Left - .Width - 200
    End With
    
    With vsfList
        .Left = 50
        .Top = txtLeechdom.Top + txtLeechdom.Height + 50
        .Width = Me.ScaleWidth - 50
        .Height = Me.ScaleHeight - .Top - cmdAllCls.Height - 200
    End With
    
    With chkAllSelect
        .Left = vsfList.Left + 70
        .Top = vsfList.Top + 60
    End With
End Sub

Private Sub InitComBox()
    '��ʼ�������б�
    With cboInputDate
        .AddItem "������"
        .AddItem "һ����"
        .AddItem "һ����"
        .AddItem "��������"
        .AddItem "һ����"
        .AddItem "�Զ���"
        .ListIndex = 0
    End With
End Sub

Private Sub txtLeechdom_Change()
    If txtLeechdom.Text = "" Then
        txtLeechdom.Tag = ""
    End If
End Sub

Private Sub txtLeechdom_GotFocus()
    zlControl.TxtSelAll txtLeechdom
End Sub

Private Sub txtLeechdom_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecReturn As ADODB.Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtLeechdom.Text) = "" Then Exit Sub
    sngLeft = Me.Left + txtLeechdom.Left
    sngTop = Me.Top + txtLeechdom.Top + txtLeechdom.Height + 500
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - txtLeechdom.Height - 3630
    End If
    
    strkey = Trim(txtLeechdom.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "ҩƷ�⹺������", mlngStoreroom, mlngStoreroom)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 1, 1, txtLeechdom.Text, sngLeft, sngTop, mlngStoreroom, mlngStoreroom, mlngStoreroom, , , , , , False)
    If RecReturn.RecordCount > 0 Then
        txtLeechdom.Tag = RecReturn!ҩƷid
        txtLeechdom.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    Else
        txtLeechdom.Tag = ""
    End If
End Sub



Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNo
End Sub


Private Sub txtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txtNo) < 8 And Len(txtNo) > 0 Then
            txtNo.Text = GetFullNO(txtNo.Text, 21, mlngStoreroom)
        End If
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub


Private Sub vsfList_Click()
    With vsfList
        If .Row = 0 Or .rows = 1 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then
            If InStr(1, mstrSelectInfo, .TextMatrix(.Row, .ColIndex("ҩƷid")) & "," & Val(.TextMatrix(.Row, .ColIndex("����"))) & "|") = 0 Then
                .TextMatrix(.Row, 0) = "��"
                mstrSelectInfo = mstrSelectInfo & .TextMatrix(.Row, .ColIndex("ҩƷid")) & "," & Val(.TextMatrix(.Row, .ColIndex("����"))) & "|"
            Else
                MsgBox "������ҩƷ�Ѿ�ѡ���ˣ�����Ҫѡ���Σ�", vbInformation, gstrSysName
            End If
        Else
            .TextMatrix(.Row, 0) = ""
            If InStr(1, mstrSelectInfo, .TextMatrix(.Row, .ColIndex("ҩƷid")) & "," & Val(.TextMatrix(.Row, .ColIndex("����"))) & "|") > 0 Then
                mstrSelectInfo = Replace(mstrSelectInfo, .TextMatrix(.Row, .ColIndex("ҩƷid")) & "," & Val(.TextMatrix(.Row, .ColIndex("����"))) & "|", "")
            End If
        End If
    End With
End Sub


