VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPurchaseCardExtract 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ȡ����"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9345
   Icon            =   "frmPurchaseCardExtract.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6840
      TabIndex        =   14
      Top             =   6120
      Width           =   1100
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8040
      TabIndex        =   13
      Top             =   6120
      Width           =   1100
   End
   Begin TabDlg.SSTab sstGuide 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10821
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPurchaseCardExtract.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOption"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPurchaseCardExtract.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picView"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox picView 
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   3315
         ScaleWidth      =   3315
         TabIndex        =   15
         Top             =   120
         Width           =   3375
         Begin VSFlex8Ctl.VSFlexGrid vsfView 
            Height          =   2175
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   2055
            _cx             =   3625
            _cy             =   3836
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
      End
      Begin VB.Frame fraOption 
         Caption         =   "��ȡ����"
         Height          =   2775
         Left            =   240
         TabIndex        =   1
         Top             =   150
         Width           =   5175
         Begin VB.TextBox txtNO 
            Height          =   300
            Index           =   1
            Left            =   3480
            TabIndex        =   12
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtNO 
            Height          =   300
            Index           =   0
            Left            =   1800
            TabIndex        =   10
            Top             =   2160
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpData 
            Height          =   300
            Index           =   0
            Left            =   1800
            TabIndex        =   6
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   186580993
            CurrentDate     =   40532
         End
         Begin VB.OptionButton optExtract 
            Caption         =   "��ȡ�������(&2)"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   4
            Top             =   1200
            Width           =   2295
         End
         Begin VB.OptionButton optExtract 
            Caption         =   "��ȡ�������(&1)"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   480
            Value           =   -1  'True
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtpData 
            Height          =   300
            Index           =   1
            Left            =   3480
            TabIndex        =   8
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   186580993
            CurrentDate     =   40532
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   3240
            TabIndex        =   11
            Top             =   2210
            Width           =   180
         End
         Begin VB.Label lblNO 
            AutoSize        =   -1  'True
            Caption         =   "��ⵥ��(&N)"
            Height          =   180
            Left            =   720
            TabIndex        =   9
            Top             =   2160
            Width           =   990
         End
         Begin VB.Label lblTo 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   0
            Left            =   3240
            TabIndex        =   7
            Top             =   1730
            Width           =   180
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            Caption         =   "���ʱ��(&T)"
            Height          =   180
            Left            =   720
            TabIndex        =   5
            Top             =   1680
            Width           =   990
         End
         Begin VB.Label lblStock 
            AutoSize        =   -1  'True
            Caption         =   "�ⷿ�� xxx"
            Height          =   180
            Left            =   720
            TabIndex        =   3
            Top             =   840
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmPurchaseCardExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngProviderID As Long  '��Ӧ��ID
Private mlngStockID As Long     '�ⷿID
Private mstrStock As String
Private mintUnit As Integer     '��ʾ��λ�� 0-ɢװ; 1-��װ
Private mFMT As g_FmtString

Private Const mlngModule = 1712

Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    If optExtract(0).Value Then
        Call ExtractStockData
    ElseIf optExtract(1).Value Then
        Call ExtractInStockData
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strReg As String
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(0, g_�ۼ�)
    End With
End Sub

Private Sub Form_Activate()
    Me.Visible = False
    Call cmdȷ��_Click
End Sub

Private Sub ExtractStockData()
    '��ȡ�������
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select b.Id, '[' || b.���� || ']' || b.���� ����,e.���� ��Ʒ��, b.���, b.����,a.��׼�ĺ�, " & IIf(mintUnit = 0, "b.���㵥λ", "c.��װ��λ") & " ��λ," & vbNewLine & _
              "       Decode(b.�Ƿ���, 1, a.ʵ�ʽ�� / a.ʵ������, d.�ּ�) * " & IIf(mintUnit = 0, "1", "c.����ϵ��") & " �ۼ�," & vbNewLine & _
              "       Decode(c.���÷���, 1, a.�ϴβɹ���, (a.ʵ�ʽ�� - a.ʵ�ʲ��) / a.ʵ������) * " & IIf(mintUnit = 0, "1", "c.����ϵ��") & " �ɱ���," & vbNewLine & _
              "       a.ʵ������ / " & IIf(mintUnit = 0, "1", "c.����ϵ��") & " ����," & vbNewLine & _
              "       c.���Ч��, " & IIf(mintUnit = 0, "1", "c.����ϵ��") & " ����ϵ��, a.����, b.�Ƿ���, c.���÷���, c.ָ������� / 100 ָ�������" & vbNewLine & _
              "From ҩƷ��� A, �շ���ĿĿ¼ B, �������� C, �շѼ�Ŀ D, �շ���Ŀ���� E" & vbNewLine & _
              "Where a.ҩƷid = b.Id And a.ҩƷid = c.����id And a.ҩƷid = d.�շ�ϸĿid And a.���� = 1 And a.�ⷿid = [1] And a.ʵ������ > 0 And" & vbNewLine & _
              "      a.�ϴι�Ӧ��ID=[2] And b.��� = '4' And b.����ʱ�� >= To_Date('3000-1-1', 'yyyy-mm-dd') And d.��ֹ���� >= To_Date('3000-1-1', 'yyyy-mm-dd')" & vbNewLine & _
              GetPriceClassString("D") & " And b.Id = e.�շ�ϸĿid(+) And e.����(+) = 3" & vbNewLine & _
              " Order By a.ҩƷID, a.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Caption, mlngStockID, mlngProviderID)
    FillData rsTmp
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ExtractInStockData()
    '��ȡ�������
End Sub

Private Sub FillData(ByVal rsVal As ADODB.Recordset)
    '��д���ݵ���Ƭ��
    If rsVal.RecordCount = 0 Then
        MsgBox "δ��ȡ��������ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Dim i As Long
    
    With frmPurchaseCard
        If .mshBill.Rows > 1 And Trim(.mshBill.TextMatrix(1, 0)) <> "" Then
            If MsgBox("�˻���Ƭ�������ݽ�ȫ�����������ȡ������ݣ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        Me.MousePointer = vbHourglass
        .mshBill.Clear
        .mshBill.Rows = 2
        For i = 1 To rsVal.RecordCount
            .mshBill.TextMatrix(i, 1) = i
            .SetColValue i, rsVal!Id, rsVal!����, IIf(IsNull(rsVal!���), "", rsVal!���), IIf(IsNull(rsVal!����), "", rsVal!����) _
                , IIf(IsNull(rsVal!��λ), "", rsVal!��λ) _
                , IIf(IsNull(rsVal!�ۼ�), 0, Format(rsVal!�ۼ�, IIf(mintUnit = 0, mFMT.FM_ɢװ���ۼ�, mFMT.FM_���ۼ�))) _
                , IIf(IsNull(rsVal!�ɱ���), 0, Format(rsVal!�ɱ���, mFMT.FM_�ɱ���)) _
                , IIf(IsNull(rsVal!����), "", rsVal!����), IIf(IsNull(rsVal!���Ч��), 0, rsVal!���Ч��), "" _
                , rsVal!����ϵ��, IIf(IsNull(rsVal!����), 0, rsVal!����), IIf(IsNull(rsVal!�Ƿ���), 0, rsVal!�Ƿ���) _
                , IIf(IsNull(rsVal!���÷���), 0, rsVal!���÷���), rsVal!ָ�������, IIf(IsNull(rsVal!��׼�ĺ�), "", rsVal!��׼�ĺ�), IIf(IsNull(rsVal!��Ʒ��), "", rsVal!��Ʒ��)
            .mshBill.TextMatrix(i, 21) = Format(rsVal!����, mFMT.FM_����)
            rsVal.MoveNext
            If Not rsVal.EOF Then .mshBill.Rows = .mshBill.Rows + 1
        Next
        Me.MousePointer = vbDefault
    End With
End Sub

Public Sub EntryPort(ByVal strStock As String, ByVal lngProviderID As Long)
    mlngStockID = Mid(strStock, 1, InStr(strStock, ";") - 1)
    mstrStock = Mid(strStock, InStr(strStock, ";") + 1)
    mlngProviderID = lngProviderID
    LblStock.Caption = "�ⷿ��" & mstrStock
End Sub

