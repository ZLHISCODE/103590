VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHandBackPlanModify 
   Caption         =   "ҩƷ��ҩ�ƻ��༭"
   ClientHeight    =   8175
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11760
   Icon            =   "frmHandBackPlanModify.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11760
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraControl 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   7560
      Width           =   13095
      Begin VB.CommandButton cmdClear 
         Caption         =   "���(&D)"
         Height          =   350
         Left            =   8160
         TabIndex        =   38
         ToolTipText     =   "���������"
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2040
         TabIndex        =   27
         Top             =   145
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Left            =   10560
         TabIndex        =   26
         ToolTipText     =   "�����¼"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton CmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   11760
         TabIndex        =   25
         ToolTipText     =   "�������˳�"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "����(&R)"
         Height          =   350
         Left            =   9360
         TabIndex        =   24
         Tag             =   "����������ҩ����"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   105
         TabIndex        =   23
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblComment1 
         AutoSize        =   -1  'True
         Caption         =   "�ڱ��ڰ�F3������������"
         Height          =   180
         Left            =   3840
         TabIndex        =   39
         Top             =   205
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label lblFindType 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1440
         TabIndex        =   28
         Top             =   210
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin VB.PictureBox picBill 
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   13035
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.Frame fraComment 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         TabIndex        =   29
         Top             =   6720
         Width           =   12975
         Begin VB.TextBox Txt�������� 
            Height          =   300
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   1770
         End
         Begin VB.TextBox txt������ 
            Height          =   300
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox txtNo 
            Height          =   300
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox txtժҪ 
            Height          =   300
            Left            =   7020
            MaxLength       =   40
            TabIndex        =   30
            Top             =   240
            Width           =   5835
         End
         Begin VB.Label Lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            Height          =   180
            Left            =   1680
            TabIndex        =   34
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Lbl�������� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            Height          =   180
            Left            =   3600
            TabIndex        =   33
            Top             =   300
            Width           =   720
         End
         Begin VB.Label lblժҪ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ժҪ"
            Height          =   180
            Left            =   6480
            TabIndex        =   32
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblNo 
            AutoSize        =   -1  'True
            Caption         =   "NO"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   300
            Width           =   180
         End
      End
      Begin VB.Frame fraCondition 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   12855
         Begin VB.CommandButton CmdSelecter 
            Caption         =   "��"
            Height          =   300
            Index           =   2
            Left            =   10470
            TabIndex        =   9
            Top             =   580
            Width           =   255
         End
         Begin VB.CommandButton CmdSelecter 
            Caption         =   "��"
            Height          =   300
            Index           =   1
            Left            =   4870
            TabIndex        =   8
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton CmdSelecter 
            Caption         =   "��"
            Height          =   300
            Index           =   0
            Left            =   7400
            TabIndex        =   7
            Top             =   180
            Width           =   255
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Index           =   2
            Left            =   6315
            TabIndex        =   6
            ToolTipText     =   "���������̱��롢���������"
            Top             =   600
            Width           =   4170
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Index           =   1
            Left            =   720
            TabIndex        =   5
            ToolTipText     =   "���빩Ӧ�̱��롢���������"
            Top             =   600
            Width           =   4170
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Index           =   0
            Left            =   4320
            TabIndex        =   4
            ToolTipText     =   "����ҩƷ���롢���������"
            Top             =   180
            Width           =   3090
         End
         Begin VB.CommandButton cmdGet 
            Caption         =   "��ȡ(&G)"
            Height          =   350
            Left            =   11520
            TabIndex        =   3
            Top             =   575
            Width           =   1100
         End
         Begin VB.ComboBox cboStock 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   180
            Width           =   2610
         End
         Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
            Height          =   315
            Left            =   9120
            TabIndex        =   10
            Top             =   180
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   166658051
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   315
            Left            =   11040
            TabIndex        =   11
            Top             =   180
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   166658051
            CurrentDate     =   36263
         End
         Begin VB.Label lblInputTxt 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   2
            Left            =   5640
            TabIndex        =   20
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblInputTxt 
            AutoSize        =   -1  'True
            Caption         =   "��Ӧ��"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblInputTxt 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ"
            Height          =   180
            Index           =   0
            Left            =   3840
            TabIndex        =   18
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   10800
            TabIndex        =   17
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�������"
            Height          =   180
            Left            =   8280
            TabIndex        =   16
            Top             =   240
            Width           =   720
         End
         Begin VB.Label LblStock 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ⷿ"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblFlag 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   180
            Index           =   0
            Left            =   3360
            TabIndex        =   14
            Top             =   240
            Width           =   90
         End
         Begin VB.Label lblFlag 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   180
            Index           =   1
            Left            =   7700
            TabIndex        =   13
            Top             =   240
            Width           =   90
         End
         Begin VB.Label lblFlag 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   180
            Index           =   2
            Left            =   12660
            TabIndex        =   12
            Top             =   240
            Width           =   90
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBill 
         Height          =   3255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   4335
         _cx             =   7646
         _cy             =   5741
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlanModify.frx":038A
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
End
Attribute VB_Name = "frmHandBackPlanModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng�ⷿID As Long
Private mintUnit As Integer
Private mstrNo As String
Private mblnSuccess As Boolean
Private Const MStrCaption As String = "ҩƷ��ҩ�ƻ��༭"

Dim mlngFind As Long                            '���ڲ���
Dim mrsFindName As ADODB.Recordset              '���ڲ���

Private Enum InputType
    ҩƷ = 0
    ��Ӧ�� = 1
    ������ = 2
End Enum

'���ܣ���ϸ�б����
Private Const mconstBillHead = "ҩƷID,1,0|��Ӧ��ID,1,0|���,4,500|��Ӧ��,1,2500|ҩƷ����,1,1000|ҩƷ����,1,2000|��Ʒ��,1,2000|���,1,2000|������,1,2000|����,1,1000|Ч��,1,1000|��λ,1,800|��ҩ����,7,1000|�ɱ���,7,1000|�ɱ����,7,1000|��װ,7,0"

Private Enum ��ϸ�б�
    ҩƷid = 0
    ��Ӧ��id = 1
    ��� = 2
    ��Ӧ�� = 3
    ҩƷ���� = 4
    ҩƷ���� = 5
    ��Ʒ�� = 6
    ��� = 7
    ������ = 8
    ���� = 9
    Ч�� = 10
    ��λ = 11
    ���� = 12
    �ɱ��� = 13
    �ɱ���� = 14
    ��װ = 15
    ���� = 16
End Enum

Private Function CheckRepeat(ByVal strInfo As String, Optional ByVal intExceptCol As Integer = 0) As Boolean
    '����Ƿ��ظ�
    '������ҩƷID����Ӧ��ID�������̡�����
    'strInfo��ʽ��ҩƷID;��Ӧ��ID;������;����
    'intExceptCol���ų�����
    'CheckRepeat���أ�True-�ظ�;False-���ظ�
    
    Dim lngҩƷID As Long
    Dim lng��Ӧ��ID As Long
    Dim str������ As String
    Dim str���� As String
    Dim i As Integer
    
    If vsfBill.rows = 1 Then Exit Function
    If vsfBill.TextMatrix(1, ��ϸ�б�.ҩƷid) = "" Then Exit Function
    
    lngҩƷID = Split(strInfo, ";")(0)
    lng��Ӧ��ID = Split(strInfo, ";")(1)
    str������ = Split(strInfo, ";")(2)
    str���� = Split(strInfo, ";")(3)
    
    With vsfBill
        For i = 1 To .rows - 1
            If i <> intExceptCol And Val(.TextMatrix(i, ��ϸ�б�.ҩƷid)) = lngҩƷID And Val(.TextMatrix(i, ��ϸ�б�.��Ӧ��id)) = lng��Ӧ��ID _
                And .TextMatrix(i, ��ϸ�б�.������) = str������ And .TextMatrix(i, ��ϸ�б�.����) = str���� Then
                CheckRepeat = True
                Exit Function
            End If
        Next
    End With
End Function
Private Sub IniGrid()
    Dim i As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    
    strTemp = Split(mconstBillHead, "|")
    With vsfBill
        .Redraw = flexRDNone
        .rows = 1
        .Cols = ��ϸ�б�.����
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next

        .Redraw = flexRDDirect
    End With
End Sub
Private Sub GetDate(ByVal strNo As String)
    '��ȡ�Ѵ��ڵĵ�����ϸ
    Dim rsTmp As ADODB.Recordset
    Dim strSubUnit As String
    
    If strNo = "" Then Exit Sub
    On Error GoTo errHandle
    '��λ����װ����
    '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
    Select Case mintUnit
    Case 1
        strSubUnit = "D.���㵥λ ��λ,1 ��װ "
    Case 2
        strSubUnit = "B.���ﵥλ ��λ,B.�����װ ��װ "
    Case 3
        strSubUnit = "B.סԺ��λ ��λ,B.סԺ��װ ��װ "
    Case 4
        strSubUnit = "B.ҩ�ⵥλ ��λ,B.ҩ���װ ��װ "
    End Select
    
    gstrSQL = "Select Distinct A.���, A.ҩƷid, D.���� As ҩƷ����,D.���� As ͨ����,E.���� As ��Ʒ��, " & _
        " D.���, A.ʵ������, A.Ч��,A.�ɱ���, A.�ɱ����, A.���� As ������, A.����,A.��ҩ��λid,F.���� As ��Ӧ��, " & _
        " A.������, A.��������, A.ժҪ, " & strSubUnit & _
        " From ҩƷ��ҩ�ƻ� A, ҩƷ��� B, �շ���ĿĿ¼ D, �շ���Ŀ���� E, ��Ӧ�� F " & _
        " Where A.ҩƷid = B.ҩƷid And B.ҩƷid = D.ID And B.ҩƷid = E.�շ�ϸĿid(+) And E.����(+) = 3 And A.��ҩ��λid = F.ID And A.No = [1] " & _
        " Order By A.��� "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ�����Ϣ", strNo)
    
    vsfBill.rows = 1
    
    If rsTmp.EOF Then Exit Sub
    
    With rsTmp
        txtNo.Text = strNo
        txt������.Text = !������
        Txt��������.Text = Format(!��������, "yyyy-mm-dd hh:mm:ss")
        txtժҪ.Text = Nvl(!ժҪ)
        Do While Not .EOF
            vsfBill.rows = vsfBill.rows + 1
            
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.���) = .AbsolutePosition
            
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.ҩƷid) = !ҩƷid
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��Ӧ��id) = !��ҩ��λID
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.ҩƷ����) = !ҩƷ����
            If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.ҩƷ����) = !ͨ����
            Else
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.ҩƷ����) = IIf(IsNull(!��Ʒ��), !ͨ����, !��Ʒ��)
            End If
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��Ʒ��) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��Ӧ��) = Nvl(!��Ӧ��)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.���) = Nvl(!���)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��λ) = Nvl(!��λ)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.����) = zlStr.FormatEx(!ʵ������ / !��װ, 2, , True)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.�ɱ���) = zlStr.FormatEx(!�ɱ��� * !��װ, 5, , True)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.�ɱ����) = zlStr.FormatEx(!�ɱ����, 2, , True)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.������) = Nvl(!������)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.����) = Nvl(!����)
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��װ) = !��װ
            
            vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.Ч��) = Format(IIf(IsNull(!Ч��), "", !Ч��), "yyyy-mm-dd")
                    
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.Ч��) <> "" Then
                '����Ϊ��Ч��
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.Ч��) = Format(DateAdd("D", -1, vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.Ч��)), "yyyy-mm-dd")
            End If
            
            .MoveNext
        Loop
        
        vsfBill.Cell(flexcpForeColor, 1, ��ϸ�б�.����, vsfBill.rows - 1, ��ϸ�б�.����) = vbBlue
        vsfBill.Cell(flexcpFontBold, 1, ��ϸ�б�.����, vsfBill.rows - 1, ��ϸ�б�.����) = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetNewDate()
    '����������
    Dim lng�ⷿID As Long
    Dim lngҩƷID As Long
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim lng��Ӧ��ID As Long
    Dim str������ As String
    
    Dim rsTmp As ADODB.Recordset
    Dim strSubUnit As String
    Dim strSqlCondition As String
    
    On Error GoTo errHandle
    lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
    lngҩƷID = Val(txtInput(InputType.ҩƷ).Tag)
    str��ʼʱ�� = Format(dtp��ʼʱ��.Value, "YYYY-MM-DD") & " 00:00:01"
    str����ʱ�� = Format(dtp����ʱ��.Value, "YYYY-MM-DD") & " 23:59:59"
    If Val(txtInput(InputType.��Ӧ��).Tag) > 0 And Trim(txtInput(InputType.��Ӧ��).Text) <> "" Then
        lng��Ӧ��ID = Val(txtInput(InputType.��Ӧ��).Tag)
    End If
    str������ = Trim(txtInput(InputType.������).Text)
    
    If lng�ⷿID = 0 Or lngҩƷID = 0 Then Exit Sub
        
    strSqlCondition = " And A.�ⷿid + 0 = [1] And A.ҩƷid + 0 = [2] And A.������� Between [3] And [4] "
    
    If lng��Ӧ��ID > 0 Then
        strSqlCondition = strSqlCondition & " And A.��ҩ��λid = [5] "
    End If
    
    If str������ <> "" Then
        strSqlCondition = strSqlCondition & " And A.���� = [6] "
    End If
        
    '��λ����װ����
    '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
    Select Case mintUnit
    Case 1
        strSubUnit = "D.���㵥λ ��λ,1 ��װ "
    Case 2
        strSubUnit = "B.���ﵥλ ��λ,B.�����װ ��װ "
    Case 3
        strSubUnit = "B.סԺ��λ ��λ,B.סԺ��װ ��װ "
    Case 4
        strSubUnit = "B.ҩ�ⵥλ ��λ,B.ҩ���װ ��װ "
    End Select
    
    gstrSQL = "Select Distinct A.ҩƷid, D.���� As ҩƷ����,D.���� As ͨ����,E.���� As ��Ʒ��, " & _
        " D.���, A.Ч��, A.ʵ������, A.�ɱ���, A.�ɱ����, A.���� As ������, A.����,A.��ҩ��λid,F.���� As ��Ӧ��, " & strSubUnit & _
        " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ D, �շ���Ŀ���� E, ��Ӧ�� F " & _
        " Where A.ҩƷid = B.ҩƷid And B.ҩƷid = D.ID And B.ҩƷid = E.�շ�ϸĿid(+) And E.����(+) = 3 And A.��ҩ��λid = F.ID And A.���� = 1 " & strSqlCondition & _
        " Order By F.����"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ�����Ϣ", lng�ⷿID, lngҩƷID, CDate(str��ʼʱ��), CDate(str����ʱ��), lng��Ӧ��ID, str������)
    
    If rsTmp.EOF Then Exit Sub
    With rsTmp
        Do While Not .EOF
            '����Ƿ��ظ�
            If CheckRepeat(!ҩƷid & ";" & !��ҩ��λID & ";" & !������ & ";" & !����) = False Then
                vsfBill.rows = vsfBill.rows + 1
                
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.ҩƷid) = !ҩƷid
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��Ӧ��id) = !��ҩ��λID
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.���) = vsfBill.rows - 1
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��Ӧ��) = !��Ӧ��
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.ҩƷ����) = !ҩƷ����
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.ҩƷ����) = !ͨ����
                Else
                    vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.ҩƷ����) = IIf(IsNull(!��Ʒ��), !ͨ����, !��Ʒ��)
                End If
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��Ʒ��) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.���) = Nvl(!���)
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.������) = Nvl(!������)
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.����) = Nvl(!����)
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.Ч��) = Format(IIf(IsNull(!Ч��), "", !Ч��), "yyyy-mm-dd")
                    
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.Ч��) <> "" Then
                    '����Ϊ��Ч��
                    vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.Ч��) = Format(DateAdd("D", -1, vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.Ч��)), "yyyy-mm-dd")
                End If
                
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��λ) = Nvl(!��λ)
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.����) = ""
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.�ɱ���) = zlStr.FormatEx(!�ɱ��� * !��װ, 5, , True)
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.�ɱ����) = ""
                
                
                vsfBill.TextMatrix(vsfBill.rows - 1, ��ϸ�б�.��װ) = !��װ
            End If
            
            .MoveNext
        Loop
        vsfBill.Cell(flexcpForeColor, 1, ��ϸ�б�.����, vsfBill.rows - 1, ��ϸ�б�.����) = vbBlue
        vsfBill.Cell(flexcpFontBold, 1, ��ϸ�б�.����, vsfBill.rows - 1, ��ϸ�б�.����) = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadStock()
    'ȡ�ⷿ��ֻȡҩ�����ԵĿⷿ
    
    Dim rsTmp As ADODB.Recordset
    Dim lngDrugStoreIndex As Long
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct A.ID, A.���� " & _
              "From ��������˵�� C, �������ʷ��� B, ���ű� A " & _
              "Where (A.վ�� = [1] Or A.վ�� is Null) And C.�������� = B.���� And Instr('HIJ', B.����, 1) > 0 " & _
              "  And A.ID = C.����id And To_Char(A.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " & _
              "Order By A.���� "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ����ҩ�����ԵĿⷿ", gstrNodeNo)
    
    If rsTmp.EOF Then
        MsgBox "����Ӧ������һ������ҩ�����ʵĲ���,��鿴���Ź���", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    With rsTmp
        cboStock.Clear
        
        Do While Not .EOF
            cboStock.AddItem !����
            cboStock.ItemData(cboStock.NewIndex) = !id
            If !id = mlng�ⷿID Then
                lngDrugStoreIndex = intIndex
            End If
            intIndex = intIndex + 1
            .MoveNext
        Loop
        
        cboStock.ListIndex = lngDrugStoreIndex
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshSerialNumber()
    '���µ�����ţ�������ɾ���к�ʹ��
    
    Dim i As Integer
        
    With vsfBill
        If .rows = 2 Then Exit Sub
        For i = 1 To .rows - 1
            .TextMatrix(i, ��ϸ�б�.���) = i
        Next
    End With
End Sub

Private Function SelectInput(ByVal intType As Integer, ByVal strkey As String, ByVal sngX As Single, ByVal sngY As Single, ByVal sngH As Single) As String
    'ѡ������֧�ֶ�ҩƷ����Ӧ�̡������̵�ѡ��
    'intType��0-ҩƷ;1-��Ӧ��;2-������
    'strKey����-ȫ��;�ǿ�-ģ��ƥ��
    'SelectInput����ֵ����-û�ҵ�ƥ���¼;
    '                 �ǿ�-ҩƷ��ҩƷID;ҩƷ����;���;��λ;��װ��
    '                     -��Ӧ�̣���Ӧ��ID;��Ӧ�����ƣ�
    '                     -�����̣�������ID;���������ƣ�
    
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strSubUnit As String
    Dim strFindString As String
    Dim strReturn As String
    Dim strSqlҩƷ As String
    
    Err = 0: On Error GoTo ErrHand:
    
    strkey = UCase(Trim(strkey))
    
    Select Case intType
    Case InputType.ҩƷ
        '��λ����װ����
        '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
        Select Case mintUnit
        Case 1
            strSubUnit = "B.���㵥λ ��λ,1 ��װ "
        Case 2
            strSubUnit = "A.���ﵥλ ��λ,A.�����װ ��װ "
        Case 3
            strSubUnit = "A.סԺ��λ ��λ,A.סԺ��װ ��װ "
        Case 4
            strSubUnit = "A.ҩ�ⵥλ ��λ,A.ҩ���װ ��װ "
        End Select
        
        If strkey <> "" Then
            strFindString = " And (B.���� Like [1] OR B.���� Like [2] OR C.���� LIKE [2])"
            
            If IsNumeric(strkey) Then                         '10,11.����ȫ������ʱֻƥ�����'����ҩƷ,��Ҫƥ�����(����Ϊ3��������)
                If Mid(gtype_UserSysParms.P44_����ƥ��, 1, 1) = "1" Then strFindString = " And (B.���� Like [1] Or B.���� Like [2] And C.����=3)"
            ElseIf zlStr.IsCharAlpha(strkey) Then         '01,11.����ȫ����ĸʱֻƥ�����
                If Mid(gtype_UserSysParms.P44_����ƥ��, 2, 1) = "1" Then strFindString = " And C.���� Like [2] "
            ElseIf zlStr.IsCharChinese(strkey) Then
                strFindString = " And B.���� Like [2] "
            End If
        End If
        
        If strkey = "" Then
            If gintҩƷ������ʾ = 0 Then
                strSqlҩƷ = ",'['||����||']'|| ͨ���� As ҩƷ����"
            ElseIf gintҩƷ������ʾ = 1 Then
                strSqlҩƷ = ",'['||����||']'|| Nvl(��Ʒ��,ͨ����) As ҩƷ����"
            ElseIf gintҩƷ������ʾ = 2 Then
                strSqlҩƷ = ",'['||����||']'|| ͨ���� As ҩƷ����,��Ʒ��"
            End If
            
            gstrSQL = "Select Rownum As ID, ҩƷid " & strSqlҩƷ & ", ���, ���� as ������, ��λ, ��װ,��Ʒ�� " & _
                " From (Select Distinct A.ҩƷid, B.����, B.���� As ͨ����, C.���� As ��Ʒ��, B.���,B.����,  " & strSubUnit & _
                " From ҩƷ��� A, " & _
                " (Select B.ID, B.����, B.����, B.���,B.����,B.���㵥λ From �շ���ĿĿ¼ B, �շ���Ŀ���� C " & _
                " Where (B.վ�� = [3] Or B.վ�� is Null) And B.ID = C.�շ�ϸĿid And B.��� In ('5', '6', '7') " & strFindString & ") B, �շ���Ŀ���� C " & _
                " Where A.ҩƷid = B.ID And A.ҩƷid = C.�շ�ϸĿid(+) And C.����(+) = 3 "
            gstrSQL = gstrSQL & " Order By B.����)"
        Else
            strSqlҩƷ = ",'['||����||']'|| �������� As ҩƷ����"
            
            gstrSQL = "Select Rownum As ID, ҩƷid " & strSqlҩƷ & ", ���, ���� as ������, ��λ, ��װ,��Ʒ�� " & _
                " From (Select Distinct A.ҩƷid, B.����, B.��������, B.���� As ͨ����, C.���� As ��Ʒ��, B.���,B.����,  " & strSubUnit & _
                " From ҩƷ��� A, " & _
                " (Select B.ID, B.����, B.����, B.���,B.����,B.���㵥λ, C.���� As �������� From �շ���ĿĿ¼ B, �շ���Ŀ���� C " & _
                " Where (B.վ�� = [3] Or B.վ�� is Null) And B.ID = C.�շ�ϸĿid And B.��� In ('5', '6', '7') " & strFindString & ") B, �շ���Ŀ���� C " & _
                " Where A.ҩƷid = B.ID And A.ҩƷid = C.�շ�ϸĿid(+) And C.����(+) = 3 "
            gstrSQL = gstrSQL & " Order By B.����)"
        End If
        
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷѡ����", False, "", "ѡ��ҩƷ", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            strReturn = ""
        Else
            strReturn = rsTemp!ҩƷid & ";" & rsTemp!ҩƷ���� & ";" & rsTemp!��Ʒ�� & ";" & rsTemp!��� & ";" & rsTemp!��λ & ";" & rsTemp!��װ
        End If
    Case InputType.��Ӧ��
        gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) " & _
                  "  And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null ) And ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
                  "  And (���� like [1] or ���� like [2] or ���� like [2])"
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "��Ӧ��ѡ����", False, "", "ѡ��Ӧ��", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            strReturn = ""
        Else
            strReturn = rsTemp!id & ";" & rsTemp!����
        End If
    Case InputType.������
        gstrSQL = "Select Rownum As ID,����,����,���� From ҩƷ������ " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (���� like [1] Or ���� like [2] Or ���� like [2]) Order By ����"
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "������ѡ����", False, "", "ѡ��������", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
            strReturn = ""
        Else
            strReturn = rsTemp!id & ";" & rsTemp!����
        End If
    End Select
    
    SelectInput = strReturn
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub FindGridRow(ByVal strInput As String)
    Dim lngStart As Long
    Dim str���� As String, str���� As String, str���� As String
    Dim str�������� As String
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim strҩ�� As String
    
    '����ҩƷ
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '��ʾ�µĲ���
        If vsfBill.rows > 1 Then vsfBill.Row = 1
        txtFind.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.��� In ('5','6','7') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] )" & _
                  "Order By '[' || A.���� || ']' "
        Set mrsFindName = zlDataBase.OpenSQLRecord(gstrSQL, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If
    
    
    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    If mrsFindName.EOF Then mrsFindName.MoveFirst
    Do While Not mrsFindName.EOF
        strҩ�� = mrsFindName!ҩƷ����
        strҩ�� = Mid(strҩ��, 2, Len(strҩ��) - 2)
        lngFindRow = vsfBill.FindRow(strҩ��, lngStart, ��ϸ�б�.ҩƷ����, True, True)
        If lngFindRow > 0 Then
            vsfBill.Select lngFindRow, 1, lngFindRow, vsfBill.Cols - 1
            vsfBill.TopRow = lngFindRow
            mlngFind = lngFindRow
            mrsFindName.MoveNext
            If lngStart >= vsfBill.rows - 1 Then
                lngStart = 1
            Else
                lngStart = lngStart + 1
            End If
            Exit Do
        End If
        mrsFindName.MoveNext
        If mrsFindName.EOF Then
            MsgBox "���������ף��������ҽ��Ӷ���ʼ��", vbInformation, gstrSysName
        End If
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub ShowForm(FrmMain As Form, ByVal lng�ⷿID As Long, ByVal intUnit As Integer, ByRef BlnSuccess As Boolean, Optional ByVal strNo As String = "")
    mlng�ⷿID = lng�ⷿID
    mintUnit = intUnit
    mstrNo = strNo
    mblnSuccess = False
    
    Me.Show vbModal, FrmMain
    
    BlnSuccess = mblnSuccess
End Sub

Private Function GetSeleterReturn(ByVal intType As Integer, ByVal objInputObj As Object, ByVal strInput As String) As Boolean
    'ͨ��������Ŀ����ѡ������ѡ��ֵ
    'intType��0-ҩƷ;1-��Ӧ��;2-������
    'objInputObj���������͵�¼�����TextBox��VSFlexGrid
    'strInput������ֵ�������Ǳ��롢���롢����
    Dim vRect As RECT
    Dim strReturn As String
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    
    '���ݶ��������ȡ����Ĳ���λ��
    If TypeName(objInputObj) = "TextBox" Then
        vRect = zlControl.GetControlRect(objInputObj.hWnd)
        sngX = vRect.Left
        sngY = vRect.Top
        sngH = objInputObj.Height
    ElseIf TypeName(objInputObj) = "VSFlexGrid" Then
        Call CalcPosition(sngX, sngY, objInputObj)
        sngY = sngY - objInputObj.CellHeight
        sngH = objInputObj.CellHeight
    Else
        Exit Function
    End If
    
    '�õ�ѡ�����ķ���ֵ
    strReturn = SelectInput(intType, strInput, sngX, sngY, sngH)
    
    '����ʵ��ҵ����
    If TypeName(objInputObj) = "TextBox" Then
'        strReturn="ID;����"
        If strReturn = "" Then Exit Function
            
        objInputObj.Tag = Val(Split(strReturn, ";")(0))
        objInputObj.Text = Split(strReturn, ";")(1)
    Else
        If strReturn = "" Then
            Select Case intType
            Case InputType.ҩƷ
                If Val(objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.ҩƷid)) > 0 Then
                    Exit Function
                End If
            Case InputType.��Ӧ��
                If Val(objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.��Ӧ��id)) > 0 Then
                    Exit Function
                End If
            Case InputType.������
                If Trim(objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.������)) <> "" Then
                    Exit Function
                End If
            End Select
            
            objInputObj.TextMatrix(objInputObj.Row, objInputObj.Col) = objInputObj.EditText
            objInputObj.Cell(flexcpText, objInputObj.Row, objInputObj.Col) = ""
            Exit Function
        Else
            Select Case intType
            Case InputType.ҩƷ
        '        strReturn="ҩƷID;ҩƷ����;���;��λ;��װ"
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.ҩƷid) = Val(Split(strReturn, ";")(0))
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.ҩƷ����) = Split(strReturn, ";")(1)
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.��Ʒ��) = Split(strReturn, ";")(2)
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.���) = Split(strReturn, ";")(3)
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.��λ) = Split(strReturn, ";")(4)
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.��װ) = Val(Split(strReturn, ";")(5))
                
                objInputObj.EditText = Split(strReturn, ";")(1)
            Case InputType.��Ӧ��
        '        strReturn="��Ӧ��ID;��Ӧ������"
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.��Ӧ��id) = Val(Split(strReturn, ";")(0))
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.��Ӧ��) = Split(strReturn, ";")(1)
                
                objInputObj.EditText = Split(strReturn, ";")(1)
            Case InputType.������
        '        strReturn="������ID;����������"
                objInputObj.TextMatrix(objInputObj.Row, ��ϸ�б�.������) = Split(strReturn, ";")(1)
                
                objInputObj.EditText = Split(strReturn, ";")(1)
            End Select
        End If
    End If
    
    GetSeleterReturn = True
End Function

Private Sub cboStock_Click()
    Call SetSelectorRS(1, "ҩƷ�⹺������", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    '���������
    With vsfBill
        .rows = 1
    End With
End Sub

Private Sub cmdFind_Click()
    Dim blnVisible As Boolean
    '���һ������һ��
    blnVisible = lblFindType.Visible Xor True
    lblFindType.Visible = blnVisible
    txtFind.Visible = blnVisible
    lblComment1.Visible = blnVisible
    If blnVisible Then txtFind.SetFocus
End Sub

Private Sub cmdGet_Click()
    Call GetNewDate
End Sub
Private Sub cmdReset_Click()
    '�����ҩ�����ͳɱ����
    With vsfBill
        If .rows = 1 Then Exit Sub
        If .TextMatrix(1, 0) = "" Then Exit Sub
        .Cell(flexcpText, 1, ��ϸ�б�.����, .rows - 1, ��ϸ�б�.����) = ""
        .Cell(flexcpText, 1, ��ϸ�б�.�ɱ����, .rows - 1, ��ϸ�б�.�ɱ����) = ""
    End With
End Sub

Private Sub CmdSave_Click()
    Dim strNo_In As String
    Dim int���_In As Integer
    Dim lngҩƷid_In As Long
    Dim lng��ҩ��λid_In As Long
    Dim dbl����_In As Double
    Dim dbl�ɱ���_In As Double
    Dim dbl�ɱ����_In As Double
    Dim str����_In As String
    Dim str����_In As String
    Dim str������_In As String
    Dim str��������_In As String
    Dim strժҪ_In As String
    Dim strЧ�� As String
    
    Dim blnTrans As Boolean
    Dim i As Integer
    Dim intCount As Integer
    Dim arrSql As Variant
    
    On Error GoTo errHandle
    
    arrSql = Array()
    '��������е��ݣ���ɾ��
    strNo_In = Trim(txtNo.Text)
    If strNo_In <> "" Then
'        gcnOracle.BeginTrans
        blnTrans = True
    
        gstrSQL = "Zl_ҩƷ��ҩ�ƻ�_Delete('" & strNo_In & "')"
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    End If
    
    '�����µ���
    With vsfBill
        '������NO��
        If strNo_In = "" Then
            strNo_In = Sys.GetNextNo(100)
        End If
        str������_In = txt������.Text
        str��������_In = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        strժҪ_In = txtժҪ.Text
        
        For i = 1 To .rows - 1
            If Val(.TextMatrix(i, ��ϸ�б�.����)) > 0 Then
                int���_In = intCount + 1
                lngҩƷid_In = Val(.TextMatrix(i, ��ϸ�б�.ҩƷid))
                lng��ҩ��λid_In = Val(.TextMatrix(i, ��ϸ�б�.��Ӧ��id))
                dbl����_In = zlStr.FormatEx(Val(.TextMatrix(i, ��ϸ�б�.����)) * Val(.TextMatrix(i, ��ϸ�б�.��װ)), 5, , True)
                dbl�ɱ���_In = zlStr.FormatEx(Val(.TextMatrix(i, ��ϸ�б�.�ɱ���)) / Val(.TextMatrix(i, ��ϸ�б�.��װ)), 5, , True)
                dbl�ɱ����_In = Val(.TextMatrix(i, ��ϸ�б�.�ɱ����))
                str����_In = IIf(Trim(.TextMatrix(i, ��ϸ�б�.������)) = "", "", .TextMatrix(i, ��ϸ�б�.������))
                str����_In = IIf(Trim(.TextMatrix(i, ��ϸ�б�.����)) = "", "", .TextMatrix(i, ��ϸ�б�.����))
                
                strЧ�� = IIf(Trim(.TextMatrix(i, ��ϸ�б�.Ч��)) = "", "", .TextMatrix(i, ��ϸ�б�.Ч��))
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And strЧ�� <> "" Then
                    '����ΪʧЧ��������
                    strЧ�� = Format(DateAdd("D", 1, strЧ��), "yyyy-mm-dd")
                End If
                
                gstrSQL = "Zl_ҩƷ��ҩ�ƻ�_Insert("
                'NO
                gstrSQL = gstrSQL & "'" & strNo_In & "'"
                '���
                gstrSQL = gstrSQL & "," & int���_In
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngҩƷid_In
                '��ҩ��λID
                gstrSQL = gstrSQL & "," & lng��ҩ��λid_In
                '��ҩ����
                gstrSQL = gstrSQL & "," & dbl����_In
                '�ɱ���
                gstrSQL = gstrSQL & "," & dbl�ɱ���_In
                '�ɱ����
                gstrSQL = gstrSQL & "," & dbl�ɱ����_In
                '����
                gstrSQL = gstrSQL & ",'" & str����_In & "'"
                '����
                gstrSQL = gstrSQL & ",'" & str����_In & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '������
                gstrSQL = gstrSQL & ",'" & str������_In & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & str��������_In & "','yyyy-mm-dd HH24:MI:SS')"
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strժҪ_In & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                
                intCount = intCount + 1
            End If
        Next
    End With
    
    If intCount = 0 Then
        MsgBox "��¼����ҩ������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
    Next
    gcnOracle.CommitTrans
    mblnSuccess = True
    
    Unload Me
    Exit Sub
errHandle:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdSelecter_Click(Index As Integer)
    Dim RecReturn As ADODB.Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "ҩƷ�⹺������", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    
    If Index = InputType.ҩƷ Then
'        Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, 0, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
        
        Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), 0, True, True, True, 0, False)

        If RecReturn.RecordCount = 0 Then
            Call zlControl.TxtSelAll(txtInput(Index))
            Exit Sub
        End If
            
        If gintҩƷ������ʾ = 1 Then
            txtInput(Index).Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
        Else
            txtInput(Index).Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
        End If
        txtInput(Index).Tag = RecReturn!ҩƷid
    Else
        If GetSeleterReturn(Index, txtInput(Index), "") = False Then
            Call zlControl.TxtSelAll(txtInput(Index))
        End If
    End If
End Sub


Private Sub Form_Load()
    Dim dateCurr As Date
    
    Set Me.Icon = Nothing
    dateCurr = Sys.Currentdate
    dtp��ʼʱ��.Value = CDate(Format(dateCurr, "YYYY-MM") & "-01 00:00:00")
    dtp����ʱ��.Value = dateCurr
    
    Call LoadStock
    Call IniGrid
    
    If mstrNo <> "" Then
        Call GetDate(mstrNo)
    Else
        txt������.Text = UserInfo.�û�����
        Txt��������.Text = Format(dateCurr, "YYYY-MM-DD HH:MM:SS")
    End If
End Sub

Private Sub Form_Resize()
    If Me.Width < 13365 Then Me.Width = 13365
    If Me.Height < 8715 Then Me.Height = 8715
    
    With picBill
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - fraControl.Height - 100
    End With
    
    With fraControl
        .Top = Me.ScaleHeight - .Height - 100
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = 615
    End With
    
    With fraCondition
        .Top = 50
        .Left = 100
        .Width = picBill.Width - 100
    End With
    
    With vsfBill
        .Top = fraCondition.Top + fraCondition.Height + 100
        .Left = fraCondition.Left
        .Width = fraCondition.Width
        .Height = picBill.Height - fraCondition.Top - fraCondition.Height - fraComment.Height - 100
    End With
    
    With fraComment
        .Top = picBill.Height - .Height
        .Left = 0
        .Width = picBill.Width
    End With
    
    With txtժҪ
        .Width = fraComment.Width - .Top - 200
    End With
    
    RestoreWinState Me, App.ProductName, MStrCaption
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
    Call ReleaseSelectorRS
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    strInput = Trim(UCase(txtFind.Text))
    If strInput = "" Then Exit Sub
    
    Call FindGridRow(strInput)
End Sub


Private Sub txtInput_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtInput(Index)
End Sub
Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtInput(Index).Text) = "" Then Exit Sub
    
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If Index = InputType.ҩƷ Then
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(txtInput(Index).Text) = "" Then Exit Sub
        sngLeft = Me.Left + fraCondition.Left + txtInput(Index).Left
        sngTop = Me.Top + fraCondition.Top + txtInput(Index).Top + txtInput(Index).Height + Me.Height - Me.ScaleHeight '  50
        If sngTop + 3630 > Screen.Height Then
            sngTop = sngTop - txtInput(Index).Height - 3630
        End If
        
        strkey = Trim(txtInput(Index).Text)
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(1, "ҩƷ�⹺������", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
        End If
        
'        Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), strkey, sngLeft, sngTop)
        Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), 0, True, True, True, 0, False)
        
        If RecReturn.RecordCount = 0 Then
            Call zlControl.TxtSelAll(txtInput(Index))
            Exit Sub
        End If
        
        If gintҩƷ������ʾ = 1 Then
            txtInput(Index).Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
        Else
            txtInput(Index).Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
        End If
        txtInput(Index).Tag = RecReturn!ҩƷid
    Else
        If GetSeleterReturn(Index, txtInput(Index), Trim(txtInput(Index).Text)) = False Then
            Call zlControl.TxtSelAll(txtInput(Index))
        End If
    End If
End Sub


Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr(" ';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    With vsfBill
'        Select Case Col
'            Case ��ϸ�б�.ҩƷ����, ��ϸ�б�.��Ӧ��, ��ϸ�б�.������
'                .ColComboList(Col) = "..."
'        End Select
'    End With
End Sub

Private Sub vsfBill_AfterSort(ByVal Col As Long, Order As Integer)
    Call RefreshSerialNumber
End Sub

Private Sub vsfBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    With vsfBill
'        Select Case Col
'            Case ��ϸ�б�.ҩƷ����
'                Call GetSeleterReturn(InputType.ҩƷ, vsfBill, "")
'            Case ��ϸ�б�.��Ӧ��
'                Call GetSeleterReturn(InputType.��Ӧ��, vsfBill, "")
'            Case ��ϸ�б�.������
'                Call GetSeleterReturn(InputType.������, vsfBill, "")
'        End Select
'    End With
End Sub
Private Sub vsfBill_EnterCell()
    With vsfBill
        .Editable = flexEDNone
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, ��ϸ�б�.ҩƷid) = "" Then Exit Sub
        
        Select Case .Col
'        Case ��ϸ�б�.ҩƷ����, ��ϸ�б�.��Ӧ��, ��ϸ�б�.������
'            .ColComboList(.Col) = "..."
'            .Editable = flexEDKbdMouse
        Case ��ϸ�б�.���� ', ��ϸ�б�.�ɱ���, ��ϸ�б�.����
            .Editable = flexEDKbdMouse
        End Select
    End With
End Sub

Private Sub vsfBill_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBill
        If KeyCode = vbKeyDelete Then
            If .Row < 1 Then Exit Sub
            If .TextMatrix(.Row, ��ϸ�б�.ҩƷid) = "" Then Exit Sub
            
            If MsgBox("�Ƿ�ɾ����" & .Row & "�е���ҩ��¼��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                .RemoveItem .Row
                Call RefreshSerialNumber
            End If
        End If
        
'        Select Case .Col
'        Case ��ϸ�б�.ҩƷ����, ��ϸ�б�.��Ӧ��, ��ϸ�б�.������
'            If KeyCode <> vbKeyReturn Then
'                .ColComboList(.Col) = ""
'            End If
'        End Select
        
        If txtFind.Visible And KeyCode = vbKeyF3 Then
            Call txtFind_KeyDown(vbKeyReturn, 0)
        End If
    End With
End Sub


Private Sub vsfBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsfBill
        If Trim(.EditText) = "" Then Exit Sub

'        Select Case Col
'            Case ��ϸ�б�.ҩƷ����
'                Call GetSeleterReturn(InputType.ҩƷ, vsfBill, Trim(.EditText))
'            Case ��ϸ�б�.��Ӧ��
'                Call GetSeleterReturn(InputType.��Ӧ��, vsfBill, Trim(.EditText))
'            Case ��ϸ�б�.������
'                Call GetSeleterReturn(InputType.������, vsfBill, Trim(.EditText))
'        End Select
    End With
End Sub

Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" ';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    'ֻ����������
    If Col = ��ϸ�б�.���� Then
        If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    
'    'ֻ���������֣�С����
'    If Col = ��ϸ�б�.�ɱ��� Then
'        If InStr(".1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
'            KeyAscii = 0
'        End If
'    End If
End Sub
Private Sub vsfBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfBill
        Select Case Col
        Case ��ϸ�б�.����
            .TextMatrix(Row, ��ϸ�б�.����) = Val(.EditText)
            .TextMatrix(Row, ��ϸ�б�.�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(Row, ��ϸ�б�.����)) * Val(.TextMatrix(Row, ��ϸ�б�.�ɱ���)), 2, , True)
'        Case ��ϸ�б�.�ɱ���
'            .EditText = zlStr.FormatEx(Val(.EditText), 5)
'            .TextMatrix(Row, ��ϸ�б�.�ɱ���) = .EditText
'            .TextMatrix(Row, ��ϸ�б�.�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(Row, ��ϸ�б�.����)) * Val(.TextMatrix(Row, ��ϸ�б�.�ɱ���)), 2)
        End Select
    End With
End Sub


