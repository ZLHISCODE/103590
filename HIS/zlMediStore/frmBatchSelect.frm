VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmBatchSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ����ѡ��"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12045
   DrawStyle       =   4  'Dash-Dot-Dot
   Icon            =   "frmBatchSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   15644.17
   ScaleMode       =   0  'User
   ScaleWidth      =   14442.45
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picInit 
      Height          =   6975
      Left            =   0
      ScaleHeight     =   6915
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   600
      Width           =   12015
      Begin VB.TextBox txtClass 
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   3495
      End
      Begin VB.TextBox txtPingZhong 
         Height          =   300
         Left            =   6000
         TabIndex        =   3
         Top             =   120
         Width           =   4020
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   2
         Top             =   6600
         Width           =   1335
      End
      Begin VB.CommandButton cmdClass 
         Caption         =   "��"
         Height          =   300
         Left            =   4440
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSelectDrug 
         Height          =   6045
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   11895
         _cx             =   20981
         _cy             =   10663
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
      Begin VB.Label lblPingZhong 
         AutoSize        =   -1  'True
         Caption         =   "Ʒ�ּ���"
         Height          =   180
         Left            =   5160
         TabIndex        =   8
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   6660
         Width           =   360
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   4680
      Top             =   6600
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
            Key             =   "pingzhong"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":6E7D
            Key             =   "���U"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgList 
      Bindings        =   "frmBatchSelect.frx":7417
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmBatchSelect.frx":742B
   End
End
Attribute VB_Name = "frmBatchSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintUnit As Integer '��ģ�������õ���ʾ��λ 0-ҩ�ⵥλ;1-���ﵥλ;2-סԺ��λ;3-�ۼ۵�λ
Private Const mlngRowHeight As Long = 300 '����и����и�
Private mrsReturn As ADODB.Recordset        '����ѡ��ҩƷ����
Private mblnOK As Boolean   '��¼�Ƿ��ǵ����ȷ����ť
Private mrsFindName As ADODB.Recordset '��¼��ѯ���ݼ�
Private mstrMatch  As String '0-˫��ƥ�� 1-������ƥ��
Private mint����ģʽ  As Integer '0-���۽��룬1-�����������۽���


'����λ
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
Private Const MStrCaption As String = "ҩƷ����ѡ��"

'���ܰ�ť
Private Const mconMenu_Save = 100 '���
Private Const mconMenu_Quit = 101 'ȡ��
Private Const mconMenu_ClearAll = 102 '����б�
Private Const mconMenu_Find = 103 '����

Private Enum vsfSelectDrugCol
    ҩƷid = 0
    ҩƷ��Ϣ = 1
    ҩƷ����
    ��Ʒ��
    ͨ����
    ���
    ����
    ��λ
    �ۼ۵�λ
    ���ﵥλ
    ����ϵ��
    סԺ��λ
    סԺϵ��
    ҩ�ⵥλ
    ҩ��ϵ��
    ����
    �ۼ�
    �ɱ���
    ָ������
    ָ���ۼ�
    ������
End Enum

Public Sub ShowME(ByVal frmParent As Form, ByRef rsTemp As ADODB.Recordset, ByRef blnOK As Boolean, Optional int����ģʽ As Integer = 0)
    mint����ģʽ = int����ģʽ
    Me.Show vbModal, frmParent
    blnOK = mblnOK
    Set rsTemp = mrsReturn
End Sub

Private Sub initVsflexgrid()
    With vsfSelectDrug
        .Editable = flexEDNone
        .Cols = vsfSelectDrugCol.������
        .rows = 1
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMove '�ƶ�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��

        '�����п�
        .ColWidth(vsfSelectDrugCol.ҩƷid) = 0
        .ColWidth(vsfSelectDrugCol.ҩƷ��Ϣ) = 3000
        .ColWidth(vsfSelectDrugCol.ҩƷ����) = 0
        .ColWidth(vsfSelectDrugCol.��Ʒ��) = 0
        .ColWidth(vsfSelectDrugCol.ͨ����) = 0
        .ColWidth(vsfSelectDrugCol.����) = 1500
        .ColWidth(vsfSelectDrugCol.��λ) = 800

        .ColWidth(vsfSelectDrugCol.�ۼ۵�λ) = 0
        .ColWidth(vsfSelectDrugCol.���ﵥλ) = 0
        .ColWidth(vsfSelectDrugCol.����ϵ��) = 0
        .ColWidth(vsfSelectDrugCol.סԺ��λ) = 0
        .ColWidth(vsfSelectDrugCol.סԺϵ��) = 0
        .ColWidth(vsfSelectDrugCol.ҩ�ⵥλ) = 0
        .ColWidth(vsfSelectDrugCol.ҩ��ϵ��) = 0

        .ColWidth(vsfSelectDrugCol.����) = 1000
        .ColWidth(vsfSelectDrugCol.�ۼ�) = 1500
        .ColWidth(vsfSelectDrugCol.�ɱ���) = 1500
        .ColWidth(vsfSelectDrugCol.ָ������) = 1500
        .ColWidth(vsfSelectDrugCol.ָ���ۼ�) = 1500
        '������ͷ
        .TextMatrix(0, vsfSelectDrugCol.ҩƷid) = "ҩƷid"
        .TextMatrix(0, vsfSelectDrugCol.ҩƷ��Ϣ) = "ҩƷ"
        .TextMatrix(0, vsfSelectDrugCol.ҩƷ����) = "ҩƷ����"
        .TextMatrix(0, vsfSelectDrugCol.��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, vsfSelectDrugCol.ͨ����) = "ͨ����"
        .TextMatrix(0, vsfSelectDrugCol.���) = "���"
        .TextMatrix(0, vsfSelectDrugCol.����) = "������"
        .TextMatrix(0, vsfSelectDrugCol.��λ) = "��λ"

        .TextMatrix(0, vsfSelectDrugCol.�ۼ۵�λ) = "�ۼ۵�λ"
        .TextMatrix(0, vsfSelectDrugCol.���ﵥλ) = "���ﵥλ"
        .TextMatrix(0, vsfSelectDrugCol.����ϵ��) = "����ϵ��"
        .TextMatrix(0, vsfSelectDrugCol.סԺ��λ) = "סԺ��λ"
        .TextMatrix(0, vsfSelectDrugCol.סԺϵ��) = "סԺϵ��"
        .TextMatrix(0, vsfSelectDrugCol.ҩ�ⵥλ) = "ҩ�ⵥλ"
        .TextMatrix(0, vsfSelectDrugCol.ҩ��ϵ��) = "ҩ��ϵ��"

        .TextMatrix(0, vsfSelectDrugCol.����) = "����"
        .TextMatrix(0, vsfSelectDrugCol.�ۼ�) = "�ۼ�"
        .TextMatrix(0, vsfSelectDrugCol.�ɱ���) = "�ɱ���"
        .TextMatrix(0, vsfSelectDrugCol.ָ������) = "ָ������"
        .TextMatrix(0, vsfSelectDrugCol.ָ���ۼ�) = "ָ���ۼ�"

        .ColAlignment(vsfSelectDrugCol.ҩƷid) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.ҩƷ��Ϣ) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.ҩƷ����) = flexAlignLeftCenter
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

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case mconMenu_Save  '���
            Call Save
        Case mconMenu_ClearAll  '���
            Call ClearAll
        Case mconMenu_Find '����
            txtFind.SetFocus
            If Trim(txtFind.Text) = "" Then Exit Sub
            Call FindGridRow(UCase(Trim(txtFind.Text)))
        Case mconMenu_Quit  'ȡ��
            Call Quit
    End Select
End Sub

Private Sub ClearAll()
    With vsfSelectDrug
        If MsgBox("ȷ��Ҫ��������Ѿ�ѡ���ҩƷ��", vbYesNo, gstrSysName) = vbYes Then
            .rows = 1
        End If
    End With
End Sub

Private Sub cmdClass_Click()
    Dim rsProvider As Recordset
    Dim strsql����ģʽ As String
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txtClass.hWnd)
    On Error GoTo ErrHandle
    
    If mint����ģʽ = 1 Then
        strsql����ģʽ = " a.�Ƿ����۹���=1 And "
    Else
        strsql����ģʽ = ""
    End If
    
    '����
    gstrSQL = "Select Level, ID, �ϼ�id, ����, ����, ����" & vbNewLine & _
                    "From (Select -1 As ID, Null As �ϼ�id, '001' As ����, '����ҩ' As ����, '����ҩ' As ����" & vbNewLine & _
                    "       From Dual" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select -2 As ID, Null As �ϼ�id, '002' As ����, '�г�ҩ' As ����, '�г�ҩ' As ����" & vbNewLine & _
                    "       From Dual" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select -3 As ID, Null As �ϼ�id, '003' As ����, '�в�ҩ' As ����, '�в�ҩ' As ����" & vbNewLine & _
                    "       From Dual" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select ID, Nvl(�ϼ�id, Decode(����, 1, -1, Decode(����, 2, -2, -3))) As �ϼ�id, ����, ����," & vbNewLine & _
                    "              Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����" & vbNewLine & _
                    "       From ���Ʒ���Ŀ¼" & vbNewLine & _
                    "       Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01')" & vbNewLine & _
                    "Start With �ϼ�id Is Null" & vbNewLine & _
                    "Connect By Prior ID = �ϼ�id" & vbNewLine & _
                    "Order By Level"

    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "����", False, "", "����ѡ��", False, False, _
    True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider Is Nothing Then
        Exit Sub
    End If
    
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select Distinct a.ҩƷid, c.���� As ҩƷ����, c.���� As ͨ����, d.��Ʒ��, c.���, c.�Ƿ��� As ʱ��, c.����, c.���㵥λ As �ۼ۵�λ, a.���ﵥλ, a.�����װ," & vbNewLine & _
                "                a.סԺ��λ, a.סԺ��װ, a.ҩ�ⵥλ, a.ҩ���װ, a.�ɱ���, e.�ּ�, a.ָ��������, a.ָ�����ۼ�" & vbNewLine & _
                "From ҩƷ��� A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) D, �շѼ�Ŀ E" & vbNewLine & _
                "Where a.ҩ��id = b.Id And a.ҩƷid = c.Id And c.Id = d.�շ�ϸĿid(+) And a.ҩƷid = e.�շ�ϸĿid And Sysdate Between e.ִ������ And" & vbNewLine & _
                "      e.��ֹ���� And (c.����ʱ�� = to_date('3000-01-01','yyyy-mm-dd') or c.����ʱ�� is null ) " & GetPriceClassString("E") & _
                "And " & strsql����ģʽ & "b.����id In (Select ID" & vbNewLine & _
                "                            From ���Ʒ���Ŀ¼" & vbNewLine & _
                "                            Where ���� In (1, 2, 3) And Nvl(To_Char(����ʱ��, 'Yyyy-Mm-Dd'), '3000-01-01') = '3000-01-01'" & vbNewLine & _
                "                            Start With ID = [1]" & vbNewLine & _
                "                            Connect By Prior ID = �ϼ�id)" & vbNewLine & _
                "Order By c.����"
    
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "����", rsProvider!id)
    
    If rsTemp.RecordCount = 0 And rsProvider!id > 0 Then
        If mint����ģʽ = 1 Then
            MsgBox "û���ҵ��÷��������۹����ҩƷ��", vbInformation, gstrSysName
            Exit Sub
        Else
            MsgBox "û���ҵ��÷����µ�ҩƷ��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
           
    Call GetDetails(rsTemp)
        
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Save()
    Dim intRow As Integer
    Set mrsReturn = New ADODB.Recordset

    With mrsReturn
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ͨ����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ʱ��", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable

        .Fields.Append "�ۼ۵�λ", adLongVarChar, 8, adFldIsNullable
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

    With vsfSelectDrug
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, vsfSelectDrugCol.ҩƷid) = "" Then Exit For
            mrsReturn.AddNew
            mrsReturn!ҩƷid = .TextMatrix(intRow, vsfSelectDrugCol.ҩƷid)
            mrsReturn!ҩƷ���� = .TextMatrix(intRow, vsfSelectDrugCol.ҩƷ����)
            mrsReturn!��Ʒ�� = .TextMatrix(intRow, vsfSelectDrugCol.��Ʒ��)
            mrsReturn!ͨ���� = .TextMatrix(intRow, vsfSelectDrugCol.ͨ����)
            mrsReturn!��� = .TextMatrix(intRow, vsfSelectDrugCol.���)
            mrsReturn!ʱ�� = IIf(.TextMatrix(intRow, vsfSelectDrugCol.����) = "ʱ��", 1, 0)
            mrsReturn!���� = .TextMatrix(intRow, vsfSelectDrugCol.����)
            mrsReturn!�ۼ۵�λ = .TextMatrix(intRow, vsfSelectDrugCol.�ۼ۵�λ)
            mrsReturn!���ﵥλ = .TextMatrix(intRow, vsfSelectDrugCol.���ﵥλ)
            mrsReturn!�����װ = .TextMatrix(intRow, vsfSelectDrugCol.����ϵ��)
            mrsReturn!סԺ��λ = .TextMatrix(intRow, vsfSelectDrugCol.סԺ��λ)
            mrsReturn!סԺ��װ = .TextMatrix(intRow, vsfSelectDrugCol.סԺϵ��)
            mrsReturn!ҩ�ⵥλ = .TextMatrix(intRow, vsfSelectDrugCol.ҩ�ⵥλ)
            mrsReturn!ҩ���װ = .TextMatrix(intRow, vsfSelectDrugCol.ҩ��ϵ��)

            mrsReturn.Update
        Next
    End With
    mblnOK = True

    Unload Me
End Sub

Private Sub Quit()
    mblnOK = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intUnitTemp As Integer
    
    '��ȡ���õĵ�λ
    mintUnit = Val(zlDataBase.GetPara("ҩƷ��λ", glngSys, 1333, 1))
    Select Case mintUnit
        Case 0 'ҩ��
            intUnitTemp = 4
        Case 1 'סԺ
            intUnitTemp = 3
        Case 2 '����
            intUnitTemp = 2
        Case 3 '�ۼ�
            intUnitTemp = 1
    End Select
    '��ȡ������λ����
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)

    mstrMatch = IIf(zlDataBase.GetPara("����ƥ��", , , 0) = "0", "%", "")
    mblnOK = False
    Call initCommandBars
    Call initVsflexgrid
    Call RestoreWinState(Me, App.ProductName, MStrCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, MStrCaption)
End Sub

Private Sub txtClass_GotFocus()
    zlControl.TxtSelAll txtClass
End Sub

Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As Recordset
    Dim rsTemp As ADODB.Recordset
    Dim strsql����ģʽ As String
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtClass.hWnd)
    On Error GoTo ErrHandle
    
    If KeyCode = vbKeyReturn Then
    
        If mint����ģʽ = 1 Then
            strsql����ģʽ = " a.�Ƿ����۹���=1 And "
        Else
            strsql����ģʽ = ""
        End If
        '����
        
        If Trim(txtClass.Text) = "" Then Exit Sub
        
        gstrSQL = "Select id,����,����" & vbNewLine & _
                    "From ���Ʒ���Ŀ¼" & vbNewLine & _
                    "Where ���� In (1, 2, 3) And (Sysdate Between ����ʱ�� And ����ʱ�� Or ����ʱ�� Is Null) And" & vbNewLine & _
                    "      (���� Like '" & "%" & UCase(txtClass.Text) & mstrMatch & "' Or ���� Like '" & "%" & UCase(txtClass.Text) & mstrMatch & "' Or" & vbNewLine & _
                    "       ���� Like '" & "%" & UCase(txtClass.Text) & mstrMatch & "')" & vbNewLine & _
                    "Order By ID"
    
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "����ѡ��", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If blnCancel = True Then Exit Sub '��ѡ����ʱ����Esc�������´���
        
        If rsProvider Is Nothing Then
            MsgBox "û���ҵ��÷����µ�ҩƷ�������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            Exit Sub
        End If
        
        gstrSQL = "Select Distinct a.ҩƷid, c.���� As ҩƷ����, c.���� As ͨ����, d.��Ʒ��, c.���, c.�Ƿ��� As ʱ��, c.����, c.���㵥λ As �ۼ۵�λ, a.���ﵥλ, a.�����װ," & vbNewLine & _
                    "                a.סԺ��λ, a.סԺ��װ, a.ҩ�ⵥλ, a.ҩ���װ, a.�ɱ���, e.�ּ�, a.ָ��������, a.ָ�����ۼ�" & vbNewLine & _
                    "From ҩƷ��� A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) D, �շѼ�Ŀ E" & vbNewLine & _
                    "Where a.ҩ��id = b.Id And a.ҩƷid = c.Id And c.Id = d.�շ�ϸĿid(+) And a.ҩƷid = e.�շ�ϸĿid And Sysdate Between e.ִ������ And" & vbNewLine & _
                    "      e.��ֹ����  And (c.����ʱ�� = to_date('3000-01-01','yyyy-mm-dd') or c.����ʱ�� is null ) " & GetPriceClassString("E") & _
                    " And " & strsql����ģʽ & "b.����id In (Select ID" & vbNewLine & _
                    "                            From ���Ʒ���Ŀ¼" & vbNewLine & _
                    "                            Where ���� In (1, 2, 3) And Nvl(To_Char(����ʱ��, 'Yyyy-Mm-Dd'), '3000-01-01') = '3000-01-01'" & vbNewLine & _
                    "                            Start With ID = [1]" & vbNewLine & _
                    "                            Connect By Prior ID = �ϼ�id)" & vbNewLine & _
                    "Order By c.����"
        
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "����", rsProvider!id)
        
        If rsTemp.RecordCount = 0 Then
            If mint����ģʽ = 1 Then
                MsgBox "û���ҵ��÷��������۹����ҩƷ��", vbInformation, gstrSysName
                Exit Sub
            Else
                MsgBox "û���ҵ��÷����µ�ҩƷ��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        Call GetDetails(rsTemp)
    End If
        
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub

    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim strҩ�� As String
    Dim lngRow As Long

    '����ҩƷ
    On Error GoTo ErrHandle
    If strInput <> txtFind.Tag Then
        '��ʾ�µĲ���
        txtFind.Tag = strInput

        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.��� In ('5','6','7') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDataBase.OpenSQLRecord(gstrSQL, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = mrsFindName!ҩƷ���� & mrsFindName!ͨ����
        Else
            strҩ�� = mrsFindName!ҩƷ���� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)
        End If

        For lngRow = 1 To vsfSelectDrug.rows - 1
            lngFindRow = vsfSelectDrug.FindRow(strҩ��, lngRow, CLng(vsfSelectDrugCol.ҩƷ��Ϣ), True, True)
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

Private Sub txtPingZhong_GotFocus()
    zlControl.TxtSelAll txtPingZhong
End Sub

Private Sub txtPingZhong_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As Recordset
    Dim rsTemp As ADODB.Recordset
    Dim strsql����ģʽ As String
    Dim vRect As RECT, blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txtPingZhong.hWnd)
    On Error GoTo ErrHandle
    
    If KeyCode = vbKeyReturn Then

        If mint����ģʽ = 1 Then
            strsql����ģʽ = " a.�Ƿ����۹���=1 And "
        Else
            strsql����ģʽ = ""
        End If
        
        If Trim(txtPingZhong.Text) = "" Then Exit Sub

        gstrSQL = "Select Distinct a.id,a.����,a.����" & vbNewLine & _
                  "  From ������ĿĿ¼ A, ������Ŀ���� B" & vbNewLine & _
                    " Where a.Id = b.������Ŀid(+) And a.��� In ('5', '6', '7') And Sysdate Between ����ʱ�� And ����ʱ�� And " & vbNewLine & _
                         " (a.���� Like '" & "%" & UCase(txtPingZhong.Text) & mstrMatch & "' Or " & vbNewLine & _
                         "a.���� Like '" & "%" & UCase(txtPingZhong.Text) & mstrMatch & "'  Or " & vbNewLine & _
                         "b.���� Like '" & "%" & UCase(txtPingZhong.Text) & mstrMatch & "'  Or " & vbNewLine & _
                         "b.���� Like '" & "%" & UCase(txtPingZhong.Text) & mstrMatch & "' )"
    
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "����ѡ��", False, False, _
        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If blnCancel = True Then Exit Sub '��ѡ����ʱ����Esc�������´���
        
        If rsProvider Is Nothing Then
            MsgBox "û���ҵ���Ʒ���µ�ҩƷ�������䣡", vbOKOnly + vbInformation, gstrSysName
            KeyCode = 0
            Exit Sub
        End If
        
        gstrSQL = "Select Distinct a.ҩƷid, c.���� As ҩƷ����, c.���� As ͨ����, d.��Ʒ��, c.���, c.�Ƿ��� As ʱ��, c.����, c.���㵥λ As �ۼ۵�λ, a.���ﵥλ, a.�����װ," & _
                                  " a.סԺ��λ , a.סԺ��װ, a.ҩ�ⵥλ, a.ҩ���װ, a.�ɱ���, e.�ּ�, a.ָ��������, a.ָ�����ۼ�" & _
                  " From ҩƷ��� A, ������ĿĿ¼ B, �շ���ĿĿ¼ C, (Select ���� As ��Ʒ��, �շ�ϸĿid From �շ���Ŀ���� Where ���� = 3) D,�շѼ�Ŀ E" & _
                  " Where a.ҩ��id = b.Id And a.ҩƷid = c.Id And c.Id = d.�շ�ϸĿid(+) and a.ҩƷid=e.�շ�ϸĿid and sysdate between e.ִ������ and e.��ֹ����  And (c.����ʱ�� = to_date('3000-01-01','yyyy-mm-dd') or c.����ʱ�� is null ) " & _
                  GetPriceClassString("E") & " and " & strsql����ģʽ & "b.id=[1] order by c.����"
                  
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Ʒ��", rsProvider!id)
               
        If rsTemp.RecordCount = 0 Then
            If mint����ģʽ = 1 Then
                MsgBox "û���ҵ���Ʒ�������۹����ҩƷ��", vbInformation, gstrSysName
                Exit Sub
            Else
                MsgBox "û���ҵ���Ʒ���µ�ҩƷ��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
               
        Call GetDetails(rsTemp)
        
    End If

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetDetails(ByVal rsTemp As ADODB.Recordset)
    Dim lngID As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim blnDou As Boolean '�ظ�����
    Dim dbl����ϵ�� As Double
    Dim strUnit As String   '��λ

    With vsfSelectDrug
        For intRow = 0 To rsTemp.RecordCount - 1
            blnDou = False
            For i = 1 To .rows - 1
                If .TextMatrix(i, vsfSelectDrugCol.ҩƷid) = rsTemp!ҩƷid Then
                    blnDou = True
                End If
            Next
            If blnDou = False Then
                .rows = .rows + 1
                .RowHeight(.rows - 1) = mlngRowHeight

                Select Case mintUnit
                    Case 0
                        dbl����ϵ�� = rsTemp!ҩ���װ
                        strUnit = rsTemp!ҩ�ⵥλ
                    Case 1
                        dbl����ϵ�� = rsTemp!סԺ��װ
                        strUnit = rsTemp!סԺ��λ
                    Case 2
                        dbl����ϵ�� = rsTemp!�����װ
                        strUnit = rsTemp!���ﵥλ
                    Case 3
                        dbl����ϵ�� = 1
                        strUnit = rsTemp!�ۼ۵�λ
                End Select

                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷid) = rsTemp!ҩƷid
                If gintҩƷ������ʾ = 1 Then
                    .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ��Ϣ) = "[" & rsTemp!ҩƷ���� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)
                Else
                    .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ��Ϣ) = "[" & rsTemp!ҩƷ���� & "]" & rsTemp!ͨ����
                End If

                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩƷ����) = rsTemp!ҩƷ����
                .TextMatrix(.rows - 1, vsfSelectDrugCol.��Ʒ��) = IIf(IsNull(rsTemp!��Ʒ��), "", rsTemp!��Ʒ��)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ͨ����) = IIf(IsNull(rsTemp!ͨ����), "", rsTemp!ͨ����)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.��λ) = strUnit

                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ۼ۵�λ) = rsTemp!�ۼ۵�λ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.���ﵥλ) = rsTemp!���ﵥλ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.����ϵ��) = rsTemp!�����װ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.סԺ��λ) = rsTemp!סԺ��λ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.סԺϵ��) = rsTemp!סԺ��װ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩ�ⵥλ) = rsTemp!ҩ�ⵥλ
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ҩ��ϵ��) = rsTemp!ҩ���װ


                .TextMatrix(.rows - 1, vsfSelectDrugCol.����) = IIf(rsTemp!ʱ�� = 1, "ʱ��", "����")
                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ۼ�) = zlStr.FormatEx(dbl����ϵ�� * rsTemp!�ּ�, mintPriceDigit, , True)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.�ɱ���) = zlStr.FormatEx(dbl����ϵ�� * rsTemp!�ɱ���, mintCostDigit, , True)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ָ������) = zlStr.FormatEx(dbl����ϵ�� * rsTemp!ָ��������, mintCostDigit, , True)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.ָ���ۼ�) = zlStr.FormatEx(dbl����ϵ�� * rsTemp!ָ�����ۼ�, mintPriceDigit, , True)

            End If
            rsTemp.MoveNext
        Next
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub initCommandBars()
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    Dim cbrControlPopu As CommandBarControl
    Dim lngCount As Integer
    
    With CommandBarsGlobalSettings
        .App = App
        .CompanyName = "����������Ϣ��ҵ�������ι�˾" '��˾����
        .ResourceFile = .OcxPath & "\XTPResourceZhCn.dll" '��������������Դ�ļ�
        .ColorManager.SystemTheme = xtpSystemThemeAuto  '�ؼ��������ɫ����
    End With

    With cbsMain.Options
        .ShowExpandButtonAlways = False '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False '�����õĲ˵���������
        .UseFadedIcons = True 'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24 '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16 '����Сͼ��ĳߴ�
    End With

    With cbsMain
        .VisualTheme = xtpThemeOffice2003 '���ÿؼ���ʾ���
        .EnableCustomization False '�Ƿ������Զ�������
        Set .Icons = imgList.Icons '���ù�����ͼ��ؼ�
        .ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap '����仯ʱ�������ʾ����˵�Ҳ������
        .ActiveMenuBar.Title = "�˵�"
    End With
    
    'ɾ�����ڵĹ������������˵���
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 1 Step -1
        cbsMain(lngCount).Delete
    Next
    
    '����������
    Set cbrToolBar = cbsMain.Add("������", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ContextMenuPresent = False

    With cbrToolBar
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_ClearAll, "���")
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Find, "����")
        cbrControl.Visible = False

        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Save, "���")
        cbrControl.BeginGroup = True
        Set cbrControl = .Controls.Add(xtpControlButton, mconMenu_Quit, "ȡ��")
                
    End With

    For Each cbrControl In cbrToolBar.Controls  '�ù������а�ťͬʱ��ʾͼ�������
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconMenu_Find
    End With

End Sub

