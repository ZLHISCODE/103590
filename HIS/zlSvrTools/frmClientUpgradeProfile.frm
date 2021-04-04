VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClientUpgradeProfile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11850
   FillColor       =   &H80000000&
   Icon            =   "frmClientUpgradeProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFlash 
      Caption         =   "ˢ��(&F)"
      Height          =   350
      Left            =   975
      TabIndex        =   9
      Top             =   6465
      Width           =   1100
   End
   Begin zlSvrStudio.ucPieChart pcUpgrade 
      Height          =   1800
      Left            =   45
      TabIndex        =   8
      Top             =   0
      Width           =   1950
      _extentx        =   3201
      _extenty        =   2858
      showtype        =   0
      symboltype      =   0
      title           =   ""
      linecolor       =   0
      titlefont       =   "frmClientUpgradeProfile.frx":6852
      itemfont        =   "frmClientUpgradeProfile.frx":6878
      itemcolor       =   0
      titlecolor      =   0
      legend          =   -1
      levcolor        =   0
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      FillColor       =   &H80000000&
      Height          =   6405
      Left            =   2370
      ScaleHeight     =   6405
      ScaleWidth      =   8580
      TabIndex        =   0
      Top             =   0
      Width           =   8580
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   1
         Left            =   4140
         TabIndex        =   7
         Top             =   1725
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   222429187
         CurrentDate     =   42892
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "ˢ��(&R)"
         Height          =   350
         Left            =   6720
         TabIndex        =   5
         Top             =   1707
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   2
         Top             =   1725
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   222429187
         CurrentDate     =   42892
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfReport 
         Height          =   3840
         Left            =   150
         TabIndex        =   1
         Top             =   2265
         Width           =   5265
         _cx             =   9287
         _cy             =   6773
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483644
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483634
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   260
         RowHeightMax    =   260
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClientUpgradeProfile.frx":689C
         ScrollTrack     =   0   'False
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
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   3825
         TabIndex        =   6
         Top             =   1792
         Width           =   180
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "�ͻ�����������ͳ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   4
         Top             =   100
         Width           =   2160
      End
      Begin VB.Label lblPrecision 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ�䷶Χ"
         Height          =   180
         Left            =   150
         TabIndex        =   3
         Top             =   1785
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmClientUpgradeProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum PicColor
    PC_δ���� = &HF2D3C1
    PC_�����ɹ� = &H5AD0B5
    PC_����ʧ�� = &H7D4AAC
    PC_�������� = &HF7995F
End Enum

Private Enum DateBetween
    DB_��ʼʱ�� = 0
    DB_����ʱ�� = 1
End Enum

Private Enum vsfReport_Column
    VC_�������� = 0
    VC_����Ƶ�� = 1
End Enum

Private Sub cmdFilter_Click()
    Call FillData
    If VScrollVisible(vsfReport) Then
        vsfReport.ColWidth(VC_��������) = vsfReport.Width - vsfReport.ColWidth(VC_����Ƶ��) - 300
    Else
        vsfReport.ColWidth(VC_��������) = vsfReport.Width - vsfReport.ColWidth(VC_����Ƶ��) - 50
    End If
End Sub

Private Sub FillData()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
 
    On Error GoTo errH:
    strSQL = "Select ����, Count(����) as Ƶ��" & vbNewLine & _
            "From (Select Upper(����) ����" & vbNewLine & _
            "       From Zlclientupdatelog" & vbNewLine & _
            "       Where �������� > [1] And �������� < [2] And ���� not like '����:%' And ���� not like '���:%')" & vbNewLine & _
            "Group By ����" & vbNewLine & _
            "Order By Ƶ�� Desc"

    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, dtpDate(DB_��ʼʱ��).value, dtpDate(DB_����ʱ��).value)
    With rsTemp
        vsfReport.Rows = vsfReport.FixedRows
        vsfReport.Rows = .RecordCount + 1
        If .RecordCount = 0 Then vsfReport.Rows = 2
        For i = 1 To .RecordCount
            vsfReport.TextMatrix(i, VC_��������) = !����
            vsfReport.TextMatrix(i, VC_����Ƶ��) = !Ƶ��
            .MoveNext
        Next
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub cmdFlash_Click()
    Call FillPic
End Sub

Private Sub Form_Resize()
    pcUpgrade.Width = Me.Width / 5 * 2
    pcUpgrade.Height = Me.Height
    picRight.Left = pcUpgrade.Width
    picRight.Width = Me.Width / 5 * 3
    picRight.Height = Me.Height
End Sub

Public Sub RefreshData()
    Call LoadData
End Sub

Private Sub LoadData()
    On Error GoTo errH
    '��ʼ��dtpDate
    dtpDate(DB_����ʱ��).value = CurrentDate()
    dtpDate(DB_��ʼʱ��).value = DateAdd("d", -3, dtpDate(DB_����ʱ��).value)
    
    '��ͼ
    Call FillPic
        
    '���
    Call FillData
    Exit Sub
errH:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
    If False Then
        Resume
    End If
End Sub

Private Sub FillPic()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
    pcUpgrade.Clear
    strSQL = "Select Nvl(Sum(δ����), 0) δ����, Nvl(Sum(�������), 0) �������, Nvl(Sum(����ʧ��), 0) ����ʧ��, Nvl(Sum(��������), 0) ��������, Nvl(Count(1), 0) ����" & vbNewLine & _
            "From (Select Decode(�������, 0, 1, 0) δ����, Decode(�������, 1, 1, 0) �������, Decode(�������, 2, 1, 0) ����ʧ��," & vbNewLine & _
            "              Decode(�������, 3, 1, 0) ��������" & vbNewLine & _
            "       From Zlclients)"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    pcUpgrade.Tag = "�ͻ���������" & rsTemp!���� & _
                    "����δ��������" & rsTemp!δ���� & _
                    "���������ɹ�����" & rsTemp!������� & _
                    "��������ʧ������" & rsTemp!����ʧ�� & _
                    "����������������" & rsTemp!�������� & "��"
    pcUpgrade.Title = "�ͻ����������" & "(��" & rsTemp!���� & "��)"
    pcUpgrade.addItem "δ����", PC_δ����, rsTemp!δ����
    pcUpgrade.addItem "�����ɹ�", PC_�����ɹ�, rsTemp!�������
    pcUpgrade.addItem "����ʧ��", PC_����ʧ��, rsTemp!����ʧ��
    pcUpgrade.addItem "��������", PC_��������, rsTemp!��������
    pcUpgrade.PaintChart
    Exit Sub
errH:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
    If False Then
        Resume
    End If
End Sub

Public Sub SetMenu()
    frmMDIMain.stbThis.Panels(2).Text = pcUpgrade.Tag
End Sub

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = False
End Function

Private Sub picRight_Resize()
    On Error Resume Next
    lblTitle.Left = (picRight.Width - lblTitle.Width) / 2
    lblPrecision.Top = lblTitle.Top + lblTitle.Height + 300
    dtpDate(DB_��ʼʱ��).Top = lblPrecision.Top + lblPrecision.Height / 2 - dtpDate(DB_��ʼʱ��).Height / 2
    dtpDate(DB_����ʱ��).Top = dtpDate(DB_��ʼʱ��).Top
    lblEnd.Top = lblPrecision.Top
    cmdFilter.Top = dtpDate(DB_��ʼʱ��).Top - 10
    vsfReport.Top = lblPrecision.Top + lblPrecision.Height + 200
    vsfReport.Width = picRight.Width - vsfReport.Left - 100
    vsfReport.Height = picRight.Height - vsfReport.Top - 100
    cmdFlash.Left = pcUpgrade.Width - cmdFlash.Width
    cmdFlash.Top = pcUpgrade.Height - cmdFlash.Height - 100
    If VScrollVisible(vsfReport) Then
        vsfReport.ColWidth(VC_��������) = vsfReport.Width - vsfReport.ColWidth(VC_����Ƶ��) - 300
    Else
        vsfReport.ColWidth(VC_��������) = vsfReport.Width - vsfReport.ColWidth(VC_����Ƶ��) - 50
    End If
    err.Clear
End Sub

Public Sub SetControlEnable(ByVal strProgFunc As String)
'����Ȩ���ַ������ÿؼ�״̬
'strProgFunc:Ȩ���ַ���
End Sub

