VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInDoctorAdvice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8955
   ControlBox      =   0   'False
   Icon            =   "frmInDoctorAdvice.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7155
      TabIndex        =   6
      Top             =   210
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frmInDoctorAdvice.frx":000C
         ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.Frame fraAdviceUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   240
      MousePointer    =   7  'Size N S
      TabIndex        =   4
      Top             =   5040
      Width           =   7275
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4875
      Left            =   210
      TabIndex        =   0
      Top             =   165
      Width           =   6240
      _cx             =   11007
      _cy             =   8599
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInDoctorAdvice.frx":055A
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList imgFlag 
         Left            =   345
         Top             =   645
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":05F5
               Key             =   "����"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":080F
               Key             =   "��¼"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":0D29
               Key             =   "δ����"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":1243
               Key             =   "������"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgPass 
         Left            =   975
         Top             =   645
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":175D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":1A57
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":1D51
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":204B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":2345
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSign 
         Left            =   1635
         Top             =   645
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInDoctorAdvice.frx":263F
               Key             =   "ǩ��"
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   45
      ScaleHeight     =   600
      ScaleWidth      =   630
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   630
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAppend 
      Height          =   1425
      Left            =   225
      TabIndex        =   2
      Top             =   5505
      Width           =   6270
      _cx             =   11060
      _cy             =   2514
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
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid vsColumn 
      Height          =   3495
      Left            =   6885
      TabIndex        =   5
      Top             =   450
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   6165
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInDoctorAdvice.frx":2991
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
   Begin MSComctlLib.TabStrip tabAppend 
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   529
      MultiRow        =   -1  'True
      Style           =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҽ��������ϸ(&S)"
            Key             =   "ҽ��������ϸ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҽ��ǩ����¼(&G)"
            Key             =   "ҽ��ǩ����¼"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmInDoctorAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mfrmParent As Object
Public mstrPrivs As String
Private WithEvents mfrmEdit As Form
Attribute mfrmEdit.VB_VarHelpID = -1

'�ϴ�ˢ������ʱ�Ĳ�����Ϣ
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mlng����ID As Long
Private mbln��Ժ As Boolean
Private mlngǰ��ID As Long
Private mblnShowAll As Boolean

Private mblnMoved As Boolean
Private mvInDate As Date

Private Enum Menu_Advice
    mnu�¿�ҽ�� = 0
    mnu��¼ҽ�� = 1
    mnu�޸�ҽ�� = 2
    mnuɾ��ҽ�� = 3
    mnuҽ����� = 5 '-
    mnuҽ��ֹͣ = 7 '-
    mnuҽ������ = 8
    mnuҽ����ͣ = 9 '-
    mnuҽ������ = 11
    mnu�������� = 13 '-
    mnuҽ������ = 14
    mnu���Ƶ��ı� = 16 '-
End Enum

'����˵�������
Private Enum Menu_Report
    mnu����ҽ���� = 0
    mnu��ʱҽ���� = 1
    mnuҽ����¼�� = 2
End Enum

'�̶���
Private Const COL_F��־ = 0
Private Const COL_F���� = 1
'������
Private Const COL_ID = 2
Private Const COL_���ID = COL_ID + 1
Private Const COL_��ID = COL_ID + 2
Private Const COL_��� = COL_ID + 3
Private Const COL_Ӥ��ID = COL_ID + 4
Private Const COL_ҽ��״̬ = COL_ID + 5
Private Const COL_������� = COL_ID + 6
Private Const COL_�������� = COL_ID + 7
Private Const COL_������� = COL_ID + 8
Private Const COL_��־ = COL_ID + 9
'�ɼ���
Private Const COL_��ʾ = COL_ID + 10 'Pass
Private Const COL_��Ч = COL_ID + 11
Private Const COL_��ʼʱ�� = COL_ID + 12
Private Const COL_ҽ������ = COL_ID + 13
Private Const COL_Ƥ�� = COL_ID + 14
Private Const COL_���� = COL_ID + 15
Private Const COL_���� = COL_ID + 16
Private Const COL_Ƶ�� = COL_ID + 17
Private Const COL_�÷� = COL_ID + 18
Private Const COL_ҽ������ = COL_ID + 19
Private Const COL_ִ��ʱ�� = COL_ID + 20
Private Const COL_��ֹʱ�� = COL_ID + 21
Private Const COL_ִ�п��� = COL_ID + 22
Private Const COL_ִ������ = COL_ID + 23
Private Const COL_�ϴ�ִ�� = COL_ID + 24
Private Const COL_״̬ = COL_ID + 25
Private Const COL_����ҽ�� = COL_ID + 26
Private Const COL_����ʱ�� = COL_ID + 27
Private Const COL_У�Ի�ʿ = COL_ID + 28
Private Const COL_У��ʱ�� = COL_ID + 29
Private Const COL_ͣ��ҽ�� = COL_ID + 30
Private Const COL_ͣ��ʱ�� = COL_ID + 31
Private Const COL_ͣ����ʿ = COL_ID + 32
Private Const COL_ȷ��ͣ��ʱ�� = COL_ID + 33
'������
Private Const COL_����ID = COL_ID + 34 '��Ӧ�����ļ�Ŀ¼.ID
Private Const COL_������ = COL_ID + 35 '���Ƶ����Ƿ���������
Private Const COL_������ = COL_ID + 36 '���Ƶ����Ƿ��б�����
Private Const COL_����ID = COL_ID + 37 '��Ӧ���˲�����¼.ID
Private Const COL_ǰ��ID = COL_ID + 38
Private Const COL_ǩ���� = COL_ID + 39

Private Enum COL�����嵥
    cs���ͺ� = 0
    cs����ʱ�� = 1
    cs����ҽ�� = 2
    cs���ݺ� = 3
    cs�շ���Ŀ = 4
    cs���� = 5
    cs�Ʒ�״̬ = 6
    csִ��״̬ = 7
    csִ�п��� = 8
    cs�״�ʱ�� = 9
    csĩ��ʱ�� = 10
    cs������ = 11
    cs��¼���� = 12
End Enum

Public Function zlRefresh(lng����ID As Long, lng��ҳID As Long, lng����ID As Long, _
    lng����ID As Long, Optional bln��Ժ As Boolean, Optional ByVal lngǰ��ID As Long = 0, Optional ByVal ifShowAll As Boolean = True) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID
    mlng����ID = lng����ID: mlng����ID = lng����ID
    mbln��Ժ = bln��Ժ: mlngǰ��ID = lngǰ��ID
    mblnShowAll = ifShowAll
    
    '�жϲ����Ƿ���ת������Ժʱ��
    '��Ϊ�ú������ⶼ�ڵ���,�������ñ�,ֱ�Ӷ�ȡ
    mblnMoved = False
    If lng����ID <> 0 Then
        strSQL = "Select ��Ժ����,����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng����ID, lng��ҳID)
        On Error GoTo 0
        mblnMoved = Nvl(rsTmp!����ת��, 0) <> 0
        mvInDate = rsTmp!��Ժ����
    End If
    
    If mlng����ID = 0 Then
        '���ҽ���嵥
        Call ClearAdviceData
        Call ClearAppendData
    Else
        '��ʾҽ���嵥
        Call LoadAdvice
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlButtonClick(objButton As Button) As Boolean
'���ܣ�ִ��ҽ����ť����
    Select Case objButton.Key
        Case "�¿�"
            Call FuncAdviceAdd
        Case "�޸�"
            Call FuncAdviceModi
        Case "ɾ��"
            Call FuncAdviceDel
        Case "���"
            Call FuncAdviceAudit
        Case "ֹͣ"
            Call FuncAdviceStop
        Case "����"
            Call FuncAdviceRevoke
        Case "����"
            Call FuncAdviceSend
        Case "����"
            '���������д�����
        Case "ǩ��"
            Call FuncAdviceSign
    End Select
End Function

Public Function zlMenuClick(objMenu As Menu) As Boolean
'���ܣ�ִ��ҽ���˵�����
    Dim strText As String
    
    If objMenu.Caption Like "*(&*)*" Then
        strText = Split(objMenu.Caption, "(")(0)
    Else
        strText = objMenu.Caption
    End If
        
    If objMenu.Name = "mnuAdviceFuncRoll" Then
        '�����Ӳ˵��Ĵ���
        Call FuncAdviceRoll(objMenu.Index)
    ElseIf objMenu.Name = "mnuViewAdviceAppend" Then
        '��ʾ/���ظ��ӱ��
        objMenu.Checked = Not objMenu.Checked
        fraAdviceUD.Visible = objMenu.Checked
        vsAppend.Visible = objMenu.Checked
        Call Form_Resize
        
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Else
        Select Case strText
            Case "�¿�ҽ��"
                Call FuncAdviceAdd
            Case "��¼ҽ��"
                Call FuncAdviceSupply
            Case "�޸�ҽ��"
                Call FuncAdviceModi
            Case "ɾ��ҽ��"
                Call FuncAdviceDel
            Case "ҽ�����"
                Call FuncAdviceAudit
            Case "ҽ��ֹͣ"
                Call FuncAdviceStop
            Case "ҽ������"
                Call FuncAdviceRevoke
            Case "ҽ����ͣ"
                Call FuncAdvicePause
            Case "ҽ������"
                Call FuncAdviceResume
            Case "��������"
                Call FuncAdviceSend
            Case "����ҽ����"
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1204_1", mfrmParent, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID)
            Case "��ʱҽ����"
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1204_2", mfrmParent, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID)
            Case "ҽ����¼��"
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1204_3", mfrmParent, "���˿���=" & mlng����ID)
            Case "���Ƶ��ı�"
                Call FuncCopyToText
            Case "����ǩ��"
                Call FuncAdviceSign
            Case "��֤ǩ��"
                Call FuncAdviceSignVerify
            Case "ȡ��ǩ��"
                Call FuncAdviceSignErase
        End Select
    End If
    
    zlMenuClick = True
End Function

Private Sub SetFuncEnabled()
'���ܣ����ݵ�ǰ���˻�������������ù��ܿ�����
    Dim blnAdvice As Boolean, blnEnabled As Boolean
    
    '���������ط�����
    On Error Resume Next
    
    With mfrmParent
        '1.�޲��˵����:mlng����ID <> 0
        '2.�����ѳ�Ժ�����:Not mbln��Ժ
        '3.�����ݵ����
        blnAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
        
        .mnuAdviceFunc(mnu�¿�ҽ��).Enabled = mlng����ID <> 0 And Not mbln��Ժ
        .mnuAdviceFunc(mnu��¼ҽ��).Enabled = mlng����ID <> 0 And Not mbln��Ժ
        
        'δУ�Բſ����޸�
        blnEnabled = mlng����ID <> 0 And Not mbln��Ժ And blnAdvice
        If blnEnabled Then
            If InStr(",1,2,", vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 0 Then blnEnabled = False
        End If
        If blnEnabled Then 'δǩ��
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 Then blnEnabled = False
        End If
        .mnuAdviceFunc(mnu�޸�ҽ��).Enabled = blnEnabled
        
        'δУ�Բſ���ɾ��
        blnEnabled = mlng����ID <> 0 And blnAdvice
        If blnEnabled Then
            If InStr(",1,2,", vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 0 Then blnEnabled = False
        End If
        If blnEnabled Then 'δǩ��
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 Then blnEnabled = False
        End If
        .mnuAdviceFunc(mnuɾ��ҽ��).Enabled = blnEnabled
        
        'ҽ�����:���¿�һ��
        .mnuAdviceFunc(mnuҽ�����).Enabled = mlng����ID <> 0 And Not mbln��Ժ
        
        .mnuAdviceFunc(mnuҽ��ֹͣ).Enabled = mlng����ID <> 0
        .mnuAdviceFunc(mnuҽ������).Enabled = mlng����ID <> 0
        .mnuAdviceFunc(mnuҽ����ͣ).Enabled = mlng����ID <> 0 And Not mbln��Ժ
        .mnuAdviceFunc(mnuҽ������).Enabled = mlng����ID <> 0 And Not mbln��Ժ
        
        '��Ժ���˲�������˲���:Ԥ��Ժ���˿��Ի��˳�Ժҽ������
        blnEnabled = False
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "Z" _
            And InStr(",5,6,11,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������))) > 0 Then
            If Val(mfrmParent.lvwPati.SelectedItem.ListSubItems(4).Tag) = 3 Then
                If mfrmParent.mnuAdviceFuncRoll(0).Tag <> "" Then
                    If Val(Split(mfrmParent.mnuAdviceFuncRoll(0).Tag, "|")(0)) = 0 _
                        And Val(Split(mfrmParent.mnuAdviceFuncRoll(0).Tag, "|")(1)) <> 0 Then
                        blnEnabled = True
                    End If
                End If
            End If
        End If
        .mnuAdviceFunc(mnuҽ������).Enabled = mlng����ID <> 0 And blnAdvice And (Not mbln��Ժ Or blnEnabled)
        .mnuAdviceFunc(mnu��������).Enabled = mlng����ID <> 0 And Not mbln��Ժ
        
        .mnuReportItem(mnu����ҽ����).Enabled = mlng����ID <> 0
        .mnuReportItem(mnu��ʱҽ����).Enabled = mlng����ID <> 0
        .mnuReportItem(mnuҽ����¼��).Enabled = mlng����ID <> 0
        .mnuAdviceFunc(mnu���Ƶ��ı�).Enabled = mlng����ID <> 0 And blnAdvice
                
        '----------------------------------------------------------------------
        '����ǩ������
        blnEnabled = mlng����ID <> 0 And blnAdvice And tabAppend.SelectedItem.Index = 2
        If blnEnabled Then
            If vsAppend.RowData(vsAppend.Row) = 0 Then blnEnabled = False
        End If
        .mnuSignVerify.Enabled = blnEnabled
        .mnuSignErase.Enabled = blnEnabled
        .mnuSignNew.Enabled = mlng����ID <> 0
        .tbrSys.Buttons("ǩ��").Enabled = .mnuSignNew.Enabled
        '----------------------------------------------------------------------
        .tbrSys.Buttons("�¿�").Enabled = .mnuAdviceFunc(mnu�¿�ҽ��).Enabled
        .tbrSys.Buttons("�޸�").Enabled = .mnuAdviceFunc(mnu�޸�ҽ��).Enabled
        .tbrSys.Buttons("ɾ��").Enabled = .mnuAdviceFunc(mnuɾ��ҽ��).Enabled
        .tbrSys.Buttons("���").Enabled = .mnuAdviceFunc(mnuҽ�����).Enabled
        .tbrSys.Buttons("ֹͣ").Enabled = .mnuAdviceFunc(mnuҽ��ֹͣ).Enabled
        .tbrSys.Buttons("����").Enabled = .mnuAdviceFunc(mnuҽ������).Enabled
        .tbrSys.Buttons("����").Enabled = .mnuAdviceFunc(mnu��������).Enabled
        .tbrSys.Buttons("����").Enabled = .mnuAdviceFunc(mnuҽ������).Enabled
    End With
End Sub

Private Sub FuncAdviceSend()
'���ܣ���������
    Dim blnRefresh As Boolean
    
    If frmInAdviceSend.ShowMe(mfrmParent, mstrPrivs, mlng����ID, mlng��ҳID, mlngǰ��ID, blnRefresh) Then
        If blnRefresh Then
            Call mfrmParent.mnuViewRefresh_Click
        Else
            Call LoadAdvice
        End If
    End If
End Sub

Private Sub FuncAdviceRoll(Index As Integer)
'���ܣ�ҽ������
'������Index=���������ڲ˵��ϵ�����
    Dim strSQL As String, strOper As String
    Dim lngFlag As Long, blnBat As Boolean
    Dim int���� As Integer, lngҽ��ID As Long, lng���ͺ� As Long
    Dim vOperDate As Date, vOperName As String
    Dim lngǩ��ID As Long, strSign As String
    
    If Val(mfrmParent.mnuAdviceFunc(mnuҽ������).Tag) = 0 Then Exit Sub
    If mfrmParent.mnuAdviceFuncRoll(Index).Tag = "" Then Exit Sub
    
    '(��ID)ȡһ��ҽ�������IDΪ�յ�ҽ��ID(��ҩ;��,��ҩ�÷�,��Ҫ����,�����Ŀ,������ҽ��)
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID)) <> 0 Then
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID))
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    If lngҽ��ID = 0 Then Exit Sub
    
    int���� = Val(Split(mfrmParent.mnuAdviceFuncRoll(Index).Tag, "|")(0))
    lng���ͺ� = Val(Split(mfrmParent.mnuAdviceFuncRoll(Index).Tag, "|")(1))
    vOperDate = CDate(Split(mfrmParent.mnuAdviceFuncRoll(Index).Tag, "|")(2))
    vOperName = Split(mfrmParent.mnuAdviceFuncRoll(Index).Tag, "|")(3)
        
    'ҽ��ֻ�ܻ������ѵĲ���,�Ե���ǩ��ͬʱҲ�ж����Ƿ���˱��˵�ǩ��
    If vOperName <> UserInfo.���� Then
        MsgBox "�㲻�ܻ��������˶�ҽ���Ĳ�����" & vbCrLf & vbCrLf & mfrmParent.mnuAdviceFuncRoll(Index).Caption & vbTab, vbInformation, gstrSysName
        Exit Sub
    End If
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
        
    '����ҽ���Ƿ�������ǩ����������ʾ
    '-------------------------------------------------------
    blnBat = AdviceCanBatchRoll(lngҽ��ID, int����, lng���ͺ�, vOperDate, lngǩ��ID) '����һ�������ҽ���Ƿ���ǩ��
    strOper = Decode(int����, 0, "����", 4, "����", 5, "����", 6, "��ͣ", 7, "����", 8, "ֹͣ", 9, "ȷ��ֹͣ", 10, "��дƤ�Խ��")
    If MsgBox("ȷʵҪ�������²�����" & vbCrLf & vbCrLf & mfrmParent.mnuAdviceFuncRoll(Index).Caption & vbTab & _
        IIF(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 And lngǩ��ID <> 0, _
            vbCrLf & vbCrLf & "��ʾ����ҽ��" & strOper & "ʱ��ǩ������ͬʱ��������һ��" & strOper & "��ǩ��������ҽ����", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    '����������ʾ
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 And lngǩ��ID <> 0 Then
        '��ǰ������ҽ��һ�������ǩ��,�̶�һ�����(blnBat=True)
    Else
        If blnBat Then
            If MsgBox("��������ҽ���͵�ǰҽ��һ��ͬʱ" & strOper & "��Ҫͬʱ������Щҽ����", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnBat = False
        End If
    End If
    
    '��ҽ�����õĽ���������м��
    If int���� = 0 And lng���ͺ� <> 0 Then
        If Not CheckAdviceBalanceRoll(lng���ͺ�, lngҽ��ID, blnBat) Then Exit Sub
    End If
    
    If int���� = 8 Then '��������ֱ�ӻ����Զ�ֹͣ
        If blnBat Then
            If MsgBox("Ҫ������Щҽ����ִ����ֹʱ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                lngFlag = 1
            End If
        Else
            If RowIs�䷽��(vsAdvice.Row) Then
                lngFlag = 1 '��ҩ�䷽ʼ�ձ���ִ����ֹʱ��
            Else
                If MsgBox("Ҫ����ҽ����ִ����ֹʱ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    lngFlag = 1
                End If
            End If
        End If
    End If
    
    '���漰������ǩ���Ĳ�������ȡ��ǩ��
    '-------------------------------------------------------
    If blnBat Then
        If lngǩ��ID = 0 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 Then
            lngǩ��ID = GetAdviceSign(lngҽ��ID, int����, vOperName, vOperDate)
        End If
    Else
        lngǩ��ID = 0
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 Then
            lngǩ��ID = GetAdviceSign(lngҽ��ID, int����, vOperName, vOperDate)
        End If
    End If
    If lngǩ��ID <> 0 Then
        strSign = "zl_ҽ��ǩ����¼_Delete(" & lngǩ��ID & ")"
    End If
    
    '����ܷ����ǩ��
    If strSign <> "" Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "ϵͳû�����õ���ǩ����֤���ģ����˲������ܼ�����", vbInformation, gstrSysName
            Else
                MsgBox "����ǩ������δ����ȷ��װ�����˲������ܼ�����", vbInformation, gstrSysName
            End If
            Exit Sub
        Else
            If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    End If
    
    '����ǻ��˷������ѼƷ�,��δ���������ϴ�����(1.�����ǲ�������,2.Ҳ�ɲ���,Ԥ��ʱ�Զ��ϴ�)
    If blnBat Then
        strSQL = "zl_����ҽ����¼_��������(" & lngҽ��ID & "," & int���� & "," & _
            "To_Date('" & Format(vOperDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & lng���ͺ� & "," & lngFlag & ")"
    Else
        strSQL = "zl_����ҽ����¼_����(" & lngҽ��ID & "," & lngFlag & ")"
    End If
    
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans
    If strSign <> "" Then
        Call zlDatabase.ExecuteProcedure(strSign, Me.Name)
    End If
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Name)
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "Z" _
        And InStr(",5,6,11,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������))) > 0 _
        And int���� = 0 And lng���ͺ� <> 0 Then
        '���˳�Ժҽ��ˢ��������
        Call mfrmParent.mnuViewRefresh_Click
    Else
        Call LoadAdvice
    End If
    Screen.MousePointer = 0
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function AdviceCanBatchRoll(ByVal lngҽ��ID As Long, ByVal int���� As Integer, ByVal lng���ͺ� As Long, ByVal datʱ�� As Date, lngǩ��ID As Long) As Boolean
'���ܣ����ָ��ҽ����ǰ�����Ƿ�������ҽ��һ������ִ�е�,���ж��Ƿ������������
'������lngҽ��ID=���IDΪ�յ�ҽ����ID(һ��ҽ����ID)
'      int����=ҽ����������
'      datʱ��=ҽ��������ʱ��
'���أ��Ƿ��п���һ����˵�����ҽ��
'      lngǩ��ID=��ЩҪ���˵�ҽ���Ƿ���ǩ��(����,ֹͣ),�����򷵻�ǩ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    lngǩ��ID = 0
    If int���� = 0 Then
        strSQL = "Select ҽ��ID From ����ҽ������ A Where ���ͺ�=[2]" & _
            " And Not Exists(Select ID From ����ҽ����¼ B Where B.ID=A.ҽ��ID And (ID=[1] Or ���ID=[1]))"
    Else
        strSQL = "Select ��������,����ʱ��,������Ա From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=[3] And ����ʱ��=[4]"
        strSQL = "Select ҽ��ID,Nvl(ǩ��ID,0) as ǩ��ID From ����ҽ��״̬ A Where (��������,����ʱ��,������Ա)=(" & strSQL & ")" & _
            " And Not Exists(Select ID From ����ҽ����¼ B Where B.ID=A.ҽ��ID And (ID=[1] Or ���ID=[1] Or (A.��������=8 And ҽ����Ч=1)))"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lngҽ��ID, lng���ͺ�, int����, datʱ��)
    If Not rsTmp.EOF Then
        If int���� = 0 Then
'            '����ͨ�����������ѳ�Ժ��Ԥ��Ժ���˵�ҽ������
'            strSQL = "Select C.����ID,C.��ҳID From ����ҽ������ A,����ҽ����¼ B,������ҳ C" & _
'                " Where A.ҽ��ID=B.ID And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
'                " And (C.��Ժ���� is Not NULL Or C.״̬=3) And A.���ͺ�=[1] And Rownum=1"
'            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng���ͺ�)
'            If Not rsTmp.EOF Then Exit Function
        ElseIf int���� <> 0 Then
            rsTmp.Filter = "ǩ��ID<>0"
            If Not rsTmp.EOF Then lngǩ��ID = rsTmp!ǩ��ID
        End If
        AdviceCanBatchRoll = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncCopyToText()
    Dim strCopy As String, intRow As Integer
    
    With vsAdvice
        strCopy = ""
        For intRow = .FixedRows To .Rows - 1
            If InStr(",5,6,", .TextMatrix(intRow, COL_�������)) > 0 Then
                strCopy = strCopy & .TextMatrix(intRow, COL_ҽ������) _
                        & " " & .TextMatrix(intRow, COL_����) _
                        & " " & .TextMatrix(intRow, COL_Ƶ��) _
                        & " " & .TextMatrix(intRow, COL_�÷�) _
                        & vbCrLf
            Else
                strCopy = strCopy & .TextMatrix(intRow, COL_ҽ������) & vbCrLf
            End If
        Next
    End With
    If strCopy <> "" Then
        VB.Clipboard.Clear
        VB.Clipboard.SetText strCopy
    End If
End Sub

Private Function RowIs�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���ҩ�䷽��
'˵����ָ����Ϊ��ʾ��,�����="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where Rownum=1 And �������='7' And ���ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs�䷽�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���������
'˵����ָ����Ϊ��ʾ��,�����="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where Rownum=1 And �������='C' And ���ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlItemRef()
'���ܣ��������Ʋο�
    Dim lng������ĿID As Long, i As Long

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) <> 0 Then
            If .TextMatrix(.Row, COL_�������) = "E" And (RowIs�䷽��(.Row) Or RowIs������(.Row)) Then
                lng������ĿID = Get������ĿID(Val(.TextMatrix(.Row, COL_ID)), True)
            Else
                lng������ĿID = Get������ĿID(Val(.TextMatrix(.Row, COL_ID)), False)
            End If
        End If
    End With
    Call ShowClinicHelp(0, mfrmParent, lng������ĿID)
End Sub

Public Sub zlPrintSetup()
    Call zlPrintSet
End Sub

Public Sub zlExcel()
    Call OutputList(3)
End Sub

Public Sub zlPreview()
    Call OutputList(2)
End Sub

Public Sub zlPrint()
    Call OutputList(1)
End Sub

Private Sub Form_Activate()
    If vsColumn.Visible Then
        vsColumn.SetFocus '��ѡ����
    Else
        picFocus.SetFocus '�������ú󱾴����ڵĽ���˳�����Ч
        vsAdvice.SetFocus
    End If
End Sub

Private Sub Form_Deactivate()
    vsColumn.Visible = False '��ѡ����
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objMenu As Object
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    
    If KeyCode = vbKeyEscape Then '��ѡ����
        If vsColumn.Visible Then
            vsColumn.Visible = False
            vsAdvice.SetFocus
        End If
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then
        Call imgColSel_MouseUp(1, 0, 0, 0)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu�¿�ҽ��)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyM Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnu�޸�ҽ��)
    ElseIf KeyCode = vbKeyDelete Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnuɾ��ҽ��)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        Set objMenu = mfrmParent.mnuAdviceFunc(mnuҽ��ֹͣ)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyZ Then
        Set objMenu = mfrmParent.mnuAdviceFuncRoll(0)
    ElseIf KeyCode = vbKeyF5 Then
        Call LoadAdvice
    ElseIf KeyCode = vbKeyF6 Then
        Call zlItemRef
    ElseIf KeyCode = vbKeyF9 Then
        If mfrmParent.mnuViewAdviceFilter.Visible And mfrmParent.mnuViewAdviceFilter.Enabled Then
            Call mfrmParent.mnuViewAdviceFiler_Click
        End If
    ElseIf KeyCode = vbKeyF8 Then
        Call mfrmParent.mnuViewAdviceCyc_Click
    End If
    
    If Not objMenu Is Nothing Then
        If objMenu.Enabled And objMenu.Visible Then
            Call zlMenuClick(objMenu)
        End If
    End If
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsAppend.Height - y < 60 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + y
        vsAdvice.Height = vsAdvice.Height + y
        tabAppend.Top = tabAppend.Top + y
        vsAppend.Top = vsAppend.Top + y
        vsAppend.Height = vsAppend.Height - y
        Me.Refresh
    End If
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsColumn
            If .Visible Then
                .Visible = False
                vsAdvice.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vsAdvice.ColHidden(.RowData(i)) Or vsAdvice.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub mfrmEdit_Unload(Cancel As Integer)
    If Not Cancel Then
        If frmInAdviceEdit.mblnOK Then Call LoadAdvice
        Set mfrmEdit = Nothing
        
        '��PACS���ü���
        On Error Resume Next
        
        If mfrmParent.tabFunc.SelectedItem.Key = "ҽ��" Then
            Call BringWindowToTop(Me.Hwnd)
        End If
    End If
End Sub

Private Function CheckWindow() As Boolean
'���ܣ����ҽ���༭�����Ƿ��Ѿ���
    If Not mfrmEdit Is Nothing Then
        '��ǰ���ڴ���
        MsgBox "ҽ���༭�����Ѿ��򿪣�������ɵ�ǰ��������ִ�С�", vbInformation, gstrSysName
        '��λ����ǰ�Ĵ���
        If mfrmEdit.WindowState = vbMinimized Then mfrmEdit.WindowState = vbNormal
        If mfrmEdit.Visible Then mfrmEdit.SetFocus
        Exit Function
    Else
        '�������ڴ���
        If Not CheckAdviceWindow("סԺҽ���༭") Then Exit Function
    End If
    CheckWindow = True
End Function

Private Sub FuncAdviceDel()
'ɾ����ɾ����ǰҽ��
'˵������������ɾ��,�Լ�����,�������,��ҩ�䷽,������ɾ��,һ����ҩֻɾ����ǰҩƷ
    Dim strSQL As String, lngҽ��ID As Long
    Dim blnGroup As Boolean, i As Long
    Dim lngRow As Long
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ������ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If

        If mblnMoved Then
            MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '����Ƿ����ɾ��
        If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
            MsgBox "�㲻��ɾ����ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If InStr(",1,2,", .TextMatrix(.Row, COL_ҽ��״̬)) = 0 Then
            MsgBox "��ǰѡ���ҽ���Ѿ���У�ԣ�����ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ǩ����ҽ������ɾ��
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ�ǩ��������ɾ��������ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If

        '��ִҵ�ʸ��ҽ��ֻ��ɾ���޸�δ��˵�ҽ����
        If Not HaveAuditPriv And HaveAuditPriv(GetAuditName(.TextMatrix(.Row, COL_����ҽ��))) Then
            MsgBox "��û���ʸ�ɾ����ǰѡ���ҽ�������ߵ�ǰѡ���ҽ���Ѿ�����ˣ�����ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If

        If InStr(",5,6,", .TextMatrix(.Row, COL_�������)) > 0 Then
            If .Row - 1 >= .FixedRows Then
                If Val(.TextMatrix(.Row - 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then blnGroup = True
            End If
            If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                If Val(.TextMatrix(.Row + 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then blnGroup = True
            End If
            If blnGroup Then
                If MsgBox("ҽ��""" & .TextMatrix(.Row, COL_ҽ������) & """������ҩƷһ����ҩ,ȷʵҪɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        If Not blnGroup Then
            If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, COL_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        strSQL = "ZL_����ҽ����¼_Delete(" & lngҽ��ID & ",1)"
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    With vsAdvice
        '������ֱ��ɾ��
        .Redraw = False
        
        'ɾ��һ����ҩ��һ��ʱ����ʾ����
        If blnGroup And .Row + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(.Row, COL_���ID)) = Val(.TextMatrix(.Row + 1, COL_���ID)) Then
                If .TextMatrix(.Row, COL_��ʼʱ��) <> "" And .TextMatrix(.Row + 1, COL_��ʼʱ��) = "" Then
                    .TextMatrix(.Row + 1, COL_��Ч) = .TextMatrix(.Row, COL_��Ч)
                    .TextMatrix(.Row + 1, COL_��ʼʱ��) = .TextMatrix(.Row, COL_��ʼʱ��)
                    .TextMatrix(.Row + 1, COL_Ƶ��) = .TextMatrix(.Row, COL_Ƶ��)
                    .TextMatrix(.Row + 1, COL_�÷�) = .TextMatrix(.Row, COL_�÷�)
                End If
            End If
        End If
        
        lngRow = .Row
        .RemoveItem .Row
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        
        Call .ShowCell(.Row, .Col)
        .Redraw = True
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col) '��ɫ���������
    End With
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSign()
'���ܣ���ҽ�����е���ǩ��
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lngǩ��ID As Long, lng֤��ID As Long
    Dim intRule As Integer
    
    If mlng����ID = 0 Then Exit Sub
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjESign Is Nothing Then Exit Sub
    
    '��ȡǩ��ҽ��Դ��
    intRule = ReadAdviceSignSource(1, mlng����ID, mlng��ҳID, strIDs, 0, mblnMoved, strSource, mlngǰ��ID)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "�ò���Ŀǰû�п���ǩ����ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID)
    If strSign <> "" Then
        lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
        strSQL = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        Call LoadAdvice 'ˢ�½���
        MsgBox "����ɵ���ǩ����", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignVerify()
'���ܣ�У��ҽ���ĵ���ǩ��(�ɶ���ת�Ƶ�����)
    Dim strSource As String
    
    If mlng����ID = 0 Then Exit Sub
    If gobjESign Is Nothing Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tabAppend.SelectedItem.Index <> 2 Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "��ǰѡ���ҽ��û��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ȡǩ��ҽ��Դ��
        If ReadAdviceSignSource(.Cell(flexcpData, .Row, 0), 0, 0, "", .RowData(.Row), mblnMoved, strSource) = 0 Then Exit Sub
        
        '��֤ǩ��
        Call gobjESign.VerifySignature(strSource, .RowData(.Row), 1)
    End With
End Sub

Private Sub FuncAdviceSignErase()
'���ܣ�ȡ��ҽ���ĵ���ǩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If mlng����ID = 0 Then Exit Sub
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjESign Is Nothing Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tabAppend.SelectedItem.Index <> 2 Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "��ǰѡ���ҽ��û��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '���Ϻ�ֹͣҽ����ǩ������ȡ��
        If InStr(",4,8,", .Cell(flexcpData, .Row, 0)) > 0 Then
            MsgBox "����ֱ��ȡ�����ϻ�ֹͣҽ����ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        '�¿�ǩ�����������¿���У������״̬
        If .Cell(flexcpData, .Row, 0) = 1 Then
            If InStr(",1,2,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬))) = 0 Then
                MsgBox "����ҽ���Ѿ�����У�ԣ���ǩ������ȡ����", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '����ȡ��ҽ���´��ǩ��
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) <> 0 Then
            MsgBox "�㲻��ȡ��ҽ�������´�ҽ����ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        'ֻ��ȡ������ǩ����
        If .TextMatrix(.Row, 2) <> UserInfo.���� Then
            MsgBox "��ǩ���˲����㱾�ˣ�����ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("ȷʵҪȡ�����ǩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        
        strSQL = "zl_ҽ��ǩ����¼_Delete(" & .RowData(.Row) & ")"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End With
    
    Call LoadAdvice 'ˢ�½���
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function IsConsultation() As Boolean
'���ܣ��жϵ�ǰִ�й��ܵĲ����Ƿ���ﲡ��
    On Error Resume Next
    IsConsultation = mfrmParent.tabPati.SelectedItem.Key = "���ﲡ��"
    Err.Clear: On Error GoTo 0
End Function

Private Sub FuncAdviceAudit()
'���ܣ����ҽ��
    If Not CheckWindow Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
    
    If Not HaveAuditPriv Then
        MsgBox "�㲻�������ҽ�����ʸ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng����ID, mlng��ҳID, mlngǰ��ID, , , , , , IsConsultation, True)
End Sub

Private Sub FuncAdviceAdd()
'���ܣ������µ�ҽ��
    If Not CheckWindow Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
        
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng����ID, mlng��ҳID, mlngǰ��ID, , , , , , IsConsultation)
End Sub

Private Sub FuncAdviceModi()
'���ܣ��޸ĵ�ǰҽ��
    Dim lngҽ��ID As Long
    
    If Not CheckWindow Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        If lngҽ��ID = 0 Then Exit Sub
        
        If mblnMoved Then
            MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        'ҽ���´��ҽ��
        If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
            MsgBox "�㲻���޸ĸ�ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��У�Ի��ѷ�ֹ
        If InStr(",4,8,9,", .TextMatrix(.Row, COL_ҽ��״̬)) > 0 Then
            MsgBox "��ǰѡ���ҽ���Ѿ����ϻ�ֹͣ�������޸ġ�", vbInformation, gstrSysName
            Exit Sub
        ElseIf InStr(",1,2,", .TextMatrix(.Row, COL_ҽ��״̬)) = 0 Then
            MsgBox "��ǰѡ���ҽ���Ѿ���У�ԣ������޸ġ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ǩ����ҽ�������޸�
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ�ǩ���������޸ġ�����ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ִҵ�ʸ��ҽ��ֻ��ɾ���޸�δ��˵�ҽ����
        If Not HaveAuditPriv And HaveAuditPriv(GetAuditName(.TextMatrix(.Row, COL_����ҽ��))) Then
            MsgBox "��û���ʸ��޸ĵ�ǰѡ���ҽ�������ߵ�ǰѡ���ҽ���Ѿ�����ˣ������޸ġ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Set mfrmEdit = frmInAdviceEdit
        Call frmInAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng����ID, mlng��ҳID, mlngǰ��ID, , , Val(.TextMatrix(.Row, COL_Ӥ��ID)), lngҽ��ID, , IsConsultation)
    End With
End Sub

Private Sub FuncAdviceRevoke()
'���ܣ�ҽ������
    If mlng����ID = 0 Then Exit Sub
    
    If mlngǰ��ID = 0 Then '����ҽ��վ
        If mblnMoved Then
            MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        frmAdviceOperate.mstrPrivs = mstrPrivs
        frmAdviceOperate.mint���� = 0
        frmAdviceOperate.mlng����ID = mlng����ID
        frmAdviceOperate.mlng����ID = mlng����ID
        frmAdviceOperate.mlng��ҳID = mlng��ҳID
        frmAdviceOperate.mlngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
        frmAdviceOperate.Show 1, Me
        If frmAdviceOperate.mblnOK Then Call LoadAdvice
    Else
        If FuncAdviceRevoke0(2) Then Call LoadAdvice
    End If
End Sub

Private Function FuncAdviceRevoke0(ByVal int������Դ As Integer) As Boolean
'ɾ������ǰҽ������(һ��ҽ������)
    Dim strSQL As String, lngҽ��ID As Long
    
    Dim strҽ��ID As String, intRule As Integer
    Dim lngǩ��ID As Long, lng֤��ID As Long
    Dim strSource As String, strSign As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ���������ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If mblnMoved Then
            MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '����Ƿ��������
        If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
            MsgBox "��ҽ�������������", vbInformation, gstrSysName
            Exit Function
        End If
        
        If int������Դ = 1 Then
            If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 8 Then
                MsgBox "��ǰѡ�������ҽ����δ���ͻ��Ѿ����ϡ�", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If InStr(",1,2,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) > 0 Then
                MsgBox "��ǰѡ���סԺҽ����δУ�ԣ���ֱ��ɾ����", vbInformation, gstrSysName
                Exit Function
            End If
            If InStr(",4,8,9,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) > 0 Then
                MsgBox "��ǰѡ���סԺҽ���Ѿ����ϻ�ֹͣ��", vbInformation, gstrSysName
                Exit Function
            End If
            If .TextMatrix(.Row, COL_�ϴ�ִ��) <> "" Then
                MsgBox "��ǰѡ���סԺҽ���Ѿ����ͣ����������ϡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '����ǩ��������ʾ
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ��������ϡ�", vbInformation, gstrSysName
                Else
                    MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ���������ϡ�", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
        End If
        
        If RowInһ����ҩ(.Row, 0, 0) Then
            If MsgBox("����һ����ҩ��ҽ������һ�����ϣ�ȷʵҪ������" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("ȷʵҪ����ҽ��""" & .TextMatrix(.Row, COL_ҽ������) & """��" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        strSQL = "ZL_����ҽ����¼_����(" & lngҽ��ID & ")"
        
        '����ʱ�ĵ���ǩ��
        '----------------------------------------------------------------------------------
        If strSign <> "" Then
            '��ȡǩ��ҽ��Դ��
            strҽ��ID = lngҽ��ID '��ID,����Ϊ��ϸID
            intRule = ReadAdviceSignSource(4, mlng����ID, mlng��ҳID, strҽ��ID, 0, mblnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "���ܶ�ȡ��Ҫ���ϵ���ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID)
            If strSign <> "" Then
                lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
                strSign = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strҽ��ID & "')"
            Else
                Exit Function
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    FuncAdviceRevoke0 = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceStop()
'���ܣ�ֹͣҽ��
    If mlng����ID = 0 Then Exit Sub
    
    If mlngǰ��ID = 0 Then '����ҽ��վ
        If mblnMoved Then
            MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Sub
        End If
                
        frmAdviceOperate.mstrPrivs = mstrPrivs
        frmAdviceOperate.mint���� = 1
        frmAdviceOperate.mlng����ID = mlng����ID
        frmAdviceOperate.mlng����ID = mlng����ID
        frmAdviceOperate.mlng��ҳID = mlng��ҳID
        frmAdviceOperate.mlngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
        frmAdviceOperate.Show 1, Me
        If frmAdviceOperate.mblnOK Then Call LoadAdvice
    Else
        If FuncAdviceStop0() Then Call LoadAdvice
    End If
End Sub

Private Sub FuncAdvicePause()
'���ܣ���ͣҽ��
    If mlng����ID = 0 Then Exit Sub
    
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    frmAdviceOperate.mstrPrivs = mstrPrivs
    frmAdviceOperate.mint���� = 5
    frmAdviceOperate.mlng����ID = mlng����ID
    frmAdviceOperate.mlng����ID = mlng����ID
    frmAdviceOperate.mlng��ҳID = mlng��ҳID
    frmAdviceOperate.mlngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    frmAdviceOperate.mbln��ʿվ = True
    frmAdviceOperate.Show 1, Me
    If frmAdviceOperate.mblnOK Then Call LoadAdvice
End Sub

Private Sub FuncAdviceResume()
'���ܣ�����ҽ��
    If mlng����ID = 0 Then Exit Sub
    
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    frmAdviceOperate.mstrPrivs = mstrPrivs
    frmAdviceOperate.mint���� = 6
    frmAdviceOperate.mlng����ID = mlng����ID
    frmAdviceOperate.mlng����ID = mlng����ID
    frmAdviceOperate.mlng��ҳID = mlng��ҳID
    frmAdviceOperate.mlngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    frmAdviceOperate.mbln��ʿվ = True
    frmAdviceOperate.Show 1, Me
    If frmAdviceOperate.mblnOK Then Call LoadAdvice
End Sub

Private Function FuncAdviceStop0() As Boolean
'ɾ������ǰҽ��ֹͣ(������סԺ����)
    Dim strSQL As String, lngҽ��ID As Long
    Dim strStopTime As String
    
    Dim strҽ��ID As String, intRule As Integer
    Dim lngǩ��ID As Long, lng֤��ID As Long
    Dim strSource As String, strSign As String
    Dim colStopTime As New Collection
    
    With vsAdvice
        '����Ƿ��������
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ������ֹͣ��", vbInformation, gstrSysName
            Exit Function
        End If
                        
        If mblnMoved Then
            MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
                "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
            Exit Function
        End If
                        
        If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
            MsgBox "��ҽ�������������", vbInformation, gstrSysName
            Exit Function
        End If
                        
        '���
        If .TextMatrix(.Row, COL_��Ч) <> "����" Then
            MsgBox "��ǰѡ���ҽ������סԺ����ҽ����", vbInformation, gstrSysName
            Exit Function
        End If
        If .TextMatrix(.Row, COL_����) <> "" Then
            MsgBox "��ҩ�䷽�ڷ��ͺ���Զ�ֹͣ��", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",1,2,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) > 0 Then
            MsgBox "��ǰѡ���סԺҽ����δУ�ԣ���ֱ��ɾ����", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",4,8,9,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) > 0 Then
            MsgBox "��ǰѡ���סԺҽ���Ѿ����ϻ�ֹͣ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        '����ǩ��������ʾ
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "ֹͣ��ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ�����ֹͣ��", vbInformation, gstrSysName
                Else
                    MsgBox "ֹͣ��ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ������ֹͣ��", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
        End If
        
        If RowInһ����ҩ(.Row, 0, 0) Then
            If MsgBox("����һ����ҩ��ҽ������һ��ֹͣ��ȷʵҪֹͣ��" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("ȷʵҪֹͣҽ��""" & .TextMatrix(.Row, COL_ҽ������) & """��" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        'ͣ��ʱȱʡ��ҽ����ֹʱ��
        If .TextMatrix(.Row, COL_��ֹʱ��) = "" Then
            If gbln����ҽ��������Ч Then
                strStopTime = Format(zlDatabase.Currentdate + 1, "yyyy-MM-dd 00:00")
            Else
                strStopTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            End If
        Else
            strStopTime = .TextMatrix(.Row, COL_��ֹʱ��)
        End If
        strSQL = "ZL_����ҽ����¼_ֹͣ(" & lngҽ��ID & ",To_Date('" & strStopTime & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.���� & "')"
        
        'ֹͣʱ�ĵ���ǩ��
        '----------------------------------------------------------------------------------
        If strSign <> "" Then
            '��ȡǩ��ҽ��Դ��
            strҽ��ID = lngҽ��ID '��ID,����Ϊ��ϸID
            colStopTime.Add Format(strStopTime, "yyyy-MM-dd HH:mm:00"), "_" & lngҽ��ID
            intRule = ReadAdviceSignSource(8, mlng����ID, mlng��ҳID, strҽ��ID, 0, mblnMoved, strSource, , colStopTime)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "���ܶ�ȡ��Ҫֹͣ����ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID)
            If strSign <> "" Then
                lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
                strSign = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & ",8," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strҽ��ID & "')"
            Else
                Exit Function
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    FuncAdviceStop0 = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceSupply()
'���ܣ���¼ҽ��
    If Not CheckWindow Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
    
    If mblnMoved Then
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mstrPrivs, mlng����ID, mlng��ҳID, mlngǰ��ID, , True, , , , IsConsultation)
End Sub

Private Sub tabAppend_Click()
    If Val(vsAppend.Tag) = tabAppend.SelectedItem.Index Then Exit Sub
    
    If Visible Then
        Call SaveFlexState(vsAppend, App.ProductName & "\" & Me.Name)
    End If
        
    vsAppend.Tag = tabAppend.SelectedItem.Index
    If tabAppend.SelectedItem.Index = 1 Then
        Call InitSendTable
    ElseIf tabAppend.SelectedItem.Index = 2 Then
        Call InitSignTable
    End If
    
    If Visible Then
        Call RestoreFlexState(vsAppend, App.ProductName & "\" & Me.Name)
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    If Visible Then vsAdvice.SetFocus
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next 'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��

    If NewRow = OldRow Then Exit Sub
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_��ʼʱ��)
    End If
    If vsAdvice.Redraw <> flexRDNone Then
        If Val(vsAdvice.TextMatrix(NewRow, COL_ID)) <> 0 Then
            '��ʾҽ�����ӱ�������
            If mfrmParent.mnuViewAdviceAppend.Checked Then
                If tabAppend.SelectedItem.Index = 1 Then
                    Call ShowSendList(NewRow)
                ElseIf tabAppend.SelectedItem.Index = 2 Then
                    Call ShowSignList(NewRow)
                End If
            End If
            '��ʾҽ���ɻ�������
            Call ShowRollList(NewRow)
        ElseIf mfrmParent.mnuViewAdviceAppend.Checked Then
            Call ClearAppendData
            vsAppend.Row = vsAppend.FixedRows
        End If
        
        Call SetFuncEnabled
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    If Col = COL_ҽ������ Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_Ƥ�� Then
            Cancel = True
        ElseIf Col = COL_��ʾ Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, SysColor2RGB(.BackColorFixed)

            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '����һ����ҩ������еı��߼�����
            lngLeft = COL_��Ч: lngRight = COL_��ʼʱ��
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_Ƶ��: lngRight = COL_�÷�
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '������߱����
            vRect.Right = Right - 1 '�����ұ߱����
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '���б����±���
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                'Ϊ��֧��Ԥ�����
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Ϊ���ⲿϵͳ�������ӣ�By����ͮ��
    On Error Resume Next
    If Button = 2 And mfrmParent.mnuAdvice.Visible Then PopupMenu mfrmParent.mnuAdvice, 2
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    
    If mlng����ID = 0 Then Exit Sub
    
    '��ͷ
    objOut.Title.Text = "����ҽ���嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With mfrmParent.lvwPati.SelectedItem
        Set objRow = New zlTabAppRow
        objRow.Add "���ˣ�" & .Text & " �Ա�" & .SubItems(4) & " ���䣺" & .SubItems(5)
        objRow.Add "סԺ�ţ�" & .SubItems(1) & " ���ţ�" & .SubItems(2)
        objOut.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "��Ժ���ڣ�" & .SubItems(8)
        objRow.Add "��Ժ���ڣ�" & .SubItems(9)
        objOut.UnderAppRows.Add objRow
    End With
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = vsAdvice
    
    '���
    vsAdvice.Redraw = False
    lngRow = vsAdvice.Row: lngCol = vsAdvice.Col
        
    strWidth = ""
    For i = 0 To vsAdvice.FixedCols - 1
        strWidth = strWidth & "," & vsAdvice.ColWidth(i)
        vsAdvice.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsAdvice.FixedCols - 1
        vsAdvice.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    vsAdvice.Row = lngRow: vsAdvice.Col = lngCol
    vsAdvice.Redraw = True
End Sub

Private Sub Form_Load()
    Call InitAdviceTable
    Call InitColumnSelect '��ѡ����
    Call tabAppend_Click
    Call RestoreWinState(Me, App.ProductName)
    
    On Error Resume Next
    fraAdviceUD.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    tabAppend.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    vsAppend.Visible = mfrmParent.mnuViewAdviceAppend.Checked
    Err.Clear: On Error GoTo 0
    '����ǩ����¼
    If gobjESign Is Nothing Then tabAppend.Visible = False
    
    Set mfrmEdit = Nothing
    Call InitSysPar '��ʼ��ϵͳ����
End Sub

Private Sub Form_Resize()
    Dim PriceH As Long
    
    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    
    PriceH = IIF(vsAppend.Visible, vsAppend.Height + fraAdviceUD.Height + IIF(tabAppend.Visible, tabAppend.Height, 0), 0)
    
    vsAdvice.Left = 0
    vsAdvice.Top = 0
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - PriceH
    
    '��ѡ����
    With vsAdvice
        fraColSel.Left = .Left + (.ColWidth(0) + .ColWidth(1) - fraColSel.Width) / 2 + 30
        fraColSel.Top = .Top + (.RowHeight(0) - fraColSel.Height) / 2 + 30
    End With
    
    fraAdviceUD.Left = 0
    fraAdviceUD.Top = vsAdvice.Top + vsAdvice.Height
    fraAdviceUD.Width = Me.ScaleWidth
    
    tabAppend.Left = 0
    tabAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tabAppend.Width = Me.ScaleWidth
    
    vsAppend.Left = 0
    If tabAppend.Visible Then
        vsAppend.Top = tabAppend.Top + tabAppend.Height
    Else
        vsAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    End If
    vsAppend.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmEdit = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub GetAdviceWhere(intӤ�� As Integer, str��Ч As String, str״̬ As String, bln���� As Boolean)
'���ܣ���ȡҽ��������������,����ҽ����¼�ı����Ϊ"A"
'������bln����=�����Ƿ�ֻ��ʾ����
    Dim strWhere As String, strReg As String
    Dim strTmp As String, i As Long
    
    intӤ�� = -1: str��Ч = "": str״̬ = "": bln���� = False
    
    'Ӥ������
    strReg = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\סԺҽ������", "����Ӥ��", "1")
    If Val(strReg) = 1 Then
        intӤ�� = 0
    ElseIf Val(strReg) > 1 Then
        intӤ�� = Val(strReg) - 1
    End If
    
    'ҽ����Ч
    strReg = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\סԺҽ������", "ҽ����Ч", "11")
    If strReg <> "11" Then
        strTmp = ""
        For i = 1 To Len(strReg)
            If Val(Mid(strReg, i, 1)) = 1 Then
                strTmp = strTmp & "," & i - 1
            End If
        Next
        If strTmp <> "" Then str��Ч = Mid(strTmp, 2)
        If strReg = "01" Then bln���� = True
    End If
            
    'ҽ��״̬
    strReg = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\סԺҽ������", "ҽ��״̬", "111")
    If strReg <> "111" Then
        strTmp = ""
        For i = 1 To Len(strReg)
            If Val(Mid(strReg, i, 1)) = 1 Then
                If i = 1 Then strTmp = strTmp & ",1,2"  'δУ��
                If i = 2 Then strTmp = strTmp & ",3,5,6,7"  '��У��
                If i = 3 Then strTmp = strTmp & ",4,8,9"  '�ѷ�ֹ
            End If
        Next
        If strTmp <> "" Then str״̬ = Mid(strTmp, 2)
    End If
End Sub

Private Sub ClearAdviceData()
'���ܣ����ҽ���嵥����
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Editable = flexEDNone
End Sub

Private Sub InitColumnSelect()
'���ܣ�����ҽ���嵥ԭʼ����ʾ״̬��ʼ����ѡ����
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vsAdvice
        For i = .FixedCols To .Cols - 1
            If Not (.ColHidden(i) Or .ColWidth(i) = 0) Then
                If .TextMatrix(0, i) <> "" Then '�����,Ƥ��
                    vsColumn.Rows = vsColumn.Rows + 1
                    lngRow = vsColumn.Rows - 1
                    vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                    vsColumn.RowData(lngRow) = i
                    
                    '�̶���ʾ��
                    If InStr(",��ʼʱ��,ҽ������,����ҽ��,", "," & .TextMatrix(0, i) & ",") > 0 Then
                        vsColumn.TextMatrix(lngRow, 0) = 1
                        vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                    End If
                End If
            End If
        Next
    End With
    vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 130
    vsColumn.Row = 1
End Sub

Private Sub InitAdviceTable()
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ID;���ID;��ID;���;Ӥ��ID;ҽ��״̬;�������;��������;�������;��־;" & _
        ",240,4;��Ч,500,4;��ʼʱ��,1080,1;ҽ������,3000,1;,375,4;" & _
        "����,850,1;����,850,1;Ƶ��,1000,1;�÷�,1000,1;ҽ������,1000,1;ִ��ʱ��,1000,1;" & _
        "��ֹʱ��,1560,1;ִ�п���,850,1;ִ������,850,1;�ϴ�ִ��,1560,1;״̬,500,4;" & _
        "����ҽ��,850,1;����ʱ��,1080,1;У�Ի�ʿ,850,1;У��ʱ��,1080,1;ͣ��ҽ��,850,1;" & _
        "ͣ��ʱ��,1080,1;ͣ����ʿ,850,1;ȷ��ͣ��ʱ��,1180,1;����ID;������;������;����ID;ǰ��ID;ǩ����"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 2
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i) '��¼ԭʼ�п�������ѡ����
        Next
        .ColHidden(COL_��ʾ) = Not (gblnPass And InStr(mstrPrivs, "������ҩ���") > 0) 'Pass
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(0) = 9 * Screen.TwipsPerPixelX
        .ColWidth(1) = 11 * Screen.TwipsPerPixelX
    End With
End Sub

Private Sub ClearAppendData()
'���ܣ�������ӱ������
    vsAppend.Rows = vsAppend.FixedRows
    vsAppend.Rows = vsAppend.FixedRows + 1
End Sub

Private Sub InitSendTable()
'���ܣ���ʼ�������嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "���ͺ�;����ʱ��,1080,1;����ҽ��,1800,1;���ݺ�,850,1;�շ���Ŀ,1800,1;��������,850,1;�Ʒ�״̬,850,1;ִ��״̬,850,1;ִ�п���,850,1;�״�ʱ��,1080,1;ĩ��ʱ��,1080,1;������,800,1;��¼����"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = True
        .MergeCol(1) = True
    End With
End Sub

Private Sub InitSignTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "ǩ������,1150,1;ǩ��ʱ��,1900,1;ǩ����,800,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCol(0) = False
        .MergeCol(1) = False
    End With
End Sub

Private Function LoadAdvice() As Boolean
'���ܣ����ݵ�ǰ�������ö�ȡ����ʾҽ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim lngTop As Long, i As Long, j As Long
    Dim strFormat As String, strTmp As String
    Dim bln��ҩ;�� As Boolean, bln��ҩ�÷� As Boolean
    Dim bln�ɼ����� As Boolean, bln������ As Boolean, bln������ As Boolean
    Dim str״̬ As String, lngҽ��ID As Long
    Dim blnFirst As Boolean, strBill As String
    Dim strҽ����Ч As String, strҽ��״̬ As String
    Dim intӤ�� As Integer, bln���� As Boolean
    Dim bln���� As Boolean, dat���� As Date
    Dim blnDo As Boolean, strCurr As String, strTime As String
    
    If mlng����ID = 0 Then Exit Function
    
    Screen.MousePointer = 11
        
    On Error GoTo errH
    
    lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) '��¼��ǰ��
    
    'ҽ����������
    Call GetAdviceWhere(intӤ��, strҽ����Ч, strҽ��״̬, bln����)
    strWhere = ""
    If intӤ�� <> -1 Then
        strWhere = strWhere & " And Nvl(A.Ӥ��,0)=[4]"
    End If
    If strҽ����Ч <> "" Then
        strWhere = strWhere & " And Instr([5],','||Nvl(A.ҽ����Ч,0)||',')>0"
    End If
    If strҽ��״̬ <> "" Then
        strWhere = strWhere & " And Instr([6],','||Nvl(A.ҽ��״̬,0)||',')>0"
    End If
    
    bln���� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\סԺҽ������", "����ҽ��", "0")) <> 0
    If bln���� Then
        '�����������ʱ��
        strSQL = "Select Max(B.����ʱ��) as ʱ�� From ����ҽ����¼ A,����ҽ��״̬ B" & _
            " Where A.ID=B.ҽ��ID And B.��������=5 And A.����ID=[1] And A.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!ʱ��) Then
                dat���� = Format(rsTmp!ʱ��, "yyyy-MM-dd HH:mm:ss")
            End If
        End If
        If dat���� = CDate(0) Then
            bln���� = False
        Else
            strSQL = _
                " Select Distinct" & _
                "   A.ID,A.���ID,A.����ID,A.��ҳID,A.���,A.Ӥ��,A.ҽ��״̬,A.ҽ����Ч,A.�������," & _
                "   A.������־,A.�����,A.��ʼִ��ʱ��,A.ҽ������,A.Ƥ�Խ��,A.�ܸ�����,A.��������," & _
                "   A.ִ��Ƶ��,A.ҽ������,A.ִ��ʱ�䷽��,A.ִ����ֹʱ��,A.ִ������,A.�ϴ�ִ��ʱ��," & _
                "   A.����ҽ��,A.����ʱ��,A.У�Ի�ʿ,A.У��ʱ��,A.ͣ��ҽ��,A.ͣ��ʱ��,A.ȷ��ͣ��ʱ��," & _
                "   A.������ĿID,A.ִ�п���ID,A.�շ�ϸĿID,A.����ID,A.ǰ��ID,A.������Դ" & _
                " From ����ҽ����¼ A,����ҽ��״̬ B" & _
                " Where A.ID=B.ҽ��ID And (B.����ʱ��>=[3] Or A.ҽ��״̬ IN(1,2))" & _
                " And A.����ID=[1] And A.��ҳID=[2]" & strWhere
            strSQL = "(" & strSQL & ")"
        End If
    End If
    
    '���Ƶ��ݣ���Ӧ���Ƶ���,��������,������
    strBill = "Select A.ID as ҽ��ID,B.�����ļ�ID as ����ID," & _
        " Max(Decode(C.��дʱ��,1,1,0)) as ������," & _
        " Max(Decode(C.��дʱ��,2,1,0)) as ������" & _
        " From ����ҽ����¼ A,���Ƶ���Ӧ�� B,�����ļ���� C" & _
        " Where A.������ĿID=B.������ĿID And B.Ӧ�ó���=2 And B.�����ļ�ID=C.�����ļ�ID(+)" & _
        " And A.����ID=[1] And A.��ҳID=[2]" & strWhere & _
        " And Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL)" & _
        " Group by A.ID,B.�����ļ�ID"
        
    'ҽ����¼��������������,��������,��鲿λ,��ҩ�巨
    str״̬ = "Decode(A.ҽ��״̬,1,'�¿�',2,'����',3,'У��',4,'����',5,'����',6,'��ͣ',7,'����',8,'ֹͣ',9,'ȷ��ֹͣ')"
    strSQL = _
        "Select /*+ RULE */ A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���," & _
            " Nvl(A.Ӥ��,0) as Ӥ��ID,A.ҽ��״̬,Nvl(A.�������,'*') as �������,B.��������,C.�������,A.������־ as ��־," & _
            " A.�����,Decode(Nvl(A.ҽ����Ч,0),0,'����','����') as ��Ч," & _
            " To_Char(A.��ʼִ��ʱ��,'MM-DD HH24:MI') as ��ʼʱ��,A.ҽ������,A.Ƥ�Խ�� as Ƥ��," & _
            " Decode(A.�ܸ�����,NULL,NULL,Decode(A.�������,'E',Decode(B.��������,'4',A.�ܸ�����||'��',A.�ܸ�����||B.���㵥λ),'5',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,'6',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,A.�ܸ�����||B.���㵥λ)) as ����," & _
            " Decode(A.��������,NULL,NULL,A.��������||B.���㵥λ) as ����," & _
            " A.ִ��Ƶ�� as Ƶ��,Decode(A.�������,'E',Decode(Instr('246',Nvl(B.��������,'0')),0,NULL,B.����),NULL) as �÷�," & _
            " A.ҽ������,A.ִ��ʱ�䷽�� as ִ��ʱ��,To_Char(A.ִ����ֹʱ��,'YYYY-MM-DD HH24:MI') as ��ֹʱ��," & _
            " Nvl(E.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'<Ժ��ִ��>')) as ִ�п���," & _
            " Decode(Instr('567E',Nvl(A.�������,'*')),0,NULL,A.ִ������) as ִ������," & _
            " To_Char(A.�ϴ�ִ��ʱ��,'YYYY-MM-DD HH24:MI') as �ϴ�ִ��," & str״̬ & " as ״̬," & _
            " A.����ҽ��,To_Char(A.����ʱ��,'MM-DD HH24:MI') as ����ʱ��,A.У�Ի�ʿ,To_Char(A.У��ʱ��,'MM-DD HH24:MI') as У��ʱ��," & _
            " A.ͣ��ҽ��,To_Char(A.ͣ��ʱ��,'MM-DD HH24:MI') as ͣ��ʱ��,F.������Ա as ͣ����ʿ," & _
            " To_Char(A.ȷ��ͣ��ʱ��,'MM-DD HH24:MI') as ȷ��ͣ��ʱ��," & _
            " Y.����ID,Y.������,Y.������,A.����ID,A.ǰ��ID,Decode(S.ǩ��ID,NULL,0,1) as ǩ����" & _
        " From " & IIF(bln����, strSQL, "����ҽ����¼") & " A,���ű� E,ҩƷ���� C,ҩƷ��� D,������ĿĿ¼ B," & _
            " ����ҽ��״̬ F,����ҽ��״̬ S,����ҽ����¼ X,(" & strBill & ") Y" & _
        " Where A.������ĿID=B.ID(+) And A.ִ�п���ID=E.ID(+) And A.������ĿID=C.ҩ��ID(+)" & _
            " And A.�շ�ϸĿID=D.ҩƷID(+) And A.���ID=X.ID(+) And A.ID=Y.ҽ��ID(+)" & _
            " And Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL)" & _
            " And A.ID=F.ҽ��ID(+) And F.��������(+)=9 And A.ID=S.ҽ��ID And S.��������=1" & _
            " And A.����ID=[1] And A.��ҳID=[2] And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & strWhere & _
            IIF(mlngǰ��ID = 0 Or mblnShowAll, "", " And A.ǰ��ID=[7]") & _
        " Order by Nvl(A.Ӥ��,0),���,��ID,A.���"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mlng��ҳID, dat����, intӤ��, "," & strҽ����Ч & ",", "," & strҽ��״̬ & ",", mlngǰ��ID)
    
    If Not rsTmp.EOF Then
        strCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        With vsAdvice
            .Redraw = False
            
            '��ʱ�����ʱ��FormatString�ָ�һЩȱʡֵ(�̶����������̶��������ּ����ж���,�ߴ�,�ɼ�)
            'FormatString������ʱ��ֵ��Ч
            '���AutoResize=True,�������п���и߱��Զ�����(����AutoSizeMode)
            '���WordWrap=True,���и߻ᱻ�Զ�����
            .WordWrap = False
            strFormat = GetColFormat(vsAdvice)
            Call ClearAdviceData
            .ScrollBars = flexScrollBarNone
            Set .DataSource = rsTmp
            .ScrollBars = flexScrollBarBoth
            If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
                gcnOracle.Errors.Clear '��,��ʱ�̶��д˴���
            End If
            Call SetColFormat(vsAdvice, strFormat)
            .TextMatrix(0, COL_Ƥ��) = ""
            .TextMatrix(0, COL_��ʾ) = "" 'Pass
            
            '�Զ������и�
            .WordWrap = True
            .AutoSize COL_ҽ������
            
            '����ÿ��ҽ��
            i = .FixedRows
            Do While i <= .Rows - 1
                '��ҩ����ҩ��һЩ����
                bln��ҩ;�� = False: bln��ҩ�÷� = False
                bln�ɼ����� = False: bln������ = False: bln������ = False '�����ڼ������
                If .TextMatrix(i, COL_�������) = "E" Then
                    If Val(.TextMatrix(i - 1, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                            bln��ҩ;�� = True
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    '��ʾ��ҩ�ĸ�ҩ;��
                                    .TextMatrix(j, COL_�÷�) = .TextMatrix(i, COL_�÷�)
                                    '��ʾ��ҩ��ִ������
                                    If Val(.TextMatrix(j, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                        .TextMatrix(j, COL_ִ������) = "�Ա�ҩ"
                                    ElseIf Val(.TextMatrix(j, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                        .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ"
                                    Else
                                        .TextMatrix(j, COL_ִ������) = ""
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                            bln��ҩ�÷� = .TextMatrix(i - 1, COL_�������) = "7" '��ҩ�÷���
                            bln�ɼ����� = .TextMatrix(i - 1, COL_�������) = "C" '�ɼ�������
                            
                            '��ʾ��ҩ�䷽�������ϵ�ִ�п���
                            .TextMatrix(i, COL_ִ�п���) = .TextMatrix(i - 1, COL_ִ�п���)
                            
                            If bln��ҩ�÷� Then
                                '��ʾ��ҩ�䷽ִ������
                                If Val(.TextMatrix(i - 1, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                    .TextMatrix(i, COL_ִ������) = "�Ա�ҩ"
                                ElseIf Val(.TextMatrix(i - 1, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                    .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ"
                                Else
                                    .TextMatrix(i, COL_ִ������) = ""
                                End If
                            Else
                                .TextMatrix(i, COL_ִ������) = ""
                            End If
                            
                            'ɾ����ζ��ҩ��,�Լ���������еļ�����Ŀ;ͬʱ�жϼ�������
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    If .TextMatrix(j, COL_�������) = "C" Then
                                        If Val(.TextMatrix(j, COL_������)) = 1 Then
                                            bln������ = True
                                            If Val(.TextMatrix(j, COL_����ID)) <> 0 Then
                                                bln������ = True
                                            End If
                                        End If
                                    End If
                                    .RemoveItem j: i = i - 1
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    Else
                        .TextMatrix(i, COL_ִ������) = ""
                    End If
                End If
                                                                
                '����ɼ��еĵ�һЩ��ʶ:�ſ����ɼ�����ʱδɾ������
                If Not bln��ҩ;�� And .TextMatrix(i, COL_�������) <> "7" Then
                
                    '�иߣ�Ϊ��֧��zl9PrintMode:Resize֮��,ȡRowHeight����С��RowHeightMin
                    If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                    
                    '����С��������,��δ�뵽�취
                    If Left(.TextMatrix(i, COL_����), 1) = "." Then
                        .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                    End If
                    If Left(.TextMatrix(i, COL_����), 1) = "." Then
                        .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                    End If
                    
                    '������ҽ����ʶ(����ҩƷ�����ҽ��,��ֻ����Ҫҽ��)
                    If Not bln��ҩ�÷� And InStr(",5,6,", .TextMatrix(i, COL_�������)) = 0 Then
                        If bln�ɼ����� Then '����ǰ��ȡ�Ľ��
                            If bln������ Then
                                If Not bln������ Then
                                    Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("δ����").Picture
                                Else
                                    Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("������").Picture
                                End If
                            End If
                        ElseIf Val(.TextMatrix(i, COL_������)) = 1 Then
                            If Val(.TextMatrix(i, COL_����ID)) = 0 Then
                                Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("δ����").Picture
                            Else
                                Set .Cell(flexcpPicture, i, COL_F����) = imgFlag.ListImages("������").Picture
                            End If
                        End If
                    End If
                    
                    'ҽ����ɫ
                    blnDo = False
                    If Val(.TextMatrix(i, COL_ҽ��״̬)) = 2 Then
                        'У������
                        If lngTop = 0 Then lngTop = i '��ɾ����Ҳ����Ӱ��ȡֵ
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H80& '���
                        blnDo = True
                    ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 4 Then
                        '������
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '��ɫ
                        .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                        blnDo = True
                    ElseIf InStr(",8,9,", Val(.TextMatrix(i, COL_ҽ��״̬))) > 0 Then
                        '��ֹͣ,��ȷ��ֹͣ:����������ֹʱ������ж�
                        If strCurr >= .TextMatrix(i, COL_��ֹʱ��) Or .TextMatrix(i, COL_��Ч) = "����" Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080 '��ɫ
                            blnDo = True
                        End If
                    ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 6 Then
                        '����ͣ
                        strTime = Format(GetAdviceTime(Val(.TextMatrix(i, COL_ID)), 6), "yyyy-MM-dd HH:mm")
                        If strCurr >= strTime Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000& '����
                            blnDo = True
                        End If
                    ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 7 Then
                        '������
                        strTime = Format(GetAdviceTime(Val(.TextMatrix(i, COL_ID)), 7), "yyyy-MM-dd HH:mm")
                        If strCurr < strTime Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000& '����
                            blnDo = True
                        End If
                    End If
                    If Not blnDo Then
                        If lngTop = 0 Then lngTop = i
                        If Val(.TextMatrix(i, COL_ҽ��״̬)) <> 1 Then
                            '��ͨ��У��(Ҳ���������Ķ��״̬)
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000 '����
                        End If
                    End If
                    
                    'У�Ժ�����ҽ����ɫ��ʾ
                    If .TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) = 4 _
                        And InStr(",1,2,4,", Val(.TextMatrix(i, COL_ҽ��״̬))) = 0 Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed '��ɫ
                    End If
                    
                    '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                    If .TextMatrix(i, COL_�������) <> "" Then
                        If InStr(",����ҩ,����ҩ,����ҩ,", .TextMatrix(i, COL_�������)) > 0 Then
                            .Cell(flexcpFontBold, i, COL_ҽ������) = True
                        End If
                    End If
                    
                    'Ƥ�Խ����ʶ
                    If .TextMatrix(i, COL_Ƥ��) = "(+)" Then
                        .Cell(flexcpForeColor, i, COL_Ƥ��) = vbRed
                    ElseIf .TextMatrix(i, COL_Ƥ��) = "(-)" Then
                        .Cell(flexcpForeColor, i, COL_Ƥ��) = vbBlue
                    End If
                    
                    '������־:һ����ҩֻ��ʾ�ڵ�һ��
                    blnFirst = True
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = False
                        End If
                    End If
                    If blnFirst Then
                        If Val(.TextMatrix(i, COL_��־)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = imgFlag.ListImages("����").Picture
                        ElseIf Val(.TextMatrix(i, COL_��־)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = imgFlag.ListImages("��¼").Picture
                        End If
                    End If
                    
                    'Pass:�����������ʾ��ʾ��
                    If .TextMatrix(i, COL_��ʾ) <> "" Then
                        Set .Cell(flexcpPicture, i, COL_��ʾ) = imgPass.ListImages(Val(.TextMatrix(i, COL_��ʾ)) + 1).Picture
                        .TextMatrix(i, COL_��ʾ) = ""
                    End If
                    
                    '����ǩ����ʶ
                    If Val(.TextMatrix(i, COL_ǩ����)) = 1 Then
                        Set .Cell(flexcpPicture, i, COL_ҽ������) = imgSign.ListImages(1).Picture
                    End If
                End If
                
                If bln��ҩ;�� Then
                    .RemoveItem i
                Else
                    i = i + 1
                End If
            Loop
            
            '�̶���ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '����ǩ��ͼ�����
            .Cell(flexcpPictureAlignment, .FixedRows, COL_ҽ������, .Rows - 1, COL_ҽ������) = 0
            .Redraw = True
        End With
    Else
        Call ClearAdviceData
        Call ClearAppendData
    End If
    
    'ֻ������ʱ���ú�ɫ����
    vsAdvice.GridColor = IIF(bln����, &H8080FF, vsAdvice.GridColorFixed)
        
    'ȱʡ��λ
    vsAdvice.Redraw = flexRDNone
    If lngҽ��ID <> 0 Then
        lngҽ��ID = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
        If lngҽ��ID <> -1 Then vsAdvice.Row = lngҽ��ID
    End If
    If lngҽ��ID = -1 Or lngҽ��ID = 0 Then
        If lngTop <> 0 Then
            vsAdvice.Row = lngTop
            vsAdvice.TopRow = lngTop
        Else
            vsAdvice.Row = vsAdvice.FixedRows
        End If
    End If
    vsAdvice.Col = vsAdvice.FixedCols
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    vsAdvice.Redraw = flexRDDirect
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    vsAdvice.Refresh
    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAppend_GotFocus()
    vsAppend.BackColorSel = &HFFCC99
End Sub

Private Sub vsAppend_LostFocus()
    vsAppend.BackColorSel = &HFFEBD7
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Function ShowSendList(ByVal lngRow As Long) As Boolean
'���ܣ���ʾָ����ҽ���ķ��ͼ�¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strExe1 As String, strExe2 As String, strState As String
    Dim bln�䷽�� As Boolean, bln������ As Boolean
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeRestrictRows
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "E" Then
            bln�䷽�� = RowIs�䷽��(lngRow)
            bln������ = RowIs������(lngRow)
        End If
                
        strExe1 = "Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��')"
        strExe2 = "Decode(Nvl(B.ִ��״̬,0),0,'δִ��',1,'ִ�����',2,'�ܾ�ִ��',3,'����ִ��')"
        strState = "Decode(A.��¼����,1,Decode(A.��¼״̬,0,'�շѻ���',1,'���շ�',3,'���˷�'),2,Decode(A.��¼״̬,0,'���ʻ���',1,'�Ѽ���',3,'������'),'�ѼƷ�')"
        
        'ҩ����Ӧ��ҩƷ�Ƽ۰�סԺ��װ��ʾ,��ҩ����Ӧ��ҩƷ�Ƽ۰����۵�λ��ʾ
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            If Not RowInһ����ҩ(lngRow, lngBegin, lngEnd) Then lngBegin = lngRow
            '��ҩ����:��д�˷��ͼ�¼,�������޶�Ӧ����(���Ա�ҩ,��ҽ���й��)
            strSub = "Select A.*,B.סԺ��װ,B.סԺ��λ" & _
                " From ���˷��ü�¼ A,ҩƷ��� B" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL And A.�շ���� IN('5','6','7')" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.ҽ�����=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
            ElseIf MovedByDate(mvInDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
            End If
                
            strSQL = _
                " Select C.���ID,C.�걾��λ,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                " Nvl(A.סԺ��λ,D.סԺ��λ) as ��λ," & _
                " Nvl(A.����/Nvl(A.סԺ��װ,1),B.��������/Nvl(D.����ϵ��,1)/Nvl(D.סԺ��װ,1)) as ��������," & _
                " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID,Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬," & _
                " B.�״�ʱ��,B.ĩ��ʱ��,Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�',1," & strState & ") as �Ʒ�״̬," & _
                " B.������,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������" & _
                " From (" & strSub & ") A,����ҽ������ B,����ҽ����¼ C,ҩƷ��� D" & _
                " Where B.ҽ��ID=C.ID And C.�շ�ϸĿID=D.ҩƷID(+)" & _
                " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And A.ҽ�����(+)=B.ҽ��ID" & _
                " And C.ID=[1]"

            '��һ����ҩ�����в���ʾ��ҩ;���ķ���
            If lngRow = lngBegin Then
                '��ҩ;������:��д�˷��ͼ�¼(������),����һ���з���
                strSub = "Select A.*,B.סԺ��װ,B.סԺ��λ" & _
                    " From ���˷��ü�¼ A,ҩƷ��� B" & _
                    " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                    " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=[2]"
                If mblnMoved Then
                    strSub = Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
                ElseIf MovedByDate(mvInDate) Then
                    strSub = strSub & " Union ALL " & Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
                End If
                    
                strSQL = strSQL & " Union ALL " & _
                    " Select C.���ID,C.�걾��λ,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                    " Decode(Nvl(Instr('567',A.�շ����),0),0,D.���㵥λ,Nvl(A.סԺ��λ,E.סԺ��λ)) as ��λ," & _
                    " Decode(Nvl(Instr('567',A.�շ����),0),0,B.��������," & _
                    "   Nvl(A.����/Nvl(A.סԺ��װ,1),B.��������/Nvl(E.����ϵ��,1)/Nvl(E.סԺ��װ,1))) as ��������," & _
                    " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID,Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬," & _
                    " B.�״�ʱ��,B.ĩ��ʱ��,Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�',1," & strState & ") as �Ʒ�״̬," & _
                    " B.������,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������" & _
                    " From (" & strSub & ") A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ D,ҩƷ��� E" & _
                    " Where B.ҽ��ID=C.ID And C.������ĿID=D.ID And C.�շ�ϸĿID=E.ҩƷID(+)" & _
                    " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And 0+A.ҽ�����(+)=B.ҽ��ID" & _
                    " And C.ID=[2]"
            End If
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
        Else
            '����ҽ��(�����䷽����飬����һ��ҽ��):��д�˷��ͼ�¼(������),����һ���з���
            '��ҩ�Ա�ҩҲ���޶�Ӧ����(��ҽ���й��)
            strSub = _
                " Select A.*,B.סԺ��װ,B.סԺ��λ" & _
                " From ���˷��ü�¼ A,ҩƷ��� B" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=[1]"
            strSub = strSub & " Union ALL " & _
                " Select A.*,B.סԺ��װ,B.סԺ��λ" & _
                " From ���˷��ü�¼ A,ҩƷ��� B,����ҽ����¼ C" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=C.ID" & _
                " And C.���ID=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
            ElseIf MovedByDate(mvInDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, "���˷��ü�¼", "H���˷��ü�¼")
            End If
            
            strSQL = _
                " Select * From ����ҽ����¼ Where ID=[1]" & _
                " Union ALL " & _
                " Select * From ����ҽ����¼ Where ���ID=[1]"
            strSQL = _
                " Select C.���ID,C.�걾��λ,B.����ʱ��,B.NO,B.��¼����,A.�շ�ϸĿID," & _
                " Decode(Nvl(Instr('567',A.�շ����),0),0,D.���㵥λ,Nvl(A.סԺ��λ,E.סԺ��λ)) as ��λ," & _
                " Decode(Nvl(Instr('567',A.�շ����),0),0,B.��������," & _
                "   Nvl(Nvl(A.����,1)*A.����/Nvl(A.סԺ��װ,1),B.��������/Nvl(E.����ϵ��,1)/Nvl(E.סԺ��װ,1))) as ��������," & _
                " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID,Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬," & _
                " B.�״�ʱ��,B.ĩ��ʱ��,Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�',1," & strState & ") as �Ʒ�״̬," & _
                " B.������,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������" & _
                " From (" & strSub & ") A,����ҽ������ B,(" & strSQL & ") C,������ĿĿ¼ D,ҩƷ��� E" & _
                " Where B.ҽ��ID=C.ID And C.������ĿID=D.ID And C.�շ�ϸĿID=E.ҩƷID(+)" & _
                " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And 0+A.ҽ�����(+)=B.ҽ��ID"
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
        End If
        
        strSQL = "Select /*+ RULE */ A.�������,A.�������," & _
            " A.���ID,A.�������,F.���� as �������,D.���� as ������Ŀ,A.�걾��λ,A.����ʱ��,A.NO,A.��¼����," & _
            " Nvl(G.����,B.����)||Decode(B.����,NULL,NULL,'('||B.����||')')||Decode(B.���,NULL,NULL,' '||B.���) as �շ���Ŀ," & _
            " A.��λ,A.�������� as ����,C.���� as ִ�п���,A.ִ��״̬,A.�״�ʱ��,A.ĩ��ʱ��,A.�Ʒ�״̬,A.������,A.���ͺ�" & _
            " From (" & strSQL & ") A,�շ���ĿĿ¼ B,���ű� C,������ĿĿ¼ D,������Ŀ��� F,�շ���Ŀ���� G" & _
            " Where A.�շ�ϸĿID=B.ID(+) And A.ִ�в���ID=C.ID(+)" & _
            " And A.������ĿID=D.ID And A.�������=F.����" & _
            " And A.�շ�ϸĿID=G.�շ�ϸĿID(+) And G.����(+)=1 And G.����(+)=" & IIF(gbln��Ʒ��, 3, 1) & _
            " Order by A.���ͺ� Desc,A.�������,A.�������,A.�������"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID)))
        
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, cs���ͺ�) = Nvl(rsTmp!���ͺ�, 0)
                .TextMatrix(i, cs����ʱ��) = Format(Nvl(rsTmp!����ʱ��), "MM-dd HH:mm")
                
                '����ҽ��
                If InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, cs����ҽ��) = "ҩƷҽ��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    .TextMatrix(i, cs����ҽ��) = "��ҩ;��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And (bln�䷽�� Or bln������) Then
                    If bln������ Then
                        .TextMatrix(i, cs����ҽ��) = "�ɼ�����-" & rsTmp!������Ŀ
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        .TextMatrix(i, cs����ҽ��) = "��ҩ�巨-" & rsTmp!������Ŀ
                    Else
                        .TextMatrix(i, cs����ҽ��) = "��ҩ�÷�-" & rsTmp!������Ŀ
                    End If
                ElseIf Not IsNull(rsTmp!���ID) Then
                    If rsTmp!������� = "C" Then
                        .TextMatrix(i, cs����ҽ��) = "������Ŀ-" & rsTmp!������Ŀ
                    ElseIf rsTmp!������� = "D" Then
                        .TextMatrix(i, cs����ҽ��) = "��鲿λ-" & Nvl(rsTmp!�걾��λ)
                    ElseIf rsTmp!������� = "F" Then
                        .TextMatrix(i, cs����ҽ��) = "��������-" & rsTmp!������Ŀ
                    ElseIf rsTmp!������� = "G" Then
                        .TextMatrix(i, cs����ҽ��) = "������Ŀ-" & rsTmp!������Ŀ
                    End If
                Else
                    .TextMatrix(i, cs����ҽ��) = rsTmp!������� & "ҽ��-" & rsTmp!������Ŀ
                End If
               
                .TextMatrix(i, cs���ݺ�) = Nvl(rsTmp!NO)
                .TextMatrix(i, cs�շ���Ŀ) = Nvl(rsTmp!�շ���Ŀ)
                .TextMatrix(i, cs����) = FormatEx(Nvl(rsTmp!����), 5) & Nvl(rsTmp!��λ)
                .TextMatrix(i, cs�Ʒ�״̬) = Nvl(rsTmp!�Ʒ�״̬)
                .TextMatrix(i, csִ��״̬) = Nvl(rsTmp!ִ��״̬)
                .TextMatrix(i, csִ�п���) = Nvl(rsTmp!ִ�п���)
                .TextMatrix(i, cs�״�ʱ��) = Format(Nvl(rsTmp!�״�ʱ��), "MM-dd HH:mm")
                .TextMatrix(i, csĩ��ʱ��) = Format(Nvl(rsTmp!ĩ��ʱ��), "MM-dd HH:mm")
                .TextMatrix(i, cs������) = Nvl(rsTmp!������)
                .TextMatrix(i, cs��¼����) = Nvl(rsTmp!��¼����)
                rsTmp.MoveNext
            Next
        End If
        
        .Row = 1: .Col = cs����ҽ��
        .Redraw = True
    End With
    ShowSendList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSignList(ByVal lngRow As Long) As Boolean
'���ܣ���ʾָ����ҽ����ǩ����¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .MergeCells = flexMergeNever
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSignList = True: Exit Function
        End If
        
        strSQL = "Select A.ǩ��ID,A.��������,B.ǩ��ʱ��,B.ǩ����," & _
            " Decode(A.��������,1,'�¿�ҽ��',4,'����ҽ��',8,'ֹͣҽ��','��������') as ǩ������" & _
            " From ����ҽ��״̬ A,ҽ��ǩ����¼ B Where A.ҽ��ID=[1] And A.ǩ��ID=B.ID Order by B.ǩ��ʱ��"
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
            strSQL = Replace(strSQL, "ҽ��ǩ����¼", "Hҽ��ǩ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!ǩ��ID)
                .TextMatrix(i, 0) = rsTmp!ǩ������
                .Cell(flexcpData, i, 0) = Val(rsTmp!��������)
                .TextMatrix(i, 1) = Format(rsTmp!ǩ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 2) = rsTmp!ǩ����
                Set .Cell(flexcpPicture, i, 0) = imgSign.ListImages(1).Picture
                rsTmp.MoveNext
            Next
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 0
        .Row = 1
        .Redraw = True
    End With
    ShowSignList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowRollList(ByVal lngRow As Long) As Boolean
'���ܣ���ʾָ����ҽ�����Ի��˵������ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objMenu As Menu
    
    For i = mfrmParent.mnuAdviceFuncRoll.UBound To 0 Step -1
        mfrmParent.mnuAdviceFuncRoll(i).Tag = ""
        If i = 0 Then
            mfrmParent.mnuAdviceFuncRoll(i).Caption = "<������>"
        Else
            On Error Resume Next
            Unload mfrmParent.mnuAdviceFuncRoll(i)
            If Err.Number <> 0 Then
                Err.Clear
                mfrmParent.mnuAdviceFuncRoll(i).Visible = False
                mfrmParent.mnuAdviceFuncRoll(i).Tag = ""
            End If
            On Error GoTo 0
        End If
    Next
    mfrmParent.mnuAdviceFunc(mnuҽ������).Tag = ""
    If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
        ShowRollList = True: Exit Function
    End If
    
    On Error GoTo errH
    
    '�ɻ���ҽ�������ͷ���,ҽ������Ĳ�������(�緢�ͺ��Զ�ֹͣ)
    '�������ɻ����Զ�ֹͣ,���˷���ʱ�Զ�����ֹͣ
    strSQL = " And (A.ID=[1] Or A.���ID=[1])"
    strSQL = _
        " Select Distinct 0 as ���ͺ�,B.������Ա as ��Ա,B.����ʱ�� as ʱ��,B.��������," & _
        " Decode(B.��������,4,'����ҽ��',5,'����ҽ��',6,'��ͣҽ��',7,'����ҽ��',8,'ֹͣҽ��',9,'ȷ��ֹͣ',10,'Ƥ�Խ��') as ����" & _
        " From ����ҽ����¼ A,����ҽ��״̬ B" & _
        " Where A.ID=B.ҽ��ID" & strSQL & _
        " And (Nvl(A.ҽ����Ч,0)=0 And B.�������� Not IN(1,2,3)" & _
            " Or Nvl(A.ҽ����Ч,0)=1 And B.�������� Not IN(1,2,3,8))" & _
        " Union ALL" & _
        " Select Distinct B.���ͺ�,B.������ as ��Ա,B.����ʱ�� as ʱ��,0 as ��������,'����ҽ��' as ����" & _
        " From ����ҽ����¼ A,����ҽ������ B" & _
        " Where A.ID=B.ҽ��ID" & strSQL & _
        " Order by ʱ�� Desc,���ͺ�"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_��ID)))
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If i > 1 Then
                If mfrmParent.mnuAdviceFuncRoll.UBound >= i - 1 Then
                    mfrmParent.mnuAdviceFuncRoll(i - 1).Visible = True
                Else
                    Load mfrmParent.mnuAdviceFuncRoll(i - 1)
                End If
            End If
            Set objMenu = mfrmParent.mnuAdviceFuncRoll(mfrmParent.mnuAdviceFuncRoll.UBound)
            objMenu.Caption = "������:" & rsTmp!��Ա & ",ʱ��:" & Format(rsTmp!ʱ��, "MM-dd HH:mm") & ",����:" & rsTmp!����
            '��¼��������,���ͺ�,����ʱ��,��Ա
            objMenu.Tag = rsTmp!�������� & "|" & rsTmp!���ͺ� & "|" & Format(rsTmp!ʱ��, "yyyy-MM-dd HH:mm:ss") & "|" & rsTmp!��Ա
            If i = 1 Then
                'ҽ��ֻ�ܻ������ϡ�ֹͣ�����ѵ��������Ͳ���
                If InStr(",0,4,8,", Nvl(rsTmp!��������, 0)) > 0 Then
                    objMenu.Enabled = True
                Else
                    objMenu.Enabled = False
                End If
            Else
                objMenu.Enabled = False
            End If
            rsTmp.MoveNext
        Next
        mfrmParent.mnuAdviceFunc(mnuҽ������).Tag = rsTmp.RecordCount
    End If
    
    ShowRollList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            vsAdvice.ColWidth(lngCol) = vsAdvice.ColData(lngCol)
            vsAdvice.ColHidden(lngCol) = False
        Else
            vsAdvice.ColWidth(lngCol) = 0
            vsAdvice.ColHidden(lngCol) = True
        End If
    End If
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub
