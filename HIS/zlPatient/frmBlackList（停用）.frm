VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBlackList 
   AutoRedraw      =   -1  'True
   Caption         =   "���ⲡ�˹���"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "frmBlackList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9885
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�Ǽ�"
               Key             =   "Add"
               Description     =   "�Ǽ�"
               Object.ToolTipText     =   "�Ǽ����ⲡ��"
               Object.Tag             =   "�Ǽ�"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modi"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸����ⲡ��"
               Object.Tag             =   "�޸�"
               ImageKey        =   "Modi"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Del"
               Description     =   "����"
               Object.ToolTipText     =   "�������˵ĵǼ�"
               Object.Tag             =   "����"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��Ƭ"
               Key             =   "View"
               Description     =   "��Ƭ"
               Object.ToolTipText     =   "�Կ�Ƭ��ʽ���ĵ�ǰ������Ϣ"
               Object.Tag             =   "��Ƭ"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "�ڵ�ǰ�����嵥�й������������Ĳ���"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������Ĳ�����"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
         Begin VB.CheckBox chkShowDel 
            Caption         =   "��ʾ�ѳ�������"
            Height          =   195
            Left            =   8175
            TabIndex        =   4
            Top             =   240
            Width           =   1560
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   6675
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBlackList.frx":06EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12356
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   5970
      Left            =   0
      TabIndex        =   0
      Top             =   705
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   10530
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorSel    =   12632256
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmBlackList.frx":0F7C
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   45
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":1296
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":14B0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":16CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":18E4
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":1AFE
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":1D18
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":2412
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":2B0C
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3206
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3420
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":363A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3854
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   645
      Top             =   450
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3A6E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3C88
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":3EA2
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":40BC
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":42D6
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":44F0
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":4BEA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":52E4
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":59DE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":5BF8
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":5E12
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBlackList.frx":602C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBlackList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstrPrivs As String
Private mrsPati As ADODB.Recordset
Private mstrFilter As String
Private mfrmFilter As frmPatiFilter
Private mfrmFind As frmPatiFind
Private mlngGo As Long, mblnGo As Boolean
Private mblnDown As Boolean
Private mlngCurRow As Long, mlngTopRow As Long

Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    �Ǽ�ʱ��B As Date
    �Ǽ�ʱ��E As Date
    ����ʱ��B As Date
    ����ʱ��E As Date
    ��Ժʱ��B As Date
    ��Ժʱ��E As Date
    ��Ժʱ��B As Date
    ��Ժʱ��E As Date
    סԺ�� As String
    �Ա� As String
    ��� As String
    ���� As String
    Patient As String
End Type
Private SQLCondition As Type_SQLCondition

Private Sub chkShowDel_Click()
    If Visible Then Call ShowPatis(mstrFilter)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF5
            Call ShowPatis(mstrFilter)
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ����
            If tbr.Buttons("Go").Enabled Then Call SeekPati(False)
        Case vbKeyReturn
            Call tbr_ButtonClick(tbr.Buttons("View"))
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub Form_Load()
    Call SetHeader
    RestoreWinState Me, App.ProductName
    
    Set mfrmFilter = New frmPatiFilter
    Set mfrmFind = New frmPatiFind
    
    If InStr(mstrPrivs, "���ⲡ�˹���") = 0 Then
        tbr.Buttons("Add").Visible = False
        tbr.Buttons("Modi").Visible = False
        tbr.Buttons("Del").Visible = False
        tbr.Buttons("Edit_").Visible = False
    End If
    
    'ˢ������
    mstrFilter = ""
    Call ShowPatis(mstrFilter)
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long '������ռ�ø߶�
    Dim staH As Long '״̬��ռ�ø߶�
    
    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub

    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    With mshPati
        .Left = 0
        .Top = cbrH
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - cbrH - staH
    End With
    
    chkShowDel.Top = tbr.Top + (tbr.Height - chkShowDel.Height) / 2
    If Me.ScaleWidth - chkShowDel.Width - 100 < 6000 Then
        chkShowDel.Left = 6000
    Else
        chkShowDel.Left = Me.ScaleWidth - chkShowDel.Width - 100
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmFind
    Unload mfrmFilter
    
    Set mrsPati = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshPati_DblClick()
    Call tbr_ButtonClick(tbr.Buttons("View"))
End Sub

Private Sub mshPati_EnterCell()
    mshPati.ForeColorSel = mshPati.CellForeColor
    Call SetMenuEnabled
    mlngGo = mshPati.Row
    mlngCurRow = mshPati.Row: mlngTopRow = mshPati.TopRow
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshPati.Cols - 1
        If mshPati.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Sub SetMenuEnabled()
'���ܣ����ݵ�ǰ��¼������ò˵�����״̬
    Dim lng����ID As Long
    
    If glngSys Like "8??" Then
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")))
    Else
        lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
    End If
    
    tbr.Buttons("Print").Enabled = lng����ID <> 0
    tbr.Buttons("Preview").Enabled = lng����ID <> 0
    tbr.Buttons("Modi").Enabled = lng����ID <> 0
    tbr.Buttons("Del").Enabled = lng����ID <> 0 And mshPati.TextMatrix(mshPati.Row, GetColNum("����ʱ��")) = ""
    tbr.Buttons("View").Enabled = lng����ID <> 0
    tbr.Buttons("Go").Enabled = lng����ID <> 0
End Sub

Private Sub mshPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And tbr.Buttons("Del").Enabled And tbr.Buttons("Del").Visible Then
        Call tbr_ButtonClick(tbr.Buttons("Del"))
    End If
End Sub

Private Sub mshPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mblnDown = True
End Sub

Private Sub mshPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPati.MouseRow = 0 Then
        mshPati.MousePointer = 99
    Else
        mshPati.MousePointer = 0
    End If
End Sub

Private Sub mshPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshPati.MouseCol
    
    If Button = 1 And mshPati.MousePointer = 99 And mblnDown Then '˫�����ʱ��ִ��
        mblnDown = False
        If mshPati.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshPati.TextMatrix(1, GetColNum("����ID")) = "" Then Exit Sub
        Set mshPati.DataSource = Nothing
        mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
        mshPati.ColData(lngCol) = (mshPati.ColData(lngCol) + 1) Mod 2
        Call ShowPatis("", True)
    End If
End Sub

Private Sub DeletePati(ByVal lngRow As Long)
    Dim lng��� As Long
    
    lng��� = Val(mshPati.TextMatrix(lngRow, GetColNum("���")))
    If lng��� = 0 Then Exit Sub
    
    If frmBlackListEdit.ShowMe(Me, mstrPrivs, lng���, True) Then
        Call ShowPatis(mstrFilter)
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lng��� As Long, lng����ID As Long, blnOK As Boolean
    
    Select Case Button.Key
        Case "Preview"
            Call OutputList(2)
        Case "Print"
            Call OutputList(1)
        Case "Add"
            If frmBlackListEdit.ShowMe(Me, mstrPrivs) Then
                Call ShowPatis(mstrFilter)
            End If
        Case "Modi"
            lng��� = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("���")))
            If lng��� <> 0 Then
                If frmBlackListEdit.ShowMe(Me, mstrPrivs, lng���) Then
                    Call ShowPatis(mstrFilter)
                End If
            End If
        Case "Del"
            Call DeletePati(mshPati.Row)
        Case "View"
            lng����ID = Val(mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")))
            If lng����ID <> 0 Then
                If CreatePublicPatient() Then
                    Call gobjPublicPatient.ReadPatiDegreeCard(Me, lng����ID, Val(mshPati.TextMatrix(mshPati.Row, GetColNum("��ҳID"))))
                End If
                mshPati.Refresh
            End If
        Case "Go"
            blnOK = gblnOK
            mfrmFind.mbytType = 0 '�����в�����������
            mfrmFind.Show 1, Me
            If gblnOK Then Call SeekPati(mfrmFind.optHead)
            gblnOK = blnOK
        Case "Filter"
            blnOK = gblnOK
            
            mfrmFilter.mbytType = 0 '�����в�����������
            mfrmFilter.mbytInFun = 1
            mfrmFilter.Show 1, Me
            If gblnOK Then
                 With mfrmFilter
                    mstrFilter = .mstrFilter
                    SQLCondition.�Ǽ�ʱ��B = .dtp�Ǽ�B
                    SQLCondition.�Ǽ�ʱ��E = .dtp�Ǽ�E
                    SQLCondition.����ʱ��B = .dtp����B
                    SQLCondition.����ʱ��E = .dtp����E
                    
                    SQLCondition.��Ժʱ��B = .dtp��ԺB
                    SQLCondition.��Ժʱ��E = .dtp��ԺE
                    SQLCondition.��Ժʱ��B = .dtp��ԺB
                    SQLCondition.��Ժʱ��E = .dtp��ԺE
                    
                    SQLCondition.סԺ�� = Trim(.txtסԺ��.Text)
                    SQLCondition.�Ա� = zlCommFun.GetNeedName(.cbo�Ա�.Text)
                    SQLCondition.��� = Trim(.txt���.Text)
                    SQLCondition.���� = zlCommFun.GetNeedName(.txt����.Text)
                    
                    If .PatiIdentify.GetCurCard.���� = "����" And .mlngPatiId = 0 And (.chk�Ǽ�.Value = 1 Or .chk��Ժ.Value = 1 Or .chk��Ժ.Value = 1) Then    '����
                        SQLCondition.Patient = Trim(.PatiIdentify.Text) & "%"
                    Else
                        SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
                    End If
                End With
                
                Call ShowPatis(mstrFilter)
            End If
            gblnOK = blnOK
        Case "Help"
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case "Quit"
            Unload Me
    End Select
End Sub

Private Function GetPatiType(ByVal lngRow As Long) As Byte
'���ܣ�0-����,1-��Ժ,2-��Ժ,3-����
    With mshPati
        If .TextMatrix(.Row, GetColNum("��Ժʱ��")) <> "" And .TextMatrix(.Row, GetColNum("��Ժʱ��")) = "" Then
            GetPatiType = 1
        ElseIf .TextMatrix(.Row, GetColNum("��Ժʱ��")) <> "" Then
            GetPatiType = 2
        Else
            GetPatiType = 3
        End If
    End With
End Function

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshPati.Row
    
    '��ͷ
    objOut.Title.Text = "���ⲡ������"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshPati.Redraw = False
    Set objOut.Body = mshPati
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshPati.Row = intRow
    mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
    mshPati.Redraw = True
End Sub

Private Sub SetHeader(Optional blnWidth As Boolean = True)
    Dim strHead As String
    Dim i As Integer
    
    strHead = "���,1,600|����ID,1,750|��ʶ��,1,750|����,1,700|�Ա�,1,500|����,1,500|����ԭ��,1,2500|����ʱ��,1,1100|�Ǽ���,1,700|����ԭ��,1,2500|����ʱ��,1,1100|������,1,700|�ѱ�,1,850|����,1,850|����,1,500|��Ժʱ��,1,1000|��Ժʱ��,1,1000|סԺ����,4,850|��ҳID,1,0|���֤��,1,0|�����,1,0|סԺ��,1,0|���￨��,1,0"

    With mshPati
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible And blnWidth Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        .RowHeight(0) = 320
        
        '�ָ��ϴ���
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
        Call mshPati_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub ShowPatis(ByVal strIF As String, Optional blnSort As Boolean)
    Dim Curdate As Date, strSQL As String
    
    On Error GoTo errH
    
    If Not blnSort Then
        'ȱʡ��ʾ�������ⲡ��
        If strIF = "" Then
            strIF = " And C.����ʱ�� Between trunc(sysdate-30) And trunc(sysdate)+1"
        Else
            strIF = Replace(strIF, "A.�Ǽ�ʱ��", "C.����ʱ��")
        End If
        
        If chkShowDel.Value = 0 Then
            strIF = strIF & " And C.����ʱ�� is NULL"
        End If
        
        strSQL = _
            "Select C.���,A.����ID,Decode(Nvl(A.סԺ����,0),0,A.�����,A.סԺ��) as ��ʶ��," & _
            " A.����,A.�Ա�,A.����,C.����ԭ��,To_Char(C.����ʱ��,'MM-DD HH24:MI') as ����ʱ��,C.�Ǽ���," & _
            " C.����ԭ��,To_Char(C.����ʱ��,'MM-DD HH24:MI') as ����ʱ��,C.������," & _
            " Decode(Nvl(A.��ҳID,0),0,A.�ѱ�,P.�ѱ�) as �ѱ�,D.���� as ����,P.��Ժ���� as ����," & _
            " To_Char(P.��Ժ����,'YYYY-MM-DD') as ��Ժʱ��,To_Char(P.��Ժ����,'YYYY-MM-DD') as ��Ժʱ��," & _
            " A.סԺ����,P.��ҳID,A.���֤��,A.�����,A.סԺ��,A.���￨��" & _
            " From ������ҳ P,������Ϣ A,���ⲡ�� C,���ű� D" & _
            " Where A.����ID=P.����ID(+) And Nvl(A.��ҳID,0)=P.��ҳID(+)" & _
            " And A.����ID=C.����ID And P.��Ժ����ID=D.ID(+)" & strIF & _
            " Order by C.����ʱ�� Desc"

        Call zlCommFun.ShowFlash("���ڶ�ȡ���ⲡ������,���Ժ� ...", Me)
        Me.Refresh
        With SQLCondition
            Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 0, "", .�Ǽ�ʱ��B, .�Ǽ�ʱ��E, .����ʱ��B, .����ʱ��E, _
                .��Ժʱ��B, .��Ժʱ��E, .��Ժʱ��B, .��Ժʱ��E, .סԺ��, .�Ա�, .����, .���, .Patient)
        End With
    End If
    
    mshPati.Clear
    mshPati.Rows = 2
    
    If mrsPati.EOF Then
        Call SetHeader(False)
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κβ���"
    Else
        Set mshPati.DataSource = mrsPati
        Call SetHeader(False)
        mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
        stbThis.Panels(2) = "�� " & mrsPati.RecordCount & " ������"
    End If
    Call mshPati_EnterCell
    
    If Not blnSort Then Call zlCommFun.StopFlash
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SeekPati(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������Ĳ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With mfrmFind
            If .txt����ID.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����ID")) = .txt����ID.Text
            End If
            If .txt���￨.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("���￨��")) = .txt���￨.Text
            End If
            If .txt�����.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("�����")) = .txt�����.Text
            End If
            If .txtסԺ��.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("סԺ��")) = .txtסԺ��.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����")) = .txt����.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshPati.TextMatrix(i, GetColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
            If .txt���֤.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("���֤��")) = .txt���֤.Text
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mlngGo = i + 1
            mshPati.Row = i: mshPati.TopRow = i
            mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
            stbThis.Panels(2).Text = "�ҵ�һ����¼"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            stbThis.Panels(2).Text = "�û�ȡ����λ����"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "�Ѷ�λ������β��"
    Screen.MousePointer = 0
End Sub
