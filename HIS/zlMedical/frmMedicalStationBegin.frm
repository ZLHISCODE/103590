VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmMedicalStationBegin 
   Caption         =   "��ʼ���"
   ClientHeight    =   5880
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9750
   Icon            =   "frmMedicalStationBegin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9750
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5520
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStationBegin.frx":076A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12118
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9750
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   1270
         ButtonWidth     =   1402
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Alt+A)"
               Object.Tag             =   "&A.ȫѡ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Alt+C)"
               Object.Tag             =   "&C.ȫ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&B.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+B)"
               Object.Tag             =   "&B.����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&B.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+B)"
               Object.Tag             =   "&B.����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����(Alt+H)"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�(Alt+X)"
               Object.Tag             =   "&X.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8145
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":1778
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":1EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":266C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":288C
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7485
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":2AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":3226
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":39A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":411A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStationBegin.frx":433A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4875
      Left            =   6300
      TabIndex        =   10
      Top             =   630
      Width           =   3405
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   45
         Picture         =   "frmMedicalStationBegin.frx":455A
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   960
         Width           =   285
      End
      Begin VB.CheckBox chk 
         Caption         =   "�ҵ�����Ϊ������(&5)"
         Height          =   240
         Index           =   1
         Left            =   855
         TabIndex        =   14
         Top             =   2250
         Width           =   2370
      End
      Begin VB.CheckBox chk 
         Caption         =   "�ҵ�����Ϊ����(&4)"
         Height          =   240
         Index           =   0
         Left            =   855
         TabIndex        =   13
         Top             =   1935
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   1305
         TabIndex        =   12
         Text            =   "cbo"
         Top             =   195
         Width           =   1995
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   930
         Width           =   1995
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   1125
         TabIndex        =   5
         Top             =   1365
         Width           =   1470
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&3.��������λ"
         Height          =   240
         Left            =   4710
         TabIndex        =   11
         Top             =   210
         Width           =   1425
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   1995
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.�� �� ��"
         Height          =   180
         Index           =   3
         Left            =   375
         TabIndex        =   3
         Tag             =   "�����"
         Top             =   1005
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.��    ��"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   1
         Top             =   630
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.������λ"
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   0
         Top             =   255
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   4740
      Left            =   60
      TabIndex        =   6
      Top             =   735
      Width           =   6210
      _cx             =   10954
      _cy             =   8361
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   -30
         X2              =   1755
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileSelectAll 
         Caption         =   "ȫѡ(&A)"
      End
      Begin VB.Menu mnuFileClearAll 
         Caption         =   "ȫ��(&C)"
      End
      Begin VB.Menu mnuFile_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileCome 
         Caption         =   "����(&B)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalStationBegin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mblnStarted As Boolean
Private mlngFindRow As Long
Private mintSort As Integer

Private Enum mCol
    ���� = 0
    ����
    �����
    ������
    ���￨��
    ���֤��
    �Ա�
    ������λ
    ���
End Enum

Public WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1

'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long

    mnuFileSave.Enabled = True
        
    If vData = False Then
        mnuFileSave.Enabled = False
    End If

    If mblnStarted = False Then
        tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    Else
        tbrThis.Buttons("����").Enabled = mnuFileSave.Enabled
    End If
    
End Property

Private Sub RefreshState()
    
    Dim lngLoop As Long
    Dim intCount As Integer
    
    intCount = 0
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, 0))) = 1 Then
            intCount = intCount + 1
        End If
    Next
    
    stbThis.Panels(2).Text = "��ǰѡ�� " & intCount & " ��"
End Sub

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next



    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef lngKey As Long, Optional lng����id As Long = 0, Optional blnStarted As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    mblnStarted = blnStarted
    mlngKey = lngKey
    Set mfrmMain = frmMain
        
    If InitData = False Then Exit Function
    If ReadData(mlngKey) = False Then Exit Function
    
    '����ǵ�����,ֱ�Ӵ���,����������
    If lng����id > 0 Then
        vsf.TextMatrix(1, mCol.����) = 1
        
        If MsgBox("���Ҫ���������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveEdit() Then ShowEdit = True
        End If
        
        Exit Function
    End If
    
    EditChanged = (Val(vsf.RowData(1)) > 0)
    
    Call RefreshState
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK

End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset

    On Error GoTo errHand

    gstrSQL = "SELECT 0 AS ����,A.����id AS ID,B.����,B.�Ա�,B.������λ,B.�����,b.������,b.���￨��,b.���֤��,A.������� AS ���,'' AS δ��ԭ�� " & _
                "FROM �����Ա���� A,������Ϣ B " & _
                "WHERE A.���״̬ IN (1,4) AND A.��챨��=0 AND A.����id=B.����id and A.�Ǽ�id=[1]"
                
    gstrSQL = gstrSQL & " Order By �����"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
        Call AppendRows(vsf, lnX, lnY)
    End If
    
    gstrSQL = "SELECT Distinct B.������λ " & _
                "FROM �����Ա���� A,������Ϣ B " & _
                "WHERE B.������λ Is Not Null And A.���״̬ IN (1,4) AND A.��챨�� In ([2],[3]) AND A.����id=B.����id and A.�Ǽ�id=[1]"
                    
    cbo(1).Clear
    cbo(1).AddItem ""
    cbo(1).ListIndex = 0
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, 0, 0)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(1).AddItem rs("������λ").Value
            rs.MoveNext
        Loop
    End If
    
    gstrSQL = "SELECT Distinct A.������� " & _
                "FROM �����Ա���� A " & _
                "WHERE A.������� Is Not Null And A.���״̬ IN (1,4) AND A.��챨��=0 AND A.�Ǽ�id=[1]"
                    
    cbo(0).Clear
    cbo(0).AddItem ""
    cbo(0).ListIndex = 0
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo(0).AddItem rs("�������").Value
            rs.MoveNext
        Loop
    End If
    
    ReadData = True
    
    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mlngFindRow = 1
    mintSort = 0
    
    strVsf = "����,450,1,1,1,;����,900,1,1,1,;�����,900,7,1,1,;������,900,7,1,1,;���￨��,900,1,1,1,;���֤��,1200,1,1,0,;�Ա�,810,1,1,1,;������λ,2280,1,1,1,;���,1200,1,1,1,"
    
    Call CreateVsf(vsf, strVsf)
    vsf.Cols = vsf.Cols + 1
    vsf.ColWidth(vsf.Cols - 1) = 15
    vsf.ColDataType(0) = flexDTBoolean
    vsf.Editable = True
    
    Call AppendRows(vsf, lnX, lnY)
    
    tbrThis.Buttons("����").Visible = True
    tbrThis.Buttons("����").Visible = True
    
    If mblnStarted = False Then
        Me.Caption = "�������"
        tbrThis.Buttons("����").Visible = False
        mnuFileCome.Visible = False
    Else
        Me.Caption = "��Ա����"
        
        tbrThis.Buttons("����").Visible = False
        mnuFileSave.Visible = False
    End If
    
    cbo(0).Clear
    cbo(0).AddItem ""
    cbo(0).ListIndex = 0
    
    cbo(1).Clear
    cbo(1).AddItem ""
    cbo(1).ListIndex = 0
    
    
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function


Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL  As String
    Dim lngCount As Long
    Dim lngDept As Long
    Dim lngSendNo As Long
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim str�ɼ�No As String
    Dim strNO As String
    Dim lngTotal As Long
    Dim strTmp As String
    Dim strSample As String
    Dim strDeptID As String
    Dim blnVerfiy As Boolean
    Dim blnCheck As Boolean
    Dim varTmp As Variant
    
    On Error GoTo errHand

    Me.Enabled = False

    Call frmWait.OpenWait(Me, IIf(tbrThis.Buttons("����").Visible, "�������", "��챨��"))
    frmWait.WaitInfo = "���ڽ��б�������..."

    '������ʼ������
    
    
    '��ȡ�Զ���ӡ���뵥��ز���
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\������뵥����", "����1", "")
    
    If Trim(strTmp) <> "" And InStr(strTmp, "'") > 0 Then
        
        strTmp = Mid(strTmp, InStr(strTmp, "'") + 1)

        varTmp = Split(strTmp, "'")
        
        On Error Resume Next
        
        blnCheck = (Val(varTmp(0)) = 1)
        blnVerfiy = (Val(varTmp(1)) = 1)
        
        strSample = ""
        For lngLoop = 2 To UBound(varTmp) - 1
            strSample = strSample & "''" & varTmp(lngLoop)
        Next
        If strSample <> "" Then strSample = strSample & "''"
        
    End If
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\������뵥����", "��ӡִ�п���", "")
    If strTmp <> "" Then
        strTmp = "'" & strTmp & "'"
        varTmp = Split(strTmp, "'")
        
        strDeptID = ""
        For lngLoop = 0 To UBound(varTmp)
            strDeptID = strDeptID & "," & Val(varTmp(lngLoop))
        Next
        If strDeptID <> "" Then strDeptID = Mid(strDeptID, 2)
        
    End If
    
    
    lngDept = mfrmMain.cboDept.ItemData(mfrmMain.cboDept.ListIndex)

    
    lngTotal = vsf.Rows - 1
    lngSendNo = GetNextNo(10)
    
    frmWait.WaitInfo = "���ڽ��б�������..."
    
    strSQL = "Select a.����id,a.�嵥id,b.����;��,b.�ɼ���ʽid From �����Ŀҽ�� a,�����Ŀ�嵥 b Where a.�嵥id=b.id and b.�Ǽ�id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
    If rs.BOF Then GoTo errHand
    
    For lngLoop = 1 To lngTotal
        
        If Val(vsf.RowData(lngLoop)) > 0 And Abs(Val(vsf.TextMatrix(lngLoop, mCol.����))) = 1 Then
            
            Call SQLRecord(rsSQL)

            frmWait.WaitInfo = "���ڽ��б������ܡ�" & vsf.TextMatrix(lngLoop, mCol.����) & " ��..." & Format(100 * lngLoop / lngTotal, "0.00") & "%"
            
            '�������õ��ݺ�
            rs.Filter = ""
            rs.Filter = "����id=" & Val(vsf.RowData(lngLoop))
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    
                    str�ɼ�No = ""
                    strNO = ""
                    
                    If zlCommFun.NVL(rs("����;��").Value, 1) = 1 Then
                        '����
                        strNO = GetNextNo(14)
                    Else
                        strNO = GetNextNo(13)
                    End If
                    
                    If zlCommFun.NVL(rs("�ɼ���ʽid").Value, 0) > 0 Then
                        '�ɼ�
                        If zlCommFun.NVL(rs("����;��").Value, 1) = 1 Then
                            '����
                            str�ɼ�No = GetNextNo(14)
                        Else
                            str�ɼ�No = GetNextNo(13)
                        End If
                    End If
                    
                    strSQL = "ZL_�����Ŀҽ��_NO(" & zlCommFun.NVL(rs("�嵥id").Value, 0) & "," & Val(vsf.RowData(lngLoop)) & ",'" & strNO & "','" & str�ɼ�No & "')"
                    Call SQLRecordAdd(rsSQL, strSQL)
                    
                    rs.MoveNext
                Loop
            End If
            
            '��ʼִ��
            blnTran = True
            gcnOracle.BeginTrans
            If rsSQL.RecordCount > 0 Then rsSQL.MoveFirst
            For lngCount = 1 To rsSQL.RecordCount
                Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
                rsSQL.MoveNext
            Next
            
            '���ܻ򱨵���ʼ
            strSQL = "zl_�����Ա����_Accept(" & mlngKey & "," & lngSendNo & "," & Val(vsf.RowData(lngLoop)) & "," & lngDept & ",NULL,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '������ط���
            If MakeMedicalCharge(rsSQL, mlngKey) = False Then
                frmWait.CloseWait
                Me.Enabled = True
                gcnOracle.RollbackTrans
                blnTran = False
                Exit Function
            End If
            
            '���ܻ򱨵�����
            strSQL = "zl_�����Ա����_Accept(" & mlngKey & "," & lngSendNo & "," & Val(vsf.RowData(lngLoop)) & "," & lngDept & ",NULL,2)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            gcnOracle.CommitTrans
            blnTran = False
            
            If Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ���ӡָ����", 0)) = 1 Or Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ���ӡ���뵥", 0)) = 1 Then
                
                Call frmWait.HideWait
                
            
                '�Զ���ӡָ����
                If Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ���ӡָ����", 0)) = 1 Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1861", Me, "�Ǽ�id=" & mlngKey, "����id=" & Val(vsf.RowData(lngLoop)), 2)
                End If
                
                '�Զ���ӡ���뵥
                If Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "�Զ���ӡ���뵥", 0)) = 1 Then
                    Call OutPutQuestBill(Me, mlngKey, Val(vsf.RowData(lngLoop)), strDeptID, strSample, blnVerfiy, blnCheck, 2)
                End If
                
                Call frmWait.ShowWait
            End If
            
        End If
    Next


    frmWait.CloseWait
    Me.Enabled = True
        
    SaveEdit = True

    Exit Function

errHand:

    frmWait.CloseWait
    Me.Enabled = True

    If ErrCenter = 1 Then
        Resume
    End If

    If blnTran Then
        gcnOracle.RollbackTrans
        ShowSimpleMsg "δ��ȫ�����ɹ��򲿷ݽ��ܳɹ���"
    End If

End Function


Private Sub cbo_Click(Index As Integer)
    mlngFindRow = 0
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Function FindData() As Boolean
    
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnFind1 As Boolean
    Dim blnFind2 As Boolean
    Dim blnFind3 As Boolean
    Dim blnFind4 As Boolean
    Dim strCol As String
    
    FindData = True
    
    If mlngFindRow >= vsf.Rows - 1 Then mlngFindRow = 0

    strCol = Mid(lbl(3).Caption, 4)
    lngCol = GetCol(vsf, strCol)
    
    For lngRow = mlngFindRow + 1 To vsf.Rows - 1
        
        blnFind1 = True
        blnFind2 = True
        blnFind3 = True
        blnFind4 = True
        
        If cbo(1).Text <> "" Then
            blnFind1 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.������λ)), UCase(cbo(1).Text)) > 0 Then
                blnFind1 = True
            End If
        End If
        
        If txt(1).Text <> "" Then
            blnFind2 = False
            
            Select Case strCol
            Case "�� �� ��"
                If UCase(vsf.TextMatrix(lngRow, mCol.�����)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "�� �� ��"
                If UCase(vsf.TextMatrix(lngRow, mCol.������)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "���￨��"
                If UCase(vsf.TextMatrix(lngRow, mCol.���￨��)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "���֤��"
                If UCase(vsf.TextMatrix(lngRow, mCol.���֤��)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "��    ��"
                If UCase(vsf.TextMatrix(lngRow, mCol.����)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "����ƴ��"
                If zlGetSymbol(UCase(vsf.TextMatrix(lngRow, mCol.����))) = UCase(txt(1).Text) Then blnFind2 = True
            Case "�������"
                If zlGetSymbol(UCase(vsf.TextMatrix(lngRow, mCol.����)), 1) = UCase(txt(1).Text) Then blnFind2 = True
            End Select

        End If
        
        
        If cbo(0).Text <> "" Then
            blnFind4 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.���)), UCase(cbo(0).Text)) > 0 Then
                blnFind4 = True
            End If
        End If
        
        If blnFind1 And blnFind2 And blnFind3 And blnFind4 Then
            mlngFindRow = lngRow
            
            vsf.Row = mlngFindRow
            vsf.ShowCell vsf.Row, vsf.Col
            vsf.SetFocus
            
            Exit Function
        End If
    Next
    
    For lngRow = 1 To mlngFindRow
        
        blnFind1 = True
        blnFind2 = True
        blnFind3 = True
        blnFind4 = True
        
        If cbo(1).Text <> "" Then
            blnFind1 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.������λ)), UCase(cbo(1).Text)) > 0 Then
                blnFind1 = True
            End If
        End If
        
        If txt(1).Text <> "" Then
            blnFind2 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.�����)), UCase(txt(1).Text)) > 0 Then
                blnFind2 = True
            End If
            
            Select Case strCol
            Case "�� �� ��"
                If UCase(vsf.TextMatrix(lngRow, mCol.�����)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "�� �� ��"
                If UCase(vsf.TextMatrix(lngRow, mCol.������)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "���￨��"
                If UCase(vsf.TextMatrix(lngRow, mCol.���￨��)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "���֤��"
                If UCase(vsf.TextMatrix(lngRow, mCol.���֤��)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "��    ��"
                If UCase(vsf.TextMatrix(lngRow, mCol.����)) = UCase(txt(1).Text) Then blnFind2 = True
            Case "����ƴ��"
                If zlGetSymbol(UCase(vsf.TextMatrix(lngRow, mCol.����))) = UCase(txt(1).Text) Then blnFind2 = True
            Case "�������"
                If zlGetSymbol(UCase(vsf.TextMatrix(lngRow, mCol.����)), 1) = UCase(txt(1).Text) Then blnFind2 = True
            End Select
            
        End If
        
        If cbo(0).Text <> "" Then
            blnFind4 = False
            If InStr(UCase(vsf.TextMatrix(lngRow, mCol.���)), UCase(cbo(0).Text)) > 0 Then
                blnFind4 = True
            End If
        End If
        
        If blnFind1 And blnFind2 And blnFind3 And blnFind4 Then
            mlngFindRow = lngRow
            
            vsf.Row = mlngFindRow
            vsf.ShowCell vsf.Row, vsf.Col
            vsf.SetFocus
            
            Exit Function
        End If
        

    Next
    FindData = False
    
End Function

Private Sub chk_Click(Index As Integer)
    zlControl.TxtSelAll txt(1)
    txt(1).SetFocus
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 300 * 3)
    
    txt(1).Text = ""
    LocationObj txt(1)
End Sub

Private Sub cmdSelect_Click()
    
    If FindData Then
    
        If chk(0).Value = 1 Then
            vsf.TextMatrix(vsf.Row, mCol.����) = 1
            EditChanged = True
            Call RefreshState
        End If
        
        If chk(1).Value = 1 Then
            vsf.TextMatrix(vsf.Row, mCol.����) = 0
            EditChanged = True
            Call RefreshState
        End If
    End If
    zlControl.TxtSelAll txt(1)
    txt(1).SetFocus
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyA
            If tbrThis.Buttons("ȫѡ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫѡ"))
        Case vbKeyC
            If tbrThis.Buttons("ȫ��").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("ȫ��"))
        Case vbKeyB
            If tbrThis.Buttons("��ʼ").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("��ʼ"))
        Case vbKeyH
            If tbrThis.Buttons("����").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("����"))
        Case vbKeyX
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("�˳�").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("�˳�"))
        End If
    End If
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub Form_Load()

    glngFormW = 9870
    glngFormH = 6570
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    
    lbl(3).Caption = "&3." & (GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", "�� �� ��"))
    lbl(3).Tag = Mid(lbl(3).Caption, 4)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    With vsf
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = Me.ScaleWidth - .Left - fra.Width - 15
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra
        .Left = vsf.Left + vsf.Width + 15
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = vsf.Height + 90
    End With
    
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", lbl(3).Tag)
End Sub


Private Sub mnuFileClearAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.����) = 0
        End If
    Next
    
    EditChanged = False
End Sub

Private Sub mnuFileCome_Click()
    Call mnuFileSave_Click
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileSave_Click()
    
    If ValidEdit = False Then Exit Sub

    If SaveEdit() Then
        EditChanged = False
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub mnuFileSelectAll_Click()
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            vsf.TextMatrix(lngLoop, mCol.����) = 1
            EditChanged = True
        End If
    Next
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer

    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize

End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    mobjPopMenu.Add 1, "&1.�� �� ��", , , True, , (lbl(3).Tag = "�� �� ��")
    mobjPopMenu.Add 2, "&2.�� �� ��", , , True, , (lbl(3).Tag = "�� �� ��")
    mobjPopMenu.Add 3, "&3.���￨��", , , True, , (lbl(3).Tag = "���￨��")
    mobjPopMenu.Add 4, "&4.��    ��", , , True, , (lbl(3).Tag = "��    ��")
    mobjPopMenu.Add 5, "&5.����ƴ��", , , True, , (lbl(3).Tag = "����ƴ��")
    mobjPopMenu.Add 6, "&6.�������", , , True, , (lbl(3).Tag = "�������")
    mobjPopMenu.Add 7, "&7.���֤��", , , True, , (lbl(3).Tag = "���֤��")
    
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    
    Caption = Mid(Caption, 4)
    
    lbl(3).Caption = "&3." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
    lbl(3).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ȫѡ"
        Call mnuFileSelectAll_Click
    Case "ȫ��"
        Call mnuFileClearAll_Click
    Case "����"
        Call mnuFileSave_Click
    Case "����"
        Call mnuFileSave_Click
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_Change(Index As Integer)
    mlngFindRow = 0
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    
    cmdSelect.Default = True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCol As String
    Dim lngCol As Long
    Dim blnCard As Boolean
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    
    strCol = Mid(lbl(3).Caption, 4)
    lngCol = GetCol(vsf, strCol)
            
    If strCol = "���￨��" And KeyAscii <> vbKeyReturn Then
        '���￨�ţ��Զ�ʶ��

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.���￨���볤�� - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = 0
            Call cmdSelect_Click
        End If

    End If
    
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        zlCommFun.OpenIme False
    End Select
    
    cmdSelect.Default = False
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngLoop As Long
    
    If Abs(Val(vsf.TextMatrix(Row, mCol.����))) = 1 Then
        EditChanged = True
        Call RefreshState
        Exit Sub
    End If
        
    For lngLoop = 1 To vsf.Rows - 1
        If Abs(Val(vsf.TextMatrix(lngLoop, mCol.����))) = 1 Then
            EditChanged = True
            Call RefreshState
            Exit Sub
        End If
    Next
    
    If lngLoop = vsf.Rows Then EditChanged = False
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeySpace And vsf.Col <> mCol.���� Then
    
        If Abs(Val(vsf.TextMatrix(vsf.Row, mCol.����))) = 1 Then
            vsf.TextMatrix(vsf.Row, mCol.����) = 0
        Else
            vsf.TextMatrix(vsf.Row, mCol.����) = 1
        End If
        
        EditChanged = True
        
        Call RefreshState
            
    End If
End Sub

Private Sub vsf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    If vsf.MouseRow = 0 Then
        If mintSort = flexSortGenericAscending Then
            mintSort = flexSortGenericDescending
        Else
            mintSort = flexSortGenericAscending
        End If
        
        vsf.Sort = mintSort
    End If
    
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.���� Or Val(vsf.RowData(Row)) <= 0 Then
        Cancel = True
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

