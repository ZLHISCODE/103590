VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisStationAdjust 
   Caption         =   "�����������"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   Icon            =   "frmLisStationAdjust.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9015
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9015
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
         TabIndex        =   20
         Top             =   30
         Width           =   8895
         _ExtentX        =   15690
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
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&A.ȫѡ"
               Key             =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ"
               Object.Tag             =   "&A.ȫѡ"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&C.ȫ��"
               Key             =   "ȫ��"
               Object.ToolTipText     =   "ȫ��"
               Object.Tag             =   "&C.ȫ��"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&D.����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "&D.����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "&H.����"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&E.�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "&E.�˳�"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   4470
      Top             =   750
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
            Picture         =   "frmLisStationAdjust.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAdjust.frx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAdjust.frx":2166
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAdjust.frx":28E0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAdjust.frx":2B00
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   5460
      Top             =   750
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
            Picture         =   "frmLisStationAdjust.frx":2D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAdjust.frx":349A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAdjust.frx":3C14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAdjust.frx":438E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisStationAdjust.frx":45AE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   21
      Top             =   5415
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLisStationAdjust.frx":47CE
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10821
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
   Begin VSFlex8Ctl.VSFlexGrid Vsf 
      Height          =   2730
      Left            =   3615
      TabIndex        =   18
      Top             =   1440
      Width           =   4455
      _cx             =   7858
      _cy             =   4815
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
      BackColorSel    =   16768667
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
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
      OwnerDraw       =   0
      Editable        =   2
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fra 
      Height          =   4770
      Left            =   90
      TabIndex        =   22
      Top             =   615
      Width           =   3255
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   270
         TabIndex        =   5
         Top             =   1095
         Width           =   915
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3960
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.CommandButton cmdCalc 
         Caption         =   "����ֵ��ʽ��д(&J)"
         Height          =   350
         Left            =   270
         TabIndex        =   17
         Top             =   4320
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&P"
         Height          =   300
         Left            =   2835
         TabIndex        =   13
         Top             =   2895
         Width           =   300
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   270
         TabIndex        =   16
         Top             =   3960
         Width           =   2880
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   2880
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "��������(&S)"
         Height          =   350
         Left            =   270
         TabIndex        =   14
         Top             =   3255
         Width           =   1185
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2280
         Width           =   2880
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   270
         TabIndex        =   12
         Top             =   2895
         Width           =   2535
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1260
         TabIndex        =   6
         Top             =   1095
         Width           =   1890
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   435
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   204734467
         CurrentDate     =   38229
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Index           =   1
         Left            =   1860
         TabIndex        =   3
         Top             =   435
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   204734467
         CurrentDate     =   38229
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.�������ֵ��ʽ"
         Height          =   180
         Index           =   5
         Left            =   90
         TabIndex        =   15
         Top             =   3735
         Width           =   1620
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.�������"
         Height          =   180
         Index           =   6
         Left            =   90
         TabIndex        =   7
         Top             =   1470
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   4
         Left            =   1620
         TabIndex        =   2
         Top             =   495
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.������Ŀ"
         Height          =   180
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   2670
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.����       �걾��� "
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   855
         Width           =   1980
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.�걾ʱ��"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.��������"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   2055
         Width           =   900
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&A"
      Height          =   350
      Index           =   0
      Left            =   405
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&C"
      Height          =   350
      Index           =   1
      Left            =   405
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&D"
      Height          =   350
      Index           =   2
      Left            =   405
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2505
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&H"
      Height          =   350
      Index           =   3
      Left            =   405
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2850
      Width           =   1100
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&E"
      Height          =   350
      Index           =   4
      Left            =   405
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3210
      Width           =   1100
   End
End
Attribute VB_Name = "frmLisStationAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Private mblnOK As Boolean
Private mblnStartUp As Boolean
Private mfrmMain As Form
Private mlngLoop As Long
Private mRs As New ADODB.Recordset
Private mstrSql As String
Private mblnChangeEdit As Boolean
Private mstrName As String
Private mstrPrivs As String
Private mbyt������� As Byte
Private mlngDeptID As Long

Private Enum mCol
    ѡ�� = 0
    ����
    �걾��
    �걾ʱ��
    ԭʼ���
    ������
    �����־
    ����ο�
    ����걾id
    ��������
    �ϴν��
    ��������
    ��������
End Enum

Private Function ShowOpenTree() As Byte
    '-----------------------------------------------------------------------------------------
    '����:������+�б�ṹ��������Ŀ����
    '����:������2;�ɹ�����1;ȡ������0
    '-----------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim objPoint As POINTAPI
    
    On Error GoTo ErrHand
    
    strLvw = "����,1200,0,1;����,2700,0,0;Ӣ����д,900,0,0"

    ShowOpenTree = 2
    
    strSQL = "select * " & _
             "from (Select DISTINCT ID,�ϼ�ID,0 as ĩ��,����,���� ,'' as Ӣ������,'' as Ӣ����д,NULL+0 AS �������, " & _
                                   "DECODE(�ϼ�ID, Null, ID * POWER(10, 20), �ϼ�ID * POWER(10, 20) + ID) As ���� " & _
                     "From ���Ʒ���Ŀ¼ " & _
                    "Where ���� = 5 " & _
                    "Start With ID IN (SELECT DISTINCT ����id FROM ������ĿĿ¼ WHERE ��� = 'C') " & _
                   "Connect by Prior �ϼ�ID = ID " & _
                   "Union All " & _
                     "Select C.ID,A.����id AS �ϼ�ID,1 as ĩ��, A.����,A.����,C.Ӣ���� AS Ӣ������,D.��д AS Ӣ����д,D.�������, " & _
                            "1 AS ���� " & _
                       "FROM ������ĿĿ¼ A,���鱨����Ŀ B,����������Ŀ C,������Ŀ D " & _
                      "Where D.��Ŀ���<>2 AND A.ID=B.������Ŀid AND B.������Ŀid=C.ID AND C.ID=D.������Ŀid AND Nvl(A.�����Ŀ,0)=0 AND A.��� = 'C' AND (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) " & _
                   ") A " & _
            "ORDER BY A.ĩ��, A.����, A.����"
            
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rs.BOF Then Exit Function
    
    Call ClientToScreen(txt(2).hWnd, objPoint)
    
    If frmSelectExplorer.ShowSelect(Me, _
                            rs, _
                            objPoint.X * 15 - 30, objPoint.Y * 15 + txt(2).Height - 30, 5400, 2400, _
                            txt(2).Height, _
                            "������Ŀ����ѡ��", _
                            strLvw, _
                            "��ѡ��һ��������Ŀ") Then
        
        txt(2).Text = zlCommFun.Nvl(rs("����").Value)
        cmdOpen.Tag = zlCommFun.Nvl(rs("ID").Value)
        mbyt������� = zlCommFun.Nvl(rs("�������").Value)
        Select Case zlCommFun.Nvl(rs("�������").Value)
        Case 1      '������
            cbo(2).Visible = False
            txt(1).Visible = True
        Case 2      '����
            cbo(2).Visible = False
            txt(1).Visible = True
        Case 3      '����
            cbo(2).Visible = True
            txt(1).Visible = False
        End Select
        
        If mbyt������� = 1 Then
            txt(1).MaxLength = 0
            LocationObj txt(1)
        ElseIf mbyt������� = 2 Then
            txt(1).MaxLength = GetMaxLength("������ͨ���", "������")
            LocationObj txt(1)
        End If
        
        txt(2).Tag = ""
        ShowOpenTree = 1
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function OpenSelect(ByVal strText As String) As Byte
    '-----------------------------------------------------------------------------------------
    '����:���б�ṹ��������Ŀ����
    '����:������2;�ɹ�����1;ȡ������0
    '-----------------------------------------------------------------------------------------
    Dim strInput As String
    Dim rs As New ADODB.Recordset
    Dim strLvw As String
    Dim objPoint As POINTAPI
    
    On Error GoTo ErrHand
    
    OpenSelect = 2
    
    strLvw = "����,900,0,1;������Ŀ,3600,0,0;Ӣ����д,900,0,0"
    
    strInput = "%" & UCase(strText) & "%"
    mstrSql = "SELECT C.ID,A.����,A.���� AS ������Ŀ,C.Ӣ���� AS Ӣ������,D.��д AS Ӣ����д,D.�������  " & _
                "FROM ������ĿĿ¼ A,���鱨����Ŀ B,����������Ŀ C,������Ŀ D " & _
                "WHERE Nvl(A.�����Ŀ,0)=0 AND A.���='C' " & _
                    "AND A.ID=B.������ĿID " & _
                    "AND B.������ĿID=C.ID " & _
                    "AND C.ID=D.������Ŀid AND D.��Ŀ���<>2 " & _
                    "AND (A.���� LIKE [1] OR UPPER(D.��д) LIKE [1] OR A.���� LIKE [1] OR A.ID IN (" & _
                                                                                "SELECT ������Ŀid " & _
                                                                                    "FROM ������Ŀ���� " & _
                                                                                    "WHERE (���� LIKE [1] OR UPPER(����) LIKE UPPER([1]) )))"
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, strInput)
    If rs.BOF Then
        OpenSelect = 0
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
            
    Call ClientToScreen(txt(2).hWnd, objPoint)
    If frmSelectList.ShowSelect(Me, rs, strLvw, objPoint.X * 15 - 30, objPoint.Y * 15 + txt(2).Height - 30, 6000, 4200, Me.Name & "\������Ŀѡ��", "����±���ѡ��һ����Ŀ") Then
        GoTo Over
    End If
    Exit Function
Over:
    txt(2).Text = zlCommFun.Nvl(rs("������Ŀ").Value)
    cmdOpen.Tag = zlCommFun.Nvl(rs("ID").Value)
    mbyt������� = zlCommFun.Nvl(rs("�������").Value)
    Select Case zlCommFun.Nvl(rs("�������").Value)
    Case 1      '������
        cbo(2).Visible = False
        txt(1).Visible = True
    Case 2      '����
        cbo(2).Visible = False
        txt(1).Visible = True
    Case 3      '����
        cbo(2).Visible = True
        txt(1).Visible = False
    End Select
    
    If mbyt������� = 1 Then
        txt(1).MaxLength = 0
        LocationObj txt(1)
    ElseIf mbyt������� = 2 Then
        txt(1).MaxLength = GetMaxLength("������ͨ���", "������")
        LocationObj txt(1)
    End If
    txt(2).Tag = ""
    
    OpenSelect = 1
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub AdjustEnableState()
    '-----------------------------------------------------------------------------------------
    '����:�����޸�״̬���ð�ť���˵��ȵĿ���״̬
    '-----------------------------------------------------------------------------------------
    cmd(2).Enabled = True
        
    If mblnChangeEdit = False Then cmd(2).Enabled = False
        
    tbrThis.Buttons("����").Enabled = cmd(2).Enabled
        
End Sub

Private Sub RefreshStatus()
    '-----------------------------------------------------------------------------------------
    '����:
    '-----------------------------------------------------------------------------------------
    If Vsf.Rows = 2 And Trim(Vsf.TextMatrix(1, 2)) = "" Then
        stbThis.Panels(2).Text = "û�б걾��Ϣ��"
    Else
        stbThis.Panels(2).Text = "���ҵ� " & Vsf.Rows - 1 & " ���걾��Ϣ��"
    End If
    
End Sub

Public Function ShowEdit(ByVal frmMain As Form, Optional ByVal lngDeptID As Long = 0, Optional ByVal strPrivs As String) As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ���༭����
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    mlngDeptID = lngDeptID
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
                    
    mblnChangeEdit = False
    Call AdjustEnableState
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    Vsf.Cols = 0
    Call NewColumn(Vsf, "ѡ��", 510, 4)
    Call NewColumn(Vsf, "����", 500, 1)
    Call NewColumn(Vsf, "�걾��", 900, 1)
    Call NewColumn(Vsf, "�걾ʱ��", 1200, 1)
    Call NewColumn(Vsf, "ԭʼ���", 900, 1)
    Call NewColumn(Vsf, "������", 900, 1)
    Call NewColumn(Vsf, "�����־", 900, 1)
    Call NewColumn(Vsf, "����ο�", 1200, 1)
    Call NewColumn(Vsf, "����걾id", 0, 1)
    Call NewColumn(Vsf, "��������", 0, 1)
    Call NewColumn(Vsf, "�ϴν��", 0, 1)
    Call NewColumn(Vsf, "��������", 0, 1)
    Call NewColumn(Vsf, "��������", 0, 1)
    
    Vsf.ColDataType(mCol.ѡ��) = flexDTBoolean
    Vsf.SelectionMode = flexSelectionByRow
    
    cbo(2).AddItem "����"
    cbo(2).AddItem "����"
    cbo(2).ListIndex = 0
    
    mbyt������� = 1
    InitData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strWhere As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim lngloop As Long
    Dim varBetween As Variant
    
    Dim strStartDate As String
    Dim strEndDate As String
    Dim lngMachineID As Long
    Dim lngExecDept As Long
    
    Dim strSampleNoBegin As String
    Dim strSampleNoEnd As String
    
    Dim lngItem As Long
    Dim strName As String
    Dim strNO As String
    Dim strOtherNo As String
    Dim strErr As String
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    Vsf.Rows = 2
    Vsf.RowData(1) = 0
    Vsf.Cell(flexcpText, 1, 0, 1, Vsf.Cols - 1) = ""
   
    strStartDate = Format(dtp(0).Value, "yyyy-mm-dd 00:00:00")
    strEndDate = Format(dtp(1).Value, "yyyy-mm-dd 23:59:59")
        
    strWhere = " AND A.����ʱ�� BETWEEN [1] AND [2] "
    
    '����ID
    If cbo(0).ListCount > 0 Then
        If cbo(0).ListIndex = 0 Then
            strWhere = strWhere & " AND A.����id IS Null"
        Else
            strWhere = strWhere & " AND A.����id= [3] "
            lngMachineID = cbo(0).ItemData(cbo(0).ListIndex)
        End If
    End If
     
    '����
    If cbo(1).ListIndex > 0 Then
        strWhere = strWhere & " AND B.ִ�п���id+0= [4] "
        lngExecDept = cbo(1).ItemData(cbo(1).ListIndex)
    End If

    '------------------------  Beging           ����걾��
    strTmp = " 1=2 "
    If Trim(txt(0).Text) <> "" Then
        If Check_Sample = False Then Exit Function '10861
        varTmp2 = Split(Trim(txt(0).Text), ",")
        
        For lngloop = 0 To UBound(varTmp2)
            varBetween = Split(varTmp2(lngloop), "~")
            If UBound(varBetween) > 0 Then
                strSampleNoBegin = IIf(Trim(Me.txt(3)) <> "", TransSampleNO(Me.txt(3) & "-" & varBetween(0)), varBetween(0))
                strSampleNoEnd = IIf(Trim(Me.txt(3)) <> "", TransSampleNO(Me.txt(3) & "-" & varBetween(1)), varBetween(1))
                
                '�걾��Ŷ����Ϊ�ַ���
                strNO = GetSampleNOStr(strSampleNoBegin, strSampleNoEnd, strErr)
                If strErr <> "" Then
                    MsgBox strErr: Exit Function
                End If
                
                If InStr(strNO, ";") > 0 Then
                    For i = 0 To UBound(Split(strNO, ";"))
                        If Split(strNO, ";")(i) <> "" Then
                            strTmp = strTmp & " OR a.�걾��� In (Select /*+cardinality(b,10)*/ * From Table(Cast(f_Str2list('" & Split(strNO, ";")(i) & "') As Zltools.t_Strlist)) B) "
                        End If
                    Next
                Else
                    strTmp = strTmp & " OR a.�걾��� In (Select /*+cardinality(b,10)*/ * From Table(Cast(f_Str2list('" & strNO & "') As Zltools.t_Strlist)) B) "
                End If
            Else
                strSampleNoBegin = IIf(Trim(Me.txt(3)) <> "", TransSampleNO(Me.txt(3) & "-" & varTmp2(lngloop)), varTmp2(lngloop))
                strOtherNo = strOtherNo & "," & strSampleNoBegin
            End If
        Next
        
        If Mid(strOtherNo, 2) <> "" Then
            strTmp = strTmp & " OR a.�걾��� In (Select /*+cardinality(b,10)*/ * From Table(Cast(f_Str2list('" & Mid(strOtherNo, 2) & "') As Zltools.t_Strlist)) B) "
        End If
    Else
        If Trim(Me.txt(3)) <> "" Then
            MsgBox "�����ţ�����д�걾��ţ�"
            Exit Function
        End If
    End If
    If strTmp <> " 1=2 " Then strWhere = strWhere & " AND (" & strTmp & ")"
    '------------------------  End           ����걾��
    
    
    strWhere = strWhere & " AND F.������Ŀid= [5] "
    lngItem = Val(cmdOpen.Tag)
    
    If InStr(mstrPrivs, "�޸����˽��") = 0 Then
        strWhere = strWhere & " AND A.������=[6] "
        strName = UserInfo.����
    End If
    
    If InStr(mstrPrivs, "�޸����ս��") = 0 Then
        strWhere = strWhere & " AND TRUNC(A.����ʱ��)=TRUNC(SYSDATE)"
    End If
        
    mstrSql = "select DISTINCT A.������ AS ��������,F.����걾id,F.������ĿID AS ID,0 AS ѡ��,A.�걾��� AS �걾��," & _
                      "F.ԭʼ���," & _
                      "TO_CHAR(A.����ʱ��,'MM-DD HH24:MI') AS �걾ʱ��," & _
                      "DECODE(F.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') AS �����־," & _
                      "zlGetReference(F.������ĿID,A.�걾����,DECODE(E.�Ա�,'��',1,'Ů',2,0),E.��������,A.����ID,a.����,a.�������id) AS ����ο�," & _
                      "F.������,F.������ AS �ϴν��,Decode(A.�걾���,1,'��','') As ����,C.��������,C.�������� " & _
                      ",zl_to_number(zl_Get_Reference(1,F.������ĿID,A.�걾����,DECODE(E.�Ա�,'��',1,'Ů',2,0),E.��������,A.����ID)) as �ο�id " & _
                 "from ����걾��¼ A, ����ҽ����¼ B, �������� D,������ͨ��� F,������Ŀ C,������Ϣ E " & _
                "WHERE A.ҽ��ID = B.���ID AND A.ID=F.����걾id AND A.������=F.��¼���� AND " & _
                      "B.����id=E.����id AND " & _
                      "A.����ID = D.ID(+) AND A.����״̬ = 1 AND C.������ĿID=F.������Ŀid " & strWhere
    
    strWhere = " AND A.����ʱ�� BETWEEN [1] AND [2] "
    
    If cbo(0).ListCount > 0 Then strWhere = strWhere & _
        IIf(cbo(0).ListIndex = 0, " AND A.����id IS Null", " AND A.����id=[3] ")
    If cbo(1).ListIndex > 0 Then strWhere = strWhere & " AND A.ִ�п���id+0=[4] "
    
    If strTmp <> " 1=2 " Then strWhere = strWhere & " AND (" & strTmp & ")"
    strWhere = strWhere & " AND F.������Ŀid=[5] "
    
    mstrSql = mstrSql & " UNION ALL " & _
              "select A.������ AS ��������,F.����걾id,F.������ĿID AS ID,0 AS ѡ��,A.�걾��� AS �걾��," & _
                      "F.ԭʼ���," & _
                      "TO_CHAR(A.����ʱ��,'MM-DD HH24:MI') AS �걾ʱ��," & _
                      "DECODE(F.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') AS �����־," & _
                      "zlGetReference(F.������ĿID,A.�걾����,0,NULL,A.����ID,a.����,a.�������id) AS ����ο�," & _
                      "F.������,F.������ AS �ϴν��,Decode(A.�걾���,1,'��','') As ����,C.��������,C.�������� " & _
                      ",zl_to_number(zl_Get_Reference(1,F.������ĿID,A.�걾����,0,NULL,A.����ID)) as �ο�id " & _
                 "from ����걾��¼ A, �������� D,������ͨ��� F,������Ŀ C " & _
                "WHERE A.ҽ��id IS NULL AND A.ID=F.����걾id AND A.������=F.��¼���� AND " & _
                      "A.����ID = D.ID(+) AND A.����״̬ = 1 AND C.������ĿID=F.������Ŀid " & strWhere
    mstrSql = "SELECT a.��������,a.����걾id,a.ID,a.ѡ��,a.�걾��,a.ԭʼ���,a.�걾ʱ��,a.�����־,a.����ο�,a.������,a.�ϴν��,a.����,f.��ʾ���� as ��������,f.��ʾ���� as �������� FROM (" & mstrSql & ") A,������Ŀ�ο� F Where a.�ο�id=F.ID(+) ORDER BY A.�걾ʱ��,A.�걾�� "
    
    Set rs = zlDatabase.OpenSQLRecord(mstrSql, Me.Caption, CDate(strStartDate), CDate(strEndDate), lngMachineID, lngExecDept, lngItem, strName)
    If rs.BOF = False Then
        Call FillGrid(Vsf, rs, Array("", "", "", "", "0.0000", "0.0000"))
    End If
            
    For mlngLoop = 1 To Vsf.Rows - 1
        Call ApplyResultColor(Vsf, mlngLoop, mCol.������, _
            Decode(Vsf.TextMatrix(mlngLoop, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
    Next
    
    ReadData = True
    
    Exit Function
    
ErrHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ValidData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim strError As String
    Dim rs As New ADODB.Recordset
    
    '�����������ֵ�Ƿ���ȷ,��Ҫ�Ǽ��鹫ʽ
    
    On Error GoTo ErrHand
    If Trim(txt(1)) = "" Then
        MsgBox "�����������ʽ!", vbInformation, Me.Caption
        Me.txt(1).SetFocus
        Exit Function
    End If
    
    Set rs = zlDatabase.OpenSQLRecord("SELECT " & ReplaceAll(txt(1).Text, "R", "10") & " FROM DUAL", Me.Caption)
            
    ValidData = True
    
    Exit Function
ErrHand:
    LocationObj txt(1)
    strError = "�������ֵ��ʽ���Ϸ���"
    MsgBox strError, vbInformation, gstrSysName
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strNow As String
    Dim strResult As String
    Dim bytResultFlag As Byte
    Dim strSQL() As String
        
    On Error GoTo ErrHand
    ReDim strSQL(1 To 1)
    
    strNow = Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    For mlngLoop = 1 To Vsf.Rows - 1
        If Abs(Val(Vsf.TextMatrix(mlngLoop, mCol.ѡ��))) = 1 And Val(Vsf.RowData(mlngLoop)) > 0 Then
            
            bytResultFlag = Decode(Vsf.TextMatrix(mlngLoop, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1)
            If mbyt������� = 1 Then
                strResult = Val(Vsf.TextMatrix(mlngLoop, mCol.������))
                strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_������д(" & Val(Vsf.TextMatrix(mlngLoop, mCol.����걾id)) & "," & Val(Vsf.RowData(mlngLoop)) & _
                                            "," & Val(Vsf.TextMatrix(mlngLoop, mCol.��������)) & ",'" & strResult & _
                                            "',TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & bytResultFlag & ",'" & _
                                            Vsf.TextMatrix(mlngLoop, mCol.����ο�) & "',2,NULL,1,0,Null,Null,Null,Null,Null,Null,Null,Null,'" & UserInfo.���� & "')"
            Else
                strResult = Vsf.TextMatrix(mlngLoop, mCol.������)
                strSQL(ReDimArray(strSQL)) = "ZL_����걾��¼_������д(" & Val(Vsf.TextMatrix(mlngLoop, mCol.����걾id)) & "," & Val(Vsf.RowData(mlngLoop)) & _
                                            "," & Val(Vsf.TextMatrix(mlngLoop, mCol.��������)) & ",'" & strResult & _
                                            "',TO_DATE('" & strNow & "','yyyy-mm-dd hh24:mi:ss')," & bytResultFlag & ",'" & _
                                            Vsf.TextMatrix(mlngLoop, mCol.����ο�) & "',2,NULL,0,0,Null,Null,Null,Null,Null,Null,Null,Null,'" & UserInfo.���� & "')"
            End If
        End If
    Next
    
    blnTran = True
    
'    gcnOracle.BeginTrans
    For mlngLoop = 1 To UBound(strSQL)
        If strSQL(mlngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(mlngLoop), Me.Caption)
    Next
'    gcnOracle.CommitTrans
    
    SaveData = True
    
    Exit Function
    
ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    
End Function


Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd_Click(Index As Integer)
    
    Select Case Index
    Case 0
        For mlngLoop = 1 To Vsf.Rows - 1
            If Val(Vsf.RowData(mlngLoop)) > 0 Then
                Vsf.TextMatrix(mlngLoop, mCol.ѡ��) = 1
            End If
        Next
        
        mblnChangeEdit = True
        Call AdjustEnableState
    Case 1
        For mlngLoop = 1 To Vsf.Rows - 1
            Vsf.TextMatrix(mlngLoop, mCol.ѡ��) = 0
        Next
        mblnChangeEdit = False
        Call AdjustEnableState
    Case 2
        If mblnChangeEdit Then
            
            If MsgBox("��ȷ��Ҫ����ʽ<" & cbo(2).Text & ">���е�����", vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
            
            If AdjustResult = False Then Exit Sub
            If SaveData() = False Then Exit Sub
            
            mblnOK = True
            
            mblnChangeEdit = False
            Call AdjustEnableState

            ShowSimpleMsg "���������ɹ���"
            
            Exit Sub
        End If
        
    Case 3
        ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
    Case 4
        Unload Me
    End Select
End Sub

Private Function CalcNewValue(ByVal str���ʽ As String, ByVal lngRow As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    
    Dim strResult As String
    Dim strReference As String
    Dim rs As New ADODB.Recordset
        
    If mbyt������� = 1 Then
        strResult = ReplaceAll(str���ʽ, "R", Val(Vsf.TextMatrix(lngRow, mCol.�ϴν��)))
        
        Set rs = zlDatabase.OpenSQLRecord("SELECT " & strResult & " FROM DUAL", Me.Caption)
        If rs.BOF = False Then strResult = zlCommFun.Nvl(rs.Fields(0).Value, "")
        
        strResult = Format(strResult, "0.0000")
    Else
        strResult = str���ʽ
    End If
                
    Vsf.TextMatrix(lngRow, mCol.������) = strResult
            
    'Ҫ���¼���������ͼ����־
    strReference = Trim(Vsf.TextMatrix(lngRow, mCol.����ο�))
    Vsf.TextMatrix(lngRow, mCol.�����־) = ""
    
    If InStr(strReference, vbCrLf) > 0 Then strReference = Mid(strReference, 1, InStr(strReference, vbCrLf) - 1)
                
    If mbyt������� = 1 Then
        
        Vsf.TextMatrix(lngRow, mCol.�����־) = ""
        
        '�����ֵ�;����ֵ����
        If Trim(Vsf.TextMatrix(lngRow, mCol.��������)) <> "" Then
            If Val(strResult) > Val(Vsf.TextMatrix(lngRow, mCol.��������)) Then
                Vsf.TextMatrix(lngRow, mCol.�����־) = "����"
            End If
        End If
        If Trim(Vsf.TextMatrix(lngRow, mCol.��������)) <> "" Then
            If Val(strResult) < Val(Vsf.TextMatrix(lngRow, mCol.��������)) Then
                Vsf.TextMatrix(lngRow, mCol.�����־) = "����"
            End If
        End If
        
        If InStr(strReference, "��") > 0 And Vsf.TextMatrix(lngRow, mCol.�����־) = "" Then
            
            '���С�ڲο���ֵ
            If Val(strResult) < Val(Mid(strReference, 1, InStr(strReference, "��") - 1)) Then Vsf.TextMatrix(lngRow, mCol.�����־) = "��"
                            
            '������ڲο���ֵ
            If Val(strResult) > Val(Mid(strReference, InStr(strReference, "��") + 1)) Then Vsf.TextMatrix(lngRow, mCol.�����־) = "��"
                                
        End If
    ElseIf mbyt������� = 3 Then
        If Trim(strResult) = Trim(strReference) Then
            Vsf.TextMatrix(lngRow, mCol.�����־) = ""
        Else
            Vsf.TextMatrix(lngRow, mCol.�����־) = "�쳣"
        End If
    End If
    
    Call ApplyResultColor(Vsf, lngRow, mCol.������, _
        Decode(Vsf.TextMatrix(lngRow, mCol.�����־), "��", 3, "��", 2, "�쳣", 4, "����", 5, "����", 6, 1))
            
End Function

Private Function AdjustResult() As Boolean
    Dim strResult As String
    Dim strReference As String
    Dim rs As New ADODB.Recordset
    
    '��Ч��
    If mbyt������� = 1 Then
        If ValidData = False Then Exit Function
    End If
    
    For mlngLoop = 1 To Vsf.Rows - 1
        If Abs(Val(Vsf.TextMatrix(mlngLoop, mCol.ѡ��))) = 1 And Val(Vsf.RowData(mlngLoop)) > 0 Then
            
            '����ʽ����
            Select Case mbyt�������
            Case 1, 2
                Call CalcNewValue(txt(1).Text, mlngLoop)
            Case Else
                Call CalcNewValue(cbo(2).Text, mlngLoop)
            End Select

        End If
    Next
    
    AdjustResult = True
    
End Function

Private Sub cmdCalc_Click()
    Call AdjustResult
End Sub

Private Sub cmdOpen_Click()
    If ShowOpenTree = 1 Then mstrName = txt(2).Text
    txt(2).SetFocus
End Sub

Private Sub cmdRefresh_Click()
    'û����д��Ŀʱ��ʾ
    If Trim(Me.txt(2).Text) = "" Then
        MsgBox "����дһ����Ҫ��������Ŀ!", vbInformation, Me.Caption
        Me.txt(2).SetFocus
        Exit Sub
    End If
    Call ReadData
        
    mblnChangeEdit = False
    Call AdjustEnableState
    Call RefreshStatus
    
    Vsf.Col = 1
    Vsf.SetFocus
    Vsf.Col = 0
End Sub


Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Activate()
    Dim rs As New ADODB.Recordset
    Dim lngDefaultDev As Long
    Dim ControlcboDept As CommandBarComboBox
    Dim strSQL As String
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    '���鲿��
    strSQL = "select A.����||'-'||A.����,ID from ���ű� a where A.id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID)
    cbo(1).Clear
    If rs.EOF = False Then
        Call AddComboData(cbo(1), rs, False)
        cbo(1).ListIndex = 0
    End If

        
    '��������
    strSQL = "SELECT A.����||'-'||A.����,ID FROM �������� A where ʹ��С��id = [1] ORDER BY A.����||'-'||A.����"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID)
    cbo(0).AddItem "�ֹ�"
    If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
    lngDefaultDev = Val(Split(GetConnectDevs & ";1", ";")(0))
    cbo(0).ListIndex = FindComboItem(cbo(0), lngDefaultDev)
    If cbo(0).ListCount > 0 And cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
            
    dtp(0).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtp(1).Value = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    
    If cbo(1).ListIndex = -1 Then
        zlControl.CboLocate cbo(1), UserInfo.����ID, True
        If cbo(1).ListCount > 0 And cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    End If
    
    txt(0).Text = ""
    txt(2).Text = ""
    dtp(0).SetFocus
    
End Sub

Private Sub Form_Load()
    
    Call RestoreWinState(Me, App.ProductName)
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fra
        .Left = 0
        .Top = cbrThis.Height - 90
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
    
    With Vsf
        .Left = fra.Left + fra.Width
        .Top = cbrThis.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    If mblnChangeEdit Then
'        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
'        If Cancel Then Exit Sub
'    End If
    Me.txt(2).Text = "": Me.txt(2).Tag = ""
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "ȫѡ"
        Call cmd_Click(0)
    Case "ȫ��"
        Call cmd_Click(1)
    Case "����"
        Call cmd_Click(2)
    Case "����"
        Call cmd_Click(3)
    Case "�˳�"
        Call cmd_Click(4)
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 2 Then txt(2).Tag = "Changed"
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strInput As String
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    If KeyAscii = vbKeyReturn Then
        If Index = 2 Then
            If txt(2).Tag <> "" Then
                txt(2).Tag = ""
                Select Case OpenSelect(txt(2).Text)
                Case 0
                    'û��ƥ�����Ŀ
                    MsgBox "û���ҵ���ƥ��Ľ����", vbInformation, gstrSysName
                    txt(2).Text = mstrName
                    
                Case 1
                    'ѡȡ��һ����Ŀ
                    mstrName = txt(2).Text
                Case 2
                    'ȡ���˱���ѡ��
                    txt(2).Text = mstrName
                End Select
            Else
                zlCommFun.PressKey vbKeyTab
                zlCommFun.PressKey vbKeyTab
            End If
            txt(2).Tag = ""
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case Index
        Case 0
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "01234567890,~")
        Case 1
            If mbyt������� = 1 Then
                KeyAscii = FilterKeyAscii(KeyAscii, 99, "R0123456789.+-*/)(")
            End If
        End Select
    End If
    
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Index = 0 Then Cancel = Not Check_Sample
    
    If Index = 2 Then
        If (txt(2).Tag = "Changed") Then txt(2).Text = mstrName
    End If
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    If Abs(Val(Vsf.TextMatrix(Row, mCol.ѡ��))) = 1 Then
        mblnChangeEdit = True
        Call AdjustEnableState
        Exit Sub
    End If
    
    For mlngLoop = 1 To Vsf.Rows - 1
        If Abs(Val(Vsf.TextMatrix(mlngLoop, mCol.ѡ��))) = 1 Then
            mblnChangeEdit = True
            Call AdjustEnableState
            Exit Sub
        End If
    Next
    
    If mlngLoop = Vsf.Rows Then
        mblnChangeEdit = False
        Call AdjustEnableState
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    
'    If NewRow + 1 > Vsf.FixedRows And OldRow + 1 > Vsf.FixedRows Then
'        Vsf.Cell(flexcpBackColor, OldRow, 0, OldRow, Vsf.Cols - 1) = Vsf.BackColor
'        Vsf.Cell(flexcpBackColor, NewRow, 0, NewRow, Vsf.Cols - 1) = Vsf.BackColorSel
'    End If
    
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(Vsf.RowData(Row)) = 0 Then Cancel = True
    If Col <> 0 Then Cancel = True
    
End Sub

Private Function Check_Sample() As Boolean
    '   10861 �������Ƿ�Ӧ����%,?���ַ�������
    Dim i As Long, str�ַ� As String
    str�ַ� = ""
    If Len(txt(0)) > 0 Then
        For i = 1 To Len(txt(0))
            If InStr("0123456789,~", Mid(txt(0), i, 1)) <= 0 Then
                str�ַ� = str�ַ� & Mid(txt(0), i, 1)
            End If
        Next
    End If

    If str�ַ� <> "" Then
        MsgBox "��������" & str�ַ�, vbQuestion, gstrSysName
        Check_Sample = False
    Else
        Check_Sample = True
    End If

End Function
