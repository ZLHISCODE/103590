VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStuffRxList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList imgList 
      Left            =   4800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":0000
            Key             =   "��ӡ11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":039A
            Key             =   "��ǰ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":6BFC
            Key             =   "ָʾ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":D45E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":D9F8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":DD92
            Key             =   "��־"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":E12C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":E4C6
            Key             =   "ͼ��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":E860
            Key             =   "ѡ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":F272
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":15AD4
            Key             =   "δ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":1C336
            Key             =   "�ڼ�"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":22B98
            Key             =   "�Ѽ�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":293FA
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":2FC5C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":2FFF6
            Key             =   "����_ѡ��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":30390
            Key             =   "�ײ�"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":36BF2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":3D454
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":43CB6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":4A518
            Key             =   "ָ��"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":50D7A
            Key             =   "���"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":575DC
            Key             =   "������ʽ"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":5DE3E
            Key             =   "�����ļ�"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":646A0
            Key             =   "����"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":6AF02
            Key             =   "�շ�"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":6B914
            Key             =   "���"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":72176
            Key             =   "����"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":789D8
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":7F23A
            Key             =   "��ʼ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":85A9C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8C2FE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8C698
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8CA32
            Key             =   "�����ܼ�"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8CDCC
            Key             =   "ȫ���ܼ�"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8D166
            Key             =   "�ܼ�"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8D500
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8DF12
            Key             =   "�Ѿ���ӡ"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8E924
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   1455
      Left            =   4080
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   2566
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
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmStuffRxList.frx":8EEBE
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2160
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2400
      _cx             =   4233
      _cy             =   3810
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStuffRxList.frx":8EF0C
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
   Begin VB.Image imgSel 
      Height          =   195
      Left            =   4320
      Picture         =   "frmStuffRxList.frx":8EF81
      ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
      Top             =   840
      Width           =   195
   End
End
Attribute VB_Name = "frmStuffRxList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��������
Private Const mconIntCol���� = 25
Private mIntCol��ǰ�� As Integer
Private mIntCol�������� As Integer
Private mIntCol��־ As Integer
Private mIntCol���� As Integer
Private mIntCol���� As Integer
Private mIntCol�շ� As Integer
Private mIntColNO As Integer
Private mIntCol���� As Integer
Private mIntCol��� As Integer
Private mIntCol���� As Integer
Private mIntCol�ɲ��� As Integer
Private mIntCol˵�� As Integer
Private mIntCol���￨�� As Integer
Private mIntCol����� As Integer
Private mIntCol���֤ As Integer
Private mIntColIC�� As Integer
Private mIntCol����ID As Integer
Private mIntColҽ���� As Integer
Private mIntColסԺ�� As Integer
Private mIntColʵ�ս��  As Integer
Private mIntCol�����־  As Integer
Private mIntCol��¼����  As Integer
Private mIntCol�շ����  As Integer
Private mIntCol�ⷿID  As Integer
Private mIntCol��¼״̬  As Integer

Private Const glng��ҩ As Long = &HC0&
Private Const glng��ҩ As Long = &HC00000
Private Const glng���� As Long = &H80000008

'�����е����صı���
Private mstrUnallowSetColHide As String
Private mstrUnallowShow As String

'ҳ�泣��
Private Enum mListType
    ������ = 0
    ���� = 1
End Enum

Private Enum mFindType
    ���ݺ� = 0
    ����� = 1
    ���� = 2
    ���֤ = 3
    IC�� = 4
    ҽ���� = 5
    סԺ�� = 6
End Enum

Private Type FindProcess
    FindType As Integer
    FindContent As String
    StartRow As Integer
End Type
Private mFindProcess As FindProcess

Private mintType As Integer               '��ǰҳ������
Private mlng�ⷿid As Long
Private mFMT As g_FmtString

Private Sub Form_Load()
    Call InitVsfList
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.vsfList.Move 0, 0, Me.Width, Me.Height

    err.Clear
End Sub

Private Sub InitVsfList()
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol��ǰ�� = 0
    mIntCol�������� = 1
    mIntCol��־ = 2
    mIntCol���� = 3
    mIntCol���� = 4
    mIntCol�շ� = 5
    mIntColNO = 6
    mIntCol���� = 7
    mIntCol��� = 8
    mIntCol���� = 9
    mIntCol�ɲ��� = 10
    mIntCol˵�� = 11
    mIntCol���￨�� = 12
    mIntCol����� = 13
    mIntCol���֤ = 14
    mIntColIC�� = 15
    mIntCol����ID = 16
    mIntColҽ���� = 17
    mIntColסԺ�� = 18
    mIntColʵ�ս�� = 19
    mIntCol�����־ = 20
    mIntCol��¼���� = 21
    mIntCol�շ���� = 22
    mIntCol�ⷿID = 23
    mIntCol��¼״̬ = 24
    
    '�ָ��û��Զ�����˳��
    str������ = LoadListColState
    If str������ <> "" Then
        arr������ = Split(str������, "|")
        If UBound(arr������) + 1 <> mconIntCol���� Then
            str������ = ""
        Else
            For n = 0 To UBound(arr������)
                SetColumnValue Split(arr������(n), ",")(0), n
            Next
        End If
    End If
     
    '��ʼ��δ��ҩ�嵥
    With vsfList
        .Redraw = flexRDNone
        
        .Rows = 1
        .Rows = 2
        .Cols = mconIntCol����
        
        .Cell(flexcpPicture, 1, mIntCol��ǰ��, 1, mIntCol��ǰ��) = Me.imgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol��ǰ��, .Rows - 1, mIntCol��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList, mIntCol��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"

        VsfGridColFormat vsfList, mIntCol��������, "��������", 0, flexAlignCenterCenter, "��������"
        VsfGridColFormat vsfList, mIntCol��־, "1", 0, flexAlignCenterCenter, "��־"
        VsfGridColFormat vsfList, mIntCol����, "���", 1000, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfList, mIntCol����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList, mIntCol�շ�, "�շ�", 0, flexAlignCenterCenter, "�շ�"
        VsfGridColFormat vsfList, mIntColNO, "NO", 800, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList, mIntCol����, "����", 800, flexAlignLeftCenter, "����"
        
        VsfGridColFormat vsfList, mIntCol���, "���", 1200, flexAlignRightCenter, "���"
        VsfGridColFormat vsfList, mIntCol����, "����", 1500, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList, mIntCol�ɲ���, "�ɲ���", 0, flexAlignCenterCenter, "�ɲ���"
        VsfGridColFormat vsfList, mIntCol˵��, "˵��", 1500, flexAlignLeftCenter, "˵��"
        VsfGridColFormat vsfList, mIntCol���￨��, "���￨��", 1000, flexAlignLeftCenter, "���￨��"
        VsfGridColFormat vsfList, mIntCol�����, "�����", 1000, flexAlignLeftCenter, "�����"
        VsfGridColFormat vsfList, mIntCol���֤, "���֤", 1600, flexAlignLeftCenter, "���֤"
        VsfGridColFormat vsfList, mIntColIC��, "IC��", 1600, flexAlignLeftCenter, "IC��"
        VsfGridColFormat vsfList, mIntCol����ID, "����ID", 0, flexAlignCenterCenter, "����ID"
        VsfGridColFormat vsfList, mIntColҽ����, "ҽ����", 1500, flexAlignLeftCenter, "ҽ����"
        VsfGridColFormat vsfList, mIntColסԺ��, "סԺ��", 1000, flexAlignLeftCenter, "סԺ��"
        
        VsfGridColFormat vsfList, mIntColʵ�ս��, "ʵ�ս��", 0, flexAlignCenterCenter, "ʵ�ս��"
        VsfGridColFormat vsfList, mIntCol�����־, "�����־", 0, flexAlignCenterCenter, "�����־"
        VsfGridColFormat vsfList, mIntCol��¼����, "��¼����", 0, flexAlignCenterCenter, "��¼����"
        VsfGridColFormat vsfList, mIntCol�շ����, "�շ�����", 0, flexAlignCenterCenter, "�շ�����"
        VsfGridColFormat vsfList, mIntCol�ⷿID, "�ⷿid", 0, flexAlignCenterCenter, "�ⷿid"
        VsfGridColFormat vsfList, mIntCol��¼״̬, "��¼״̬", 0, flexAlignCenterCenter, "��¼״̬"
        
        
        mstrUnallowSetColHide = "NO"
        mstrUnallowShow = "��ǰ��;��������;��־;����;�շ�;�ɲ���;����ID;δ���;ʵ�ս��;�����־;��¼����;�շ�����;�ⷿid"

        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If InStr(mstrUnallowShow, Split(arr������(n), ",")(0)) > 0 Then
                    For i = 0 To vsfList.Cols - 1
                        If Split(arr������(n), ",")(0) = vsfList.ColKey(i) Then
                            vsfList.ColWidth(i) = Val(Split(arr������(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        .RowSel = 1
        .Redraw = flexRDDirect
    End With
End Sub


Private Function LoadListColState() As String
    Dim strType As String
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    
    Select Case mintType
        Case mListType.������
            strType = "������"
        Case mListType.����
            strType = "����"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name & "\" & TypeName(vsfList), strType, "")
End Function


Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer)
    Select Case str����
        Case "���"
            mIntCol���� = intValue
        Case "NO"
            mIntColNO = intValue
        Case "����"
            mIntCol���� = intValue
            
        Case "���"
            mIntCol��� = intValue
        Case "����"
            mIntCol���� = intValue

        Case "˵��"
            mIntCol˵�� = intValue
        Case "���￨��"
            mIntCol���￨�� = intValue
        Case "�����"
            mIntCol����� = intValue
    
        Case "���֤"
            mIntCol���֤ = intValue
        Case "IC��"
            mIntColIC�� = intValue
    End Select
                
End Sub


Public Sub RefreshList(ByVal rsTemp As Recordset, ByVal intType As Integer)
'���ܣ�ʵ�ֽ����ݱ��ֵ�vsf���
'����1��rsTemp��ˢ���б�����ݼ�
'����2��intType�Ǵ���ǰ��ҵ�����ͣ�1-����ҩ��2-�ѷ�ҩ
    Dim intCount As Integer
    Dim strType As String
    Dim lngColor As Long
    Dim int�ɲ��� As Integer
    Dim dblMoney As Double
    Dim str�ϼ� As String
    
    On Error GoTo ErrHandle
    
    mintType = intType
    
    With Me.vsfList
        .Rows = 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        intCount = 1
        
        Do While Not rsTemp.EOF
            .TextMatrix(intCount, .ColIndex("��ǰ��")) = intCount
            If rsTemp!���� = 24 Then
                strType = "�շ�"
            Else
                strType = "����"
            End If
            
            If rsTemp!���շ� = 0 Then strType = "��δ��" & strType
            
            .TextMatrix(intCount, .ColIndex("���")) = strType
            .TextMatrix(intCount, .ColIndex("��־")) = rsTemp!�����־
            .TextMatrix(intCount, .ColIndex("����")) = rsTemp!����
            .TextMatrix(intCount, .ColIndex("�շ�")) = rsTemp!���շ�
            
            .TextMatrix(intCount, .ColIndex("NO")) = rsTemp!NO
            .TextMatrix(intCount, .ColIndex("����")) = rsTemp!����
            .TextMatrix(intCount, .ColIndex("���")) = rsTemp!���
            .TextMatrix(intCount, .ColIndex("����")) = rsTemp!����
            .TextMatrix(intCount, .ColIndex("���￨��")) = NVL(rsTemp!���￨��)
            .TextMatrix(intCount, .ColIndex("�����")) = NVL(rsTemp!�����)
            .TextMatrix(intCount, .ColIndex("���֤")) = NVL(rsTemp!���֤��)
            .TextMatrix(intCount, .ColIndex("IC��")) = NVL(rsTemp!IC����)
            .TextMatrix(intCount, .ColIndex("����id")) = NVL(rsTemp!����id)
            .TextMatrix(intCount, .ColIndex("ҽ����")) = NVL(rsTemp!ҽ����)
            .TextMatrix(intCount, .ColIndex("סԺ��")) = NVL(rsTemp!סԺ��)
            .TextMatrix(intCount, .ColIndex("�����־")) = rsTemp!�����־
            .TextMatrix(intCount, .ColIndex("ʵ�ս��")) = Format(rsTemp!���, mFMT.FM_���)
            .TextMatrix(intCount, .ColIndex("��¼����")) = rsTemp!��¼����
            .TextMatrix(intCount, .ColIndex("�ⷿid")) = rsTemp!�ⷿID
            .TextMatrix(intCount, .ColIndex("��¼״̬")) = rsTemp!��¼״̬
            .TextMatrix(intCount, .ColIndex("˵��")) = rsTemp!˵��
            
            dblMoney = dblMoney + Val(rsTemp!���)
            
            '�жϵ�ǰ����ҩ����
            If rsTemp!��¼״̬ = 1 Or rsTemp!��¼״̬ Mod 3 = 0 Then
                int�ɲ��� = 1
            Else
                int�ɲ��� = rsTemp!��¼״̬ Mod 3 + 1
            End If
            .TextMatrix(intCount, .ColIndex("�ɲ���")) = int�ɲ���
            
            
            '������������ɫ
            '������ɫ
            lngColor = IIf(intType = 0 Or int�ɲ��� = 0, &H80000008, IIf(int�ɲ��� = 1, glng����, IIf(int�ɲ��� = 2, glng��ҩ, glng��ҩ)))
            .Cell(flexcpForeColor, intCount, 1, intCount, .Cols - 1) = lngColor
              
            rsTemp.MoveNext
            intCount = intCount + 1
        Loop
        .Row = 1
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfList_EnterCell()
'ˢ����ϸ��Ϣ
    Dim rsTemp As Recordset
    
    With Me.vsfList
        If .Row = 0 Then Exit Sub
        Call frmStuffRxSend.RefreshSendData(Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("�ⷿid"))), Val(.TextMatrix(.Row, .ColIndex("��¼״̬"))), Val(.TextMatrix(.Row, .ColIndex("�ɲ���"))))
    End With
End Sub


Public Sub SetFontSize(ByVal intFont As Integer)
    Me.vsfList.FontSize = intFont
End Sub

Public Function FindSpecialRow(ByVal intFindType As Integer, ByVal strFindContent As String) As Boolean
    '��������п�����strFindContent��ʽΪ����ID|����
    Dim intCol As Integer
    Dim intFindRow As Integer
    
    With mFindProcess
        .FindType = intFindType
        .FindContent = UCase(strFindContent)
        .StartRow = 1
    End With
    
    With vsfList
        Select Case mFindProcess.FindType
            Case mFindType.����
                intCol = mIntCol����
                
                If zlCommFun.IsCharAlpha(mFindProcess.FindContent) Then
'                    'ȫ��ĸʱƥ�����
'                    If zlDatabase.GetPara("���뷽ʽ") = 0 Then
'                        intCol = mIntColƴ������
'                    Else
'                        intCol = mIntCol��ʼ���
'                    End If
                End If
            Case mFindType.���ݺ�
                intCol = mIntColNO
            Case mFindType.�����
                intCol = mIntCol�����
            Case mFindType.���֤
                intCol = mIntCol���֤
            Case mFindType.IC��
                intCol = mIntCol����ID
            Case mFindType.ҽ����
                intCol = mIntColҽ����
            Case mFindType.סԺ��
                intCol = mIntColסԺ��
            Case Else
                '����Ϊ���ѿ���������ID����
                intCol = mIntCol����ID
                mFindProcess.FindContent = zlfuncCard_GetPatiID(Val(Split(strFindContent, "|")(0)), Split(strFindContent, "|")(1))
        End Select
        
        mFindProcess.StartRow = .FindRow(mFindProcess.FindContent, mFindProcess.StartRow, intCol)
        
        If mFindProcess.StartRow > 0 Then
            .Row = mFindProcess.StartRow
            .TopRow = .Row
            FindSpecialRow = True
            If mFindProcess.StartRow + 1 >= .Rows Then
                mFindProcess.StartRow = 1
            Else
                mFindProcess.StartRow = mFindProcess.StartRow + 1
            End If
        Else
            mFindProcess.StartRow = 1
        End If
        
    End With
End Function
