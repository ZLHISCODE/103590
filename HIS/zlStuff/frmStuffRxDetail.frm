VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStuffRxDetail 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   9135
   End
   Begin VB.ComboBox cboPre 
      Height          =   300
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   97
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   1920
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
            Picture         =   "frmStuffRxDetail.frx":0000
            Key             =   "��ӡ11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":039A
            Key             =   "��ǰ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":6BFC
            Key             =   "ָʾ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":D45E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":D9F8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":DD92
            Key             =   "��־"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":E12C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":E4C6
            Key             =   "ͼ��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":E860
            Key             =   "ѡ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":F272
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":15AD4
            Key             =   "δ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":1C336
            Key             =   "�ڼ�"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":22B98
            Key             =   "�Ѽ�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":293FA
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":2FC5C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":2FFF6
            Key             =   "����_ѡ��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":30390
            Key             =   "�ײ�"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":36BF2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":3D454
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":43CB6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":4A518
            Key             =   "ָ��"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":50D7A
            Key             =   "���"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":575DC
            Key             =   "������ʽ"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":5DE3E
            Key             =   "�����ļ�"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":646A0
            Key             =   "����"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":6AF02
            Key             =   "�շ�"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":6B914
            Key             =   "���"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":72176
            Key             =   "����"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":789D8
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":7F23A
            Key             =   "��ʼ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":85A9C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8C2FE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8C698
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8CA32
            Key             =   "�����ܼ�"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8CDCC
            Key             =   "ȫ���ܼ�"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8D166
            Key             =   "�ܼ�"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8D500
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8DF12
            Key             =   "�Ѿ���ӡ"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxDetail.frx":8E924
            Key             =   "����"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFDetail 
      Height          =   5640
      Left            =   30
      TabIndex        =   7
      Top             =   600
      Width           =   8760
      _cx             =   15452
      _cy             =   9948
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStuffRxDetail.frx":8EEBE
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
      ExplorerBar     =   2
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
   Begin VB.Label lblPre 
      Caption         =   "������"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblDept 
      Caption         =   "���ң�"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblAge 
      Caption         =   "���䣺"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblSex 
      Caption         =   "�Ա�"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblName 
      Caption         =   "������"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmStuffRxDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�������б���
'�������ƣ���񣬲��أ����ţ����������������ۣ����
Private Const mconIntCol���� = 32
Private mIntCol��ǰ�� As Integer
Private mIntCol˳��� As Integer
Private mIntCol�������� As Integer
Private mIntCol������ As Integer
Private mIntColӢ���� As Integer
Private mintcol��� As Integer
Private mintcol��� As Integer
Private mintcol���� As Integer
Private mIntColId As Integer
Private mintcol����id As Integer
Private mintcol���� As Integer
Private mintcol��λ As Integer
Private mIntCol���� As Integer
Private mintcol���� As Integer
Private mIntCol��� As Integer
Private mIntCol����� As Integer
Private mIntCol��λ As Integer
Private mIntCol������ As Integer
Private mIntCol׼���� As Integer
Private mIntCol��ҩ�� As Integer
Private mIntCol���� As Integer
Private mIntCol������ As Integer
Private mIntCol��Ч�� As Integer
Private mIntCol�²��� As Integer
Private mIntCol��ע As Integer
Private mIntColҽ��id As Integer
Private mIntColʵ������ As Integer
Private mIntCol��װ As Integer
Private mIntCol���� As Integer
Private mIntColNO As Integer
Private mIntCol�����־ As Integer
Private mIntCol��¼���� As Integer
Private mIntCol�ɲ��� As Integer

Private mintType As Integer   '��ǰҳ��
Private mstrUnallowSetColHide  As String   '�����������ص���
Private mstrUnallowShow As String     '������ʾ����

Private mrsWork As Recordset '��ǰ���������ݼ�
Private mstrVBMoneyForamt As String
Private mintMoneyDigit As Integer
Private mFMT As g_FmtString
Private mintUnit As Integer

Private Enum mListType
    ������ = 0
    ���� = 1
End Enum

Private Sub Form_Load()
    '��ȡ����������С��λ��
    mintUnit = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, glngModul, "0"))
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    
    mintMoneyDigit = GetDigit
    Call GetMoneyFormat
    Call LoadPrepare
    Call InitVSFDetail(mintType)
End Sub

Private Sub Form_Resize()
    With Me.fraLine
        .Left = 0
        .Width = Me.Width
    End With
    
    Me.VSFDetail.Move 80, VSFDetail.Top, Me.Width - 2 * Me.lblName.Left, Me.Height - VSFDetail.Top
End Sub


Public Sub InitVSFDetail(ByVal intType As Integer)
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������ As Variant
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol��ǰ�� = 0
    mIntCol˳��� = 1
    mIntCol�������� = 2
    mIntCol������ = 3
    mIntColӢ���� = 4
    mintcol��� = 5
    mintcol��� = 6
    mintcol���� = 7
    mIntColId = 8
    mintcol����id = 9
    mintcol���� = 10
    mintcol��λ = 11
    mIntCol���� = 12
    mintcol���� = 13
    mIntCol��� = 14
    mIntCol��λ = 15
    mIntCol������ = 16
    mIntCol׼���� = 17
    mIntCol��ҩ�� = 18
    mIntCol���� = 19
    mIntCol������ = 20
    mIntCol��Ч�� = 21
    mIntCol�²��� = 22
    mIntCol��ע = 23
    mIntColҽ��id = 24
    mIntColʵ������ = 25
    mIntCol��װ = 26
    mIntCol���� = 27
    mIntColNO = 28
    mIntCol�����־ = 29
    mIntCol��¼���� = 30
    mIntCol�ɲ��� = 31
    
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
    With VSFDetail
        .Redraw = flexRDNone
        
        .Rows = 1
        .Rows = 2
        .Cols = mconIntCol����
        
        .Cell(flexcpPicture, 1, mIntCol��ǰ��, 1, mIntCol��ǰ��) = Me.imgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol��ǰ��, .Rows - 1, mIntCol��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat VSFDetail, mIntCol��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"

        VsfGridColFormat VSFDetail, mIntCol˳���, "˳���", 0, flexAlignRightCenter, "˳���"
        VsfGridColFormat VSFDetail, mIntCol��������, "��������", 2000, flexAlignLeftCenter, "��������"
        VsfGridColFormat VSFDetail, mIntCol������, "������", 1000, flexAlignLeftCenter, "������"
        VsfGridColFormat VSFDetail, mIntColӢ����, "Ӣ����", 0, flexAlignCenterCenter, "Ӣ����"
        VsfGridColFormat VSFDetail, mintcol���, "���", 0, flexAlignCenterCenter, "���"
        VsfGridColFormat VSFDetail, mintcol���, "���", 1200, flexAlignLeftCenter, "���"
        VsfGridColFormat VSFDetail, mintcol����, "����", 800, flexAlignLeftCenter, "����"
        
        VsfGridColFormat VSFDetail, mIntColId, "Id", 0, flexAlignRightCenter, "Id"
        VsfGridColFormat VSFDetail, mintcol����id, "����id", 0, flexAlignLeftCenter, "����id"
        VsfGridColFormat VSFDetail, mintcol����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat VSFDetail, mintcol��λ, "��λ", 800, flexAlignLeftCenter, "��λ"
        VsfGridColFormat VSFDetail, mIntCol����, "����", 1000, flexAlignRightCenter, "����"
        VsfGridColFormat VSFDetail, mintcol����, "����", 1000, flexAlignRightCenter, "����"
        VsfGridColFormat VSFDetail, mIntCol���, "���", 1600, flexAlignRightCenter, "���"
        VsfGridColFormat VSFDetail, mIntCol��λ, "��λ", 0, flexAlignCenterCenter, "��λ"
        
        
        VsfGridColFormat VSFDetail, mIntCol������, "������", IIf(intType = 1, 1000, 0), flexAlignRightCenter, "������"
        VsfGridColFormat VSFDetail, mIntCol׼����, "׼����", IIf(intType = 1, 1000, 0), flexAlignRightCenter, "׼����"
        
        VsfGridColFormat VSFDetail, mIntCol��ҩ��, "��ҩ��", IIf(intType = 1, 1000, 0), flexAlignRightCenter, "��ҩ��"
        VsfGridColFormat VSFDetail, mIntCol����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat VSFDetail, mIntCol������, "������", 0, flexAlignCenterCenter, "������"
        VsfGridColFormat VSFDetail, mIntCol��Ч��, "��Ч��", 0, flexAlignCenterCenter, "��Ч��"
        VsfGridColFormat VSFDetail, mIntCol�²���, "�²���", 0, flexAlignCenterCenter, "�²���"
        VsfGridColFormat VSFDetail, mIntCol��ע, "��ע", 0, flexAlignCenterCenter, "��ע"
        VsfGridColFormat VSFDetail, mIntColҽ��id, "ҽ��id", 0, flexAlignCenterCenter, "ҽ��id"
        VsfGridColFormat VSFDetail, mIntColʵ������, "ʵ������", 0, flexAlignRightCenter, "ʵ������"
        VsfGridColFormat VSFDetail, mIntCol��װ, "��װ", 0, flexAlignCenterCenter, "��װ"
        VsfGridColFormat VSFDetail, mIntCol����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat VSFDetail, mIntColNO, "NO", 0, flexAlignCenterCenter, "NO"
        VsfGridColFormat VSFDetail, mIntCol�����־, "�����־", 0, flexAlignCenterCenter, "�����־"
        VsfGridColFormat VSFDetail, mIntCol��¼����, "��¼����", 0, flexAlignCenterCenter, "��¼����"
        VsfGridColFormat VSFDetail, mIntCol�ɲ���, "�ɲ���", 0, flexAlignCenterCenter, "�ɲ���"
        
        mstrUnallowSetColHide = "��������;���;��λ;����;����;���"
        mstrUnallowShow = "�ⷿID;��¼����;�����־;NO;����;��װ;����id;id;����;ҽ��id"
        If intType = 1 Then mstrUnallowShow = mstrUnallowShow & ";��ҩ��;������;׼����"

        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If InStr(mstrUnallowShow, Split(arr������(n), ",")(0)) > 0 Then
                    For i = 0 To VSFDetail.Cols - 1
                        If Split(arr������(n), ",")(0) = VSFDetail.ColKey(i) Then
                            VSFDetail.ColWidth(i) = Val(Split(arr������(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        '������������
        .Select 0, 0, .Rows - 1, .Cols - 1
        .CellBorder &H9D9D9D, 1, 1, 1, 1, 1, 1
        
        .RowSel = 1
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer)
    Select Case str����
        Case "��������"
            mIntCol�������� = intValue
        Case "������"
            mIntCol������ = intValue
        Case "Ӣ����"
            mIntColӢ���� = intValue
            
        Case "���"
            mintcol��� = intValue
        Case "���"
            mintcol��� = intValue

        Case "����"
            mintcol���� = intValue
        Case "��λ"
            mintcol��λ = intValue
        Case "����"
            mIntCol���� = intValue
    
        Case "����"
            mintcol���� = intValue
    End Select
                
End Sub


Public Sub WriteSendList(ByVal intType As Integer, ByVal rsTemp As Recordset, ByVal int�ɲ��� As Integer)
    Dim i As Integer
    Dim intCount As Integer
    Dim dblMoney As Double
    Dim str�ϼ� As String
    
    Set mrsWork = rsTemp
    
    With mrsWork
        Call InitVSFDetail(intType)
        
        If .RecordCount = 0 Then Exit Sub
        
        VSFDetail.Redraw = flexRDNone
        
        VSFDetail.Rows = .RecordCount + 1
        
        '��䲡�˻�����Ϣ
        Me.lblName.Caption = "������" & Nvl(!����)
        Me.lblAge.Caption = "���䣺" & Nvl(!����)
        Me.lblSex.Caption = "�Ա�" & Nvl(!�Ա�)
        Me.lblDept.Caption = "���ң�" & Nvl(!����)
        
        For i = 1 To .RecordCount
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("˳���")) = i
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("��������")) = Nvl(!��������)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("������")) = Nvl(!������)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("���")) = Nvl(!���)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("���")) = Nvl(!���)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("����")) = Nvl(!����)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("Id")) = !Id
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("����id")) = !����ID
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("����")) = Nvl(!����)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("��λ")) = Nvl(!��λ)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("����")) = Format(!����, mFMT.FM_���ۼ�)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("����")) = Format(!����, mFMT.FM_����)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("���")) = Format(!���, mFMT.FM_���)
    
            If intType = 1 Then
                VSFDetail.TextMatrix(i, VSFDetail.ColIndex("������")) = Format(!��������, mFMT.FM_����)
                VSFDetail.TextMatrix(i, VSFDetail.ColIndex("׼����")) = Format(!׼����, mFMT.FM_����)
                VSFDetail.TextMatrix(i, VSFDetail.ColIndex("��ҩ��")) = Format(!׼����, mFMT.FM_����)
            End If
            
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("����")) = !����
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("��ע")) = Nvl(!˵��)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("ҽ��id")) = Nvl(!ҽ��id)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("ʵ������")) = Format(!����, mFMT.FM_����)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("����")) = !����
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("NO")) = Nvl(!NO)
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("�����־")) = !�����־
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("��¼����")) = !��¼����
            VSFDetail.TextMatrix(i, VSFDetail.ColIndex("�ɲ���")) = int�ɲ���

            dblMoney = dblMoney + Val(!���)

            .MoveNext
        Next
        
         '��Ӻϼ���
        '���հ�����ʾ���ϼ�
        str�ϼ� = zlStr.ChineseMoney(dblMoney)

        intCount = i
        VSFDetail.Rows = intCount + 1
        
        str�ϼ� = "���ϼƣ�" & Format(dblMoney, mstrVBMoneyForamt) & "  ��д��" & str�ϼ�
        
        VSFDetail.Cell(flexcpText, intCount, 1, intCount, VSFDetail.Cols - 1) = str�ϼ�
        VSFDetail.Cell(flexcpAlignment, intCount, mIntCol˳���, intCount, VSFDetail.Cols - 1) = flexAlignLeftCenter
        VSFDetail.Cell(flexcpFontBold, intCount, mIntCol˳���, intCount, VSFDetail.Cols - 1) = True
        
        VSFDetail.MergeCells = flexMergeRestrictRows
        VSFDetail.MergeRow(VSFDetail.Rows - 1) = True
        
        '������������
        VSFDetail.Select 0, 0, VSFDetail.Rows - 1, VSFDetail.Cols - 1
        VSFDetail.CellBorder &H9D9D9D, 1, 1, 1, 1, 1, 1

        VSFDetail.Redraw = flexRDBuffered
        VSFDetail.Refresh
        
        VSFDetail.Row = VSFDetail.Rows - 1
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
    
    LoadListColState = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & Me.Name & "\" & TypeName(VSFDetail), strType, "")
End Function

Public Function GetWorkRs(ByVal intType As Integer, ByRef str��ҩ���� As String) As Recordset
'��ȡ��ǰ���������ݼ�
'������intType��0-���ϣ�1-����,str��ҩ����,��ҩʱΪ��ҩ��������ҩΪ��ҩ��
    Dim i As Integer
    
    Set GetWorkRs = mrsWork
    
    '��ȡ��ҩ����
    If intType = 1 Then
        With VSFDetail
            For i = 1 To .Rows - 1
                str��ҩ���� = str��ҩ���� & "," & .TextMatrix(i, mIntColId) & "," & .TextMatrix(i, mIntCol��ҩ��) & "|"
            Next
        End With
    Else
        str��ҩ���� = cboPre.Text
    End If
    
End Function


Public Sub SetFontSize(ByVal intFont As Integer)
    Me.VSFDetail.FontSize = intFont
End Sub

Private Sub LoadPrepare()
    Dim rsTemp As Recordset
    Dim lng���ϲ���ID As Long
    
    Set rsTemp = LoadPerson(UserInfo.Id)
    
    'װ�뷢�ϲ�������
    With cboPre
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = UserInfo.Id Then
                .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
    End With
End Sub

Private Sub VSFDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mIntCol��ҩ�� Then
        With VSFDetail
            If Val(.TextMatrix(Row, Col)) > Val(.TextMatrix(Row, mIntCol׼����)) Then
               .TextMatrix(Row, Col) = .TextMatrix(Row, mIntCol׼����)
            ElseIf Val(.TextMatrix(Row, Col)) < 0 Then
                .TextMatrix(Row, Col) = 0
            End If
        End With
        
    End If
End Sub

Private Sub VSFDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mIntCol��ҩ�� Or Row = 0 Then
        Cancel = True
        Exit Sub
    End If
    If Val(VSFDetail.TextMatrix(Row, Col)) <> 1 Then Cancel = True
End Sub

Private Sub VSFDetail_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = mIntCol�������� Then
        Position = mIntCol��������
    End If
End Sub

Private Sub vsfDetail_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mIntCol��ǰ�� Then Cancel = True
End Sub

Private Sub VSFDetail_EnterCell()
    With VSFDetail
        If .Row = 0 Then Exit Sub
        
        .Cell(flexcpPicture, 1, 0, .Rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.imgList.ListImages(2).Picture
    End With
End Sub

Private Sub VSFDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With VSFDetail
        If Col = mIntCol��ҩ�� Then
            If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            ElseIf KeyAscii = Asc(".") Then
                If InStr(.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                End If
            End If
        End If
    End With
End Sub


Private Sub GetMoneyFormat()
    Dim n As Integer
    Dim strOracleTmp As String
    Dim strVbTmp As String
    
    strOracleTmp = "999999990."
    strVbTmp = "########0."
    For n = 1 To mintMoneyDigit
        strOracleTmp = strOracleTmp & "0"
        strVbTmp = strVbTmp & "0"
    Next
    
    mstrVBMoneyForamt = strVbTmp
    
End Sub
