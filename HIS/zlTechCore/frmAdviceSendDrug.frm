VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmAdviceSendDrug 
   AutoRedraw      =   -1  'True
   Caption         =   "ҩ��ҽ������"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmAdviceSendDrug.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmAdviceSendDrug.frx":058A
   ScaleHeight     =   6510
   ScaleWidth      =   9540
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7290
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   5
      Top             =   6210
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   6150
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceSendDrug.frx":0B14
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11430
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   900
      BandCount       =   1
      _CBWidth        =   9540
      _CBHeight       =   510
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   450
      Width1          =   3525
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   450
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   794
         ButtonWidth     =   1561
         ButtonHeight    =   794
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫѡ"
               Key             =   "ȫѡ"
               Description     =   "ȫѡ"
               Object.ToolTipText     =   "ȫѡ(Ctrl+A)"
               Object.Tag             =   "ȫѡ"
               ImageKey        =   "ȫѡ"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȫ��"
               Key             =   "ȫ��"
               Description     =   "ȫ��"
               Object.ToolTipText     =   "ȫ��(Ctrl+R)"
               Object.Tag             =   "ȫ��"
               ImageKey        =   "ȫ��"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����ѡ���ҽ��(Ctrl+E)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "�����������������������嵥(F12)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����(F1)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�(ALT+X)"
               Object.Tag             =   "�˳�"
               ImageKey        =   "�˳�"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   8
      Top             =   4605
      Width           =   9495
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   1425
      Left            =   0
      TabIndex        =   1
      Top             =   4725
      Width           =   9540
      _cx             =   16828
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceSendDrug.frx":13A8
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
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   6
      Top             =   525
      Width           =   9435
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   90
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   825
      Width           =   9540
      _cx             =   16828
      _cy             =   6641
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
      BackColorSel    =   16764057
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
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
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceSendDrug.frx":1443
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin MSComctlLib.ImageList img16 
         Left            =   3435
         Top             =   1905
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceSendDrug.frx":14DE
               Key             =   "T"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdviceSendDrug.frx":1A78
               Key             =   "F"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   360
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2012
            Key             =   "ȫѡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":222C
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2446
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2660
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":287A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2A94
            Key             =   "�˳�"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   960
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2CAE
            Key             =   "ȫѡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":2EC8
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":30E2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":32FC
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":3516
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdviceSendDrug.frx":3730
            Key             =   "�˳�"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAdviceSendDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String 'IN
Public mlng����ID As Long 'IN:���ڼ�¼������Ĳ������ϴη��Ͳ���
Public mlng����ID As Long 'IN
Public mblnSend As Boolean 'OUT:�Ƿ�ɹ����͹���

Private mcolStock As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mrsҩ�� As ADODB.Recordset
Private mrsBill As ADODB.Recordset
Private mstrEnd As String '���η��͵Ľ���ʱ��
Private mint��Ч As Integer '���η��͵�ҽ����Ч
Private mblnLimit As Boolean '���η��͸�ҩ;�������Ƿ��Խ���ʱ������
Private mlngҩƷ���ID As Long 'ҩƷ������ID
Private mlng�������ID As Long
Private mblnAutoExe As Boolean '����ִ���Զ����
Private mstrLike As String
Private mblnFirst As Boolean

'----------------------------------------------
Private Const COL_ѡ�� = 0
Private Const COL_���� = 1
Private Const COL_���� = 2
Private Const COL_סԺ�� = 3
Private Const COL_���� = 4
Private Const COL_�ѱ� = 5
Private Const COL_Ӥ�� = 6
Private Const COL_ҽ������ = 7
Private Const COL_��� = 8
Private Const COL_���� = 9
Private Const COL_������λ = 10
Private Const COL_���� = 11
Private Const COL_������λ = 12
Private Const COL_��� = 13
Private Const COL_Ƶ�� = 14
Private Const COL_�÷� = 15
Private Const COL_ҽ������ = 16
Private Const COL_ִ��ʱ�� = 17
Private Const COL_�״�ʱ�� = 18
Private Const COL_ĩ��ʱ�� = 19
Private Const COL_ִ�п��� = 20
Private Const COL_����ִ�� = 21
Private Const COL_ִ������ = 22
Private Const COL_����ID = 23 '������
Private Const COL_��ҳID = 24
Private Const COL_�Ա� = 25
Private Const COL_���� = 26
Private Const COL_ID = 27
Private Const COL_���ID = 28
Private Const COL_���˿���ID = 29
Private Const COL_��������ID = 30
Private Const COL_����ҽ�� = 31
Private Const COL_������� = 32
Private Const COL_������ĿID = 33
Private Const COL_�Ƽ����� = 34
Private Const COL_ִ������ID = 35
Private Const COL_ִ�п���ID = 36
Private Const COL_ҩƷID = 37
Private Const COL_����ϵ�� = 38
Private Const COL_סԺ��װ = 39
Private Const COL_סԺ��λ = 40
Private Const COL_�ɷ���� = 41
Private Const COL_ҩ������ = 42
Private Const COL_�Ƿ��� = 43
Private Const COL_��� = 44
Private Const COL_���� = 45
Private Const COL_�ֽ�ʱ�� = 46
'-------------------------------------------------
Private Const COLP_�Ƽ�ҽ�� = 0
Private Const COLP_��� = 1
Private Const COLP_�շ���Ŀ = 2
Private Const COLP_���� = 3
Private Const COLP_���� = 4
Private Const COLP_��λ = 5
Private Const COLP_���� = 6
Private Const COLP_Ӧ�ս�� = 7
Private Const COLP_ʵ�ս�� = 8
Private Const COLP_ִ�п��� = 9
Private Const COLP_�������� = 10
Private Const COLP_���� = 11
Private Const COLP_�к� = 12

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.Value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.Value = vNewValue
        txtPer.Text = CInt(psb.Value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property

Private Sub Form_Activate()
    If mblnFirst Then
        mblnFirst = False
        
        mlngҩƷ���ID = ExistIOClass(9)
        If mlngҩƷ���ID = 0 Then
            MsgBox "����ȷ��ҩƷ�������ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        mlng�������ID = ExistIOClass(41) '����ȷ���Ƿ�ʹ���������շ�,�������ж�
        
        If Not ResetSend Then Unload Me: Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call tbr_ButtonClick(tbr.Buttons("����"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbr_ButtonClick(tbr.Buttons("�˳�"))
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ȫѡ"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("ȫ��"))
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("����"))
    ElseIf KeyCode = vbKeyF12 Then
        Call tbr_ButtonClick(tbr.Buttons("����"))
    End If
End Sub

Private Sub Form_Load()
    Call InitAdviceTable
    Call InitPriceTable
    Call RestoreWinState(Me, App.ProductName)
        
    mblnSend = False
    mblnFirst = True
    mstrLike = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    mblnAutoExe = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "����ִ���Զ����", 0)) <> 0
    
    '��ʼ��ȡһЩ����
    '�����ⷿҩƷ�����鷽ʽ,�������ϲ���
    Set mcolStock = InitStockCheck(2, True)
    Call Initҩ���û�
End Sub

Private Function GetStockCheck(ByVal lng�ⷿID As Long) As Integer
'���ܣ���ȡָ���ⷿ�ĳ������鷽ʽ
    Dim intStyle As Integer
    On Error Resume Next
    intStyle = mcolStock("_" & lng�ⷿID)
    Err.Clear: On Error GoTo 0
    GetStockCheck = intStyle
End Function

Private Function Initҩ���û�() As Boolean
'���ܣ�'��ʼ��ȡһЩ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    '��ȡ����ҩ����������:����ҩ���û�
    Set mrsҩ�� = New ADODB.Recordset
    mrsҩ��.Fields.Append "ID", adBigInt
    mrsҩ��.Fields.Append "����", adVarChar, 10
    mrsҩ��.Fields.Append "����", adVarChar, 20
    mrsҩ��.Fields.Append "��ID", adBigInt
    mrsҩ��.CursorLocation = adUseClient
    mrsҩ��.LockType = adLockOptimistic
    mrsҩ��.CursorType = adOpenStatic
    mrsҩ��.Open
    
    strSQL = _
        "Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " AND B.����ID=A.ID And B.������� IN(2,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��')" & _
        " Order by A.����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        mrsҩ��.AddNew
        mrsҩ��!ID = rsTmp!ID
        mrsҩ��!���� = rsTmp!����
        mrsҩ��!���� = rsTmp!����
        mrsҩ��!��ID = rsTmp!ID
        mrsҩ��.Update
        rsTmp.MoveNext
    Next
    
    Initҩ���û� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    fraInfo.Top = cbr.Height
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraInfo.Top + fraInfo.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraInfo.Height - vsPrice.Height - fraUD.Height - cbr.Height - stbThis.Height
    
    fraUD.Top = vsAdvice.Top + vsAdvice.Height
    fraUD.Left = 0
    fraUD.Width = Me.ScaleWidth
    
    vsPrice.Left = 0
    vsPrice.Top = fraUD.Top + fraUD.Height
    vsPrice.Width = Me.ScaleWidth
    
    psb.Top = stbThis.Top + 60
    psb.Width = stbThis.Panels(2).Width - txtPer.Width - 100
    psb.Left = stbThis.Panels(2).Left + 30
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    
    '�ͷ�˽�м�IN����
    mstrPrivs = ""
    mlng����ID = 0
    mlng����ID = 0
    mstrEnd = ""
    mint��Ч = 0
    mblnLimit = False
    mlngҩƷ���ID = 0
    mlng�������ID = 0
    Set mrsҩ�� = Nothing
    Set mrsBill = Nothing
    Set mcolStock = Nothing
    
    gbln�Ӱ�Ӽ� = False
End Sub

Private Function ResetSend() As Boolean
'���ܣ����÷�������
    With frmAdviceSendDrugCond
        .mstrPrivs = mstrPrivs
        .mlng����ID = mlng����ID
        .mlng����ID = mlng����ID
        Set .mrsҩ�� = mrsҩ��
        .Show 1, Me
        If .mblnOK Then
            mlng����ID = .mlng����ID
            mstrEnd = .mstrEnd
            mint��Ч = .mint��Ч
            mblnLimit = .mblnLimit
            Call LoadAdviceSend(.mstrEnd, .mint��Ч, .mlngҩ��ID, .mstr����IDs, .mstr��ҩIDs)
        End If
        ResetSend = .mblnOK
    End With
End Function

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsAdvice.Height + y < 1000 Or vsPrice.Height - y < 500 Then Exit Sub
        fraUD.Top = fraUD.Top + y
        vsAdvice.Height = vsAdvice.Height + y
        vsPrice.Top = vsPrice.Top + y
        vsPrice.Height = vsPrice.Height - y
        Me.Refresh
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lng���ͺ� As Long, i As Long
    
    Select Case Button.Key
        Case "ȫѡ"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = img16.ListImages("T").Picture
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "ȫ��"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 Then
                        Set .Cell(flexcpPicture, i, COL_ѡ��) = Nothing
                    End If
                Next
            End With
            Call ShowSendTotal
        Case "����"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COL_ID)) <> 0 And .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        Exit For
                    End If
                Next
                If i > .Rows - 1 Then
                    MsgBox "��ǰû�п��Է��͵�ҽ����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End With
            If MsgBox("ȷʵҪ���͵�ǰѡ���ҽ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                lng���ͺ� = SendAdvice
                If lng���ͺ� <> 0 Then
                    mblnSend = True
                    '��ӡ���Ƶ���
                    Call frmSendBillPrint.ShowMe(lng���ͺ�, 2, Me)
                End If
            End If
        Case "����"
            Call ResetSend
        Case "����"
            ShowHelp App.ProductName, Me.Hwnd, Me.Name
        Case "�˳�"
            Unload Me
    End Select
End Sub

Private Sub RowSelectSame(ByVal lngRow As Long, ByVal lngCol As Long, _
    Optional rsSQL As ADODB.Recordset, Optional rsTotal As ADODB.Recordset, Optional rsUpload As ADODB.Recordset)
'���ܣ����ݿɼ��е�ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsAdvice
        If lngCol = COL_ѡ�� Then
            For i = lngRow + 1 To .Rows - 1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1
                If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) _
                    = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID))) Then
                    .Cell(flexcpData, i, lngCol) = .Cell(flexcpData, lngRow, lngCol)
                    Set .Cell(flexcpPicture, i, lngCol) = .Cell(flexcpPicture, lngRow, lngCol)
                Else
                    Exit For
                End If
            Next
            
            'ȡ��ѡ��ʱ
            If Not (.Cell(flexcpData, lngRow, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, lngRow, COL_ѡ��) Is Nothing) Then
                i = IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, Val(.TextMatrix(lngRow, COL_ID)), Val(.TextMatrix(lngRow, COL_���ID)))
                '1.�����Ӧ�ķ��ü����ͼ�¼��д
                If Not rsSQL Is Nothing Then
                    rsSQL.Filter = "ҽ��ID=" & i
                    Do While Not rsSQL.EOF
                        rsSQL.Delete
                        rsSQL.Update
                        rsSQL.MoveNext
                    Loop
                    rsSQL.Filter = 0 '��ΪҪʹ��BookMark����˻ָ�
                End If
                '2.�����Ӧ�ķ��ͼƼ������ۼ�
                If Not rsTotal Is Nothing Then
                    rsTotal.Filter = "ҽ��ID=" & i
                    Do While Not rsTotal.EOF
                        rsTotal.Delete
                        rsTotal.Update
                        rsTotal.MoveNext
                    Loop
                End If
                '3.�����Ӧ��ҽ���ϴ����ݺ�
                If Not rsUpload Is Nothing Then
                    rsUpload.Filter = "ҽ��ID=" & i
                    Do While Not rsUpload.EOF
                        rsUpload.Delete
                        rsUpload.Update
                        rsUpload.MoveNext
                    Loop
                End If
            End If
        End If
    End With
End Sub

Private Function GetVisibleRow(ByVal lngRow As Long, Optional ByVal blnFirst As Boolean) As Long
'���ܣ�����ָ��ҽ���У����ظ�ҽ���пɼ�����
    Dim lng��ID As Long, i As Long
    
    GetVisibleRow = lngRow
    
    With vsAdvice
        If Not .RowHidden(lngRow) Then Exit Function
        
        'һ����ҩ�Ķ�λ����һҩƷ��
        If blnFirst Then
            If .TextMatrix(lngRow, COL_�������) = "E" And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 _
                And Val(.TextMatrix(lngRow, COL_���ID)) = 0 And Val(.TextMatrix(lngRow, COL_ID)) = Val(.TextMatrix(lngRow - 1, COL_���ID)) Then
                i = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_���ID)
                If i <> -1 Then GetVisibleRow = i: Exit Function
            End If
        End If
        
        lng��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) <> 0, Val(.TextMatrix(lngRow, COL_���ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            If lng��ID = IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) Then
                If Not .RowHidden(i) Then GetVisibleRow = i: Exit Function
            Else
                Exit For
            End If
        Next
    End With
End Function

Private Function ShowSendPrice(ByVal lngRow As Long) As Boolean
'���ܣ���ʾ��ǰ����ҽ���еļ��ʷ�����Ϣ(��סԺ��λ,���ѱ���ۣ�����ҽ�������ܶ���)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngTmp As Long, strTmp As String
    Dim str�ѱ� As String, str�к� As String, dbl���� As Double
    Dim dbl���� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim dbl��ǰ���� As Double, cur��ǰӦ�� As Currency, cur��ǰʵ�� As Currency
    Dim lngִ�п���ID As Long, lng���˿���ID As Long
    Dim lng����ID As Long, lng��ҳID As Long
    
    Dim rsMain As New ADODB.Recordset
    Dim rsClone As New ADODB.Recordset
    Dim strHaveSub As String, strNoneSub As String
    
    On Error GoTo errH
    
    '���ڻ��ܼ����ۿ۵���ʱ��¼��
    rsMain.Fields.Append "ҽ���к�", adBigInt
    rsMain.Fields.Append "�����к�", adBigInt
    rsMain.Fields.Append "������ID", adBigInt
    rsMain.Fields.Append "ҽ���ϼ�", adCurrency, , adFldIsNullable
    rsMain.CursorLocation = adUseClient
    rsMain.LockType = adLockOptimistic
    rsMain.CursorType = adOpenStatic
    rsMain.Open
    
    With vsAdvice
        lng����ID = Val(.TextMatrix(lngRow, COL_����ID))
        lng��ҳID = Val(.TextMatrix(lngRow, COL_��ҳID))
        str�ѱ� = .TextMatrix(lngRow, COL_�ѱ�)
        
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '��ҩƷ�Ƽ۲���:ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
            If Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                strSQL = "Select 'ҩƷҽ��-" & Replace(.Cell(flexcpData, lngRow, COL_ҽ������), "'", "''") & "' as �Ƽ�ҽ��," & _
                    lngRow & " as �к�,A.ID,A.���,Nvl(B.����,A.����)||Decode(A.����,NULL,NULL,'('||A.����||')')||Decode(A.���,NULL,NULL,' '||A.���) as ����," & _
                    "'" & .TextMatrix(lngRow, COL_������λ) & "' as ��λ,A.�Ƿ���,A.�Ӱ�Ӽ�,0 as ����,1 as ����," & _
                    Val(.TextMatrix(lngRow, COL_����)) & " as ����," & Val(.TextMatrix(lngRow, COL_ִ�п���ID)) & " as ִ�п���ID," & _
                    " A.���ηѱ�,A.��������,0 as ����" & _
                    " From �շ���ĿĿ¼ A,�շ���Ŀ���� B" & _
                    " Where A.ID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIF(gbln��Ʒ��, 3, 1) & _
                    " And A.ID=" & Val(.TextMatrix(lngRow, COL_ҩƷID))
            End If
            '��ҩ;���Ƽ۲���:һ����ҩ(�����)�еĵ�һ��ҩ��,����ʾ��ҩ;������
            If Val(.TextMatrix(lngRow - 1, COL_���ID)) <> Val(.TextMatrix(lngRow, COL_���ID)) Then
                '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�У���ҽ������ȡ
                lngTmp = .FindRow(.TextMatrix(lngRow, COL_���ID), lngRow, COL_ID)
                If Val(.TextMatrix(lngTmp, COL_�Ƽ�����)) = 0 _
                    And InStr(",0,5,", Val(.TextMatrix(lngTmp, COL_ִ������ID))) = 0 Then
                    strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                        "Select '��ҩ;��-" & Replace(.Cell(flexcpData, lngTmp, COL_ҽ������), "'", "''") & "' as �Ƽ�ҽ��," & _
                        lngTmp & " as �к�,B.ID,B.���,B.����,B.���㵥λ as ��λ,B.�Ƿ���,B.�Ӱ�Ӽ�,A.����,1 as ����," & _
                        " Nvl(A.����,0)*" & Val(.TextMatrix(lngTmp, COL_����)) & " as ����," & _
                        " Nvl(A.ִ�п���ID," & Val(.TextMatrix(lngTmp, COL_ִ�п���ID)) & ") as ִ�п���ID," & _
                        " B.���ηѱ�,B.��������,Nvl(A.����,0) as ����" & _
                        " From ����ҽ���Ƽ� A,�շ���ĿĿ¼ B" & _
                        " Where A.�շ�ϸĿID=B.ID And Nvl(A.����,0)<>0 And A.ҽ��ID=" & Val(.TextMatrix(lngTmp, COL_ID))
                End If
            End If
            vsPrice.ColHidden(COLP_����) = True
        ElseIf .TextMatrix(lngRow, COL_�������) = "E" Then
            lngTmp = .FindRow(.TextMatrix(lngRow, COL_ID), , COL_���ID)
            For i = lngTmp To lngRow
                If .TextMatrix(i, COL_�������) = "7" Then
                    '�����ҩ
                    If Val(.TextMatrix(i, COL_ִ������ID)) <> 5 Then
                        '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                        If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                            dbl���� = Format(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)), "0.00000")
                        Else
                            dbl���� = IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                        End If
                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                            "Select 'ҩƷҽ��-" & Replace(.Cell(flexcpData, i, COL_ҽ������), "'", "''") & "' as �Ƽ�ҽ��," & i & " as �к�," & _
                            " A.ID,A.���,A.����||Decode(A.����,NULL,NULL,'('||A.����||')')||Decode(A.���,NULL,NULL,' '||A.���) as ����," & _
                            " B.סԺ��λ as ��λ,A.�Ƿ���,A.�Ӱ�Ӽ�,0 as ����," & Val(.TextMatrix(i, COL_����)) & " as ����," & _
                            dbl���� & " as ����," & Val(.TextMatrix(i, COL_ִ�п���ID)) & " as ִ�п���ID," & _
                            " A.���ηѱ�,A.��������,0 as ����" & _
                            " From �շ���ĿĿ¼ A,ҩƷ��� B Where A.ID=B.ҩƷID And A.ID=" & Val(.TextMatrix(i, COL_ҩƷID))
                    End If
                Else
                    '��ҩ�巨,�÷�
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 _
                        And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) = 0 Then
                            strTmp = "��ҩ�÷�-" & Replace(.Cell(flexcpData, i, COL_ҽ������), "'", "''")
                        Else
                            strTmp = "��ҩ�巨-" & Replace(.Cell(flexcpData, i, COL_ҽ������), "'", "''")
                        End If
                        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
                            "Select '" & strTmp & "' as �Ƽ�ҽ��," & _
                            i & " as �к�,B.ID,B.���,B.����,B.���㵥λ as ��λ,B.�Ƿ���,B.�Ӱ�Ӽ�,A.����,1 as ����," & _
                            " Nvl(A.����,0)*" & Val(.TextMatrix(i, COL_����)) & " as ����," & _
                            " Nvl(A.ִ�п���ID," & Val(.TextMatrix(i, COL_ִ�п���ID)) & ") as ִ�п���ID," & _
                            " B.���ηѱ�,B.��������,Nvl(A.����,0) as ����" & _
                            " From ����ҽ���Ƽ� A,�շ���ĿĿ¼ B" & _
                            " Where A.�շ�ϸĿID=B.ID And Nvl(A.����,0)<>0 And A.ҽ��ID=" & Val(.TextMatrix(i, COL_ID))
                    End If
                End If
            Next
            vsPrice.ColHidden(COLP_����) = False
        End If
    End With
    
    With vsPrice
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If strSQL <> "" Then
            '�����¼۸����
            strSQL = _
                " Select A.�к�,B.������ĿID,A.�Ƽ�ҽ��,A.ID,A.���,C.���� as �������," & _
                " A.����,A.��λ,D.סԺ��λ,A.����,A.����,A.ִ�п���ID,F.���� as ִ�п���,A.���ηѱ�,A.��������,D.סԺ��װ," & _
                " A.�Ƿ���,E.��������,A.����,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,Decode(Nvl(A.�Ƿ���,0),1,A.����,B.�ּ�) as ����" & _
                " From (" & strSQL & ") A,�շѼ�Ŀ B,�շ���Ŀ��� C,ҩƷ��� D,�������� E,���ű� F" & _
                " Where A.ID=B.�շ�ϸĿID And A.���=C.���� And A.ID=D.ҩƷID(+) And A.ID=E.����ID(+) And A.ִ�п���ID=F.ID(+)" & _
                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                " Order by A.�к�,A.����,A.ID"
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
            
            If Not rsTmp.EOF And gbln��������ۿ� Then
                Set rsClone = rsTmp.Clone
            End If
            
            For i = 1 To rsTmp.RecordCount
                If str�к� <> rsTmp!�к� & "_" & rsTmp!ID Then
                    If str�к� <> "" Then
                        .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, "0.00000")
                        .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                        .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
                    End If
                    str�к� = rsTmp!�к� & "_" & rsTmp!ID
                    dbl���� = 0: curӦ�� = 0: curʵ�� = 0
                    .Rows = .Rows + 1
                    
                    .TextMatrix(.Rows - 1, COLP_�к�) = rsTmp!�к�
                
                    .TextMatrix(.Rows - 1, COLP_�Ƽ�ҽ��) = rsTmp!�Ƽ�ҽ��
                    .TextMatrix(.Rows - 1, COLP_���) = rsTmp!�������
                    .TextMatrix(.Rows - 1, COLP_�շ���Ŀ) = rsTmp!����
                    .TextMatrix(.Rows - 1, COLP_��������) = Nvl(rsTmp!��������)
                    .TextMatrix(.Rows - 1, COLP_����) = IIF(Nvl(rsTmp!����, 0) = 0, "", "��")
                
                    If rsTmp!��� = "7" Then
                        .TextMatrix(.Rows - 1, COLP_����) = Nvl(rsTmp!����, 1)
                    End If
                    
                    'ҩƷ�Ƽ۵�������סԺ��λ��ʾ
                    dbl���� = Nvl(rsTmp!����, 0) 'ȡ�ۼ��������������
                    If InStr(",5,6,7,", rsTmp!���) > 0 Then
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            .TextMatrix(.Rows - 1, COLP_��λ) = Nvl(rsTmp!��λ)
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(Nvl(rsTmp!����, 0), 5)
                            dbl���� = dbl���� * Nvl(rsTmp!סԺ��װ, 1)
                        Else
                            .TextMatrix(.Rows - 1, COLP_��λ) = Nvl(rsTmp!סԺ��λ)
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,���ת��Ϊҩ����λ��ʾʱ���������㴦��
                            .TextMatrix(.Rows - 1, COLP_����) = FormatEx(Nvl(rsTmp!����, 0) / Nvl(rsTmp!סԺ��װ, 1), 5)
                        End If
                    Else
                        .TextMatrix(.Rows - 1, COLP_��λ) = Nvl(rsTmp!��λ)
                        .TextMatrix(.Rows - 1, COLP_����) = FormatEx(Nvl(rsTmp!����, 0), 5)
                    End If
                    
                    'ִ�п���:����ظ�������
                    .TextMatrix(.Rows - 1, COLP_ִ�п���) = Nvl(rsTmp!ִ�п���)
                    .Cell(flexcpData, .Rows - 1, COLP_���) = CStr(rsTmp!���) '�շ����
                    .Cell(flexcpData, .Rows - 1, COLP_�շ���Ŀ) = Val(rsTmp!ID) '��ĿID
                    .Cell(flexcpData, .Rows - 1, COLP_ִ�п���) = Val(Nvl(rsTmp!ִ�п���ID, 0)) 'ִ�п���ID
                    .Cell(flexcpData, .Rows - 1, COLP_��������) = Val(Nvl(rsTmp!��������, 0)) '��������
                    
                    '���¼���ҩ��ҩƷ���������ĵ���Чִ�п���
                    lngִ�п���ID = Nvl(rsTmp!ִ�п���ID, 0)
                    If rsTmp!��� = "4" And Nvl(rsTmp!��������, 0) = 1 _
                        Or InStr(",5,6,7,", rsTmp!���) > 0 And InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) = 0 Then
                        lng���˿���ID = Val(vsAdvice.TextMatrix(rsTmp!�к�, COL_���˿���ID))
                        lngִ�п���ID = Get�շ�ִ�п���ID(lng����ID, lng��ҳID, rsTmp!���, rsTmp!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID)
                        If lngִ�п���ID <> Val(Nvl(rsTmp!ִ�п���ID, 0)) Then
                            .TextMatrix(.Rows - 1, COLP_ִ�п���) = Get��������(lngִ�п���ID)
                            .Cell(flexcpData, .Rows - 1, COLP_ִ�п���) = lngִ�п���ID
                        End If
                    End If
                    
                    '��¼�����������Ϣ���Ա����
                    If gbln��������ۿ� And rsTmp!���� = 0 Then
                        If InStr(strHaveSub & ",", "," & rsTmp!�к� & ",") = 0 _
                            And InStr(strNoneSub & ",", "," & rsTmp!�к� & ",") = 0 Then
                            rsClone.Filter = "�к�=" & rsTmp!�к� & " And ����=1"
                            If Not rsClone.EOF Then
                                rsMain.AddNew
                                rsMain!ҽ���к� = rsTmp!�к�
                                rsMain!�����к� = .Rows - 1
                                rsMain!������ID = rsTmp!������ĿID
                                rsMain.Update
                                strHaveSub = strHaveSub & "," & rsTmp!�к�
                            Else
                                strNoneSub = strNoneSub & "," & rsTmp!�к�
                            End If
                        End If
                    End If
                End If
                
                '���ۼ������㴦��:��ҩ��ҩƷ�Ƽ���סԺ��λ����(�����Ｐ�Ƽ����ô���ͬ)
                If InStr(",5,6,7,", rsTmp!���) > 0 Then
                    If Nvl(rsTmp!�Ƿ���, 0) = 0 Then
                        dbl��ǰ���� = Nvl(rsTmp!����, 0)
                    Else
                        If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!ID, Val(.Cell(flexcpData, .Rows - 1, COLP_ִ�п���)), _
                                Format(Nvl(rsTmp!����, 1) * Nvl(rsTmp!����, 0) * Nvl(rsTmp!סԺ��װ, 1), "0.00000"), , True)
                        Else
                            dbl��ǰ���� = CalcDrugPrice(rsTmp!ID, Val(.Cell(flexcpData, .Rows - 1, COLP_ִ�п���)), _
                                Format(Nvl(rsTmp!����, 1) * Nvl(rsTmp!����, 0), "0.00000"), , True)
                        End If
                    End If
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(rsTmp!�к�, COL_�������)) > 0 Then
                        dbl��ǰ���� = Format(dbl��ǰ���� * Nvl(rsTmp!סԺ��װ, 1), "0.00000")
                        cur��ǰӦ�� = Nvl(rsTmp!����, 1) * Format(Nvl(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                    Else
                        cur��ǰӦ�� = Nvl(rsTmp!����, 1) * Format(Nvl(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                        dbl��ǰ���� = Format(dbl��ǰ���� * Nvl(rsTmp!סԺ��װ, 1), "0.00000")
                    End If
                ElseIf rsTmp!��� = "4" And Nvl(rsTmp!��������, 0) = 1 And Nvl(rsTmp!�Ƿ���, 0) = 1 Then
                    'ʱ�����ĵĵ��ۺ�ҩƷһ������
                    dbl��ǰ���� = CalcDrugPrice(rsTmp!ID, Val(.Cell(flexcpData, .Rows - 1, COLP_ִ�п���)), Format(Nvl(rsTmp!����, 0), "0.00000"), , True)
                    cur��ǰӦ�� = Format(Nvl(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                Else
                    dbl��ǰ���� = Format(Nvl(rsTmp!����, 0), "0.00000")
                    cur��ǰӦ�� = Nvl(rsTmp!����, 1) * Format(Nvl(rsTmp!����, 0), "0.00000") * dbl��ǰ����
                End If
                
                '����Ӱ�Ӽ�
                If gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1 Then
                    cur��ǰӦ�� = Format(cur��ǰӦ�� * (1 + Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100), gstrDec)
                Else
                    cur��ǰӦ�� = Format(cur��ǰӦ��, gstrDec)
                End If
                
                'ʵ��
                If gbln��������ۿ� And (rsTmp!���� = 1 Or InStr(strHaveSub & ",", "," & rsTmp!�к� & ",") > 0) Then
                    cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                    '�ۼ�ҽ���ϼ��������ۿ�
                    rsMain.Filter = "ҽ���к�=" & rsTmp!�к�
                    rsMain!ҽ���ϼ� = Nvl(rsMain!ҽ���ϼ�, 0) + cur��ǰʵ��
                    rsMain.Update
                ElseIf Nvl(rsTmp!���ηѱ�, 0) = 0 Then
                    cur��ǰʵ�� = Format(ActualMoney(str�ѱ�, rsTmp!������ĿID, cur��ǰӦ��, rsTmp!ID, lngִ�п���ID, dbl����, _
                        IIF(gbln�Ӱ�Ӽ� And Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1, Nvl(rsTmp!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                Else
                    cur��ǰʵ�� = Format(cur��ǰӦ��, gstrDec)
                End If
                
                dbl���� = dbl���� + dbl��ǰ����
                curӦ�� = curӦ�� + cur��ǰӦ��
                curʵ�� = curʵ�� + cur��ǰʵ��
                rsTmp.MoveNext
            Next
            If str�к� <> "" Then
                .TextMatrix(.Rows - 1, COLP_����) = Format(dbl����, "0.00000")
                .TextMatrix(.Rows - 1, COLP_Ӧ�ս��) = Format(curӦ��, gstrDec)
                .TextMatrix(.Rows - 1, COLP_ʵ�ս��) = Format(curʵ��, gstrDec)
            End If
        End If
        
        '���ܼ����ۿ�:������ܴ��۲�֧�ְ��ɱ����մ���
        If gbln��������ۿ� And strHaveSub <> "" Then
            rsMain.Filter = 0
            Do While Not rsMain.EOF
                cur��ǰʵ�� = Format(ActualMoney(str�ѱ�, rsMain!������ID, rsMain!ҽ���ϼ�), gstrDec)
                .TextMatrix(rsMain!�����к�, COLP_ʵ�ս��) = Format(Val(.TextMatrix(rsMain!�����к�, COLP_ʵ�ս��)) + (cur��ǰʵ�� - rsMain!ҽ���ϼ�), gstrDec)
                rsMain.MoveNext
            Loop
        End If
        
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = .FixedCols
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
        Call vsPrice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSendPrice = True
    Exit Function
errH:
    vsPrice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ����ĳ�ҩ���
    Dim rsDrug As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng���� As Long, lng��С���� As Long
    Dim dbl���� As Double, str�ֽ�ʱ�� As String
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim cur��� As Currency
    
    If Col = COL_����ִ�� Then
        With vsAdvice
            .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
        End With
    ElseIf Col = COL_��� Then
        With vsAdvice
            If Val(.TextMatrix(Row, COL_ҩƷID)) = .ComboData Then Exit Sub
            'ҩƷ�����Ϣ
            .TextMatrix(Row, COL_ҩƷID) = .ComboData
            Set rsDrug = GetDrugInfo(Val(.TextMatrix(Row, COL_������ĿID)), Val(.TextMatrix(Row, COL_ҩƷID)), Val(.TextMatrix(Row, COL_ִ�п���ID)))
            .TextMatrix(Row, COL_���) = rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "")
            .TextMatrix(Row, COL_����ϵ��) = rsDrug!����ϵ��
            .TextMatrix(Row, COL_סԺ��װ) = rsDrug!סԺ��װ
            .TextMatrix(Row, COL_�Ƿ���) = rsDrug!�Ƿ���
            .TextMatrix(Row, COL_ҩ������) = rsDrug!ҩ������
            
            .TextMatrix(Row, COL_������λ) = rsDrug!סԺ��λ
            .TextMatrix(Row, COL_���) = Format(Nvl(rsDrug!���, 0), "0.00000")
            
            'ҽ�������Ϣ
            strSQL = "Select ID,��ʼִ��ʱ��,�ϴ�ִ��ʱ��,ִ����ֹʱ��,ִ��ʱ�䷽��," & _
                " Ƶ�ʴ���,Ƶ�ʼ��,�����λ,�������� From ����ҽ����¼ Where ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(Row, COL_ID)))
            
            '���¼�������,����,�ֽ�ʱ��
            Call Calc��������ʱ��(dbl����, lng����, str�ֽ�ʱ��, mstrEnd, rsTmp, rsDrug)
            
            .TextMatrix(Row, COL_����) = FormatEx(dbl����, 5)
            .TextMatrix(Row, COL_������λ) = Nvl(rsDrug!סԺ��λ)
            
            .TextMatrix(Row, COL_����) = lng����
            .TextMatrix(Row, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
            .TextMatrix(Row, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "MM-dd HH:mm")
            .TextMatrix(Row, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "MM-dd HH:mm")
                        
            'ͬ�����ĸ�ҩ;���Ĵ���
            i = .FindRow(.TextMatrix(Row, COL_���ID), , COL_ID)
            .TextMatrix(i, COL_����) = .TextMatrix(Row, COL_����)
            .TextMatrix(i, COL_�ֽ�ʱ��) = .TextMatrix(Row, COL_�ֽ�ʱ��)
            .TextMatrix(i, COL_�״�ʱ��) = .TextMatrix(Row, COL_�״�ʱ��)
            .TextMatrix(i, COL_ĩ��ʱ��) = .TextMatrix(Row, COL_ĩ��ʱ��)
                                    
            'һ����ҩ�İ���С��������
            If RowInһ����ҩ(Row, lngBegin, lngEnd) Then
                For i = lngBegin To lngEnd
                    If Val(.TextMatrix(i, COL_����)) < lng��С���� Or lng��С���� = 0 Then
                        lng��С���� = Val(.TextMatrix(i, COL_����))
                    End If
                Next
                For i = lngBegin To lngEnd + 1
                    If Val(.TextMatrix(i, COL_����)) > lng��С���� Then
                        .TextMatrix(i, COL_����) = lng��С����
                        .TextMatrix(i, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(i, COL_�ֽ�ʱ��))
                        .TextMatrix(i, COL_�״�ʱ��) = Format(Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(0), "MM-dd HH:mm")
                        .TextMatrix(i, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "MM-dd HH:mm")
                    End If
                Next
            Else
                lngBegin = Row: lngEnd = Row
            End If
            
            '���¼��㲢��ʾ���
            For i = lngBegin To lngEnd + 1
                .TextMatrix(i, COL_���) = Format(Calcҽ�����ʽ��(i), gstrDec)
                If .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    cur��� = Val(.TextMatrix(i, COL_���))
                End If
            Next
            'һ����ҩ�ĵ�һ��(�����)����ʾ������ҩ;���Ľ��
            .TextMatrix(lngBegin, COL_���) = Format(Val(.TextMatrix(lngBegin, COL_���)) + cur���, gstrDec)
            
            '���ݿ�����ѡ��״̬
            If Val(.TextMatrix(Row, COL_����)) > Val(.TextMatrix(Row, COL_���)) Then
                If GetStockCheck(Val(.TextMatrix(Row, COL_ִ�п���ID))) = 2 _
                    Or Val(.TextMatrix(Row, COL_ҩ������)) = 1 Or Val(.TextMatrix(Row, COL_�Ƿ���)) = 1 Then
                    .Cell(flexcpData, Row, COL_ѡ��) = 1
                    Set .Cell(flexcpPicture, Row, COL_ѡ��) = img16.ListImages("F").Picture
                ElseIf GetStockCheck(Val(.TextMatrix(Row, COL_ִ�п���ID))) = 1 Then
                    .Cell(flexcpData, Row, COL_ѡ��) = 0
                    Set .Cell(flexcpPicture, Row, COL_ѡ��) = Nothing
                ElseIf GetStockCheck(Val(.TextMatrix(Row, COL_ִ�п���ID))) = 0 Then
                    .Cell(flexcpData, Row, COL_ѡ��) = 0
                    Set .Cell(flexcpPicture, Row, COL_ѡ��) = img16.ListImages("T").Picture
                End If
            ElseIf Val(.TextMatrix(Row, COL_����)) <= Val(.TextMatrix(Row, COL_���)) Then
                .Cell(flexcpData, Row, COL_ѡ��) = 0
                Set .Cell(flexcpPicture, Row, COL_ѡ��) = img16.ListImages("T").Picture
            End If
            Call RowSelectSame(Row, COL_ѡ��)
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
            Call ShowSendTotal
        End With
    End If
End Sub

Private Function Calcҽ�����ʽ��(ByVal lngRow As Long) As Currency
'���ܣ�����ָ��ҽ���еļ��ʽ��(��ʾ���鿴�����ʱ���),�����¼۸����
'˵����ҩƷ������̰�סԺ��װ����,�����ս�����ͬ��
'���أ�str���=�Ƽ����
    Dim str�ѱ� As String, dbl���� As Double
    Dim dbl���� As Double, cur��� As Currency
    
    With vsAdvice
        str�ѱ� = .TextMatrix(lngRow, COL_�ѱ�)
        If InStr(",5,6,7,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            'ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
            If Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                '���ۼ۵�λ������
                If .TextMatrix(lngRow, COL_�������) = "7" Then
                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                    If Val(.TextMatrix(lngRow, COL_�ɷ����)) = 0 Then
                        dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����)) '��ζ����
                        dbl���� = dbl���� / Val(.TextMatrix(lngRow, COL_����ϵ��))
                    Else
                        dbl���� = IntEx(Val(.TextMatrix(lngRow, COL_����)) / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ)))
                        dbl���� = dbl���� * Val(.TextMatrix(lngRow, COL_סԺ��װ)) * Val(.TextMatrix(lngRow, COL_����))
                    End If
                Else
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_סԺ��װ))
                End If
                dbl���� = Format(dbl����, "0.00000")
                
                '���ۼ۵�λ�ļ۸�
                If str�ѱ� = "" Then
                    If Val(.TextMatrix(lngRow, COL_�Ƿ���)) = 1 Then
                        'ҩƷʵ�۰��ۼ�������
                        dbl���� = CalcDrugPrice(Val(.TextMatrix(lngRow, COL_ҩƷID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), dbl����)
                    Else
                        dbl���� = CalcPrice(Val(.TextMatrix(lngRow, COL_ҩƷID)))
                    End If
                    cur��� = Format(dbl���� * Format(dbl����, "0.00000"), gstrDec)
                Else
                    If Val(.TextMatrix(lngRow, COL_�Ƿ���)) = 1 Then
                        'ҩƷʵ�۰��ۼ�������
                        cur��� = Format(CalcDrugPrice(Val(.TextMatrix(lngRow, COL_ҩƷID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), dbl����, str�ѱ�), gstrDec)
                    Else
                        cur��� = Format(CalcPrice(Val(.TextMatrix(lngRow, COL_ҩƷID)), str�ѱ�, dbl����, , Val(.TextMatrix(lngRow, COL_ִ�п���ID))), gstrDec)
                    End If
                End If
            End If
        Else
            '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�У���ҽ������ȡ
            If Val(.TextMatrix(lngRow, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������ID))) = 0 Then
                dbl���� = Format(Val(.TextMatrix(lngRow, COL_����)), "0.00000")
                If str�ѱ� = "" Then
                    dbl���� = Format(CalcAdvicePrice(Val(.TextMatrix(lngRow, COL_ID))), "0.00000")
                    cur��� = Format(dbl���� * dbl����, gstrDec)
                Else
                    cur��� = Format(CalcAdvicePrice(Val(.TextMatrix(lngRow, COL_ID)), str�ѱ�, , dbl����), gstrDec)
                End If
            End If
        End If
    End With
    Calcҽ�����ʽ�� = cur���
End Function

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAdvice
        '���ݿɷ�༭���ñ༭���Լ��������
        If NewCol = COL_��� Then
            .ComboList = .Cell(flexcpData, NewRow, NewCol)
            .FocusRect = flexFocusLight
        ElseIf NewCol = COL_����ִ�� Then
            If Should����ִ��(NewRow) Then
                .ComboList = "..."
                Set .CellButtonPicture = Me.Picture
                .FocusRect = flexFocusHeavy
            Else
                .ComboList = ""
                .FocusRect = flexFocusLight
            End If
        Else
            .ComboList = ""
            .FocusRect = flexFocusLight
        End If
        
        If OldRow <> NewRow And .Redraw <> flexRDNone And Not .RowHidden(NewRow) Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                Call ShowSendPrice(NewRow)
            End If
        End If
    End With
End Sub

Private Function Should����ִ��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ����ҽ����(�ɼ���)�Ƿ�������ø��ӵ�ִ�п���
    Dim lngRow2 As Long, i As Long
    
    If lngRow = 0 Then Exit Function
    
    lngRow2 = -1
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) = 0 Then Exit Function
        If .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 _
            And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) _
            And InStr(",7,E,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
            '��ҩ�÷�
            lngRow2 = lngRow
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
            '��ҩ;��
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1, COL_ID)
        End If
        
        '������Ժ��ִ��
        If lngRow2 <> -1 Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_ִ������ID))) = 0 Then
                Should����ִ�� = True
            End If
        End If
    End With
End Function

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_ѡ�� + 1 - .FixedCols Then
            .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = COL_ҽ������ Or Col = COL_��� Then
            If Not .ColHidden(COL_���) Then
                .AutoSize COL_ҽ������, COL_���
            Else
                .AutoSize COL_ҽ������
            End If
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_ѡ�� Then Cancel = True
End Sub

Private Sub vsAdvice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(2,3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " Order by A.����"
    With vsAdvice
        vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "ִ�п���", , , , , , True, vPoint.x, vPoint.y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            Call SetDeptInput(Row, Col, rsTmp)
            Call vsAdvice_AfterRowColChange(-1, -1, Row, Col) '������ʾ�Ƽ�ִ�п���
        Else
            If Not blnCancel Then
                MsgBox "û�п��õĿ������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_ChangeEdit()
    If vsAdvice.Col = COL_��� Then
        Call vsAdvice_AfterEdit(vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseCol = COL_ѡ�� And .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            Call vsAdvice_KeyPress(32)
        End If
    End With
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_Ƶ��: lngRight = COL_�÷�
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '���б����±���(���������õ��±��ߴ�Ϊ2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 127 Then '���ֱ�����뺺�ֵ�����
        Call vsAdvice_KeyPress(KeyCode)
    End If
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
    Dim i As Long
    
    With vsAdvice
        For i = lngRow + 1 To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        If i > .Rows - 1 Then .Row = .FixedRows
        If .RowHidden(.Row) Then .Row = lngRow
        Call .ShowCell(.Row, .Col)
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        ElseIf KeyAscii = 32 And .Col = COL_ѡ�� Then
            KeyAscii = 0
            If .Cell(flexcpData, .Row, COL_ѡ��) = 0 Then
                If .Cell(flexcpPicture, .Row, COL_ѡ��) Is Nothing Then
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = img16.ListImages("T").Picture
                Else
                    Set .Cell(flexcpPicture, .Row, COL_ѡ��) = Nothing
                End If
                Call RowSelectSame(.Row, .Col)
                Call ShowSendTotal
            End If
        Else
            If .Col = COL_����ִ�� And .ComboList = "..." Then
                If Should����ִ��(.Row) Then blnDo = True
            End If
            If blnDo Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsAdvice_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, StrInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Col = COL_��� Then
                Call vsAdvice_KeyPress(13)
            ElseIf Col = COL_����ִ�� And .EditText <> "" Then
                StrInput = UCase(.EditText)
                strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
                    " From ���ű� A,��������˵�� B" & _
                    " Where A.ID=B.����ID And B.������� IN(2,3)" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
                    " Order by A.����"
                With vsAdvice
                    vPoint = GetCoordPos(.Hwnd, .CellLeft, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ִ�п���", False, "", "", False, False, True, _
                        vPoint.x, vPoint.y, .CellHeight, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%")
                    If Not rsTmp Is Nothing Then
                        Call SetDeptInput(Row, Col, rsTmp)
                        .EditText = .TextMatrix(Row, Col) 'ֱ������ƥ����Ҫ
                        Call EnterNextCell(Row, Col)
                    Else
                        If Not blnCancel Then
                            MsgBox "û���ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                        End If
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Call vsAdvice_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
                    End If
                End With
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAdvice
        If Col = COL_��� Then
            If .ComboList = "" Then Cancel = True
        ElseIf Col = COL_����ִ�� Then
            If Not Should����ִ��(Row) Then Cancel = True
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    
    If NewRow <> OldRow Then
        With vsPrice
            stbThis.Panels(2).Text = ""
            lngRow = Val(.TextMatrix(NewRow, COLP_�к�))
            If lngRow <> 0 And .Cell(flexcpData, NewRow, COLP_���) <> "" Then
                If InStr(",5,6,7,", .Cell(flexcpData, NewRow, COLP_���)) > 0 _
                    Or .Cell(flexcpData, NewRow, COLP_���) = "4" And Val(.Cell(flexcpData, NewRow, COLP_��������)) = 1 Then
                    '��ʾҩƷ���������ĵĿ��
                    If InStr(",5,6,7,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                        stbThis.Panels(2).Text = vsAdvice.TextMatrix(lngRow, COL_���) & "��" & vsAdvice.TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & _
                            FormatEx(Val(vsAdvice.TextMatrix(lngRow, COL_���)), 5) & vsAdvice.TextMatrix(lngRow, COL_סԺ��λ)
                    Else
                        'ͬһ������ȡ:ҩƷ��סԺ��λ,���İ��ۼ۵�λ
                        stbThis.Panels(2).Text = .TextMatrix(NewRow, COLP_�շ���Ŀ) & "��" & .TextMatrix(NewRow, COLP_ִ�п���) & "���ÿ�棺" & _
                            FormatEx(GetStock(Val(.Cell(flexcpData, NewRow, COLP_�շ���Ŀ)), Val(.Cell(flexcpData, NewRow, COLP_ִ�п���))), 5) & .TextMatrix(NewRow, COLP_��λ)
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub vsPrice_GotFocus()
    vsPrice.BackColorSel = &HFFCC99
End Sub

Private Sub vsPrice_LostFocus()
    vsPrice.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵�������ص��кŷ�Χ��������ҩ;�����к�
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

Private Sub InitAdviceTable()
'���ܣ���ʼ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;����,850,1;����,750,1;סԺ��,750,1;����,500,4;�ѱ�,750,1;" & _
        "Ӥ��,550,1;ҽ������,2000,1;���,2000,1;����,600,7;��λ,450,1;����,600,7;��λ,450,1;���,850,7;" & _
        "Ƶ��,1000,1;�÷�,1000,1;ҽ������,1500,1;ִ��ʱ��,1000,1;�״�ʱ��,1080,1;ĩ��ʱ��,1080,1;ִ�п���,850,1;����ִ��,850,1;ִ������,850,1;" & _
        "����ID;��ҳID;�Ա�;����;ID;���ID;���˿���ID;��������ID;����ҽ��;�������;������ĿID;�Ƽ�����;ִ������ID;ִ�п���ID;" & _
        "ҩƷID;����ϵ��;סԺ��װ;סԺ��λ;�ɷ����;ҩ������;�Ƿ���;���;����;�ֽ�ʱ��"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_ѡ�� + 1 - .FixedCols
        .RowHeight(0) = 320
    End With
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2500,1;����,450,4;����,900,7;" & _
        "��λ,500,1;����,1000,7;Ӧ�ս��,1200,7;ʵ�ս��,1200,7;ִ�п���,1000,1;��������,850,1;����,450,4;�к�"
    arrHead = Split(strHead, ";")
    With vsPrice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Function Decodeҩ���û�() As String
'���ܣ�����ҩ���û�����,�����ֶ�Decode���
'˵����SQL��ѯ��,"����ҽ����¼"����Ϊ"A"
    Dim i As Long, strTmp As String
    
    mrsҩ��.Filter = 0
    If Not mrsҩ��.EOF Then
        For i = 1 To mrsҩ��.RecordCount
            If mrsҩ��!ID <> mrsҩ��!��ID Then
                strTmp = strTmp & "," & mrsҩ��!ID & "," & mrsҩ��!��ID
            End If
            mrsҩ��.MoveNext
        Next
    End If
    If strTmp <> "" Then
        Decodeҩ���û� = "Decode(A.ִ�п���ID" & strTmp & ",A.ִ�п���ID)"
    Else
        Decodeҩ���û� = "A.ִ�п���ID"
    End If
End Function

Private Sub DeleteCurRow(ByVal lngRow As Long, ByVal lng���ID As Long)
'���ܣ��ڴ���������嵥�Ĺ�����ɾ������������
    Dim i As Long
    With vsAdvice
        'ɾ����ǰ��
        .RemoveItem lngRow
        
        'ɾ���䷽��һ����ҩ���Ѿ��������
        If lng���ID <> 0 Then
            For i = .Rows - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                    .RemoveItem i
                End If
            Next
        End If
    End With
End Sub

Private Function Calc��������ʱ��(dbl���� As Double, lng���� As Long, str�ֽ�ʱ�� As String, ByVal strEnd As String, rsSend As ADODB.Recordset, rsDrug As ADODB.Recordset) As Boolean
'���ܣ��Գ��ڳ�ҩҽ����������,ִ�д���,ִ��ʱ��ֽ�
'������rsDrug=����ҩƷ���������Ϣ
'      rsSend=������ǰҩƷҽ���������Ϣ
'      strEnd=���η��͵Ľ���ʱ��
'���أ�dbl����=סԺ��װ
'      lng����=ִ�д���(��Ϊ��ҩ;����ִ�д���)
'      str�ֽ�ʱ��=�����ִ��ʱ��ֽ�
    Dim datBegin As Date, datEnd As Date, strPause As String
    
    '��ǰҽ������ͣʱ���:"��ͣʱ��,��ʼʱ��;...."
    strPause = GetAdvicePause(rsSend!ID)
    
    '��ǰҽ���ķ��ͼ���ʱ���
    datBegin = rsSend!��ʼִ��ʱ��
    If Not IsNull(rsSend!�ϴ�ִ��ʱ��) Then
        datBegin = Calc�����ڿ�ʼʱ��(rsSend!��ʼִ��ʱ��, rsSend!�ϴ�ִ��ʱ��, rsSend!Ƶ�ʼ��, rsSend!�����λ)
        
        '����������ִ�е�ʱ�䲻�ټ���,����ͨ����ͣ��ʽ������
        strPause = strPause & ";" & Format(datBegin, "yyyy-MM-dd HH:mm:ss") & "," & Format(rsSend!�ϴ�ִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
        If Left(strPause, 1) = ";" Then strPause = Mid(strPause, 2)
    End If
    datEnd = CDate(strEnd)
    If Not IsNull(rsSend!ִ����ֹʱ��) Then
        If rsSend!ִ����ֹʱ�� < CDate(strEnd) Then
            datEnd = rsSend!ִ����ֹʱ��
        End If
    End If
    
    '�Ȱ���������ʱ��μ���ֽ�ʱ�估����
    str�ֽ�ʱ�� = Calc���ڷֽ�ʱ��(datBegin, datEnd, strPause, rsSend!ִ��ʱ�䷽��, rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
    If str�ֽ�ʱ�� = "" Then Calc��������ʱ�� = True: Exit Function

    lng���� = UBound(Split(str�ֽ�ʱ��, ",")) + 1
    
    '�ٰ�ҩƷ�������Լ�������(��סԺ��λ),��ʱ�����ͷֽ�ʱ���������
    dbl���� = Calc����ҩƷ����( _
        rsSend!��ʼִ��ʱ��, lng����, str�ֽ�ʱ��, rsSend!��������, _
        rsDrug!����ϵ��, rsDrug!סԺ��װ, Nvl(rsDrug!�ɷ����, 0), _
        Nvl(rsSend!ִ����ֹʱ��, CDate("3000-01-01")), strPause, Nvl(rsSend!ִ��ʱ�䷽��), _
        rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ, mblnLimit)
    
    Calc��������ʱ�� = True
End Function

Private Function CheckStock(ByVal lngRow As Long, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean, Optional ByVal blnCurPati As Boolean) As Boolean
'���ܣ����ݿ���������鷢��ҩƷ�Ŀ��
'������lngRow=ҽ���к�
'      blnCurPati=�Ƿ�ֻ�Ե�ǰ���˽��л��ܼ��,���ڷ��͹�����,��Ϊ�ǰ������ύ,��ʱ������ȡ�Ŀ����׼ȷ��
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = GetStockCheck(Val(.TextMatrix(lngRow, COL_ִ�п���ID)))
        bln���� = Val(.TextMatrix(lngRow, COL_ҩ������)) = 1
        blnʱ�� = Val(.TextMatrix(lngRow, COL_�Ƿ���)) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = .TextMatrix(lngRow, COL_סԺ��λ) '������ʾ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����:סԺ��װ
            If .TextMatrix(lngRow, COL_�������) = "7" Then
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Val(.TextMatrix(lngRow, COL_�ɷ����)) = 0 Then
                    dbl���� = Val(.TextMatrix(lngRow, COL_����)) * Val(.TextMatrix(lngRow, COL_����))
                    dbl���� = dbl���� / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ))
                Else
                    dbl���� = IntEx(Val(.TextMatrix(lngRow, COL_����)) / Val(.TextMatrix(lngRow, COL_����ϵ��)) / Val(.TextMatrix(lngRow, COL_סԺ��װ)))
                    dbl���� = dbl���� * Val(.TextMatrix(lngRow, COL_����))
                End If
            Else
                dbl���� = Val(.TextMatrix(lngRow, COL_����))
            End If
            
            '��ǰ���ÿ��:סԺ��װ,��ȥǰ����ͬҩƷҪ���͵Ŀ��
            For i = lngRow - 1 To .FixedRows Step -1
                If blnCurPati And Val(.TextMatrix(i, COL_����ID)) = Val(.TextMatrix(lngRow, COL_����ID)) Or Not blnCurPati Then
                    blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                    If blnDo Then
                        blnDo = Val(.TextMatrix(i, COL_ҩƷID)) = Val(.TextMatrix(lngRow, COL_ҩƷID)) _
                            And Val(.TextMatrix(i, COL_ִ�п���ID)) = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
                    End If
                    If blnDo Then
                        blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                    End If
                    If blnDo Then
                        If .TextMatrix(i, COL_�������) = "7" Then
                            '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                            If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                dbl�ѷ���� = dbl�ѷ���� + _
                                    Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                    / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))
                            Else
                                dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                    * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                            End If
                        Else
                            dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            dbl���ÿ�� = Val(.TextMatrix(lngRow, COL_���))
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        strTmp = "ҩ��������ʱ��ҩƷ""" & .TextMatrix(lngRow, COL_���) & """��治�㣺" & vbCrLf & vbCrLf & _
                            .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                            IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                            "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                    Else
                        strTmp = """" & .TextMatrix(lngRow, COL_���) & """��治�㣺" & vbCrLf & vbCrLf & _
                            .TextMatrix(lngRow, COL_ִ�п���) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                            IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "��" & _
                            "���η�������" & FormatEx(dbl����, 5) & strTmp & "��"
                    End If
                    If .Cell(flexcpData, lngRow, COL_���) <> "" Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "������ڷ����嵥��ѡ���ҩƷ�����㹻�����������"
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҩƷ��"
                    End If
                    
                    strTmp = "����" & .TextMatrix(lngRow, COL_����) & "��" & vbCrLf & vbCrLf & strTmp
                    
                    .Redraw = flexRDDirect:
                    Call .ShowCell(lngRow, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckStock = True
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckStock = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Private Function CheckPriceStock(ByVal lngRow As Long, rsPrice As ADODB.Recordset, ByVal lng�ⷿID As Long, ByVal dbl���� As Double, _
    rsTotal As ADODB.Recordset, Optional bln�����ʾ As Boolean, Optional blnʱ����ʾ As Boolean, Optional blnĬ�Ϸ��� As Boolean) As Boolean
'���ܣ����͹�����ʱ���Է�ҩ��ҩƷ���������õ����ļƼ۽��п����(�ۼƼ��)
'������lngRow=ҽ���к�
'      dbl����=�Ѽ���õļƼ�����(�ۼ۵�λ)
'      rsTotal=��ǰ����ǰ�����ۼƷ��͵ļƼ�ҩƷ����������(�ۼ۵�λ)
'      bln�����ʾ,blnʱ����ʾ,blnĬ�Ϸ���=������ʾ�������ʾ����
'���أ�������ʾ���Ƿ��ѡ��״̬�����˴���
    Dim int����� As Integer, dbl���� As Double
    Dim dbl���ÿ�� As Double, dbl�ѷ���� As Double
    Dim bln����ʱ�� As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim vMsg As VbMsgBoxResult, strTmp As String
    Dim blnDo As Boolean, i As Long
    
    With vsAdvice
        'ҩƷ�����(0-�����;1-���,��������;2-��飬�����ֹ)
        int����� = GetStockCheck(lng�ⷿID)
        bln���� = Nvl(rsPrice!����, 0) = 1
        blnʱ�� = Nvl(rsPrice!�Ƿ���, 0) = 1
        
        '������ʱ��ҩƷ����Ҫ���㹻�Ŀ��,�������ݿ�����������
        If int����� <> 0 Or bln���� Or blnʱ�� Then
            strTmp = Nvl(rsPrice!סԺ��λ, Nvl(rsPrice!���㵥λ)) '������ʾ
            
            '������Ͳ����ֹʱ,����ʱ��Ͳ��ص�������
            bln����ʱ�� = int����� <> 2 And (bln���� Or blnʱ��)
            
            '��ǰҩƷ����������:סԺ��װ
            dbl���� = Format(dbl���� / Nvl(rsPrice!סԺ��װ, 1), "0.00000")
            
            '��ǰ���ÿ��:סԺ��װ,��ȥǰ����ͬҩƷҽ��Ҫ���͵Ŀ��
            If InStr(",5,6,7,", rsPrice!���) > 0 Then
                For i = lngRow - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(i, COL_����ID)) = Val(.TextMatrix(lngRow, COL_����ID)) Then
                        blnDo = InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0
                        If blnDo Then
                            blnDo = Val(.TextMatrix(i, COL_ҩƷID)) = rsPrice!ID And Val(.TextMatrix(i, COL_ִ�п���ID)) = lng�ⷿID
                        End If
                        If blnDo Then
                            blnDo = .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing
                        End If
                        If blnDo Then
                            If .TextMatrix(i, COL_�������) = "7" Then
                                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                    dbl�ѷ���� = dbl�ѷ���� + _
                                        Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) _
                                        / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ))
                                Else
                                    dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����)) _
                                        * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_סԺ��װ)))
                                End If
                            Else
                                dbl�ѷ���� = dbl�ѷ���� + Val(.TextMatrix(i, COL_����))
                            End If
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
            '�Ƽ۲���Ҫ���͵��ۼ�����
            rsTotal.Filter = "��ĿID=" & rsPrice!ID & " And �ⷿID=" & lng�ⷿID
            Do While Not rsTotal.EOF
                dbl�ѷ���� = dbl�ѷ���� + Format(rsTotal!���� / Nvl(rsPrice!סԺ��װ, 1), "0.00000")
                rsTotal.MoveNext
            Loop
            
            dbl���ÿ�� = Format(GetStock(rsPrice!ID, lng�ⷿID, 2), "0.00000")
            dbl���ÿ�� = dbl���ÿ�� - dbl�ѷ����
            
            If dbl���� > dbl���ÿ�� Then
                If (Not bln����ʱ�� And int����� <> 0 And bln�����ʾ) Or (bln����ʱ�� And blnʱ����ʾ) Then
                    '��һ��û��ѡ������ʾ,����ʾ
                    If bln����ʱ�� Then
                        strTmp = "ҽ��""" & .TextMatrix(lngRow, COL_ҽ������) & """�ķ�����ʱ�ۼƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                            vbCrLf & vbCrLf & Get��������(lng�ⷿID) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                            IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                    Else
                        strTmp = "ҽ��""" & .TextMatrix(lngRow, COL_ҽ������) & """�ļƼ���Ŀ""" & rsPrice!���� & """��治�㣺" & _
                            vbCrLf & vbCrLf & Get��������(lng�ⷿID) & "���ÿ�棺" & FormatEx(dbl���ÿ��, 5) & strTmp & _
                            IIF(dbl�ѷ���� <> 0, "(�ſ�ǰ����ͬҩƷ������)", "") & "�����η���������" & FormatEx(dbl����, 5) & strTmp & "��"
                    End If
                    If int����� = 1 And Not bln����ʱ�� Then
                        strTmp = strTmp & vbCrLf & vbCrLf & "Ҫ���͸�ҽ����"
                    End If
                    strTmp = "����" & .TextMatrix(lngRow, COL_����) & "��" & vbCrLf & vbCrLf & strTmp
                    
                    .Redraw = flexRDDirect
                    .Row = GetVisibleRow(lngRow, True)
                    Call .ShowCell(.Row, COL_ѡ��)
                    Screen.MousePointer = 0
                    vMsg = frmMsgBox.ShowMsgBox(strTmp, Me, int����� = 2 Or bln����ʱ��)
                    
                    If bln����ʱ�� Then
                        If vMsg = vbIgnore Then blnʱ����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 2 Then '����ֹ
                        If vMsg = vbIgnore Then bln�����ʾ = False
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then '�������
                        If vMsg = vbYes Or vMsg = vbIgnore Then
                            If vMsg = vbIgnore Then bln�����ʾ = False
                            blnĬ�Ϸ��� = True
                        ElseIf vMsg = vbNo Or vMsg = vbCancel Then
                            If vMsg = vbCancel Then bln�����ʾ = False
                            blnĬ�Ϸ��� = False
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                    Screen.MousePointer = 11
                    .Refresh: .Redraw = flexRDNone
                Else
                    '��һ��ѡ���˲�����ʾ
                    If int����� = 2 Or bln���� Or blnʱ�� Then
                        .Cell(flexcpData, lngRow, COL_ѡ��) = 1 '��ǰ����ֹѡ��
                        Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("F").Picture
                        CheckPriceStock = True
                    ElseIf int����� = 1 Then
                        '������һ�εĽ������
                        If Not blnĬ�Ϸ��� Then
                            Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = Nothing 'ȱʡ������
                            CheckPriceStock = True
                        End If
                    End If
                End If
            End If
        End If
        
        '���δ��ʾ��Ҫ����,�����ۼƷ�������
        If Not CheckPriceStock Then
            rsTotal.AddNew
            If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_���ID))
            Else
                rsTotal!ҽ��ID = Val(.TextMatrix(lngRow, COL_ID))
            End If
            rsTotal!��ĿID = rsPrice!ID
            rsTotal!�ⷿID = lng�ⷿID
            rsTotal!���� = dbl����
            rsTotal.Update
        End If
    End With
End Function

Private Sub DeleteDrugRow(rsSend As ADODB.Recordset, ByVal lngRow As Long, lngDel���ID As Long)
'���ܣ�ɾ����Ӧ��ҩƷ��,����ҩƷͣ�û�������ԭ���Ҳ�����Ч���ʱ
'���أ�lngDel���ID-��Ҫͬʱɾ�����������ҽ����ʶ
    Dim strMsg As String
    
    With vsAdvice
        If rsSend!������� = "7" Then
            strMsg = "���в�ҩ��Ӧ����ҩ�䷽�޷����ͣ�" & vbCrLf & vbCrLf & "����" & Nvl(rsSend!ҽ������)
        Else
            strMsg = "��ҩƷ(��һ����ҩ������ҩƷ)�޷����ͣ�" & vbCrLf & vbCrLf & "����" & Nvl(rsSend!ҽ������)
        End If
        strMsg = strMsg & vbCrLf & vbCrLf & "û�з�����Ч��ҩƷ�����Ϣ����ҩƷ�����Ѿ���ͣ�û�������סԺ���ˡ�"
        strMsg = strMsg & vbCrLf & "���ȵ�ҩƷĿ¼�����д�����[ȷ��]������������ҽ����"
        .Redraw = flexRDDirect
        Call .ShowCell(lngRow, COL_ѡ��)
        Screen.MousePointer = 0
        MsgBox strMsg, vbInformation, gstrSysName
        
        Screen.MousePointer = 11
        lngDel���ID = Nvl(rsSend!���ID, 0)
        Call DeleteCurRow(lngRow, rsSend!���ID)
        .Refresh: .Redraw = flexRDNone
    End With
End Sub

Private Sub SeekMatchDrug(rsSend As ADODB.Recordset, rsDrug As ADODB.Recordset, ByVal dbl���� As Double, vBookMark As Variant, Optional strList As String)
'���ܣ�����ҩƷ�Ķ�����λȱʡ���ʵĹ��,���������ҩƷ��Ϣ�������
'������rsSend=Ҫ���͵�ҽ����Ϣ
'      rsDrug=ҩƷ��Ϣ
'      dbl����=Ҫ���͵�ҩƷ����,Ϊ0ʱ��ʾ��δ�������
'      vBookMark=�������ڶ�λ���λ�õ���ǩ
'      strList=������Ч�ɹ�ѡ��Ĺ��,������������������
    Dim vPreBookMark As Variant
    Dim lng���� As Long
        
    vPreBookMark = 0
    If Not rsDrug.EOF And Not rsDrug.BOF Then
        vPreBookMark = rsDrug.Bookmark
    End If
    
    rsDrug.MoveFirst
    vBookMark = 0: strList = ""
    Do While Not rsDrug.EOF
        '�ſ�ͣ�õ�ҩƷ
        If Nvl(rsDrug!����ʱ��, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", Nvl(rsDrug!�������, 0)) > 0 Then
            If CInt(Nvl(rsSend!��������, 0)) <> 0 And (Nvl(rsDrug!���, 0) > dbl���� Or Nvl(rsDrug!���, 0) = dbl���� And dbl���� <> 0) Then
                'Ѱ�Ҽ�����λΪ��������С�����Ĺ��
                If rsDrug!����ϵ�� / rsSend!�������� = Int(rsDrug!����ϵ�� / rsSend!��������) Then
                    If rsDrug!����ϵ�� / rsSend!�������� < lng���� Or lng���� = 0 Then
                        vBookMark = rsDrug.Bookmark
                        lng���� = rsDrug!����ϵ�� / rsSend!��������
                    End If
                End If
            End If
            strList = strList & "|#" & rsDrug!ҩƷID & ";" & rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "") & vbTab & "���:" & Nvl(rsDrug!���, 0) & rsDrug!סԺ��λ
        End If
        rsDrug.MoveNext
    Loop
    If vBookMark = 0 Then
        rsDrug.MoveFirst
        Do While Not rsDrug.EOF
            If Nvl(rsDrug!����ʱ��, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", Nvl(rsDrug!�������, 0)) > 0 Then
                If Nvl(rsDrug!���, 0) > dbl���� Or Nvl(rsDrug!���, 0) = dbl���� And dbl���� <> 0 Then
                    vBookMark = rsDrug.Bookmark: Exit Do
                End If
            End If
            rsDrug.MoveNext
        Loop
    End If
    strList = Mid(strList, 2)
    
    If vBookMark = 0 And vPreBookMark <> 0 Then 'û�ҵ�ʱ�ָ�ԭ��λ��
        rsDrug.Bookmark = vPreBookMark
    End If
End Sub

Private Function LoadAdviceSend(ByVal strEnd As String, ByVal int��Ч As Integer, _
    ByVal lngҩ��ID As Long, ByVal str����IDs As String, ByVal str��ҩIDs As String) As Boolean
'���ܣ�����������ȡ����ʾҪ���͵�ҩƷҽ���嵥
'������strEnd=���͵��Ľ���ʱ��(yyyy-MM-dd HH:mm:ss),����û��
'      int��Ч=0-����,1-����
'      lngҩ��ID=Ҫ����ҩƷ��ҩ��ID,0��ʾ������
'      str����IDs=Ҫ����ҽ������ID��(12,23,34....)
'      str��ҩIDs=Ҫ���͵ĸ�ҩ;��Ϊָ��ID���ĳ�ҩ
'      mrsҩ��=����ҩƷ�û���Ϣ�ļ�¼��
'˵����ע��CellData�д�ŵ��и�������
'   RowData��0-δ���͵�,-1-�ѳɹ����͵�
'   COL_ѡ��0-������ѡ���,1-��ֹ�ı�ѡ��״̬��
'   COL_�������1-��ҩ;����2-��ҩ�巨��3-��ҩ�÷�
'   COL_Ӥ�������Ӥ�����
'   COL_ҽ�����ݣ����������Ŀ����,������ʾ�Ƽ�ҽ��
'   COL_��񣺴�ų�ҩ��ѡ��Ĺ����������(ComboList)
'   COL_�ֽ�ʱ��:�����޷ֽ�ʱ��ʱ,��ŷ��÷���ʱ��

    Dim rsSend As New ADODB.Recordset
    Dim rsDrug As New ADODB.Recordset
    Dim strSQL As String, str��Ч���� As String
    Dim strҩ������ As String, str��ҩ;�� As String
    Dim i As Long, j As Long, k As Long, lngTmp As Long, strTmp As String
    Dim lng����ID As Long, lng���� As Long, lng������ As Long
    Dim lngRow As Long, lngDel���ID As Long, vBookMark As Variant
    Dim strҩ���û� As String, str���� As String, bln����ʱ�� As Boolean
    Dim lng���� As Long, lng��С���� As Long, str�÷� As String
    Dim str�ֽ�ʱ�� As String, dbl���� As Double, cur��� As Currency
    Dim blnʱ����ʾ As Boolean, bln�����ʾ As Boolean, blnĬ�Ϸ��� As Boolean
    Dim blnƷ��ҩƷ As Boolean
        
    Screen.MousePointer = 11
    
    stbThis.Panels(3).Text = "": Call Form_Resize
    If int��Ч = 0 Then
        lblInfo.Caption = "���η��ͣ�����ҩƷҽ��������ʱ�䣺" & strEnd
    Else
        lblInfo.Caption = "���η��ͣ���ʱҩƷҽ��"
    End If
    
    vsPrice.Rows = vsPrice.FixedRows
    vsPrice.Rows = vsPrice.FixedRows + 1
    vsAdvice.Rows = vsAdvice.FixedRows '��ɾ���й���
    Me.Refresh
        
    blnʱ����ʾ = True: bln�����ʾ = True: blnĬ�Ϸ��� = True
    
    '��ȡ�����嵥:ÿ��ҽ����¼(���ϵ�ҽ����������ʱ��,���Ϻ���Ч)
    '----------------------------------------------------------------------------------------------------------
    '�û���ҩ��ID
    strҩ���û� = Decodeҩ���û�
    
    '��ͬ��Ч������
    If int��Ч = 0 Then
        str��Ч���� = _
            " And A.��ʼִ��ʱ��<=[1] And (A.�ϴ�ִ��ʱ��<[1] Or A.�ϴ�ִ��ʱ�� is NULL)" & _
            " And (A.ִ����ֹʱ��>A.�ϴ�ִ��ʱ�� Or A.ִ����ֹʱ�� is NULL Or A.�ϴ�ִ��ʱ�� Is NULL)" & _
            " And (A.ִ����ֹʱ��>A.��ʼִ��ʱ�� Or A.ִ����ֹʱ�� is NULL)" & _
            " And Nvl(A.ҽ��״̬,0) Not IN(1,2,4) And Nvl(A.ҽ����Ч,0)=0"
    Else
        str��Ч���� = " And Nvl(A.ҽ��״̬,0) Not IN(1,2,4,8,9) And Nvl(A.ҽ����Ч,0)=1"
    End If
    '���͵�ҽ����Χ����
    If InStr(mstrPrivs, "ȫԺҽ������") = 0 Then
        str��Ч���� = str��Ч���� & " And A.����ҽ�� In(" & _
            " Select Distinct B.����" & _
            " From ������Ա A,��Ա�� B,��Ա����˵�� C" & _
            " Where A.��ԱID=B.ID And B.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
            "   And A.����ID In(" & _
            "     Select Distinct B.����ID From ������Ա A,��λ״����¼ B" & _
            "     Where A.��ԱID=(Select ��ԱID From �ϻ���Ա�� Where �û���=User)" & _
            "       And A.����ID=B.����ID)" & _
            ")"
    End If
    
    For k = 0 To UBound(Split(str����IDs, ","))
        'ֻ����ָ��ҩ����ҩƷ:ҩ���û�֮���Ϊ׼
        If lngҩ��ID <> 0 Then
            strҩ������ = "Select ID From ����ҽ����¼ X" & _
                " Where ������� IN('5','6','7') And (X.���ID=A.���ID Or X.���ID=A.ID)" & _
                " And " & Replace(strҩ���û�, "A.ִ�п���ID", "ִ�п���ID") & "+0=[3] And ����ID=[2]"
            strҩ������ = " And Exists(" & strҩ������ & ")"
        End If
        
        '����ĸ�ҩ;������(������Ӧ�ĳ�ҩ)
        If str��ҩIDs <> "" Then
            str��ҩ;�� = "Select ID From ����ҽ����¼ X" & _
                " Where �������='E' And (X.ID=A.ID Or X.ID=A.���ID)" & _
                " And ����ID=[2] And ������ĿID+0 IN(" & str��ҩIDs & ")"
            str��ҩ;�� = " And Exists(" & str��ҩ;�� & ")"
        End If
        
        '��ȡ������ϸ:(δ�ų�����������ҽ��)
        '����������(��ҩ;��,�÷�,�巨����Ϊ),�������ȶ�ȡ����
        strSQL = "Select A.ID,A.���ID,Nvl(A.���ID,A.ID) as ��ID,Nvl(X.���,A.���) as ���," & _
            " A.�������,A.������ĿID,E.���� as ������Ŀ,A.�շ�ϸĿID,A.Ӥ��," & _
            " A.����ID,A.��ҳID,C.סԺ��,B.��Ժ���� as ����,D.���� as ����,C.����,C.�Ա�,C.����,B.�ѱ�,B.����," & _
            " A.��ʼִ��ʱ��,A.�ϴ�ִ��ʱ��,A.ҽ������,A.����,A.�ܸ�����,A.��������,E.���㵥λ,A.ִ����ֹʱ��," & _
            " A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ҽ������,A.ִ��ʱ�䷽��,A.���˿���ID,A.��������ID,A.����ҽ��," & _
            " A.�Ƽ�����,A.ִ������," & strҩ���û� & " as ִ�п���ID,F.���� as ִ�п���" & _
            " From ����ҽ����¼ A,������ҳ B,������Ϣ C,���ű� D,������ĿĿ¼ E,���ű� F,����ҽ����¼ X" & _
            " Where A.����ID=[2] And A.����ID=C.����ID And B.��Ժ����ID=D.ID" & _
            " And A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.��Ժ���� is NULL" & _
            " And A.���ID=X.ID(+) And A.������ĿID=E.ID And " & strҩ���û� & "=F.ID(+)" & _
            " And A.������� IN('5','6','7','E') And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & str��Ч���� & strҩ������ & str��ҩ;�� & _
            " Order by D.����,LPAD(B.��Ժ����,10,' '),A.Ӥ��,���,��ID,A.���"
        
        On Error GoTo errH
        Set rsSend = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(IIF(strEnd = "", "1990-01-01", strEnd)), Val(Split(str����IDs, ",")(k)), lngҩ��ID)
        
        '���㲢��ʾ�����嵥
        '----------------------------------------------------------------------------------------------------------
        If Not rsSend.EOF Then
            With vsAdvice
                .Redraw = flexRDNone
                For i = 1 To rsSend.RecordCount
                    If rsSend!������� = "E" And IsNull(rsSend!���ID) And rsSend!ID <> Val(.TextMatrix(.Rows - 1, COL_���ID)) Then
                        GoTo NextLoop '������ҩ����������ҽ�������ɼ�����
                    ElseIf (rsSend!ID = lngDel���ID Or Nvl(rsSend!���ID, 0) = lngDel���ID) And lngDel���ID <> 0 Then
                        GoTo NextLoop 'һ����ҩ���䷽�е�һ�������Ѿ����ܷ���,�����鲻�ܷ���
                    Else
                        lngDel���ID = 0
                    End If
                                                    
                    '���뵱ǰ��
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                    .Cell(flexcpPictureAlignment, lngRow, COL_ѡ��) = 4
                    Set .Cell(flexcpPicture, lngRow, COL_ѡ��) = img16.ListImages("T").Picture
                    
                    '���ظ�����
                    If rsSend!������� = "7" Then
                        .RowHidden(lngRow) = True '�в�ҩ
                    ElseIf rsSend!������� = "E" Then
                        If Not IsNull(rsSend!���ID) Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_�������) = 2 '��ҩ�巨
                        ElseIf Val(.TextMatrix(lngRow - 1, COL_���ID)) = rsSend!ID _
                            And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                            .RowHidden(lngRow) = True
                            .Cell(flexcpData, lngRow, COL_�������) = 1 '��ҩ;��
                        Else
                            .Cell(flexcpData, lngRow, COL_�������) = 3 '��ҩ�÷�
                        End If
                    End If
                    
                    'һ���и�ֵ
                    '---------------------------------------------------------------
                    .Cell(flexcpData, lngRow, COL_Ӥ��) = CLng(Nvl(rsSend!Ӥ��, 0))
                    If Nvl(rsSend!Ӥ��, 0) = 0 Then
                        .TextMatrix(lngRow, COL_Ӥ��) = "����"
                    Else
                        .TextMatrix(lngRow, COL_Ӥ��) = "Ӥ��" & rsSend!Ӥ��
                        .ColHidden(COL_Ӥ��) = False '��Ӥ��ҽ��ʱ����ʾ
                    End If
                    .TextMatrix(lngRow, COL_����) = rsSend!����
                    If InStr(str���� & ",", "," & rsSend!���� & ",") = 0 Then
                        If str���� <> "" Then .ColHidden(COL_����) = False
                        str���� = str���� & "," & rsSend!����
                    End If
                    
                    .TextMatrix(lngRow, COL_����ID) = rsSend!����ID
                    .TextMatrix(lngRow, COL_��ҳID) = rsSend!��ҳID
                    .TextMatrix(lngRow, COL_����) = rsSend!����
                    .TextMatrix(lngRow, COL_�Ա�) = Nvl(rsSend!�Ա�)
                    .TextMatrix(lngRow, COL_����) = Nvl(rsSend!����)
                    .TextMatrix(lngRow, COL_סԺ��) = Nvl(rsSend!סԺ��)
                    .TextMatrix(lngRow, COL_����) = Nvl(rsSend!����)
                    .TextMatrix(lngRow, COL_�ѱ�) = Nvl(rsSend!�ѱ�)
                    
                    .TextMatrix(lngRow, COL_ID) = rsSend!ID
                    .TextMatrix(lngRow, COL_���ID) = Nvl(rsSend!���ID)
                    .TextMatrix(lngRow, COL_�������) = rsSend!�������
                    .TextMatrix(lngRow, COL_������ĿID) = rsSend!������ĿID
                    
                    .TextMatrix(lngRow, COL_ҽ������) = Nvl(rsSend!ҽ������)
                    .Cell(flexcpData, lngRow, COL_ҽ������) = CStr(Nvl(rsSend!������Ŀ)) '������ʾ�Ƽ�ҽ��
                    
                    .TextMatrix(lngRow, COL_ҽ������) = Nvl(rsSend!ҽ������)
                    .TextMatrix(lngRow, COL_ִ��ʱ��) = Nvl(rsSend!ִ��ʱ�䷽��)
                    .TextMatrix(lngRow, COL_Ƶ��) = Nvl(rsSend!ִ��Ƶ��)
                    
                    .TextMatrix(lngRow, COL_���˿���ID) = Nvl(rsSend!���˿���ID)
                    .TextMatrix(lngRow, COL_��������ID) = Nvl(rsSend!��������ID)
                    .TextMatrix(lngRow, COL_����ҽ��) = Nvl(rsSend!����ҽ��)
                    
                    .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsSend!�Ƽ�����, 0)
                    .TextMatrix(lngRow, COL_ִ������ID) = Nvl(rsSend!ִ������, 0)
                                                        
                    '��ʾ��Ҫִ�п���
                    .TextMatrix(lngRow, COL_ִ�п���) = Nvl(rsSend!ִ�п���)
                    
                    '��ʾ����ִ�п���
                    If rsSend!������� = "E" And IsNull(rsSend!���ID) Then
                        If InStr(",7,E,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                            '��ҩ�÷�
                            .TextMatrix(lngRow, COL_����ִ��) = Nvl(rsSend!ִ�п���)
                            .Cell(flexcpData, lngRow, COL_����ִ��) = CStr(Nvl(rsSend!ִ�п���))
                        ElseIf InStr(",5,6,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                            '��ҩ;��
                            For j = lngRow - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                    .TextMatrix(j, COL_����ִ��) = Nvl(rsSend!ִ�п���)
                                    .Cell(flexcpData, j, COL_����ִ��) = CStr(Nvl(rsSend!ִ�п���))
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Nvl(rsSend!ִ�п���ID)
                                                        
                    '��ȡҩƷ�����Ϣ
                    '---------------------------------------------------------------
                    If InStr(",5,6,7", rsSend!�������) > 0 Then
                        Set rsDrug = New ADODB.Recordset
                        '�Ȱ���ͣ��ҩƷ,��ȷ��Ҫ���͵�ҽ���ټ��ͣ��
                        Set rsDrug = GetDrugInfo(rsSend!������ĿID, Nvl(rsSend!�շ�ϸĿID, 0), Nvl(rsSend!ִ�п���ID, 0), 2, False)
                        If rsDrug.EOF Then
                            'ҩƷû�ж�Ӧ�Ĺ����Ϣ
                            'ɾ����ǰ��(�������),��������һҽ��
                            Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                            lng��С���� = 0: GoTo NextLoop
                        ElseIf rsDrug.RecordCount > 1 Then
                            'Ѱ�Һ��ʵĹ��
                            Call SeekMatchDrug(rsSend, rsDrug, 0, vBookMark, strTmp)
                            If vBookMark <> 0 Then
                                rsDrug.Bookmark = vBookMark
                            Else
                                rsDrug.MoveFirst
                            End If
                            .Cell(flexcpData, lngRow, COL_���) = strTmp '��ѡ��Ĺ��
                            '���ȫ��(ָ��)���ͣ�õ�ҩƷ
                            If .Cell(flexcpData, lngRow, COL_���) = "" Then
                                Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                                lng��С���� = 0: GoTo NextLoop
                            End If
                        Else
                            '���ȫ��(ָ��)���ͣ�õ�ҩƷ
                            If Not (Nvl(rsDrug!����ʱ��, CDate("3000-01-01")) = CDate("3000-01-01") And InStr(",2,3,", Nvl(rsDrug!�������, 0)) > 0) Then
                                Call DeleteDrugRow(rsSend, lngRow, lngDel���ID)
                                lng��С���� = 0: GoTo NextLoop
                            End If
                        End If
                        .TextMatrix(lngRow, COL_���) = rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "")
                        .TextMatrix(lngRow, COL_ҩƷID) = rsDrug!ҩƷID
                        .TextMatrix(lngRow, COL_���) = Format(Nvl(rsDrug!���, 0), "0.00000") '��סԺ��װ
                        .TextMatrix(lngRow, COL_����ϵ��) = Nvl(rsDrug!����ϵ��, 1)
                        .TextMatrix(lngRow, COL_סԺ��װ) = Nvl(rsDrug!סԺ��װ, 1)
                        .TextMatrix(lngRow, COL_סԺ��λ) = Nvl(rsDrug!סԺ��λ)
                        .TextMatrix(lngRow, COL_�ɷ����) = Nvl(rsDrug!�ɷ����, 0)
                        .TextMatrix(lngRow, COL_ҩ������) = Nvl(rsDrug!ҩ������, 0)
                        .TextMatrix(lngRow, COL_�Ƿ���) = Nvl(rsDrug!�Ƿ���, 0)
                        
                        '�Ƿ����δȷ������Ʒ��ҩƷ
                        If .Cell(flexcpData, lngRow, COL_���) <> "" Then
                            .Cell(flexcpForeColor, lngRow, COL_���) = vbBlue 'ͻ����ʾ
                            blnƷ��ҩƷ = True
                        End If
                    End If
                                                                            
                    '���㷢�ʹ�����ִ�еķֽ�ʱ�䣬����
                    '---------------------------------------------------------------
                    If int��Ч = 0 Then
                        '����---------------------------------------------
                        If rsSend!������� = "7" Then
                            '��ҩ�䷽������ͣ,������Ȼ����ֹʱ�䡣�䷽��ĩ��ִ��ʱ��Ӧ������ֹʱ����ͬ
                            .TextMatrix(lngRow, COL_����) = rsSend!�ܸ�����
                            .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Calc�����ֽ�ʱ��(rsSend!�ܸ�����, rsSend!��ʼִ��ʱ��, rsSend!ִ����ֹʱ��, "", rsSend!ִ��ʱ�䷽��, rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(0), "MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(rsSend!�ܸ����� - 1), "MM-dd HH:mm")
                            
                            .TextMatrix(lngRow, COL_����) = Nvl(rsSend!��������) '����
                            .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                            .TextMatrix(lngRow, COL_����) = rsSend!�ܸ����� '����
                            .TextMatrix(lngRow, COL_������λ) = "��"
                        ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                            '��ǰҽ���ķ��ͼ���ʱ���
                            Call Calc��������ʱ��(dbl����, lng����, str�ֽ�ʱ��, strEnd, rsSend, rsDrug)
                            If str�ֽ�ʱ�� = "" Then
                                '�޷��ֽ�ʱ��(�类��ͣ��)
                                lngDel���ID = rsSend!���ID
                                Call DeleteCurRow(lngRow, rsSend!���ID)
                                lng��С���� = 0: GoTo NextLoop
                            End If
                            .TextMatrix(lngRow, COL_����) = lng����
                            .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                            .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "MM-dd HH:mm")
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "MM-dd HH:mm")
                            
                            .TextMatrix(lngRow, COL_����) = FormatEx(Nvl(rsSend!��������), 5)
                            .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                            .TextMatrix(lngRow, COL_����) = FormatEx(dbl����, 5)
                            .TextMatrix(lngRow, COL_������λ) = Nvl(rsDrug!סԺ��λ)
                            
                            If lng���� < lng��С���� Or lng��С���� = 0 Then lng��С���� = lng����
                            
                            '���ж������ѡ��ʱ�����ݿ���Ƿ��㹻�ٴζ�λ���
                            If .Cell(flexcpData, lngRow, COL_���) <> "" _
                                And Val(.TextMatrix(lngRow, COL_����)) > Val(.TextMatrix(lngRow, COL_���)) Then
                                Call SeekMatchDrug(rsSend, rsDrug, Val(.TextMatrix(lngRow, COL_����)), vBookMark)
                                If vBookMark <> 0 Then
                                    rsDrug.Bookmark = vBookMark
                                    .TextMatrix(lngRow, COL_���) = rsDrug!���� & IIF(Not IsNull(rsDrug!����), "(" & rsDrug!���� & ")", "") & IIF(Not IsNull(rsDrug!���), " " & rsDrug!���, "")
                                    .TextMatrix(lngRow, COL_ҩƷID) = rsDrug!ҩƷID
                                    .TextMatrix(lngRow, COL_���) = Format(Nvl(rsDrug!���, 0), "0.00000") '��סԺ��װ
                                    .TextMatrix(lngRow, COL_����ϵ��) = Nvl(rsDrug!����ϵ��, 1)
                                    .TextMatrix(lngRow, COL_סԺ��װ) = Nvl(rsDrug!סԺ��װ, 1)
                                    .TextMatrix(lngRow, COL_סԺ��λ) = Nvl(rsDrug!סԺ��λ)
                                    .TextMatrix(lngRow, COL_ҩ������) = Nvl(rsDrug!ҩ������, 0)
                                    .TextMatrix(lngRow, COL_�Ƿ���) = Nvl(rsDrug!�Ƿ���, 0)
                                End If
                            End If
                        Else
                            'һ����ҩ�İ���С��������(Ӱ���ҩ;���ƷѼ��ϴ�ִ��ʱ��)(������Ŀ����˷�)
                            If .Cell(flexcpData, lngRow, COL_�������) = 1 Then '��ҩ;��
                                For j = lngRow - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                        If Val(.TextMatrix(j, COL_����)) > lng��С���� Then
                                            .TextMatrix(j, COL_����) = lng��С����
                                            .TextMatrix(j, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(j, COL_�ֽ�ʱ��))
                                            .TextMatrix(j, COL_�״�ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(0), "MM-dd HH:mm")
                                            .TextMatrix(j, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "MM-dd HH:mm")
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                                lng��С���� = 0
                            End If
                            
                            .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����) '���������
                            .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                            If .Cell(flexcpData, lngRow, COL_�������) = 3 Then '��ҩ�÷�
                                .TextMatrix(lngRow, COL_������λ) = "��"
                            End If
                            
                            .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                            .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                        End If
                    Else
                        '����---------------------------------------------
                        If rsSend!������� = "7" Then
                            .TextMatrix(lngRow, COL_����) = rsSend!�ܸ�����
                            If Not IsNull(rsSend!ִ��ʱ�䷽��) Then
                                .TextMatrix(lngRow, COL_�ֽ�ʱ��) = Calc�����ֽ�ʱ��(rsSend!�ܸ�����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", rsSend!ִ��ʱ�䷽��, rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                                .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(0), "MM-dd HH:mm")
                                .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(lngRow, COL_�ֽ�ʱ��), ",")(rsSend!�ܸ����� - 1), "MM-dd HH:mm")
                            Else
                                '�޷ֽ�ʱ��(��������δ����ִ��ʱ����޷��ֽ�)
                                '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                                .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                            End If
                            .TextMatrix(lngRow, COL_����) = Nvl(rsSend!��������) '����
                            .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                            .TextMatrix(lngRow, COL_����) = rsSend!�ܸ����� '����
                            .TextMatrix(lngRow, COL_������λ) = "��"
                        ElseIf InStr(",5,6,", rsSend!�������) > 0 Then
                            '����������ҩ����
                            If Nvl(rsSend!Ƶ�ʴ���, 0) = 0 Or Nvl(rsSend!Ƶ�ʼ��, 0) = 0 Then
                                lng���� = 1 '����Ϊһ���Ե�����ҩƷ
                            ElseIf Nvl(rsSend!����, 0) <> 0 And Not IsNull(rsSend!ִ��Ƶ��) Then
                                'һ��Ƶ�����ڵĴ���
                                If rsSend!�����λ = "��" Then
                                    lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / 7))
                                ElseIf rsSend!�����λ = "��" Then
                                    lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��))
                                ElseIf rsSend!�����λ = "Сʱ" Then
                                    lng���� = IntEx(rsSend!���� * (rsSend!Ƶ�ʴ��� / rsSend!Ƶ�ʼ��) * 24)
                                End If
                            Else
                                '�ɷ���ҩƷʱ,�������Ե����ı��������ҩ;���Ĵ���,����һ��Ƶ�����ڵĴ�������
                                If Nvl(rsDrug!�ɷ����, 0) = 0 And Nvl(rsSend!��������, 0) <> 0 Then
                                    lng���� = IntEx(rsSend!�ܸ����� * rsDrug!����ϵ�� / rsSend!��������)
                                Else
                                    lng���� = Nvl(rsSend!Ƶ�ʴ���, 0)
                                End If
                            End If
                            If Not IsNull(rsSend!ִ��ʱ�䷽��) Then
                                str�ֽ�ʱ�� = Calc�����ֽ�ʱ��(lng����, rsSend!��ʼִ��ʱ��, CDate("3000-01-01"), "", rsSend!ִ��ʱ�䷽��, rsSend!Ƶ�ʴ���, rsSend!Ƶ�ʼ��, rsSend!�����λ)
                                If str�ֽ�ʱ�� <> "" Then
                                    .TextMatrix(lngRow, COL_�ֽ�ʱ��) = str�ֽ�ʱ��
                                    .TextMatrix(lngRow, COL_�״�ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(0), "MM-dd HH:mm")
                                    .TextMatrix(lngRow, COL_ĩ��ʱ��) = Format(Split(str�ֽ�ʱ��, ",")(lng���� - 1), "MM-dd HH:mm")
                                End If
                            Else
                                '�޷ֽ�ʱ��(��������δ����ִ��ʱ����޷��ֽ�)
                                '��¼���÷���ʱ��(��ҽ����ʼִ��ʱ��)
                                .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = Format(rsSend!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                            End If
                            .TextMatrix(lngRow, COL_����) = lng����
                            .TextMatrix(lngRow, COL_����) = FormatEx(Nvl(rsSend!��������), 5)
                            .TextMatrix(lngRow, COL_������λ) = Nvl(rsSend!���㵥λ)
                            .TextMatrix(lngRow, COL_����) = FormatEx(rsSend!�ܸ����� / rsDrug!סԺ��װ, 5) '��סԺ��λ��ʾ
                            .TextMatrix(lngRow, COL_������λ) = Nvl(rsDrug!סԺ��λ)
                            
                            If lng���� < lng��С���� Or lng��С���� = 0 Then lng��С���� = lng����
                        Else
                            '������һ����ҩ�İ���С��������(Ӱ���ҩ;���Ʒ�)
                            If .Cell(flexcpData, lngRow, COL_�������) = 1 Then '��ҩ;��
                                For j = lngRow - 1 To .FixedRows Step -1
                                    If Val(.TextMatrix(j, COL_���ID)) = rsSend!ID Then
                                        If Val(.TextMatrix(j, COL_����)) > lng��С���� Then
                                            .TextMatrix(j, COL_����) = lng��С����
                                            If .TextMatrix(j, COL_�ֽ�ʱ��) <> "" Then
                                                .TextMatrix(j, COL_�ֽ�ʱ��) = Trim�ֽ�ʱ��(lng��С����, .TextMatrix(j, COL_�ֽ�ʱ��))
                                                .TextMatrix(j, COL_�״�ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(0), "MM-dd HH:mm")
                                                .TextMatrix(j, COL_ĩ��ʱ��) = Format(Split(.TextMatrix(j, COL_�ֽ�ʱ��), ",")(lng��С���� - 1), "MM-dd HH:mm")
                                            End If
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                                lng��С���� = 0
                            End If
                            
                            .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����) '���������
                            .TextMatrix(lngRow, COL_����) = .TextMatrix(lngRow - 1, COL_����)
                            If .Cell(flexcpData, lngRow, COL_�������) = 3 Then '��ҩ�÷�
                                .TextMatrix(lngRow, COL_������λ) = "��"
                            End If
                            
                            .TextMatrix(lngRow, COL_�ֽ�ʱ��) = .TextMatrix(lngRow - 1, COL_�ֽ�ʱ��)
                            .Cell(flexcpData, lngRow, COL_�ֽ�ʱ��) = .Cell(flexcpData, lngRow - 1, COL_�ֽ�ʱ��)
                            .TextMatrix(lngRow, COL_�״�ʱ��) = .TextMatrix(lngRow - 1, COL_�״�ʱ��)
                            .TextMatrix(lngRow, COL_ĩ��ʱ��) = .TextMatrix(lngRow - 1, COL_ĩ��ʱ��)
                        End If
                    End If
                    
                    '������Ŀ�Ľ��:���ڲ鿴�����ʱ���
                    '---------------------------------------------------------------
                    .TextMatrix(lngRow, COL_���) = Format(Calcҽ�����ʽ��(lngRow), gstrDec)
                    
                    '�����ʱ��һЩ�����ۼ���ʾ���,��ҩ;��,�÷�,ִ�п���,ִ������
                    '---------------------------------------------------------------
                    If InStr(",1,3,", Val(.Cell(flexcpData, lngRow, COL_�������))) > 0 Then '��ҩ;������ҩ�÷�
                        cur��� = 0
                        lngTmp = .FindRow(CStr(rsSend!ID), , COL_���ID)
                        
                        If .Cell(flexcpData, lngRow, COL_�������) = 1 Then '��ҩ;��
                            'һ����ҩʱ,��ҩ;���Ľ���ۼ���ʾ�ڵ�һ����ҩ��
                            .TextMatrix(lngTmp, COL_���) = Format(Val(.TextMatrix(lngTmp, COL_���)) + Val(.TextMatrix(lngRow, COL_���)), gstrDec)
                            
                            '��ʾ��ҩ;��,ִ������
                            For j = lngTmp To lngRow - 1
                                strTmp = ""
                                If Val(.TextMatrix(j, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                                    strTmp = "�Ա�ҩ"
                                ElseIf Val(.TextMatrix(j, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                                    strTmp = "��Ժ��ҩ"
                                End If
                                .TextMatrix(j, COL_ִ������) = strTmp
                                .TextMatrix(j, COL_�÷�) = rsSend!������Ŀ
                            Next
                        Else
                            'ҩƷ��ִ������
                            strTmp = ""
                            If Val(.TextMatrix(lngTmp, COL_ִ������ID)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) <> 5 Then
                                strTmp = "�Ա�ҩ"
                            ElseIf Val(.TextMatrix(lngTmp, COL_ִ������ID)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������ID)) = 5 Then
                                strTmp = "��Ժ��ҩ"
                            End If
                            
                            '��ҩ�÷�,�巨
                            str�÷� = rsSend!������Ŀ
                            If Val(.Cell(flexcpData, lngRow - 1, COL_�������)) = 2 Then
                                str�÷� = str�÷� & "|" & GetItemField("������ĿĿ¼", Val(.TextMatrix(lngRow - 1, COL_������ĿID)), "����")
                            End If
                            For j = lngTmp To lngRow
                                .TextMatrix(j, COL_�÷�) = str�÷� '������д�շ���¼
                                cur��� = cur��� + Val(.TextMatrix(j, COL_���))
                            Next
                            .TextMatrix(lngRow, COL_���) = Format(cur���, gstrDec)
                            '��ʾִ������
                            .TextMatrix(lngRow, COL_ִ������) = strTmp
                            '��ʾ�䷽ִ�п���
                            .TextMatrix(lngRow, COL_ִ�п���) = .TextMatrix(lngTmp, COL_ִ�п���)
                        End If
                        
                        'ʹ���ҽ��ѡ��״̬��ͬ(��Ϊ����ԭ��)
                        For j = lngTmp To lngRow
                            If .Cell(flexcpData, j, COL_ѡ��) <> 0 Then
                                Call RowSelectSame(j, COL_ѡ��)
                                Exit For 'һ����ֹ,ȫ����ֹ
                            End If
                        Next
                        If j > lngRow Then
                            For j = lngRow To lngTmp Step -1
                                If InStr(",5,6,7,", .TextMatrix(j, COL_�������)) > 0 Then
                                    If .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                                        Call RowSelectSame(j, COL_ѡ��)
                                        Exit For '���ѡ,ȫ����ѡ
                                    End If
                                End If
                            Next
                        End If
                    End If
                    
                    'ҩƷ�����:�Ա�ҩ�����
                    '---------------------------------------------------------------
                    If InStr(",5,6,7,", rsSend!�������) > 0 And Nvl(rsSend!ִ������, 0) <> 5 Then
                        Call CheckStock(lngRow, bln�����ʾ, blnʱ����ʾ, blnĬ�Ϸ���)
                    End If
                    
                    '��������
                    '---------------------------------------------------------------
                    '���˼������ָ�
                    If rsSend!����ID <> lng����ID Then
                        lng������ = lng������ + 1
                        If lng����ID <> 0 Then
                            For j = lngRow - 1 To .FixedRows Step -1
                                If Not .RowHidden(j) Then
                                    .CellBorderRange j, .FixedCols, j, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    lng����ID = rsSend!����ID

NextLoop:           '---------------------------------------------------------------
                    Progress = i / rsSend.RecordCount * 100
                    rsSend.MoveNext
                Next
            End With
        End If
    Next
    
    lblInfo.Caption = lblInfo.Caption & "������" & IIF(str���� = "", " ", "(" & Mid(str����, 2) & ") ") & lng������ & " �����˵�ҽ��"
    With vsAdvice
        If Not .ColHidden(COL_���) Then
            .AutoSize COL_ҽ������, COL_���
        Else
            .AutoSize COL_ҽ������
        End If
        .RowHeight(0) = 320
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    vsAdvice.ColHidden(COL_����) = True
    vsAdvice.ColHidden(COL_Ӥ��) = True
    vsAdvice.ColHidden(COL_���) = Not blnƷ��ҩƷ Or int��Ч = 1
    vsAdvice.ColHidden(COL_�״�ʱ��) = int��Ч = 1
    vsAdvice.ColHidden(COL_ĩ��ʱ��) = int��Ч = 1
    vsAdvice.SetFocus: Call vsAdvice_GotFocus
    Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    
    LoadAdviceSend = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone
        Resume
    End If
    Call SaveErrLog
    Progress = 0
End Function

Private Sub InitBillSet()
'���ܣ���ʼ��ҽ�����ʵ������ɼ�¼��
    Set mrsBill = New ADODB.Recordset
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 8
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.Fields.Append "�������", adBigInt
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
End Sub

Private Sub InitRecordSet(rsSQL As ADODB.Recordset, rsTotal As ADODB.Recordset, rsUpload As ADODB.Recordset)
'��ʼ����¼��
    'SQL��¼��
    Set rsSQL = New ADODB.Recordset
    rsSQL.Fields.Append "����", adInteger '1-���ü�¼,2-ҽ����¼,3-���ͼ�¼,4-���ϼ�¼
    rsSQL.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsSQL.Fields.Append "��ĿID", adBigInt '�շ�ϸĿID
    rsSQL.Fields.Append "���", adBigInt '��������
    rsSQL.Fields.Append "SQL", adVarChar, 5000 'SQL
    rsSQL.CursorLocation = adUseClient
    rsSQL.LockType = adLockOptimistic
    rsSQL.CursorType = adOpenStatic
    rsSQL.Open
    
    '�Ƽ������ۼƼ�¼��
    Set rsTotal = New ADODB.Recordset
    rsTotal.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsTotal.Fields.Append "��ĿID", adBigInt
    rsTotal.Fields.Append "�ⷿID", adBigInt
    rsTotal.Fields.Append "����", adDouble
    rsTotal.CursorLocation = adUseClient
    rsTotal.LockType = adLockOptimistic
    rsTotal.CursorType = adOpenStatic
    rsTotal.Open
    
    'ҽ���ϴ����ʵ�
    Set rsUpload = New ADODB.Recordset
    rsUpload.Fields.Append "ҽ��ID", adBigInt 'һ��ҽ����ID
    rsUpload.Fields.Append "NO", adVarChar, 10
    rsUpload.CursorLocation = adUseClient
    rsUpload.LockType = adLockOptimistic
    rsUpload.CursorType = adOpenStatic
    rsUpload.Open
End Sub

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String, lng������� As Long, lng������� As Long)
'���ܣ���ȡ��ǰ���ʵ��ݵ�NO�����
'������lng�������=���ü�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'      lng�������=���ͼ�¼�е����,Ϊ-1ʱ��ʾ��ȡ�������
'˵����strKey=���ݼ��ʵ������ɹ��򶨵�Ψһ�ؼ���
'1.������ҩ��"����(����ID,��ҳID)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
'2.һ���䷽�е����в�ҩ����һ���������ݺ�
'3.����ҽ�����ҩ�ֺŹ�����ͬ��
'4.������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷�)
'5.��鲿λ�͸�������,������������Ҫҽ��������ͬ���ݺš�
'6.һ���ɼ��ļ�����Ϸ�����ͬ�ĵ��ݺţ��걾�ɼ��������䵥���ĵ��ݺ�
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        mrsBill!NO = zlDatabase.GetNextNO(14)
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill!������� = IIF(lng������� = -1, 0, 1)
        mrsBill.Update
    Else
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        If lng������� <> -1 Then
            mrsBill!������� = mrsBill!������� + 1
        End If
        mrsBill.Update
    End If
    strNO = mrsBill!NO
    If lng������� <> -1 Then lng������� = mrsBill!�������
    If lng������� <> -1 Then lng������� = mrsBill!�������
End Sub

Private Function CompletePatiSend(rsPati As ADODB.Recordset, rsSQL As ADODB.Recordset, _
    rsUpload As ADODB.Recordset, ByVal cur�ϼ� As Currency, ByVal str��� As String, ByVal str������� As String, _
    strWarn As String, intWarn As Integer, blnTran As Boolean) As Boolean
'���ܣ��ύһ�����˵�ҽ����������,����֮ǰ������ʱ���
'������rsPati=����������Ϣ�ļ�¼��,���ڼ��ʱ���
'      rsSQL=��������Ҫִ�е�SQL
'      rsUpload=����ҽ���ϴ��ļ��ʵ��ݺ�
'      cur�ϼ�=���˱���Ҫ����ҽ���ļ��ʽ��ϼ�,���ڼ��ʱ���
'      str���=���˱��η��ͼ��ʷ��õ��շ����,���ڼ��ʱ���
'      str���=���˱��η��ͼ��ʷ��õ��շ��������,���ڼ��ʱ���
'      strWarn(I/O)=���ڼ�¼��ǰ�����ѱ������
'      intWarn(I/O)=���ڼ�¼���η��ͱ�����ʾʱ��ѡ����
'˵�����������,���ڵ��ú����д���,blnTran�����Ƿ�����������
    Dim rsWarn As New ADODB.Recordset
    Dim strSQL As String, intR As Integer
    Dim cur���� As Currency, i As Long
    Dim strMsg As String
    
    '���˷��ñ���
    If cur�ϼ� > 0 Then
        strSQL = "Select Nvl(���ò���,1) as ���ò���,Nvl(��������,1) as ��������," & _
            " ����ֵ,������־1,������־2,������־3 From ���ʱ�����" & _
            " Where ����ID=[1] And Nvl(���ò���,1)=[2]"
        Set rsWarn = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsPati!��ǰ����ID), IIF(Nvl(rsPati!ҽ��, 0) = 1, 2, 1))
        If Not rsWarn.EOF Then
            If rsWarn!�������� = 2 Then cur���� = GetPatiDayMoney(rsPati!����ID)
            str������� = Mid(str�������, 2)
            For i = 1 To Len(str���)
                intR = BillingWarn(Me, mstrPrivs, rsWarn, rsPati!����, Nvl(rsPati!ʣ���, 0), cur����, cur�ϼ�, Nvl(rsPati!������, 0), Mid(str���, i, 1), Split(str�������, ",")(i - 1), strWarn, intWarn, Nvl(rsPati!ҽ��, 0) = 1)
                If InStr(",2,3,", intR) > 0 Then Exit For
            Next
        End If
    End If
    
    If InStr(",2,3,", intR) = 0 Then
        'ִ��˳��:����,ҽ��,����
        '1.����д����,��Ϊ����ʱ���ܴ������
        '2.�Է��ü�¼���շ�ϸĿID�������
        rsSQL.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst
        rsSQL.Sort = "����,��ĿID,���"
        rsUpload.Filter = 0 '�ϲ㺯������ʹ�ù�,��ʹû�ù�ҲMoveFirst
        
        gcnOracle.BeginTrans: blnTran = True
        Do While Not rsSQL.EOF
            Call zlDatabase.ExecuteProcedure(rsSQL!SQL, Me.Caption)
            rsSQL.MoveNext
        Loop
        
        'ҽ�������ϴ�
        If Not IsNull(rsPati!����) Then
            If gclsInsure.GetCapability(supportҽ���ϴ�, , rsPati!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , rsPati!����) Then
                Do While Not rsUpload.EOF
                    strMsg = "" '��Ϊ����һ��NO�ڿ϶�Ϊһ�����˵�,��������˲������Բ���
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , rsPati!����) Then
                        'δ�ύǰ�ϴ�ʧ����ع�����ֹ����
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName 'ÿ����ʾ
                        Else
                            MsgBox rsPati!���� & "�ķ����ϴ�ʧ�ܣ����Ͳ���������ֹ��", vbExclamation, gstrSysName
                        End If
                        Exit Function
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName 'ÿ����ʾ
                    End If
                    rsUpload.MoveNext
                Loop
            End If
        End If
        gcnOracle.CommitTrans: blnTran = False
        
        'ҽ�������ϴ�
        If Not IsNull(rsPati!����) Then
            If gclsInsure.GetCapability(supportҽ���ϴ�, , rsPati!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, , rsPati!����) Then
                Do While Not rsUpload.EOF
                    strMsg = ""
                    If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 1, strMsg, , rsPati!����) Then
                        '�ύ���ϴ�ʧ��,����ʾ
                        If strMsg <> "" Then
                            MsgBox strMsg, vbInformation, gstrSysName
                        Else
                            MsgBox rsPati!���� & "�ļ��ʵ�""" & rsUpload!NO & """�ϴ�ʧ�ܣ�HIS���������ύ����ȷ���������͡�", vbExclamation, gstrSysName
                        End If
                    Else
                        If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                    End If
                    rsUpload.MoveNext
                Loop
            End If
        End If
            
        '�ύ�ɹ�,������ҽ���б��Ϊ��ɾ��
        With vsAdvice
            i = .FindRow(CStr(rsPati!����ID), , COL_����ID)
            For i = i To .Rows - 1
                If Val(.TextMatrix(i, COL_����ID)) = rsPati!����ID Then
                    If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                        .RowData(i) = -1
                    End If
                Else
                    Exit For
                End If
            Next
        End With
    End If
    
    CompletePatiSend = True
End Function

Private Sub DeleteSendRow()
'���ܣ���������ҽ���嵥���ѷ��ͳɹ��ĵ���ɾ��
    Dim i As Long, blnDel As Boolean
    
    With vsAdvice
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If .RowData(i) = -1 Then .RemoveItem i: blnDel = True
        Next
        .Redraw = flexRDDirect
        
        If blnDel Then
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
            For i = .FixedRows To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: .Col = COL_ѡ��
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            
            vsPrice.Rows = vsPrice.FixedRows
            vsPrice.Rows = vsPrice.FixedRows + 1
            Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        End If
    End With
End Sub

Private Function Getʵ�ս��(ByVal strSQL As String) As Currency
    Dim lngPos As Long, strMatch As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strSQL = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    strMatch = "End" & Chr(0) & Chr(1)
    strSQL = Left(strSQL, InStr(strSQL, strMatch) - 1)
    Getʵ�ս�� = CCur(strSQL)
End Function

Private Function Setʵ�ս��(ByVal strSQL As String, ByVal cur��� As Currency) As String
    Dim strLeft As String, strRight As String
    Dim strMatch As String, strVal As String
    
    strMatch = Chr(0) & Chr(1) & "Begin"
    strLeft = Mid(strSQL, 1, InStr(strSQL, strMatch) - 1)
    strMatch = "End" & Chr(0) & Chr(1)
    strRight = Mid(strSQL, InStr(strSQL, strMatch) + Len(strMatch))
    
    Setʵ�ս�� = strLeft & cur��� & strRight
End Function

Public Function SendAdvice() As Long
'���ܣ�����ҽ������(��������м��ʱ���)
'˵����������˷����ύ
'���أ�����ɹ����򷵻ط��ͺ�
    Dim rsPati As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    Dim rsSQL As ADODB.Recordset
    Dim rsTotal As ADODB.Recordset
    Dim rsUpload As ADODB.Recordset
    
    Dim i As Long, j As Long
    Dim strSQL As String, strTmp As String
    Dim curDate As Date, blnTran As Boolean
    Dim strWarn As String, intWarn As Integer, str��� As String, str������� As String
    
    Dim lng����ID As Long, lng���ͺ� As Long, int�Ʒ�״̬ As Integer, int���� As Integer, strNO As String
    Dim lngϸĿID As Long, lng������� As Long, lng���ø��� As Long, lng������� As Long
    Dim int���� As Integer, dbl���� As Double, cur�ϼ� As Currency
    Dim dbl���� As Double, curӦ�� As Currency, curʵ�� As Currency
    Dim bln������Ŀ�� As Boolean, lng���մ���ID As Long, curͳ���� As Currency, str���ձ��� As String, str�������� As String
    Dim str�ֽ�ʱ�� As String, str�״�ʱ�� As String, strĩ��ʱ�� As String
    Dim int�䷽�� As Integer, strNOKey As String, str�Զ����� As String
    Dim str����ʱ�� As String, str�Ǽ�ʱ�� As String
    Dim dbl�������� As Double, blnFirst As Boolean '�䷽�����ֺŹؼ���
    Dim lng���˿���ID As Long, lngִ�п���ID As Long, intִ��״̬ As Integer
    Dim bln��Ժ��ҩ As Boolean, blnBool As Boolean
    
    Dim blnHaveSub As Boolean, curҽ���ϼ� As Currency
    Dim int����� As Integer, var������ As Variant
    Dim lng������ID As Long, strʵ�� As String
    
    Dim blnҩƷʱ����ʾ As Boolean, blnҩƷ�����ʾ As Boolean, blnҩƷĬ�Ϸ��� As Boolean
    Dim bln����ʱ����ʾ As Boolean, bln���Ŀ����ʾ As Boolean, bln����Ĭ�Ϸ��� As Boolean
    
    Dim rsAudit As ADODB.Recordset
    Dim strAudit As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
        
    blnҩƷʱ����ʾ = True: blnҩƷ�����ʾ = True: blnҩƷĬ�Ϸ��� = True
    bln����ʱ����ʾ = True: bln���Ŀ����ʾ = True: bln����Ĭ�Ϸ��� = True
    
    Call InitBillSet
    lng���ͺ� = zlDatabase.GetNextNO(10)
    curDate = zlDatabase.Currentdate
    intWarn = -1 '���ʱ���ʱȱʡҪ��ʾ,�벡���޹�
    int�䷽�� = 1 '��ʾ���͵ĵڼ����䷽,���ڷֵ��ݺ�
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                '�ύ��ǰ���˵�����
                '-----------------------------------------------------------------------------------------
                If Val(.TextMatrix(i, COL_����ID)) <> lng����ID Then
                    '�ύ��ǰ��������
                    If lng����ID <> 0 Then
                        If strAudit <> "" Then
                            MsgBox "����""" & rsPati!���� & """���·�����Ŀ��û�о�����������Ӧ��ҽ�����ܷ��ͣ�" & vbCrLf & strAudit, vbInformation, gstrSysName
                            GoTo errH
                        End If
                    
                        If Not CompletePatiSend(rsPati, rsSQL, rsUpload, cur�ϼ�, str���, str�������, strWarn, intWarn, blnTran) Then GoTo errH
                        SendAdvice = lng���ͺ� 'ֻҪ�ύ�ɹ����ע
                    End If
                    
                    '���ò�����ر���
                    str�Զ����� = ""
                    lng����ID = Val(.TextMatrix(i, COL_����ID))
                    Call InitRecordSet(rsSQL, rsTotal, rsUpload) '����SQL����
                    cur�ϼ� = 0:  str��� = "": str������� = "": strWarn = "" '���ñ�������
                    
                    '��ȡ��ǰ������Ϣ
                    strSQL = _
                        " Select ����ID,Ԥ�����,�������,0 as Ԥ����� From ������� Where ����=1 And ����ID=[1]" & _
                        " Union ALL" & _
                        " Select A.����ID,0,0,Sum(���) From ����ģ����� A,������ҳ B" & _
                        " Where A.����ID=B.����ID And A.��ҳID=B.��ҳID And B.���� Is Not Null And A.����ID=[1] And A.��ҳID=[2] Group by A.����ID"
                    strSQL = "Select ����ID,Nvl(Sum(Ԥ�����),0)-Nvl(Sum(�������),0)+Nvl(Sum(Ԥ�����),0) as ʣ��� From (" & strSQL & ") Group by ����ID"
                    
                    strSQL = "Select A.����ID,B.��ҳID,A.����,B.����,B.��ǰ����ID,B.��Ժ����ID," & _
                        " D.���� as ������,Decode(D.����,'1',1,Decode(Nvl(B.����,0),0,0,1)) as ҽ��,C.ʣ���," & _
                        " Decode(A.������,Null,Null,zl_PatientSurety(A.����ID,B.��ҳID)) as ������" & _
                        " From ������Ϣ A,������ҳ B,(" & strSQL & ") C,ҽ�Ƹ��ʽ D" & _
                        " Where A.����ID=B.����ID And A.����ID=C.����ID(+) And B.ҽ�Ƹ��ʽ=D.����(+)" & _
                        " And A.����ID=[1] And B.��ҳID=[2]"
                    Set rsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, Val(.TextMatrix(i, COL_��ҳID)))
                    
                    '��ȡ��ǰ���˵�������Ŀ�嵥
                    strAudit = ""
                    If Not IsNull(rsPati!����) Then
                        Set rsAudit = GetAuditRecord(lng����ID, Val(.TextMatrix(i, COL_��ҳID)))
                    Else
                        Set rsAudit = Nothing '��NothingΪ��־�ò��˲���Ҫ�ж�
                    End If
                    
                    '�����²���鵱ǰ����ҽ����ҩƷ���,�Ա�ҩ�����
                    '��Ȼ��ȡʱ�ѻ��ܼ�飬����Ʒ����ʱ������˹����ܷ����仯
                    For j = i To .Rows - 1
                        If Val(.TextMatrix(j, COL_����ID)) = lng����ID Then
                            '���ܸ���ǰ��������ʾ�Ľ�������Ѳ�����
                            If .Cell(flexcpData, j, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, j, COL_ѡ��) Is Nothing Then
                                If InStr(",5,6,7,", .TextMatrix(j, COL_�������)) > 0 And Val(.TextMatrix(j, COL_ִ������ID)) <> 5 Then
                                    '�ڲ����ֹ�������,����������ʱ��ҩƷ
                                    If GetStockCheck(Val(.TextMatrix(j, COL_ִ�п���ID))) = 2 _
                                        Or Val(.TextMatrix(j, COL_ҩ������)) = 1 Or Val(.TextMatrix(j, COL_�Ƿ���)) = 1 Then
                                        .TextMatrix(j, COL_���) = Format(GetStock(Val(.TextMatrix(j, COL_ҩƷID)), Val(.TextMatrix(j, COL_ִ�п���ID)), 2), "0.00000")
                                        If CheckStock(j, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���, True) Then
                                            Call RowSelectSame(j, COL_ѡ��)
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                                    
                '���ܸ���ǰ��������ʾ�Ľ�������Ѳ�����
                If .Cell(flexcpData, i, COL_ѡ��) = 0 And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                    '�������ݺŷ���ؼ���
                    '-----------------------------------------------------------------------------------------
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        '������ҩ��"����(����ID,��ҳID)_���˿���ID_��������ID_����ҽ��_ִ�п���ID"�ֺš�
                        strNOKey = "������ҩ_" & lng����ID & "_" & Val(.TextMatrix(i, COL_��ҳID)) & "_" & _
                            Val(.TextMatrix(i, COL_���˿���ID)) & "_" & Val(.TextMatrix(i, COL_��������ID)) & "_" & _
                            .TextMatrix(i, COL_����ҽ��) & "_" & Val(.TextMatrix(i, COL_ִ�п���ID))
                        '�ٰ�Ҫ��ӡ�����Ƶ��ݷֺ�
                        strNOKey = strNOKey & "_" & GetClinicBillID(Val(.TextMatrix(i, COL_������ĿID)), 2)
                    ElseIf .TextMatrix(i, COL_�������) = "7" Then
                        'һ���䷽�е����в�ҩ����һ���������ݺ�
                        strNOKey = "��ҩ�䷽_" & lng����ID & "_" & Val(.TextMatrix(i, COL_��ҳID)) & "_" & int�䷽��
                    Else
                        '������ҩҽ��ÿ��ҽ��һ���������ݺ�(������ҩ;�����䷽�巨���÷�)
                        strNOKey = "��ҩҽ��_" & Val(.TextMatrix(i, COL_ID))
                    End If
                    
                    '�Ƿ���Ժ��ҩ
                    bln��Ժ��ҩ = False
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        If .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ" Then bln��Ժ��ҩ = True
                    ElseIf .TextMatrix(i, COL_�������) = "7" Then
                        j = .FindRow(CStr(.TextMatrix(i, COL_���ID)), i + 1, COL_ID)
                        If j <> -1 Then
                            If .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ" Then bln��Ժ��ҩ = True
                        End If
                    End If
                    
                    '����ҽ�����ʷ���:�����¼۸����
                    '-----------------------------------------------------------------------------------------
                    strSQL = "": lngϸĿID = 0
                    If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                        'ҩƷȱʡ�̶�Ϊ�����Ƽ�,����ҽ��ʱָ����Ϊ�Ա�ҩ(Ժ��ִ��)�Ĳ���ȡ;ҩƷ������Ϊ����
                        If Val(.TextMatrix(i, COL_ִ������ID)) <> 5 Then
                            strSQL = _
                                " Select A.ID,A.���,D.���� as �������,A.����,A.���㵥λ,B.������ĿID," & _
                                " C.�վݷ�Ŀ,Y.סԺ��λ,Y.סԺ��װ,Y.����ϵ��,1 as ����,B.�ּ� as ����," & _
                                " A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,A.�Ƿ���,Y.ҩ������ as ����,0 as ��������," & _
                                " 0 as ����,[3] as ִ�п���ID,A.���ηѱ�,I.Ҫ������" & _
                                " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,ҩƷ��� Y,����֧����Ŀ I" & _
                                " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.���=D.����" & _
                                " And A.ID=Y.ҩƷID(+) And A.ID=[1] And A.ID=I.�շ�ϸĿID(+) And I.����(+)=[4]" & _
                                " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                                " Order by A.����"
                        End If
                    ElseIf Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 And InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������ID))) = 0 Then
                        '���Ƽ�,�ֹ��Ƽۣ�����,Ժ��ִ�е�ҽ������ȡ
                        strSQL = _
                            " Select A.ID,A.���,D.���� as �������,RTrim(A.����||' '||A.���) as ����," & _
                            " A.���㵥λ,B.������ĿID,C.�վݷ�Ŀ,Y.סԺ��λ,Y.סԺ��װ,Y.����ϵ��,X.����," & _
                            " Decode(A.�Ƿ���,1,X.����,B.�ּ�) as ����,A.�Ӱ�Ӽ�,B.�Ӱ�Ӽ���,A.�Ƿ���," & _
                            " Decode(A.���,'4',E.���÷���,Y.ҩ������) as ����,E.��������,Nvl(X.����,0) as ����," & _
                            " Nvl(X.ִ�п���ID,[3]) as ִ�п���ID,A.���ηѱ�,I.Ҫ������" & _
                            " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ���Ŀ��� D,�������� E,����ҽ���Ƽ� X,ҩƷ��� Y,����֧����Ŀ I" & _
                            " Where A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID And A.���=D.���� And A.ID=E.����ID(+)" & _
                            " And A.ID=Y.ҩƷID(+) And X.�շ�ϸĿID=A.ID And Nvl(X.����,0)<>0 And X.ҽ��ID=[2]" & _
                            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                            " And A.ID=I.�շ�ϸĿID(+) And I.����(+)=[4]" & _
                            " Order by ����,A.ID"
                            'һ��Ҫ����������ǰ��,�Ա��ڼ�����ڷ��ü�¼�б������ӹ�ϵ
                    End If
                    
                    '�����ۿ۱�����ʼ
                    blnHaveSub = False
                    var������ = Empty: int����� = 0
                    curҽ���ϼ� = 0: lng������ID = 0
                    
                    int�Ʒ�״̬ = IIF(Val(.TextMatrix(i, COL_�Ƽ�����)) = 1, -1, 0) '����Ʒѻ�δ�Ʒ�
                    If strSQL <> "" Then
                        Set rsPrice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(i, COL_ҩƷID)), Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_ִ�п���ID)), Val(Nvl(rsPati!����, 0)))
                        If Not rsPrice.EOF Then
                            int�Ʒ�״̬ = 1 '�ѼƷ�
                            
                            'ȷ���Ƿ����ӹ�ϵ:��ʹ�������ۿ�,ҲҪ��¼
                            rsPrice.Filter = "����=1"
                            If Not rsPrice.EOF Then blnHaveSub = True
                            rsPrice.Filter = 0
                        End If
                        
                        '����������Ŀ���ķ�����ϸ
                        For j = 1 To rsPrice.RecordCount
                            '����Ƿ���Ҫ���Ѿ�����
                            If Nvl(rsPrice!Ҫ������, 0) = 1 And Not rsAudit Is Nothing Then
                                rsAudit.Filter = "��ĿID=" & rsPrice!ID
                                If rsAudit.EOF Then
                                    If UBound(Split(strAudit, vbCrLf)) < 10 Then
                                        If InStr(strAudit, "��" & rsPrice!����) = 0 Then
                                            strAudit = strAudit & vbCrLf & "��" & rsPrice!����
                                        End If
                                    ElseIf UBound(Split(strAudit, vbCrLf)) = 10 Then
                                        strAudit = strAudit & vbCrLf & "�� ��"
                                    End If
                                End If
                            End If
                        
                            'ִ�п���ID
                            lngִ�п���ID = Nvl(rsPrice!ִ�п���ID, 0)
                            '��ԭֵ������ȡ��Ч�ķ�ҩ��ҩƷ���������ĵ�ִ�п���
                            If rsPrice!��� = "4" And Nvl(rsPrice!��������, 0) = 1 _
                                Or InStr(",5,6,7", rsPrice!���) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_�������)) = 0 Then
                                lng���˿���ID = Val(.TextMatrix(i, COL_���˿���ID))
                                lngִ�п���ID = Get�շ�ִ�п���ID(rsPati!����ID, rsPati!��ҳID, rsPrice!���, rsPrice!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID)
                                
                                '���ı�������ִ�п���
                                If lngִ�п���ID = 0 And rsPrice!��� = "4" Then
                                    .Row = GetVisibleRow(i, True)
                                    Call .ShowCell(.Row, .Col)
                                    Screen.MousePointer = 0
                                    MsgBox "ϵͳ����Ϊ�Ƽ�����""" & rsPrice!���� & """ȷ�����ʵ�ִ�п��ҡ�" & vbCrLf & _
                                        "��ʹ�üƼ۵���������Ϊȷ������""����Ŀ¼����""�м��洢�ⷿ�����Ƿ���ȷ��", vbInformation, gstrSysName
                                    Call DeleteSendRow: Call ShowSendTotal
                                    Progress = 0: Exit Function
                                End If
                            End If
                            
                            If InStr(",5,6,7", rsPrice!���) > 0 Then
                                If InStr(",5,6,7", .TextMatrix(i, COL_�������)) > 0 Then
                                    If .TextMatrix(i, COL_�������) = "7" Then
                                        int���� = Val(.TextMatrix(i, COL_����))
                                        '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                        If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                                            dbl���� = Val(.TextMatrix(i, COL_����)) / Nvl(rsPrice!����ϵ��, 1)
                                        Else
                                            dbl���� = IntEx(Val(.TextMatrix(i, COL_����)) / Nvl(rsPrice!����ϵ��, 1) / Nvl(rsPrice!סԺ��װ, 1)) * Nvl(rsPrice!סԺ��װ, 1)
                                        End If
                                    Else
                                        int���� = 1
                                        dbl���� = Val(.TextMatrix(i, COL_����)) * Nvl(rsPrice!סԺ��װ, 1)
                                    End If
                                Else
                                    int���� = 1
                                    '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                                    '��ҩ��ҩƷ�Ƽ�:��Ϊ����Ԥ�����ۼ�����,��˲��������㴦��
                                    dbl���� = Val(.TextMatrix(i, COL_����)) * Nvl(rsPrice!����, 0)
                                End If
                                dbl���� = Format(dbl����, "0.00000")
                                
                                If Nvl(rsPrice!�Ƿ���, 0) = 1 Then
                                    dbl���� = Format(CalcDrugPrice(rsPrice!ID, lngִ�п���ID, int���� * dbl����, , True), "0.00000")
                                Else
                                    dbl���� = Format(Nvl(rsPrice!����, 0), "0.00000")
                                End If
                            ElseIf rsPrice!��� = "4" And Nvl(rsPrice!��������, 0) = 1 Then
                                '�����������������
                                If mlng�������ID = 0 Then
                                    Screen.MousePointer = 0
                                    MsgBox "����ȷ���������ϵ��ݵ�������,���ȵ���������������ã�", vbInformation, gstrSysName
                                    Call DeleteSendRow: Call ShowSendTotal
                                    Progress = 0: Exit Function
                                End If
                                
                                int���� = 1
                                dbl���� = Format(Val(.TextMatrix(i, COL_����)) * Nvl(rsPrice!����, 0), "0.00000")
                                
                                '����ʱ�����ĵ���
                                If Nvl(rsPrice!�Ƿ���, 0) = 1 Then
                                    dbl���� = Format(CalcDrugPrice(rsPrice!ID, lngִ�п���ID, dbl����, , True), "0.00000")
                                Else
                                    dbl���� = Format(Nvl(rsPrice!����, 0), "0.00000")
                                End If
                            Else
                                int���� = 1
                                dbl���� = Format(Val(.TextMatrix(i, COL_����)) * Nvl(rsPrice!����, 0), "0.00000")
                                dbl���� = Format(Nvl(rsPrice!����, 0), "0.00000")
                            End If
                            
                            '��ҩ��ҩƷ���������ĵĿ����
                            If rsPrice!��� = "4" And Nvl(rsPrice!��������, 0) = 1 _
                                Or InStr(",5,6,7", rsPrice!���) > 0 And InStr(",5,6,7", .TextMatrix(i, COL_�������)) = 0 Then
                                If GetStockCheck(lngִ�п���ID) <> 0 Or Nvl(rsPrice!�Ƿ���, 0) = 1 Or Nvl(rsPrice!����, 0) = 1 Then
                                    If rsPrice!��� = "4" Then
                                        blnBool = CheckPriceStock(i, rsPrice, lngִ�п���ID, int���� * dbl����, rsTotal, bln���Ŀ����ʾ, bln����ʱ����ʾ, bln����Ĭ�Ϸ���)
                                    Else
                                        blnBool = CheckPriceStock(i, rsPrice, lngִ�п���ID, int���� * dbl����, rsTotal, blnҩƷ�����ʾ, blnҩƷʱ����ʾ, blnҩƷĬ�Ϸ���)
                                    End If
                                    If blnBool Then
                                        Call RowSelectSame(i, COL_ѡ��, rsSQL, rsTotal, rsUpload)
                                        GoTo NextAdvice
                                    End If
                                End If
                            End If
                            
                            '���ͽ��
                            curӦ�� = int���� * dbl���� * dbl����
                            
                            '����Ӱ�Ӽ�
                            If gbln�Ӱ�Ӽ� And Nvl(rsPrice!�Ӱ�Ӽ�, 0) = 1 Then
                                curӦ�� = Format(curӦ�� * (1 + Nvl(rsPrice!�Ӱ�Ӽ���, 0) / 100), gstrDec)
                            Else
                                curӦ�� = Format(curӦ��, gstrDec)
                            End If
                            
                            '��������ۿۺϼ�
                            If gbln��������ۿ� And blnHaveSub Then
                                curʵ�� = curӦ��
                                curҽ���ϼ� = curҽ���ϼ� + curʵ��
                            ElseIf Nvl(rsPrice!���ηѱ�, 0) = 0 Then
                                curʵ�� = Format(ActualMoney(.TextMatrix(i, COL_�ѱ�), rsPrice!������ĿID, curӦ��, rsPrice!ID, lngִ�п���ID, _
                                    int���� * dbl����, IIF(gbln�Ӱ�Ӽ� And Nvl(rsPrice!�Ӱ�Ӽ�, 0) = 1, Nvl(rsPrice!�Ӱ�Ӽ���, 0) / 100, 0)), gstrDec)
                            Else
                                curʵ�� = curӦ��
                            End If
                                
                            'ҽ������ֶ�
                            bln������Ŀ�� = False: lng���մ���ID = 0: curͳ���� = 0: str���ձ��� = "": str�������� = ""
                            If Not IsNull(rsPati!����) Then
                                strTmp = gclsInsure.GetItemInsure(lng����ID, rsPrice!ID, curʵ��, False, rsPati!����)
                                If strTmp <> "" Then
                                    bln������Ŀ�� = Val(Split(strTmp, ";")(0)) <> 0
                                    lng���մ���ID = Val(Split(strTmp, ";")(1))
                                    curͳ���� = Format(Val(Split(strTmp, ";")(2)), gstrDec)
                                    str���ձ��� = CStr(Split(strTmp, ";")(3))
                                    If UBound(Split(strTmp, ";")) >= 5 Then
                                        If Split(strTmp, ";")(5) <> "" Then
                                            str�������� = Split(strTmp, ";")(5)
                                        End If
                                    End If
                                End If
                            End If
                            
                            '�ռ����ʱ������
                            cur�ϼ� = cur�ϼ� + curʵ��
                            If InStr(str���, rsPrice!���) = 0 Then
                                str��� = str��� & rsPrice!���
                                str������� = str������� & "," & rsPrice!�������
                            End If
                                                        
                            'NO,���
                            Call GetCurBillSet(strNOKey, strNO, lng�������, -1)
                            rsSQL.AddNew: blnBool = False
                            If rsPrice!ID <> lngϸĿID Then
                                lng���ø��� = lng�������
                                '���ӹ�ϵʱ����¼������Ϣ
                                If rsPrice!���� = 0 And blnHaveSub Then
                                    int����� = lng�������
                                    lng������ID = rsPrice!������ĿID
                                    var������ = rsSQL.Bookmark
                                    blnBool = True
                                End If
                            End If
                            lngϸĿID = rsPrice!ID
                            
                            '�����ۿ�ʱ���������ʵ�ս�������⴦��
                            If gbln��������ۿ� And blnHaveSub And blnBool Then
                                strʵ�� = Chr(0) & Chr(1) & "Begin" & curʵ�� & "End" & Chr(0) & Chr(1)
                            Else
                                strʵ�� = curʵ��
                            End If
                            
                            '�Ƿ񻮼�
                            If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                                int���� = IIF(InStr(gstr���ͻ��۵�, "5") > 0, 1, 0)
                            Else
                                int���� = IIF(InStr(gstr���ͻ��۵�, .TextMatrix(i, COL_�������)) > 0, 1, 0)
                            End If
                            
                            '����ʱ��
                            If .TextMatrix(i, COL_�ֽ�ʱ��) <> "" Then
                                str����ʱ�� = "To_Date('" & Split(.TextMatrix(i, COL_�ֽ�ʱ��), ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str����ʱ�� = "To_Date('" & .Cell(flexcpData, i, COL_�ֽ�ʱ��) & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            
                            '�Ǽ�ʱ��
                            If int���� = 1 Then '��ǻ��۵�ʱ�������ֿ�
                                str�Ǽ�ʱ�� = "To_Date('" & Format(DateAdd("s", 1, curDate), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                str�Ǽ�ʱ�� = "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            End If
                            
                            '�ռ�ҽ���ϴ����ݺ�:mrsBill�еĲ�һ�������˷���
                            If int���� = 0 Then
                                rsUpload.Filter = "NO='" & strNO & "'"
                                If rsUpload.EOF Then
                                    rsUpload.AddNew
                                    rsUpload!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                    rsUpload!NO = strNO
                                    rsUpload.Update
                                End If
                            End If
                            
                            '��Ϊ���ڲ��Ƽ۵�ҽ������������,���Դ���ļƼ����Զ�Ϊ(0-�����Ƽ�)
                            rsSQL!���� = 1
                            rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                            rsSQL!��ĿID = rsPrice!ID
                            rsSQL!��� = i
                            rsSQL!SQL = "ZL_סԺ���ʼ�¼_Insert(" & _
                                "'" & strNO & "'," & lng������� & "," & lng����ID & "," & ZVal(.TextMatrix(i, COL_��ҳID)) & "," & _
                                ZVal(.TextMatrix(i, COL_סԺ��)) & ",'" & .TextMatrix(i, COL_����) & "'," & _
                                "'" & .TextMatrix(i, COL_�Ա�) & "','" & .TextMatrix(i, COL_����) & "'," & _
                                "'" & .TextMatrix(i, COL_����) & "','" & .TextMatrix(i, COL_�ѱ�) & "'," & _
                                rsPati!��ǰ����ID & "," & rsPati!��Ժ����ID & ",0," & Val(.Cell(flexcpData, i, COL_Ӥ��)) & "," & _
                                ZVal(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                                IIF(rsPrice!���� = 1, ZVal(int�����), "NULL") & "," & rsPrice!ID & "," & _
                                "'" & rsPrice!��� & "','" & Nvl(rsPrice!���㵥λ) & "'," & _
                                IIF(bln������Ŀ��, 1, 0) & "," & ZVal(lng���մ���ID) & ",'" & str���ձ��� & "'," & _
                                int���� & "," & dbl���� & ",NULL," & ZVal(lngִ�п���ID) & "," & _
                                IIF(lng���ø��� = lng�������, "NULL", lng���ø���) & "," & rsPrice!������ĿID & "," & _
                                "'" & Nvl(rsPrice!�վݷ�Ŀ) & "'," & dbl���� & "," & curӦ�� & "," & strʵ�� & "," & _
                                curͳ���� & "," & str����ʱ�� & "," & str�Ǽ�ʱ�� & "," & _
                                "'ҽ������'," & int���� & ",'" & UserInfo.��� & "','" & UserInfo.���� & "',0," & _
                                IIF(rsPrice!��� = "4", mlng�������ID, mlngҩƷ���ID) & "," & _
                                "NULL,'" & .TextMatrix(i, COL_ҽ������) & "',NULL," & Val(.TextMatrix(i, COL_ID)) & "," & _
                                "'" & .TextMatrix(i, COL_Ƶ��) & "'," & ZVal(.TextMatrix(i, COL_����)) & "," & _
                                "'" & .TextMatrix(i, COL_�÷�) & "'," & mint��Ч & "," & _
                                IIF(bln��Ժ��ҩ, 3, Val(.TextMatrix(i, COL_�Ƽ�����))) & ",Null,'" & str�������� & "')"
                            rsSQL.Update
                            
                            '��¼�Զ����ϵ�SQL
                            If gblnסԺ�Զ����� And int���� = 0 And lngִ�п���ID <> 0 And rsPrice!��� = "4" And Nvl(rsPrice!��������, 0) = 1 Then
                                If InStr(str�Զ����� & ";", ";" & strNO & "," & lngִ�п���ID & ";") = 0 Then
                                    rsSQL.AddNew
                                    rsSQL!���� = 4
                                    rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                                    rsSQL!��ĿID = 0
                                    rsSQL!��� = i
                                    rsSQL!SQL = "zl_�����շ���¼_��������(" & lngִ�п���ID & ",25,'" & strNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
                                    rsSQL.Update
                                    str�Զ����� = str�Զ����� & ";" & strNO & "," & lngִ�п���ID
                                End If
                            End If
                            
                            rsPrice.MoveNext
                        Next
                    End If
                    
                    '��ҽ�������л����ۿ۴���:������ܴ��۲�֧�ְ��ɱ����մ���
                    If gbln��������ۿ� And blnHaveSub And var������ <> Empty And lng������ID <> 0 Then
                        rsSQL.Bookmark = var������
                        curʵ�� = Format(ActualMoney(.TextMatrix(i, COL_�ѱ�), lng������ID, curҽ���ϼ�), gstrDec)
                        curʵ�� = curʵ�� - curҽ���ϼ� '���۲��
                        curʵ�� = Getʵ�ս��(rsSQL!SQL) + curʵ��
                        rsSQL!SQL = Setʵ�ս��(rsSQL!SQL, curʵ��)
                        rsSQL.Update
                    End If
                    
                    '����ҽ����ִ�п���
                    If .Cell(flexcpData, i, COL_ִ�п���ID) = 1 Then
                        rsSQL.AddNew
                        rsSQL!���� = 2
                        rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                        rsSQL!��ĿID = 0
                        rsSQL!��� = i
                        rsSQL!SQL = "ZL_ҽ��ִ�п���_Update(" & Val(.TextMatrix(i, COL_ID)) & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & ")"
                        rsSQL.Update
                    End If
                    
                    '����ҽ�����ͼ�¼
                    '-----------------------------------------------------------------------------------------
                    If Val(.TextMatrix(i, COL_ִ������ID)) <> 0 Then '����������(��ҩ;�����䷽�巨���÷�����Ϊ)
                        'һ��Ҫ��������NO
                        Call GetCurBillSet(strNOKey, strNO, -1, lng�������)
                                                                
                        '�Ƿ�һ��ҽ���ĵ�һҽ����
                        blnFirst = False
                        If InStr(",5,6,7,", .TextMatrix(i, COL_�������)) > 0 Then
                            If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                                blnFirst = True 'ҩ�Ʒ���ʱ,ֻ�е�һҩƷ�в�Ϊ��һҽ����
                            End If
                        End If
                        
                        'ִ��״̬
                        intִ��״̬ = 0
                        If mblnAutoExe And InStr(",5,6,7,", .TextMatrix(i, COL_�������)) = 0 Then
                            If Val(.TextMatrix(i, COL_��������ID)) = Val(.TextMatrix(i, COL_ִ�п���ID)) Then
                                intִ��״̬ = 1
                            End If
                        End If
                        
                        '��������:ҩƷΪ������λ������,����Ϊ����
                        If .TextMatrix(i, COL_�������) = "7" Then
                            dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����))
                        ElseIf InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                            dbl�������� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_סԺ��װ)) * Val(.TextMatrix(i, COL_����ϵ��))
                        Else
                            dbl�������� = Val(.TextMatrix(i, COL_����))
                        End If
                        dbl�������� = Format(dbl��������, "0.00000")
                                                                
                        '��ĩʱ��
                        str�ֽ�ʱ�� = .TextMatrix(i, COL_�ֽ�ʱ��)
                        If str�ֽ�ʱ�� <> "" Then
                            str�״�ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(0) & "','YYYY-MM-DD HH24:MI:SS')"
                            strĩ��ʱ�� = "To_Date('" & Split(str�ֽ�ʱ��, ",")(Val(.TextMatrix(i, COL_����)) - 1) & "','YYYY-MM-DD HH24:MI:SS')"
                        Else
                            str�״�ʱ�� = "NULL"
                            strĩ��ʱ�� = "NULL"
                        End If
                        
                        rsSQL.AddNew
                        rsSQL!���� = 3
                        rsSQL!ҽ��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, Val(.TextMatrix(i, COL_ID)), Val(.TextMatrix(i, COL_���ID)))
                        rsSQL!��ĿID = 0
                        rsSQL!��� = i
                        rsSQL!SQL = "ZL_����ҽ������_Insert(" & _
                            Val(.TextMatrix(i, COL_ID)) & "," & lng���ͺ� & ",2,'" & strNO & "'," & _
                            lng������� & "," & dbl�������� & "," & str�״�ʱ�� & "," & strĩ��ʱ�� & "," & _
                            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            intִ��״̬ & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & int�Ʒ�״̬ & "," & IIF(blnFirst, 1, 0) & ")"
                        rsSQL.Update
                    End If
                    
                    '������ҩ�䷽��
                    If .Cell(flexcpData, i, COL_�������) = 3 Then '��ҩ�÷�
                        int�䷽�� = int�䷽�� + 1
                    End If
                End If
            End If
NextAdvice:
            '----------------------------------------
            Progress = (i - .FixedRows + 1) / (.Rows - .FixedRows) * 100
        Next
        
        '�ύ���һ�����˵�����
        '-----------------------------------------------------------------------------------------
        If lng����ID <> 0 Then
            If strAudit <> "" Then
                MsgBox "����""" & rsPati!���� & """���·�����Ŀ��û�о�����������Ӧ��ҽ�����ܷ��ͣ�" & vbCrLf & strAudit, vbInformation, gstrSysName
                GoTo errH
            End If
        
            If Not CompletePatiSend(rsPati, rsSQL, rsUpload, cur�ϼ�, str���, str�������, strWarn, intWarn, blnTran) Then GoTo errH
            SendAdvice = lng���ͺ� 'ֻҪ�ύ�ɹ����ע
        End If
        
    End With
    'ɾ�������ѳɹ����͵���
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0: Screen.MousePointer = 0
    SendAdvice = lng���ͺ�
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If Err.Number <> 0 Then '��ҽ���ϴ�ʧ���˳�û�д���
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Call DeleteSendRow: Call ShowSendTotal
    Progress = 0
End Function

Private Sub ShowSendTotal()
'���ܣ����ݵ�ǰѡ��Ҫ���͵�ҽ�������㲢��ʾ���͵�ҽ���ϼ�
    Dim curTotal As Currency, i As Long
    
    With vsAdvice
        For i = 1 To .Rows - 1
            If Not .RowHidden(i) And .Cell(flexcpData, i, COL_ѡ��) = 0 _
                And Not .Cell(flexcpPicture, i, COL_ѡ��) Is Nothing Then
                curTotal = curTotal + Val(.TextMatrix(i, COL_���))
            End If
        Next
    End With
    stbThis.Panels(3).Text = "���ͷ��ã�" & Format(curTotal, gstrDec)
    Call Form_Resize
End Sub

Private Sub SetDeptInput(ByVal lngRow As Long, ByVal lngCol As Long, rsInput As ADODB.Recordset)
'���ܣ�����ִ�п�������ĵ�ֵ
    Dim i As Long
        
    With vsAdvice
        If lngCol = COL_����ִ�� Then
            '������ʾ�еĸ���ִ�п�����ʾ
            .TextMatrix(lngRow, COL_����ִ��) = rsInput!����
            .Cell(flexcpData, lngRow, COL_����ִ��) = .TextMatrix(lngRow, COL_����ִ��)
            
            '���ĸ�����Ŀ�е�ִ�п���
            If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) > 0 Then
                '��ҩ;��
                i = .FindRow(CStr(.TextMatrix(lngRow, COL_���ID)), lngRow + 1, COL_ID)
                .TextMatrix(i, COL_ִ�п���ID) = rsInput!ID
                .Cell(flexcpData, i, COL_ִ�п���ID) = 1
                
                For i = lngRow + 1 To .Rows - 1
                    If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                        .TextMatrix(i, COL_����ִ��) = rsInput!����
                        .Cell(flexcpData, i, COL_����ִ��) = .TextMatrix(lngRow, COL_����ִ��)
                    Else
                        Exit For
                    End If
                Next
            ElseIf .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_ID)) _
                And InStr(",7,E,", .TextMatrix(lngRow - 1, COL_�������)) > 0 Then
                '��ҩ�÷�
                .TextMatrix(lngRow, COL_ִ�п���ID) = rsInput!ID
                .Cell(flexcpData, lngRow, COL_ִ�п���ID) = 1
            End If
        End If
    End With
End Sub
