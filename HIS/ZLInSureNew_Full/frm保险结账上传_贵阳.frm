VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frm���ս����ϴ�_���� 
   Caption         =   "������ҽ�����������ϴ�"
   ClientHeight    =   8265
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14625
   Icon            =   "frm���ս����ϴ�_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   14625
   StartUpPosition =   1  '����������
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   7905
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   635
      SimpleText      =   $"frm���ս����ϴ�_����.frx":0E42
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm���ս����ϴ�_����.frx":0E89
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20717
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14625
      _ExtentX        =   25797
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   14625
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   615
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   14505
         _ExtentX        =   25585
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         DisabledImageList=   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Object.ToolTipText     =   "��ӡԤ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "New"
               Object.ToolTipText     =   "���Ӳ��˽����¼"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Object.ToolTipText     =   "��������"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5595
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":171D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":1937
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":1B51
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":1D6B
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":1F85
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":219F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":23B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":25D3
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":27ED
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":2A07
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6390
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":311A
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":3334
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":354E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":3768
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":3982
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":3B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":3DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":3FD0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":41EA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ս����ϴ�_����.frx":4404
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   6015
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   14175
      _cx             =   25003
      _cy             =   10610
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm���ս����ϴ�_����.frx":4B17
      ScrollTrack     =   -1  'True
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
      ExplorerBar     =   7
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
         Visible         =   0   'False
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
         Begin VB.Menu mnuViewToolspilt1 
            Caption         =   "-"
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
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuShort 
      Caption         =   "��ݲ˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuShortMenu 
         Caption         =   "����(&A)"
         Index           =   0
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "�޸�(&M)"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShortMenu 
         Caption         =   "ɾ��(&D)"
         Index           =   2
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frm���ս����ϴ�_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsrue      As Integer
Private rsTemp          As ADODB.Recordset
Private mstrSortID  As String
Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Dim mintColumn As Integer
Dim mstrKey As String
Dim mint���� As Integer

Const conSql = "Select /*+ rule */*" & vbNewLine & _
                "From ����_�����ϴ� " & vbNewLine & _
                "Where �ϴ�ʱ�� >= sysdate-90"

Public Property Let Insure(ByVal vNewValue As Integer)
    mintInsrue = vNewValue
End Property

Private Sub Form_Load()
    Dim strField        As String
    Dim strFieldWIDth   As String
    Dim varField        As Variant
    Dim varFieldWIDth   As Variant
    Dim i               As Integer
                                 
    Call DataLoad
    If GetPersonSet Then

        RestoreWinState Me, App.ProductName
        RestoreFlexState vsfDetail, Me.Name
        'ʹ�ø��Ի����á����ѱ���ĸ�ʽ��
        strField = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrID", vsfDetail.Name & "����", "")
        strFieldWIDth = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrID", vsfDetail.Name & "���", "")
        varField = Split(strField, ",")
        varFieldWIDth = Split(strFieldWIDth, ",")
        For i = 0 To UBound(varField)
            If varField(i) <> "" And Val(varFieldWIDth(i)) <> 0 Then
                If vsfDetail.ColIndex(varField(i)) <> -1 Then
                    vsfDetail.ColPosition(vsfDetail.ColIndex(varField(i))) = i
                    vsfDetail.ColWidth(i) = Val(varFieldWIDth(i))
                End If
            End If
        Next
        Me.WindowState = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "����", 0)
        If Me.WindowState = 0 Then
            Me.Left = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "LEFT", Me.Left)
            Me.Top = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "TOP", Me.Top)
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
     
    With vsfDetail
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Top = sngTop
        .Height = ScaleHeight - sngTop
    End With
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    SaveFlexState vsfDetail, Me.Name
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "����", Me.WindowState
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "LEFT", Me.Left
    SaveSetting "ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name, "TOP", Me.Top
End Sub

Private Sub mnuEditAdd_Click()
    Dim strID       As String
    Dim str��ҳID       As String
    
    With frm���ս����ϴ�_�����༭
        .Insure = mintInsrue
        .Show vbModal
        If Not .OkCancel Then
            Set frm���ս����ϴ�_�����༭ = Nothing
            Exit Sub
        End If
'        strID = mintInsrue & .SickID & .PageID & .FeesID
    End With
    Set frm���ս����ϴ�_�����༭ = Nothing
    Call DataLoad
    vsfSetRow vsfDetail, strID, "ID"
End Sub

Private Sub mnuEditModify_Click()
    Dim strID As String
    
'    With frm���ս����ϴ�_�����༭
'        If vsfDetail.Rows <= 1 Then Exit Sub
'        .SickID = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("ID"))
'        .Insure = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����"))
'        .PageID = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("��ҳID"))
'        .FeesID = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("���ñ���"))
'        .Show vbModal
'        strID = mintInsrue & .SickID & .PageID & .FeesID
'        If Not .OkCancel Then
'            Set frm���ս����ϴ�_�����༭ = Nothing
'            Exit Sub
'        End If
'    End With
'    Set frm���ս����ϴ�_�����༭ = Nothing
'    Call DataLoad
'    vsfSetRow vsfDetail, strID, "ID"

End Sub

Private Sub mnuEditDelete_Click()
    Dim strID            As String
    Dim str��ҳID        As String
    Dim strDelNote       As String
    
    On Error GoTo errHandle
'    If vsfDetail.Rows <= 1 Then Exit Sub
'    strID = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("ID"))
'    strID = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("ID"))
'    str��ҳID = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("��ҳID"))
'    With frmCheckDelNote
'        .DelNote = vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("ȡ��ԭ��"))
'        .Show vbModal, Me
'        If (.DelNote = "") Then
'            Set frmCheckDelNote = Nothing
'            Exit Sub
'        End If
'        strDelNote = .DelNote
'    End With
'    Set frmCheckDelNote = Nothing
'    gstrSQL = "dl_����_��������_Cancel(" & vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("����")) & ",'" & strID & "','" & str��ҳID & "','" & vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("���ñ���")) & "','" & UserInfo.���� & "','" & strDelNote & "')"
'    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
'    Call DataLoad
'    vsfSetRow vsfDetail, strID, "ID"
'
'    Call SetMenu
'    MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    MousePointer = vbDefault
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim lngLoop         As Long
    Dim objControl      As Object
    Dim objPrint        As New zlPrint1Grd
    Dim objAppRow       As zlTabAppRow
    
    If vsfDetail Is Nothing Then Exit Sub
    LockWindowUpdate 0
    '���ô�ӡ��������
    Set objPrint.Body = vsfDetail
    objPrint.Title.Text = Me.Caption
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ�ˣ�" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ�䣺" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuShortIcon_Click(Index As Integer)
    mnuViewIcon_Click Index
End Sub

Private Sub mnuShortMenu_Click(Index As Integer)
    Select Case Index
        Case 0
            mnuEditAdd_Click
        Case 1
            mnuEditModify_Click
        Case 2
            mnuEditDelete_Click
    End Select
End Sub



Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Integer
'    For i = 0 To 3
'        mnuViewIcon(i).Checked = False
'    Next
'    mnuViewIcon(Index).Checked = True
End Sub

Private Sub mnuViewRefresh_Click()
    Call DataLoad
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
    Dim lngCount As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For lngCount = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(lngCount).Caption = IIf(mnuViewToolText.Checked = True, tbrThis.Buttons(lngCount).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    cbrThis.Refresh
    Call Form_Resize
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Preview"
            mnuFilePreview_Click
        Case "Print"
            mnuFilePrint_Click
        Case "New"
            mnuEditAdd_Click
     '   Case "Delete"
     '       mnuEditDelete_Click
    '    Case "Modify"
      '      mnuEditModify_Click
        Case "View"
'            If lvwItem.View = 3 Then
'                mnuViewIcon(0).Checked = True
'                lvwItem.View = 0
'            Else
'                mnuViewIcon(lvwItem.View + 1).Checked = True
'                lvwItem.View = lvwItem.View + 1
'            End If
      '  Case "Find"
      '      mnuViewFind_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
'    For i = 0 To 3
'        mnuViewIcon(i).Checked = False
'    Next
'    mnuViewIcon(ButtonMenu.Index - 1).Checked = True
'    lvwItem.View = ButtonMenu.Index - 1
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
    
End Sub

Private Sub SetMenu()
'���ܣ����ݵ�ǰ�������ò˵��Ŀ�����
    Dim bln���Թ� As Boolean
    Dim bln������ As Boolean
    
'    Call FillItem
'    stbThis.Panels(2).Text = lvwKind_S.SelectedItem.Text & "����" & lvwItem.ListItems.Count & "�����˼�¼"
    
    tbrThis.Buttons("New").Enabled = True
    mnuEdit.Enabled = True
    mnuEditAdd.Enabled = True
    mnuShortMenu(0).Enabled = True
    
    If vsfDetail.Rows > 1 Then
     '   tbrThis.Buttons("Modify").Enabled = True
'        tbrThis.Buttons("Delete").Enabled = True
'        tbrThis.Buttons("Split1").Enabled = True
        mnuEditModify.Enabled = True
        mnuShortMenu(1).Enabled = True
        mnuShortMenu(2).Enabled = True
    Else
      '  tbrThis.Buttons("Modify").Enabled = False
      '  tbrThis.Buttons("Delete").Enabled = False
     '   tbrThis.Buttons("Split1").Enabled = False
        mnuEditModify.Enabled = False
        mnuShortMenu(1).Enabled = False
        mnuShortMenu(2).Enabled = False
    End If
End Sub

Private Sub DataLoad()
    
    gstrSQL = conSql
    gstrSQL = gstrSQL
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, mintInsrue)
    Set vsfDetail.DataSource = rsTemp

    Call SetMenu
End Sub
'==============================================================================
'=���ܣ� �����λ��¼ vsfDetail
'==============================================================================
Private Sub vsfDetail_AfterSort(ByVal COL As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
'    vsfSetRow vsfDetail, mstrSortID, "ID"
    lngRow = vsfDetail.FindRow(mstrSortID, -1, vsfDetail.ColIndex("ID"), False, True)
    If lngRow > 0 Then vsfDetail.Row = lngRow
    vsfDetail.ShowCell lngRow, 1
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetail_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    Cancel = True
End Sub

'==============================================================================
'=���ܣ� ĳ�в����ƶ�λ�� vsfDetail[ͼ��]
'==============================================================================
Private Sub vsfDetail_BeforeMoveColumn(ByVal COL As Long, Position As Long)
    If COL = vsfDetail.ColIndex("ͼ��") Then
        Position = -1
    Else
        If Position <= vsfDetail.ColIndex("ͼ��") Then Position = COL
    End If
End Sub

'==============================================================================
'=���ܣ� ����ǰ��¼ID vsfDetail
'==============================================================================
Private Sub vsfDetail_BeforeSort(ByVal COL As Long, Order As Integer)
    On Error GoTo ErrH
    mstrSortID = "" & vsfDetail.TextMatrix(vsfDetail.Row, vsfDetail.ColIndex("ID"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ĳ�в����϶���С vsfDetail[ͼ��]
'==============================================================================
Private Sub vsfDetail_BeforeUserResize(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    If COL = vsfDetail.ColIndex("ͼ��") Then Cancel = True
End Sub

'==============================================================================
'=���ܣ� ˫������޸Ĺ��� vsfDetail
'==============================================================================
Private Sub vsfDetail_DblClick()
    On Error GoTo ErrH
    If vsfDetail.MouseRow <= 0 Then Exit Sub
    mnuEditModify_Click
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �Ҽ��˵� vsfDetail
'==============================================================================
Private Sub vsfDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '�����˵�����
            PopupMenu mnuShort
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
'==============================================================================
'=���ܣ����б任ʱ
'==============================================================================
Private Sub vsfDetail_RowColChange()
    Dim rsTemp          As ADODB.Recordset
    Dim varPos          As Variant
    On Error GoTo ErrH
    DoEvents
    If vsfDetail.Rows = 1 Then
'        With frmAuditItemEdit
'            .txtTypeID.Tag = "-1"
'            .txtTypeID.Text = ""
'            .txtName.Text = ""
'            .txtCode.Text = ""
'            .txtMnemonicCode.Text = ""
'            .cboUsed.ListIndex = -1
'            .cboLink.ListIndex = -1
'            .txtDescription.Text = ""
'            .txtAudit_NotCheck.Text = ""
'            Set .vsfFiles.DataSource = Nothing
'        End With
        stbThis.Panels(2) = "��ǰ��ʾ�� 0 ����Ŀ��"
        Exit Sub
    End If
    If vsfDetail.ColIndex("ID") <= 0 Then Exit Sub
    stbThis.Panels(2) = "��ǰ��ʾ�� " & vsfDetail.Rows - 1 & " ����Ŀ��"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

