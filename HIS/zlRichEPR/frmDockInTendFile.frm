VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockInTendFile 
   BorderStyle     =   0  'None
   Caption         =   "�����¼�ļ�"
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   60
      ScaleHeight     =   2715
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   1500
      Width           =   3675
      Begin VB.PictureBox picNote 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         ScaleHeight     =   195
         ScaleWidth      =   3825
         TabIndex        =   2
         Top             =   2190
         Width           =   3825
         Begin VB.Label lblNote 
            Caption         =   "Label1"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   30
            TabIndex        =   13
            Top             =   0
            Width           =   2175
         End
      End
      Begin XtremeSuiteControls.TabControl tbcFile 
         Height          =   1830
         Left            =   660
         TabIndex        =   1
         Top             =   330
         Width           =   2100
         _Version        =   589884
         _ExtentX        =   3704
         _ExtentY        =   3228
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picRecord 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5835
      Left            =   3930
      ScaleHeight     =   5835
      ScaleWidth      =   7275
      TabIndex        =   3
      Top             =   300
      Width           =   7275
      Begin VB.PictureBox picSplit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   45
         ScaleWidth      =   6435
         TabIndex        =   12
         Top             =   1560
         Width           =   6435
      End
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   1560
         Index           =   0
         Left            =   0
         ScaleHeight     =   1560
         ScaleWidth      =   6690
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   6690
         Begin VB.Frame fra 
            Height          =   525
            Left            =   0
            TabIndex        =   8
            Top             =   -90
            Width           =   6015
            Begin VB.ComboBox cboBaby 
               Height          =   300
               Left            =   4170
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   150
               Width           =   1350
            End
            Begin VB.Label lblFile 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����¼�ļ�:(���淶��ʽ��ʾ���Ļ����¼)"
               Height          =   180
               Left            =   60
               TabIndex        =   10
               Top             =   210
               Width           =   3690
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgFile 
            Height          =   1095
            Left            =   -15
            TabIndex        =   11
            Top             =   435
            Width           =   6060
            _cx             =   10689
            _cy             =   1931
            Appearance      =   2
            BorderStyle     =   0
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
            BackColorFixed  =   14737632
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
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
         Begin MSComctlLib.ImageList imgData 
            Left            =   3810
            Top             =   4080
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
                  Picture         =   "frmDockInTendFile.frx":0000
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendFile.frx":6862
                  Key             =   "����"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDockInTendFile.frx":6DFC
                  Key             =   "��ͨ"
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picPane 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   2700
         Index           =   1
         Left            =   0
         ScaleHeight     =   2700
         ScaleWidth      =   4680
         TabIndex        =   5
         Top             =   2505
         Width           =   4680
         Begin XtremeSuiteControls.TabControl tbcSub 
            Height          =   2490
            Left            =   120
            TabIndex        =   6
            Top             =   150
            Width           =   3450
            _Version        =   589884
            _ExtentX        =   6085
            _ExtentY        =   4392
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picPane 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1455
         Index           =   2
         Left            =   5310
         ScaleHeight     =   1455
         ScaleWidth      =   1410
         TabIndex        =   4
         Top             =   3330
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmDockInTendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnMouseMove As Boolean

'######################################################################################################################

Private Enum mCol
    r��־ = 0: rID: r����ʱ��: r��¼��Ŀ: r��¼����: r����: r��ʿ: r�Ǽ�ʱ��: r����ID: r������:: rǩ����: rǩ��ʱ��: r��Ŀ���: r��ʼ�汾: rδ��˵��
    f��־ = 0: fID: f���: f�ļ�: f���ڷ�Χ: f����id: f������: f������: f�ļ�����: f����
    w��־ = 0: wID: wҳ����: wҳ������: w��������: w������: w����ʱ��: w������: w���ʱ��: w��ǰ�汾: wǩ������: w��ǰ���: w�鵵��: w�鵵����: w����ID: w������: w����״̬
End Enum

Private Enum mColWidth
        c��־ = 270: cID = 0: c��� = 600: c�ļ� = 2000: c���ڷ�Χ = 3500: c����id = 0: c������ = 1200: c������ = 810: c�ļ����� = 0: c���� = 0
End Enum

Private WithEvents mclsDockAduits As zlRichEPR.clsDockAduits
Attribute mclsDockAduits.VB_VarHelpID = -1
Private mfrmCaseTendEditForBatch As frmCaseTendEditForBatch
Private mblnNoRefresh As Boolean
Private mstrPrivs As String                             '��ǰʹ���߶Ա�����(1255)��Ȩ�޴�
Private mlngPatiId As Long                              '����id
Private mlngPageId As Long                              '��ҳid
Private mlngDeptId As Long                              '��ǰ��������id���粡�˿��Һ͵�ǰ���Ҳ�һ�£����ܲ����鵵��Ĺ���
Private mblnEdit As Boolean                             '�Ƿ����������ͨ�����ϼ�������ݵ�ǰ���������Ƿ�ǰ���˲���������
Private mblnDoctorStation As Boolean
Private mbytFontSize As Byte                            '������ʾ��С0-9������,1-12������
Private mblnRefreshFontSize As Boolean                  '��¼�Ƿ�ˢ��������Ϣ
Private rsTemp As New ADODB.Recordset
Private mintBaby As Integer
Private mfrmMain As Object
Private mblnTendArchive As Boolean

Public Event AfterDataChanged()
Public Event Activate()

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

''######################################################################################################################

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-18 15:16
    '����:51746
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    Call ReSetFontSize
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytSize As Byte
    Dim PATI_COLWIDTH As Variant
    Dim lngCol As Long, lngReDraw As Long
    bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
    
    Me.FontSize = mbytFontSize
    Me.FontName = "����"
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("��") + 20
        Case UCase("VsFlexGrid")
            objCtrl.FontSize = mbytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = mbytFontSize
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
            objCtrl.PaintManager.Layout = xtpTabLayoutAutoSize
        End Select
    Next
    '�����ؼ�λ��
    PATI_COLWIDTH = Array(c��־, cID, c���, c�ļ�, c���ڷ�Χ, c����id, c������, c������, c�ļ�����, c����)
    With vfgFile
        lngReDraw = .Redraw
        .Redraw = flexRDNone
        For lngCol = fID To .Cols - 1
            .ColWidth(lngCol) = BlowUp(CDbl(PATI_COLWIDTH(lngCol)))
        Next lngCol
        .Redraw = lngReDraw
    End With
    
    lblNote.Top = 0: lblNote.Left = 30
    picNote.Height = lblNote.Height
    lblFile.Left = 60
    cboBaby.Top = 150
    cboBaby.Width = BlowUp(1350)
    lblFile.Top = cboBaby.Top + (cboBaby.Height - lblFile.Height) \ 2
    fra.Height = cboBaby.Top + cboBaby.Height + 75
    picSplit.Top = vfgFile.Rows * vfgFile.RowHeightMin + vfgFile.Top + 100
    Call Form_Resize
    
    'ˢ������
    Call ExecuteCommand("��������")
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange + (dblChange * IIf(mbytFontSize = 12, 1, 0) / 3)
End Function

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strInfo As String
    Dim byt����ȼ� As Byte
    Dim objFrmBody As Object
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open

        With vfgFile
        
            strInfo = Val(.TextMatrix(.Row, mCol.f����id))
            
            If Val(.TextMatrix(.Row, mCol.f����)) = -1 Then
                '���µ��鿴������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
                If Not CreateBodyEditor Then Exit Sub
                Set objFrmBody = gobjBodyEditor.GetTendBody
                On Error Resume Next
                objFrmBody.Resize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
                If Err <> 0 Then Err.Clear
                On Error GoTo errHand
                Call gobjBodyEditor.GetTendBody.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";0;0;" & mintBaby, 1, mstrPrivs)
            Else
                                    
                Call frmTendFileOpen.ShowMe(Me, Val(.TextMatrix(.Row, mCol.fID)), mlngPatiId, mlngPageId, Val(strInfo), mintBaby, .TextMatrix(.Row, mCol.f���ڷ�Χ), , Val(.TextMatrix(.Row, mCol.f������)), mblnMoved_HL, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
                
            End If
        End With

        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview
        
        ''1-Ԥ��,2-��ӡ
        
        With vfgFile

            If Val(.TextMatrix(.Row, mCol.f����)) = -1 Then
                
                Call mclsDockAduits.zlPrintDocument(1, 1)

            ElseIf .TextMatrix(.Row, mCol.f�ļ�) <> "�ļ�" Then
                
                Call mclsDockAduits.zlPrintDocument(2, 1)
                
            End If
            
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print
            
        With vfgFile

            If Val(.TextMatrix(.Row, mCol.f����)) = -1 Then
                
                Call mclsDockAduits.zlPrintDocument(1, 2)

            ElseIf .TextMatrix(.Row, mCol.f�ļ�) <> "�ļ�" Then
                
                Call mclsDockAduits.zlPrintDocument(2, 2)
                
            End If
            
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
            
        With vfgFile

            If Val(.TextMatrix(.Row, mCol.f����)) = -1 Then
                
                ShowSimpleMsg "�Բ������µ���֧�������Excel��"

            ElseIf .TextMatrix(.Row, mCol.f�ļ�) <> "�ļ�" Then
                
                Call mclsDockAduits.zlPrintDocument(2, 3)
                
            End If
            
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap
        
        '������ͼ������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
        If Not CreateBodyEditor Then Exit Sub
        Set objFrmBody = gobjBodyEditor.GetTendBody
        On Error Resume Next
        objFrmBody.Resize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
        If Err <> 0 Then Err.Clear
        On Error GoTo errHand
        If objFrmBody.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";0;1;" & mintBaby, 2, mstrPrivs) Then
            
            Call ExecuteCommand("ˢ������")

            RaiseEvent AfterDataChanged

        End If
        
    Case conMenu_File_PrintDayDetail        '����¼��
        If mfrmCaseTendEditForBatch Is Nothing Then Set mfrmCaseTendEditForBatch = New frmCaseTendEditForBatch
        Call mfrmCaseTendEditForBatch.ShowMe(Me, mlngDeptId, mstrPrivs)
    Case conMenu_Tool_Sign
        Call mclsDockAduits.zlGetFormTendEdit.SignMe
        RaiseEvent AfterDataChanged
    Case conMenu_Tool_SignEarse
        Call mclsDockAduits.zlGetFormTendEdit.UnSignMe
        RaiseEvent AfterDataChanged
    Case conMenu_Edit_Archive * 10
        Call mclsDockAduits.zlGetFormTendEdit.ArchiveMe
        RaiseEvent AfterDataChanged
    Case conMenu_Edit_UnArchive
        Call mclsDockAduits.zlGetFormTendEdit.UnArchiveMe
        RaiseEvent AfterDataChanged
    Case conMenu_Tool_SignVerify
        Call mclsDockAduits.SignMarker
    Case conMenu_Edit_Save
        If mclsDockAduits.zlGetFormTendEdit.SaveME Then RaiseEvent AfterDataChanged
    Case conMenu_Edit_Transf_Cancle
        Call mclsDockAduits.CancelMe
    End Select
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Public Property Get TendArchive() As Boolean
    TendArchive = mblnTendArchive
End Property

Public Property Let TendArchive(ByVal vData As Boolean)
    mblnTendArchive = vData
End Property

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open
        Control.Visible = True
        Control.Enabled = tbcFile.Item(1).Selected And (Val(vfgFile.TextMatrix(Me.vfgFile.Row, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview, conMenu_File_Print
        Control.Enabled = tbcFile.Item(1).Selected And (Val(vfgFile.TextMatrix(Me.vfgFile.Row, mCol.fID)) <> 0)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel
        Control.Enabled = tbcFile.Item(1).Selected And (vfgFile.Rows > 1 And Val(vfgFile.TextMatrix(vfgFile.Row, mCol.f����)) <> -1)
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_EditPopup
        Control.Visible = (mblnDoctorStation = False And InStr(1, mstrPrivs, "���µ���ͼ") > 0)
        Control.Enabled = Control.Visible
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MarkMap
        Control.Visible = (InStr(1, mstrPrivs, "���µ���ͼ") > 0 And mblnDoctorStation = False)
        Control.Enabled = (mblnEdit And mlngPatiId > 0 And Control.Visible And TendArchive = False And Not mblnMoved_HL) 'And Val(vfgFile.TextMatrix(vfgFile.Row, mCol.f����)) = -1
    Case conMenu_File_PrintDayDetail
        Control.Enabled = (mblnEdit And mlngPatiId > 0)  'And (Not mclsDockAduits.zlIsPigeonhole))

        Control.Visible = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0 And mblnDoctorStation = False And Not mblnMoved_HL)
        If Control.Enabled Then Control.Enabled = (InStr(1, mstrPrivs, "�����¼�Ǽ�") > 0)
    Case conMenu_Tool_Sign  'ǩ��
        Control.Visible = Not mblnDoctorStation
        Control.Enabled = (Not mclsDockAduits.zlIsCert) And (mlngPatiId > 0) And (Not mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL And mblnEdit
    Case conMenu_Tool_SignEarse  'ȡ��ǩ��
        Control.Visible = Not mblnDoctorStation
        Control.Enabled = mclsDockAduits.zlIsCert And (mlngPatiId > 0) And (Not mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL And mblnEdit
    Case conMenu_Edit_Archive * 10 '�鵵
        Control.Visible = Not mblnDoctorStation And mblnTendArchive = False
        Control.Enabled = (Not mclsDockAduits.zlIsPigeonhole) And (mlngPatiId > 0) And (Not mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL And mblnEdit
    Case conMenu_Edit_UnArchive  'ȡ���鵵
        Control.Visible = Not mblnDoctorStation And mblnTendArchive
        Control.Enabled = mclsDockAduits.zlIsPigeonhole And (mlngPatiId > 0) And (Not mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL And mblnEdit
    Case conMenu_Edit_Save  '����
        Control.Visible = Not mblnDoctorStation
        Control.Enabled = (mclsDockAduits.zlDataChange) And (Not mblnDoctorStation) And Not mblnMoved_HL
    Case conMenu_Edit_Transf_Cancle  'ȡ��
        Control.Visible = Not mblnDoctorStation
        Control.Enabled = (mclsDockAduits.zlDataChange) And (Not mblnDoctorStation)
    Case conMenu_Tool_SignVerify
        Control.Visible = (tbcFile.Selected.Index = 0)
        Control.Enabled = Control.Visible
    End Select
    
End Sub

Private Function zlRefData() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngCol As Long, lngRow As Long
    Dim rsMain As New ADODB.Recordset
    Dim rs As New ADODB.Recordset

    Dim strSvrKey As String
    Dim int���� As Integer
    Dim strCode As String
    Dim strFile As String
    Dim strStart As String
    Dim strEnd As String
    Dim lng����ID As Long
    Dim str���� As String
    Dim str������ As String
    Dim bln������ As Boolean
    Dim blnһ���ļ� As Boolean
    Dim strTmp As String
    '��������ֱ���ʾʱ,���һ���ȼ��������ļ�IDδ�����仯ʱ,ֻ��ʾһ���ļ�
    Dim strStart_Cur As String
    Dim strEnd_Cur As String
    Dim strHLDate_Cur As String
    Dim strFile_Cur As String
    Dim str����ID_Cur As String
    Dim str������_Cur As String
    Dim str���_CUR As String
    Dim str����_CUR As String
    Dim str����_CUR As String
    Dim str����_CUR As String
    Dim blnExit As Boolean
    
    Err = 0
    On Error GoTo errHand
    '------------------------------------------------------------------------------------------------------------------
    '�����ļ�ˢ��
    
    bln������ = (Val(zlDatabase.GetPara("�����������", glngSys, 1255, "0")) = 1)
    blnһ���ļ� = (Val(zlDatabase.GetPara("��ʾһ�ݻ����ļ�", glngSys, 1255, "1")) = 1)

    If bln������ Then
        '--------------------------------------------------------------------------------------------------------------
        With vfgFile
            .Rows = 2
            .Cols = 10
            .FixedCols = 1
            
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = "ID"
            .TextMatrix(0, 2) = "���"
            .TextMatrix(0, 3) = "�ļ�"
            .TextMatrix(0, 4) = "���ڷ�Χ"
            .TextMatrix(0, 5) = "����id"
            .TextMatrix(0, 6) = "����"
            .TextMatrix(0, 7) = "������"
            .TextMatrix(0, 8) = "�ļ�����"
            .TextMatrix(0, 9) = "����"
            
            Set .Cell(flexcpPicture, 1, mCol.f��־) = Nothing
            .TextMatrix(1, mCol.fID) = ""
            .TextMatrix(1, mCol.f���) = ""
            .TextMatrix(1, mCol.f�ļ�) = ""
            .TextMatrix(1, mCol.f���ڷ�Χ) = ""
            .TextMatrix(1, mCol.f����id) = ""
            .TextMatrix(1, mCol.f������) = ""
            .TextMatrix(1, mCol.f������) = ""
            .TextMatrix(1, mCol.f�ļ�����) = ""
            .TextMatrix(1, mCol.f����) = ""
            
            .ColWidth(mCol.f��־) = mColWidth.c��־
            .ColWidth(mCol.fID) = mColWidth.cID: .ColWidth(mCol.f���) = mColWidth.c���: .ColWidth(mCol.f�ļ�) = mColWidth.c�ļ�: .ColWidth(mCol.f���ڷ�Χ) = mColWidth.c���ڷ�Χ
            .ColWidth(mCol.f����id) = mColWidth.c����id: .ColWidth(mCol.f������) = mColWidth.c������: .ColWidth(mCol.f������) = mColWidth.c������: .ColWidth(mCol.f����) = mColWidth.c����: .ColWidth(mCol.f�ļ�����) = mColWidth.c�ļ�����
    
        End With
        
        gstrSQL = "Select a.Id, a.���, a.���� As �ļ�," & _
                "        To_Char(a.��ʼ, 'yyyy-mm-dd hh24:mi') || ' �� ' || To_Char(a.��ֹ, 'yyyy-mm-dd hh24:mi') As ���ڷ�Χ," & _
                "        a.����id, b.���� As ����, 3 As ������,����" & _
                " From (" & _
                "        Select f.Id, f.���, f.����, r.��ʼ, r.��ֹ, r.����id, ����" & _
                "        From ( Select Id, ���, ����, 3 As ������, ͨ��, 0 As ����id,���� From �����ļ��б� Where ����=3 And ����<0 And NVL(����,0)=0) f," & _
                "             (Select r.����id, Nvl(Min(r.������),3) As ������, Min(r.����ʱ��) As ��ʼ, Max(r.����ʱ��) As ��ֹ" & _
                "               From ���˻����¼ r" & _
                "               Where r.������Դ = 2 And r.����ID = [1] And NVL(r.��ҳID, 0) = [2] And Nvl(r.Ӥ��,0)=[3] " & _
                "               Group By r.����id) r" & _
                "        Where f.����<0  And f.������ >= r.������) a, ���ű� b" & _
                " Where a.����ID = b.ID " & _
                " Order By a.���, To_Char(a.��ʼ, 'yyyy-mm-dd hh24:mi') || ' �� ' || To_Char(a.��ֹ, 'yyyy-mm-dd hh24:mi')"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        End If
        Set rsMain = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
        
            
        If rsMain.BOF = False Then
            
            With vfgFile
                If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""
    
                Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("����").Picture
    
                .TextMatrix(.Rows - 1, mCol.fID) = rsMain("ID").Value
                .TextMatrix(.Rows - 1, mCol.f���) = rsMain("���").Value
                .TextMatrix(.Rows - 1, mCol.f�ļ�) = rsMain("�ļ�").Value
                .TextMatrix(.Rows - 1, mCol.f���ڷ�Χ) = rsMain("���ڷ�Χ").Value
                .TextMatrix(.Rows - 1, mCol.f����id) = rsMain("����id").Value
                .TextMatrix(.Rows - 1, mCol.f������) = rsMain("����").Value
                .TextMatrix(.Rows - 1, mCol.f������) = "/"
                .TextMatrix(.Rows - 1, mCol.f�ļ�����) = "/"
                .TextMatrix(.Rows - 1, mCol.f����) = -1
                
            End With
        End If
        
        '1.��ʱ��
        gstrSQL = _
            "Select a.����id,a.����id, b.������ As ������,d.���� As ����, Min(a.��ʼʱ��) As ��ʼʱ��, Max(Nvl(a.��ֹʱ��,Sysdate+100)) As ��ֹʱ��" & vbNewLine & _
            "From ���˱䶯��¼ a," & vbNewLine & _
            "        (Select Id, ����ȼ�,Decode(�ؼ�, 'Y', 0, Decode(һ��, 'Y', 1, Decode(����, 'Y', 2, 3))) As ������" & vbNewLine & _
            "            From (Select b.Id,b.���� As ����ȼ�, Decode(Sign(Instr(b.����, '��')), 1, 'Y', Decode(Sign(Instr(b.����, '��')), 1, 'Y', 'N')) As �ؼ�," & vbNewLine & _
            "                                        Decode(Sign(Instr(b.����, 'һ')), 1, 'Y'," & vbNewLine & _
            "                                                        Decode(Sign(Instr(b.����, '1')), 1, 'Y'," & vbNewLine & _
            "                                                                        Decode(Sign(Instr(b.����, '��')), 1, 'Y', Decode(Sign(Instr(b.����, 'I')), 1, 'Y', 'N')))) As һ��," & vbNewLine & _
            "                                        Decode(Sign(Instr(b.����, '��')), 1, 'Y'," & vbNewLine & _
            "                                                        Decode(Sign(Instr(b.����, '2')), 1, 'Y'," & vbNewLine & _
            "                                                                        Decode(Sign(Instr(b.����, '��')), 1, 'Y', Decode(Sign(Instr(b.����, 'II')), 1, 'Y', 'N')))) As ����" & vbNewLine & _
            "                           From �շ���ĿĿ¼ b" & vbNewLine & _
            "                           Where b.��� = 'H')) b,���ű� d" & vbNewLine & _
            "Where a.����id = [1] And a.��ҳid = [2] And b.Id = a.����ȼ�id  And d.Id = a.����id" & vbNewLine & _
            "Group By a.����id,a.����id, b.������,d.���� "
        gstrSQL = " Select * From (" & gstrSQL & ") Order by ��ʼʱ��"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
        If rs.EOF = False Then
            Do While Not rs.EOF
                '2.��ָ�����������Ļ����ļ�,ֻȡ��һ��(����������id,������)
                gstrSQL = _
                    "Select l.Id, l.���, l.����, l.����,a.����id,f.����" & vbNewLine & _
                    "From �����ļ��б� l, ����ҳ���ʽ f, ����Ӧ�ÿ��� a" & vbNewLine & _
                    "Where l.���� = 3 And l.���� = 0 And l.���� = f.���� And l.��� = f.��� And l.Id = a.�ļ�id(+) And" & vbNewLine & _
                    "           (l.���� < 0 Or l.ͨ�� = 1 Or l.ͨ�� = 2 And a.����id = [1]) And f.���� >= [2]" & vbNewLine & _
                    "Order By f.����"
                
                If IsNull(rs("����id").Value) = False Then
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(rs("����id").Value), Val(rs("������").Value))
                Else
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(rs("����id").Value), Val(rs("������").Value))
                End If
                
                strStart_Cur = ""
                Do While Not rsTemp.EOF
                    
                    'ֻȡ��һ��
                    int���� = rsTemp("����").Value
                    strCode = rsTemp("���").Value
                    strFile = rsTemp("����").Value
                    lng����ID = rs("����id").Value
                    str���� = rs("����").Value
                    str������ = rs!������
                    
                    strStart = Format(rs("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                    strEnd = Format(rs("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                    
                    If blnһ���ļ� Then
                        Call ShowFileOnly(mlngPatiId, mlngPageId, mintBaby, strStart, strEnd, lng����ID, rs!������, rsTemp!ID, strCode, strFile, str����, int����, Val(rsTemp("����").Value), rs.AbsolutePosition = 1)
                        Exit Do
                    Else
                        '��������ֱ���ʾʱ,���һ���ȼ��������ļ�IDδ�����仯ʱ,ֻ��ʾһ���ļ�
                        If strStart_Cur = "" Then
                            strStart_Cur = strStart
                            strEnd_Cur = strEnd
                            strFile_Cur = rsTemp!ID
                            str����ID_Cur = rs!����ID
                            str������_Cur = rs!������
                            
                            str���_CUR = strCode
                            str����_CUR = strFile
                            str����_CUR = str����
                            str����_CUR = Val(rsTemp("����").Value)
                        End If
                        
                        If str����ID_Cur <> rs!����ID Or Val(str������_Cur) <> rs!������ Or strFile_Cur <> rsTemp!ID Then
                            Call ShowFile(mlngPatiId, mlngPageId, mintBaby, strStart_Cur, strEnd_Cur, str����ID_Cur, str������_Cur, strFile_Cur, str���_CUR, str����_CUR, str����_CUR, int����, Val(str����_CUR), True)
                            
                            strStart_Cur = strStart
                            strFile_Cur = rsTemp!ID
                            str����ID_Cur = rs!����ID
                            str������_Cur = rs!������
                            
                            str���_CUR = strCode
                            str����_CUR = strFile
                            str����_CUR = str����
                            str����_CUR = Val(rsTemp("����").Value)
                        End If
                        strEnd_Cur = strEnd
                    End If
                    
                    rsTemp.MoveNext
                Loop
                '���һ����¼����Ҫ���
                If Not blnһ���ļ� Then
                    Call ShowFile(mlngPatiId, mlngPageId, mintBaby, strStart_Cur, strEnd_Cur, str����ID_Cur, str������_Cur, strFile_Cur, str���_CUR, str����_CUR, str����_CUR, int����, Val(str����_CUR), True)
                End If
                
                rs.MoveNext
            Loop
        End If
        
    Else
        '--------------------------------------------------------------------------------------------------------------
        
        With vfgFile
            .Rows = 2
            .Cols = 10
            .FixedCols = 1
            
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = "ID"
            .TextMatrix(0, 2) = "���"
            .TextMatrix(0, 3) = "�ļ�"
            .TextMatrix(0, 4) = "���ڷ�Χ"
            .TextMatrix(0, 5) = "����id"
            .TextMatrix(0, 6) = "����"
            .TextMatrix(0, 7) = "������"
            .TextMatrix(0, 8) = "�ļ�����"
            .TextMatrix(0, 9) = "����"
            
            Set .Cell(flexcpPicture, 1, mCol.f��־) = Nothing
            .TextMatrix(1, mCol.fID) = ""
            .TextMatrix(1, mCol.f���) = ""
            .TextMatrix(1, mCol.f�ļ�) = ""
            .TextMatrix(1, mCol.f���ڷ�Χ) = ""
            .TextMatrix(1, mCol.f����id) = ""
            .TextMatrix(1, mCol.f������) = ""
            .TextMatrix(1, mCol.f������) = ""
            .TextMatrix(1, mCol.f�ļ�����) = ""
            .TextMatrix(1, mCol.f����) = ""
            
            .ColWidth(mCol.f��־) = mColWidth.c��־
            .ColWidth(mCol.fID) = mColWidth.cID: .ColWidth(mCol.f���) = mColWidth.c���: .ColWidth(mCol.f�ļ�) = mColWidth.c�ļ�: .ColWidth(mCol.f���ڷ�Χ) = mColWidth.c���ڷ�Χ
            .ColWidth(mCol.f����id) = mColWidth.c����id: .ColWidth(mCol.f������) = mColWidth.c������: .ColWidth(mCol.f������) = mColWidth.c������: .ColWidth(mCol.f����) = mColWidth.c����: .ColWidth(mCol.f�ļ�����) = mColWidth.c�ļ�����
    
        End With
        
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select distinct a.Id, a.���, a.���� As �ļ�," & _
                "        a.��ʼ,a.��ֹ," & _
                "        a.����id, b.���� As ����, 0 As ������,a.�ļ�����,����" & _
                " From (" & _
                "        Select f.Id, f.���, f.����, r.��ʼ, r.��ֹ, r.����id, ����,�ļ����� " & _
                "        From ( Select Id, ���, ����, 3 As �ļ�����, ͨ��, 0 As ����id,���� From �����ļ��б� Where ����=3 And ����<0 And NVL(����,0)=0" & _
                "               Union All " & _
                "               Select l.Id, l.���, l.����, f.���� As �ļ�����, l.ͨ��, a.����id,l.���� " & _
                "               From �����ļ��б� l, ����ҳ���ʽ f, ����Ӧ�ÿ��� a" & _
                "               Where l.���� = 3 And l.���� = 0 And l.���� = f.���� And l.��� = f.��� And l.Id = a.�ļ�id(+)) f," & _
                "             (Select r.����id, Nvl(Min(r.������),3) As ������, Min(r.����ʱ��) As ��ʼ, Max(r.����ʱ��) As ��ֹ" & _
                "               From ���˻����¼ r" & _
                "               Where r.������Դ = 2 And r.����ID = [1] And NVL(r.��ҳID, 0) = [2] And Nvl(r.Ӥ��,0)=[3] " & _
                "               Group By r.����id) r" & _
                "        Where (f.����<0 Or f.ͨ�� = 1 Or f.ͨ�� = 2 And r.����id In (Select t.����id From �������Ҷ�Ӧ t Where t.����id=f.����id)) And f.�ļ����� >= r.������) a, ���ű� b" & _
                " Where a.����ID = b.ID " & _
                " Order By a.����,A.�ļ�����,A.��� desc, To_Char(a.��ʼ, 'yyyy-mm-dd hh24:mi') || ' �� ' || To_Char(a.��ֹ, 'yyyy-mm-dd hh24:mi')"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        End If
        Set rsMain = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId, mintBaby)
        
        '��Ҫ���н������ݴ���(���ܳ�����һ��A�ļ�,�ڶ���A�ļ�,��һ��B�ļ�,�ڶ���B�ļ�,��ѡ��ֻ��ʾһ���ļ�ʱ,��������¼��Ӧ����ʾ����)
        Dim rsData As New ADODB.Recordset
        Set rsData = DataProcess(rsMain, blnһ���ļ�)
        
        With Me.vfgFile
            If rsData.RecordCount <> 0 Then rsData.MoveFirst
            Do While Not rsData.EOF
                If rsData!ɾ�� = 0 Then
                    int���� = rsData("����").Value
                    strCode = rsData("���").Value
                    strFile = rsData("����").Value
                    lng����ID = rsData("����id").Value
                    str���� = rsData("��������").Value
    
                    strStart = Format(rsData("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                    strEnd = Format(rsData("����ʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                    
                    Call ShowFile(mlngPatiId, mlngPageId, mintBaby, strStart, strEnd, lng����ID, 0, Val(rsData("�ļ�ID").Value), strCode, strFile, str����, int����, Val(rsData("�ļ�����").Value), (rsData.AbsolutePosition = 1))
                End If
                
                rsData.MoveNext
            Loop

            For lngRow = .FixedRows To .Rows - 1
                If Val(.TextMatrix(lngRow, mCol.f����)) = -1 Then
                    Set .Cell(flexcpPicture, lngRow, mCol.f��־) = Me.imgData.ListImages("����").Picture
                Else
                    Set .Cell(flexcpPicture, lngRow, mCol.f��־) = Me.imgData.ListImages("��ͨ").Picture
                End If
            Next
        End With
    End If
    
    If mblnEdit = True Then
        '41778,������,2012-09-06
        '��������ϰ���°����ݶ��Ѿ����ڣ������κ����ơ����ֻ���°����ݣ�û���ϰ档���ϰ岻������ļ���
        'Ӥ��Ӧ�ú�ĸ��ʹ��ͬһ��ϵͳ��
        gstrSQL = "Select 1 From ���˻����ļ� A Where a.����id = [1] And a.��ҳid = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngPatiId, mlngPageId)
        If rsTemp.RecordCount > 0 And Val(vfgFile.TextMatrix(vfgFile.Rows - 1, mCol.fID)) = 0 Then
            mblnEdit = False
        End If
    End If
    
    zlRefData = True

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DataProcess(ByVal rsMain As ADODB.Recordset, ByVal blnһ���ļ� As Boolean) As ADODB.Recordset
    Dim blnAdd As Boolean           'δ��������ʾ�����Ļ����¼��
    Dim arrFormat, intFormat As Integer
    Dim strField As String, strValue As String, str��ʼ As String, str��ֹ As String
    Dim intLocal As Integer '��ǰָ��λ��
    Dim intCount As Integer
        Dim intRecords As Integer
    Dim int������ As Integer, lng����ID As Long, int���� As Integer
    Dim rsData As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '��Ҫ���н������ݴ���(���ܳ�����һ��A�ļ�,�ڶ���A�ļ�,��һ��B�ļ�,�ڶ���B�ļ�,��ѡ��ֻ��ʾһ���ļ�ʱ,��������¼��Ӧ����ʾ����)
        
    strField = "ID," & adDouble & ",5|����," & adDouble & ",18|���," & adLongVarChar & ",50|����," & adLongVarChar & ",200|" & _
               "����ID," & adDouble & ",18|��������," & adLongVarChar & ",200|��ʼʱ��," & adLongVarChar & ",20|" & _
               "����ʱ��," & adLongVarChar & ",20|�ļ�ID," & adDouble & ",18|�ļ�����," & adDouble & ",18|ɾ��," & adDouble & ",1"
    Set rsData = New ADODB.Recordset
    Call Record_Init(rsData, strField)
    
    strField = "ID|����|���|����|����ID|��������|��ʼʱ��|����ʱ��|�ļ�ID|�ļ�����|ɾ��"
    If rsMain.RecordCount <> 0 Then rsMain.MoveFirst
    Do While Not rsMain.EOF
        str��ʼ = Format(rsMain("��ʼ").Value, "yyyy-MM-dd HH:mm:ss")
        str��ֹ = Format(rsMain("��ֹ").Value, "yyyy-MM-dd HH:mm:ss")
        blnAdd = True
        
        '�����ƻ����¼��
        If rsMain!���� <> -1 Then
            gstrSQL = " Select ��ʽ From ����ҳ���ʽ Where ����=3 And ���=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ���ʽ", CStr(rsMain!���))
            intFormat = 0
            arrFormat = Split(NVL(rsTemp!��ʽ, ";;;;;;;;"), ";")
            If UBound(arrFormat) >= 8 Then intFormat = Val(arrFormat(8))
            
            If intFormat <> 0 Then
                '1-����ǰ;2-�����
                gstrSQL = " Select MAX(����ʱ��) AS ����ʱ�� From ������������¼ Where ����ID=[1] And ��ҳID=[2] "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����������ʱ��", mlngPatiId, mlngPageId)
                If Not IsNull(rsTemp!����ʱ��) Then
                    If intFormat = 1 Then
                        str��ֹ = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
                    Else
                        str��ʼ = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
                    End If
                Else
                    blnAdd = (intFormat = 1)
                End If
            End If
        End If
        
        If blnAdd Then
                        intRecords = intRecords + 1
            strValue = intRecords & "|" & rsMain!���� & "|" & rsMain!��� & "|" & rsMain!�ļ� & "|" & rsMain!����ID & "|" & _
                    rsMain!���� & "|" & str��ʼ & "|" & str��ֹ & "|" & rsMain!ID & "|" & Val(rsMain!�ļ�����) & "|0"
            'Debug.Print strValue
            Call Record_Update(rsData, strField, strValue, "ID|" & intRecords)
        End If
        rsMain.MoveNext
    Loop
    
    If Not blnһ���ļ� Then
        Set DataProcess = rsData
        Exit Function
    End If
    
    '����ѭ�����,֮����ڻ���ȼ�\����ID��ͬ��,�Ѽ�¼ɾ��
    intCount = rsData.RecordCount
    If intCount > 0 Then
        For intLocal = 1 To intCount
            rsData.MoveFirst
            rsData.Move intLocal - 1
            
            If rsData!ɾ�� = 0 Then
                int���� = rsData!����
                int������ = rsData!�ļ�����
                lng����ID = rsData!����ID
                
                rsData.MoveFirst
                Do While Not rsData.EOF
                    If rsData.AbsolutePosition <> intLocal Then
                        If rsData!ɾ�� = 0 And rsData!�ļ����� = int������ And rsData!����ID = lng����ID And int���� = rsData!���� Then
                            Call Record_Update(rsData, "ɾ��", 1, "ID|" & rsData.AbsolutePosition)
                        End If
                    End If
                    rsData.MoveNext
                Loop
            End If
        Next
    End If
    Set DataProcess = rsData
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowFile(ByVal lngPatiID As Long, _
                        ByVal lngPageId As Long, _
                        ByVal intBaby As Integer, _
                        ByVal strStart As String, _
                        ByVal strEnd As String, _
                        ByVal lng����ID As Long, _
                        ByVal byt������ As Byte, _
                        ByVal lngId As Long, _
                        ByVal strCode As String, _
                        ByVal strFile As String, _
                        ByVal str���� As String, _
                        ByVal int���� As Integer, _
                        ByVal byt�ļ����� As Byte, _
                        Optional ByVal blnFirst As Boolean = False) As Boolean
    '******************************************************************************************************************
    '���ܣ����ָ��ʱ�������û�л����¼���ݣ�����������д������ʵ�ʵ����ڷ�Χ
    '������blnShow=False,��ʾ�����ݲ���ʾ;True,�˹���������,ǿ����ʾ
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    With vfgFile
        gstrSQL = "Select Min(r.����ʱ��) As ��ʼ, Max(r.����ʱ��) As ��ֹ From ���˻����¼ r Where r.����ID = [1] And NVL(r.��ҳID, 0) = [2] And Nvl(r.Ӥ��,0)=[3] And r.����ʱ�� between [4] And [5] And r.����id=[6] And r.������<=[7]"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        End If
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId, intBaby, CDate(strStart), CDate(strEnd), lng����ID, byt�ļ�����)
        
        If rs.EOF = False Then
            
            If strEnd >= Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then strEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            If zlCommFun.NVL(rs("��ʼ").Value, "") <> "" Then

                strStart = Format(rs("��ʼ").Value, "yyyy-MM-dd HH:mm")
                strEnd = Format(rs("��ֹ").Value, "yyyy-MM-dd HH:mm")
                
                If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""

                If int���� = -1 Then
                    Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("����").Picture
                Else
                    Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("��ͨ").Picture
                End If

                .TextMatrix(.Rows - 1, mCol.fID) = lngId
                .TextMatrix(.Rows - 1, mCol.f���) = strCode
                .TextMatrix(.Rows - 1, mCol.f�ļ�) = strFile
                .TextMatrix(.Rows - 1, mCol.f���ڷ�Χ) = Format(strStart, "yyyy-MM-dd HH:mm") & " �� " & Format(strEnd, "yyyy-MM-dd HH:mm")
                .TextMatrix(.Rows - 1, mCol.f����id) = lng����ID
                .TextMatrix(.Rows - 1, mCol.f������) = str����
                .TextMatrix(.Rows - 1, mCol.f������) = IIf(int���� = -1, "/", byt������)
                .TextMatrix(.Rows - 1, mCol.f�ļ�����) = IIf(int���� = -1, "/", byt�ļ�����)
                .TextMatrix(.Rows - 1, mCol.f����) = int����
            End If
        End If
    End With

    ShowFile = True

End Function

Private Function ShowFileOnly(ByVal lngPatiID As Long, _
                        ByVal lngPageId As Long, _
                        ByVal intBaby As Integer, _
                        ByVal strStart As String, _
                        ByVal strEnd As String, _
                        ByVal lng����ID As Long, _
                        ByVal byt������ As Byte, _
                        ByVal lngId As Long, _
                        ByVal strCode As String, _
                        ByVal strFile As String, _
                        ByVal str���� As String, _
                        ByVal int���� As Integer, _
                        ByVal byt�ļ����� As Byte, _
                        Optional ByVal blnFirst As Boolean = False) As Boolean
    '******************************************************************************************************************
    '���ܣ����ָ��ʱ�������û�л����¼���ݣ�����������д������ʵ�ʵ����ڷ�Χ
    '������blnShow=False,��ʾ�����ݲ���ʾ;True,�˹���������,ǿ����ʾ
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Static sbyt������  As Byte
    Static slng����ID As Long
    Static slng����ID As Long
    Static sintBaby As Integer
    Static sbyt�ļ����� As Byte

    If slng����ID <> lngPatiID Or sintBaby <> intBaby Or blnFirst Then
        '������˷����仯,����
        slng����ID = lngPatiID
        sintBaby = intBaby
        sbyt������ = 0
        slng����ID = 0
        sbyt�ļ����� = 0
    End If
    
    With vfgFile
        gstrSQL = "Select Min(r.����ʱ��) As ��ʼ, Max(r.����ʱ��) As ��ֹ From ���˻����¼ r Where r.����ID = [1] And NVL(r.��ҳID, 0) = [2] And Nvl(r.Ӥ��,0)=[3] And r.����ʱ�� between [4] And [5] And r.����id=[6] And r.������<=[7]"
        If mblnMoved_HL Then
            gstrSQL = Replace(gstrSQL, "���˻����¼", "H���˻����¼")
            gstrSQL = Replace(gstrSQL, "���˻�������", "H���˻�������")
        End If
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId, intBaby, CDate(strStart), CDate(strEnd), lng����ID, byt�ļ�����)
        
        If rs.EOF = False Then
            
            If strEnd >= Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then strEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            If zlCommFun.NVL(rs("��ʼ").Value, "") <> "" Then
                If (sbyt������ <> byt������ Or slng����ID <> lng����ID Or sbyt�ļ����� <> byt�ļ�����) Then
                    
                    If Val(.TextMatrix(.Rows - 1, mCol.fID)) > 0 Then .AddItem ""
    
                    If int���� = -1 Then
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("����").Picture
                    Else
                        Set .Cell(flexcpPicture, .Rows - 1, mCol.f��־) = imgData.ListImages("��ͨ").Picture
                    End If
    
                    .TextMatrix(.Rows - 1, mCol.fID) = lngId
                    .TextMatrix(.Rows - 1, mCol.f���) = strCode
                    .TextMatrix(.Rows - 1, mCol.f�ļ�) = strFile
                    .TextMatrix(.Rows - 1, mCol.f���ڷ�Χ) = Format(strStart, "yyyy-MM-dd HH:mm") & " �� " & Format(strEnd, "yyyy-MM-dd HH:mm")
                    .TextMatrix(.Rows - 1, mCol.f����id) = lng����ID
                    .TextMatrix(.Rows - 1, mCol.f������) = str����
                    .TextMatrix(.Rows - 1, mCol.f������) = IIf(int���� = -1, "/", byt������)
                    .TextMatrix(.Rows - 1, mCol.f�ļ�����) = IIf(int���� = -1, "/", byt�ļ�����)
                    .TextMatrix(.Rows - 1, mCol.f����) = int����
    '
                    sbyt������ = byt������
                    slng����ID = lng����ID
                    sbyt�ļ����� = byt�ļ�����
                Else
                    .TextMatrix(.Rows - 1, mCol.f���ڷ�Χ) = Split(.TextMatrix(.Rows - 1, mCol.f���ڷ�Χ), " �� ")(0) & " �� " & Format(strEnd, "yyyy-MM-dd HH:mm")
                End If
            End If
    
            If int���� = -1 Then
                sbyt������ = 0
                slng����ID = 0
                sbyt�ļ����� = 0
            End If
        End If
    End With

    ShowFileOnly = True

End Function

Public Function InitData(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Then Exit Function
    Call ExecuteCommand("��ע���")
    Call ExecuteCommand("�ؼ�״̬")
    
End Function

Public Function RefreshData(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngDeptId As Long, ByVal blnDoctorStation As Boolean, ByVal blnEdit As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ�ˢ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    
    mlngPatiId = lng����ID
    mlngPageId = lng��ҳID
    mlngDeptId = lngDeptId
    mblnEdit = blnEdit And Not mblnMoved_HL
    
    mblnDoctorStation = blnDoctorStation
    mblnRefreshFontSize = False
    Call ExecuteCommand("ˢ������")
    
    If mblnDoctorStation Then
        tbcFile.Item(1).Selected = True
        tbcFile.Item(0).Visible = False
    End If
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    Dim byt����ȼ� As Byte, bytSize As Byte
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        Set mclsDockAduits = New zlRichEPR.clsDockAduits
        Call FormSetCaption(mclsDockAduits.zlGetFormTendBody, False, False)
        
        With tbcSub
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .BoldSelected = True
                .ClientFrame = xtpTabFrameSingleLine
                .ShowIcons = True
                .DisableLunaColors = False
                .Position = xtpTabPositionTop
            End With

            .InsertItem 0, "", picPane(2).hWnd, 0
            .InsertItem 1, "���¼�¼��", mclsDockAduits.zlGetFormTendBody.hWnd, 0
            .InsertItem 2, "�����¼��", mclsDockAduits.zlGetFormTendFile.hWnd, 0
            .Item(0).Selected = True
            Call SetTabVisible(0)
        End With
        
        With tbcFile
            With .PaintManager
                .Appearance = xtpTabAppearancePropertyPage2003
                .BoldSelected = True
                .ClientFrame = xtpTabFrameSingleLine
                .ShowIcons = True
                .DisableLunaColors = False
                .Position = xtpTabPositionBottom
            End With

            .InsertItem 0, "����¼��", mclsDockAduits.zlGetFormTendEdit.hWnd, 0
            .InsertItem 1, "�����¼��", picRecord.hWnd, 0
            .Item(0).Selected = True
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"

        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ��״̬"
        

        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
    
        '�жϲ����Ƿ���ת��
        '��Ϊ�ú������ⶼ�ڵ���,�������ñ�,ֱ�Ӷ�ȡ
        '------------------------------------------------------------------------------------------------------------------
        
        mblnNoRefresh = True
        cboBaby.Clear
        cboBaby.AddItem "���˱���"
        gstrSQL = "Select a.���,Decode(a.Ӥ������,Null,NVL(c.����,b.����) ||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������" & _
            " From ������Ϣ b,������ҳ c,������������¼ a Where b.����id=c.����id And a.����id=c.����id And a.��ҳid=c.��ҳid And c.����id=[1] And c.��ҳid=[2]  Order By a.���"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiId, mlngPageId)
        If rs.BOF = False Then
            Do While Not rs.EOF
                cboBaby.AddItem rs("Ӥ������").Value
                rs.MoveNext
            Loop
        End If
        cboBaby.ListIndex = 0
        cboBaby.Visible = (cboBaby.ListCount > 1)
        
        Call zlRefData
        mblnNoRefresh = False
        Call ExecuteCommand("��ʾ�ļ�����", vfgFile.Row)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʾ�ļ�����"
        If tbcFile.Item(0).Selected Then
            '��ȡ�ò��˵�ʱ�Ļ���ȼ�
            gstrSQL = "select Zl_Patittendgrade([1],[2]) from dual"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngPatiId, mlngPageId)
            byt����ȼ� = rs.Fields(0).Value
            Call mclsDockAduits.zlRefreshTendEdit(mlngPatiId, mlngPageId, mlngDeptId, byt����ȼ�, 0, mstrPrivs, False, mblnEdit)
        Else
            With vfgFile
                Call mclsDockAduits.zlRefreshTendBody(mlngPatiId, mlngPageId, mlngDeptId, mintBaby)
                If Val(.TextMatrix(.Row, mCol.f����)) = -1 Then
                    '���µ��鿴������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��
                    
                    Call SetTabVisible(1)
                    tbcSub.Item(1).Caption = .TextMatrix(.Row, mCol.f�ļ�) & "(" & .TextMatrix(.Row, mCol.f���ڷ�Χ) & ")"
                
                ElseIf .TextMatrix(.Row, mCol.f�ļ�) <> "�ļ�" And .TextMatrix(.Row, mCol.f�ļ�) <> "" Then
                    
                    Call mclsDockAduits.zlRefresh(3, Val(.TextMatrix(.Row, mCol.fID)), mlngPatiId, mlngPageId, Val(.TextMatrix(.Row, mCol.f����id)), .TextMatrix(.Row, mCol.f���ڷ�Χ), Val(.TextMatrix(.Row, mCol.f�ļ�����)), mintBaby)
                    Call SetTabVisible(2)
                    tbcSub.Item(2).Caption = .TextMatrix(.Row, mCol.f�ļ�) & "(" & .TextMatrix(.Row, mCol.f���ڷ�Χ) & ")"
                    tbcSub.Item(1).Selected = True
                    tbcSub.Item(2).Selected = True
                Else
                    Call SetTabVisible(0)
                    tbcSub.Item(0).Caption = "�޿���ʾ�Ļ����ļ�"
                End If
                tbcSub.PaintManager.Layout = xtpTabLayoutAutoSize
            End With
        End If
        '������������
        If mblnRefreshFontSize = True Then Call ExecuteCommand("��������")
        mblnRefreshFontSize = True
    Case "��������"
        bytSize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
        If tbcFile.Item(0).Selected Then
            '����¼��
            Call mclsDockAduits.SetFontSize(0, bytSize)
        Else
            With vfgFile
                If Val(.TextMatrix(.Row, mCol.f����)) = -1 Then
                    '���µ���¼
                    Call mclsDockAduits.SetFontSize(1, bytSize)
                ElseIf .TextMatrix(.Row, mCol.f�ļ�) <> "�ļ�" And .TextMatrix(.Row, mCol.f�ļ�) <> "" Then
                    '��¼��
                    Call mclsDockAduits.SetFontSize(2, bytSize)
                End If
            End With
        End If
    End Select

    ExecuteCommand = True

    GoTo EndHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
EndHand:

End Function

Private Function SetTabVisible(ByVal intIndex As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    
    If tbcSub.Item(intIndex).Visible = False Then
        tbcSub.Item(intIndex).Visible = True
        tbcSub.Item(intIndex).Selected = True
    End If

    For intLoop = 0 To tbcSub.ItemCount - 1
        If intLoop <> intIndex Then
            If tbcSub.Item(intLoop).Visible = True Then tbcSub.Item(intLoop).Visible = False
        End If
    Next
    SetTabVisible = True
End Function

Private Sub cboBaby_Click()
    If mintBaby = cboBaby.ListIndex Then Exit Sub
    mintBaby = cboBaby.ListIndex
    If mblnNoRefresh = True Then Exit Sub
    Call zlRefData
End Sub

Private Sub Form_Load()
    lblNote.Caption = ""
    mblnMouseMove = False
End Sub

Private Sub Form_Resize()
    Dim intSel As Integer
    On Error Resume Next
    
    picFile.Move 0, 0, Me.ScaleWidth + 500, Me.ScaleHeight + 500
    picRecord.Move 0, 0, picFile.ScaleWidth, picFile.ScaleHeight
    picNote.Move 3000, tbcFile.Top + tbcFile.Height - picNote.Height - 50, picFile.Width - picNote.Left
    
    '����ѡ��ǰҳ��,���������,��ôҳͷ�Ϳ�����,�ֵĺ�,���ƺ͸ÿؼ�Ƕ���й�
    intSel = tbcFile.Selected.Index
    tbcFile.Item(0).Selected = True
    If intSel <> 0 Then tbcFile.Item(intSel).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnNoRefresh = False
    Set mclsDockAduits = Nothing
    If Not mfrmCaseTendEditForBatch Is Nothing Then Unload mfrmCaseTendEditForBatch
    Set mfrmCaseTendEditForBatch = Nothing
End Sub

Private Sub mclsDockAduits_ShowItemInfo(ByVal strInfo As String)
    lblNote.Width = picNote.Width
    lblNote.Caption = strInfo
End Sub

Private Sub PicFile_Resize()
    On Error Resume Next
    
    tbcFile.Move 0, 0, picFile.ScaleWidth - 500, picFile.ScaleHeight - 500
End Sub


Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        fra.Move 0, -90, picPane(Index).Width
        
        cboBaby.Move fra.Width - cboBaby.Width, cboBaby.Top
        vfgFile.Move 15, fra.Top + fra.Height + 15, picPane(Index).Width - 30, picPane(Index).Height - (fra.Top + fra.Height + 15) - 15
    Case 1
        tbcSub.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub picRecord_Resize()
    On Error Resume Next
    If picSplit.Top < 1000 Then picSplit.Top = 1000
    If picSplit.Top > picRecord.Height - 2000 Then picSplit.Top = picRecord.Height - 2000
    
    With picSplit
        .Left = 0
        .Width = picRecord.Width
    End With
    
    With picPane(0)
        .Height = picSplit.Top
        .Width = picSplit.Width
    End With
    
    With picPane(1)
        .Top = picSplit.Top + picSplit.Height
        .Height = picRecord.Height - .Top
        .Width = picSplit.Width
    End With
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMouseMove = (Button = 1)
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnMouseMove = False Then Exit Sub
    
    If picSplit.Top < 1000 Then picSplit.Top = 1000
    If picSplit.Top > picRecord.Height - 2000 Then picSplit.Top = picRecord.Height - 2000
    picSplit.Move 0, picSplit.Top + Y
    Me.Refresh
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMouseMove = False
    
    Call picRecord_Resize
End Sub

Private Sub tbcFile_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNoRefresh = True Then Exit Sub
    lblNote.Caption = ""
    Call ExecuteCommand("��ʾ�ļ�����")
End Sub

Private Sub vfgFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNoRefresh = True Then Exit Sub
    If OldRow <> NewRow Then
        
        Call ExecuteCommand("��ʾ�ļ�����", NewRow)
        DoEvents
        
        On Error Resume Next
        vfgFile.SetFocus
    End If

    
End Sub

Private Sub vfgFile_DblClick()
    Dim strInfo As String
    Dim intEdit As Integer
    Dim objFrmBody As Object
    
    On Error GoTo errHand
    
    strInfo = Val(Me.vfgFile.TextMatrix(vfgFile.Row, mCol.f����id))

    If Val(vfgFile.TextMatrix(vfgFile.Row, mCol.f����)) = -1 Then
        '���µ��鿴������ID;��ҳID;����ID;��Ժ;�༭;Ӥ��

        intEdit = 0
        If (InStr(1, mstrPrivs, "���µ���ͼ") > 0 And mblnDoctorStation = False) Then
            If (mblnEdit And mlngPatiId > 0 And TendArchive = False) Then
                intEdit = 1
            End If
        End If
        
        If Not CreateBodyEditor Then Exit Sub
        Set objFrmBody = gobjBodyEditor.GetTendBody
        On Error Resume Next
        objFrmBody.Resize = IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize))
        If Err <> 0 Then Err.Clear
        On Error GoTo errHand
        Call objFrmBody.ShowEdit(Me, mlngPatiId & ";" & mlngPageId & ";" & mlngDeptId & ";0;" & intEdit & ";" & mintBaby, 1, mstrPrivs)

    Else
        With vfgFile
            Call frmTendFileOpen.ShowMe(Me, Val(.TextMatrix(.Row, mCol.fID)), mlngPatiId, mlngPageId, Val(strInfo), mintBaby, .TextMatrix(.Row, mCol.f���ڷ�Χ), , Val(.TextMatrix(.Row, mCol.f������)), mblnMoved_HL, IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, mbytFontSize)))
        End With
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vfgFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vfgFile_DblClick
End Sub

'---------------------------------------------------------------------------------
'�����ǻ������������
'---------------------------------------------------------------------------------
Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���|ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub
